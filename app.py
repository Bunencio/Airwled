from __future__ import annotations

import io
import re
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from typing import BinaryIO, Iterable, Sequence
from xml.sax.saxutils import escape

import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import legal, letter, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Flowable, LongTable, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


APP_TITLE = "Generador profesional de listas de empaque"
PACKAGE_PATTERN = re.compile(r"paquete\s+de\s+(\d+)", re.IGNORECASE)

PAGE_OPTIONS: dict[str, tuple[float, float]] = {
    "Carta horizontal": landscape(letter),
    "Oficio horizontal": landscape(legal),
}

PAGE_MARGINS = {
    "left": 0.32 * inch,
    "right": 0.32 * inch,
    "top": 0.50 * inch,
    "bottom": 0.40 * inch,
}

COLOR_BLACK = colors.black
COLOR_WHITE = colors.white
COLOR_HEADER = colors.HexColor("#1F2937")
COLOR_SUBHEADER = colors.HexColor("#E5E7EB")
COLOR_ROW_ALT = colors.HexColor("#F9FAFB")
COLOR_ROW = colors.HexColor("#FFFFFF")
COLOR_TEXT_SOFT = colors.HexColor("#4B5563")
COLOR_PACKAGE = colors.HexColor("#EEF2FF")


@dataclass(frozen=True)
class ColumnRule:
    key: str
    fallback_index: int
    aliases: Sequence[str]
    required: bool = True


@dataclass(frozen=True)
class ResolvedColumns:
    sale: str
    state: str
    units: str
    sku: str
    title: str


@dataclass(frozen=True)
class LineItem:
    sale_id: str
    units: str
    sku: str
    title: str

    @property
    def qty_number(self) -> float | None:
        return to_number(self.units)


@dataclass(frozen=True)
class OrderGroup:
    main_sale: str
    delivery_type: str
    items: list[LineItem]

    @property
    def item_count(self) -> int:
        return len(self.items)

    @property
    def total_units(self) -> str:
        values: list[float] = []
        has_non_numeric = False
        for item in self.items:
            number = item.qty_number
            if number is None:
                if clean_value(item.units):
                    has_non_numeric = True
                continue
            values.append(number)

        if values and not has_non_numeric:
            return format_number(sum(values))
        return "-"

    @property
    def short_type(self) -> str:
        return "INDIV." if self.delivery_type == "INDIVIDUAL" else f"PAQ x{self.item_count}"

    @property
    def is_package(self) -> bool:
        return self.delivery_type != "INDIVIDUAL"


@dataclass(frozen=True)
class PreparationItem:
    sku: str
    product_name: str
    total_units: float
    order_count: int
    line_count: int


COLUMN_RULES: tuple[ColumnRule, ...] = (
    ColumnRule("sale", 0, ("venta", "pedido", "order", "folio")),
    ColumnRule("state", 2, ("estado", "status")),
    ColumnRule("units", 6, ("unidades", "cantidad", "cant", "qty")),
    ColumnRule("sku", 17, ("sku", "seller sku", "codigo", "código", "asin", "referencia")),
    ColumnRule(
        "title",
        21,
        ("titulo", "título", "producto", "descripcion", "descripción", "articulo", "artículo"),
    ),
)


class Checkbox(Flowable):
    def __init__(self, size: float = 9.2, stroke_width: float = 1.0) -> None:
        super().__init__()
        self.size = size
        self.stroke_width = stroke_width
        self.width = size
        self.height = size

    def draw(self) -> None:
        self.canv.setLineWidth(self.stroke_width)
        self.canv.rect(0, 0, self.size, self.size)


# -----------------------------------------------------------------------------
# UTILIDADES
# -----------------------------------------------------------------------------


def normalize_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    text = unicodedata.normalize("NFKD", text)
    text = text.encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"\s+", " ", text)
    return text.lower().strip()


def clean_value(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def safe_paragraph_text(value: object, default: str = "-") -> str:
    text = clean_value(value) or default
    return escape(text).replace("\n", "<br/>")


def to_number(value: object) -> float | None:
    text = clean_value(value)
    if not text:
        return None
    normalized = text.replace(",", ".")
    normalized = re.sub(r"[^0-9.\-]", "", normalized)
    if not normalized or normalized in {"-", ".", "-."}:
        return None
    try:
        return float(normalized)
    except ValueError:
        return None


def format_number(value: float) -> str:
    return str(int(value)) if float(value).is_integer() else f"{value:.2f}"


def normalize_sku(value: object) -> str:
    sku = clean_value(value)
    return sku.upper().strip()


def build_product_name_for_pdf(item: LineItem) -> str:
    sku = escape(item.sku) if clean_value(item.sku) else "SIN SKU"
    title = escape(item.title) if clean_value(item.title) else "Sin descripción"
    return f"<b>{sku}</b><br/>{title}"


def pick_canonical_product_name(names: Sequence[str]) -> str:
    cleaned = [clean_value(name) for name in names if clean_value(name)]
    if not cleaned:
        return "Sin descripción"
    cleaned.sort(key=lambda x: (-len(x), x.lower()))
    return cleaned[0]


def is_effectively_blank(value: object) -> bool:
    text = clean_value(value)
    return text == ""


def row_has_meaningful_item_data(row: pd.Series, columns: ResolvedColumns) -> bool:
    sku = clean_value(row[columns.sku])
    title = clean_value(row[columns.title])
    units = clean_value(row[columns.units])

    if sku or title:
        return True

    qty = to_number(units)
    return qty is not None and qty > 0


def looks_like_header_artifact(row: pd.Series, columns: ResolvedColumns) -> bool:
    sale = normalize_text(row[columns.sale])
    title = normalize_text(row[columns.title])
    sku = normalize_text(row[columns.sku])
    units = normalize_text(row[columns.units])
    state = normalize_text(row[columns.state])

    joined = " ".join(part for part in (sale, title, sku, units, state) if part)
    if not joined:
        return True

    suspicious_markers = (
        "# de venta",
        "venta principal",
        "titulo de la publicacion",
        "unidades x sku",
        "contenido verificado",
        "producto",
        "cantidad",
        "fecha de venta",
        "descripcion del estado",
    )
    return any(marker in joined for marker in suspicious_markers)


def looks_like_package_summary_row(row: pd.Series, columns: ResolvedColumns) -> bool:
    state = clean_value(row[columns.state])
    return extract_package_size(state) is not None


def looks_like_child_row(row: pd.Series, columns: ResolvedColumns) -> bool:
    if looks_like_header_artifact(row, columns):
        return False

    state = clean_value(row[columns.state])
    sale = clean_value(row[columns.sale])
    sku = clean_value(row[columns.sku])
    title = clean_value(row[columns.title])
    units = clean_value(row[columns.units])

    # Una fila hija normalmente no vuelve a iniciar otro paquete
    if extract_package_size(state) is not None:
        return False

    # Si no trae estado pero sí SKU/título/cantidad, es muy probable que sea hija
    if not state and (sku or title or units):
        return True

    return False


def reset_excel_pointer(excel_file: BinaryIO) -> None:
    if hasattr(excel_file, "seek"):
        excel_file.seek(0)


# -----------------------------------------------------------------------------
# LECTURA Y PARSEO
# -----------------------------------------------------------------------------


def resolve_source_columns(df: pd.DataFrame) -> tuple[ResolvedColumns, list[str]]:
    normalized_headers = [(column, normalize_text(column)) for column in df.columns]
    warnings: list[str] = []
    resolved: dict[str, str] = {}

    preferred_exact_matches: dict[str, tuple[str, ...]] = {
        "sale": ("# de venta", "venta", "pedido", "order", "folio"),
        "state": ("estado",),
        "units": ("unidades", "cantidad", "qty"),
        "sku": ("sku", "seller sku"),
        "title": ("titulo de la publicacion", "titulo", "producto", "descripcion"),
    }

    def score_header(rule: ColumnRule, normalized_header: str, column_position: int) -> tuple[int, int]:
        exacts = tuple(normalize_text(v) for v in preferred_exact_matches.get(rule.key, ()))
        aliases = tuple(normalize_text(v) for v in rule.aliases)

        if normalized_header in exacts:
            return (400, -column_position)
        if any(normalized_header.startswith(alias) for alias in exacts):
            return (350, -column_position)
        if any(re.search(rf"(^|\b){re.escape(alias)}(\b|$)", normalized_header) for alias in aliases):
            return (300, -column_position)
        if any(alias in normalized_header for alias in aliases):
            return (200, -column_position)
        return (-1, -column_position)

    for rule in COLUMN_RULES:
        best_column = None
        best_score = (-1, 0)

        for idx, (column_name, normalized_header) in enumerate(normalized_headers):
            candidate_score = score_header(rule, normalized_header, idx)
            if candidate_score > best_score:
                best_score = candidate_score
                best_column = column_name

        if best_column is None or best_score[0] < 0:
            if rule.fallback_index >= len(df.columns):
                raise ValueError(
                    f"No se pudo resolver la columna '{rule.key}' y el índice de respaldo {rule.fallback_index + 1} no existe."
                )
            best_column = df.columns[rule.fallback_index]
            warnings.append(
                f"La columna '{rule.key}' no se encontró por nombre; se usó la columna en posición {rule.fallback_index + 1} ({best_column})."
            )

        resolved[rule.key] = best_column

    warnings.append(
        "Columnas detectadas: "
        f"venta='{resolved['sale']}', estado='{resolved['state']}', unidades='{resolved['units']}', "
        f"sku='{resolved['sku']}', producto='{resolved['title']}'."
    )

    return (
        ResolvedColumns(
            sale=resolved["sale"],
            state=resolved["state"],
            units=resolved["units"],
            sku=resolved["sku"],
            title=resolved["title"],
        ),
        warnings,
    )


def detect_header_row(raw_df: pd.DataFrame, max_scan_rows: int = 15) -> int:
    best_index = -1
    best_score = -1

    for idx in range(min(len(raw_df), max_scan_rows)):
        row_values = [normalize_text(value) for value in raw_df.iloc[idx].tolist()]
        score = 0

        if any(value == "# de venta" for value in row_values):
            score += 5
        if any(value == "sku" for value in row_values):
            score += 5
        if any("titulo de la publicacion" in value for value in row_values):
            score += 5
        if any(value == "estado" for value in row_values):
            score += 3
        if sum(1 for value in row_values if value == "unidades") >= 1:
            score += 3

        if score > best_score:
            best_score = score
            best_index = idx

    if best_index < 0 or best_score < 10:
        raise ValueError("No se pudo detectar automáticamente la fila de encabezados del reporte.")

    return best_index


def load_source_dataframe(excel_file: BinaryIO) -> tuple[pd.DataFrame, ResolvedColumns, list[str]]:
    reset_excel_pointer(excel_file)
    raw_df = pd.read_excel(excel_file, header=None, dtype=str)
    raw_df = raw_df.dropna(how="all").reset_index(drop=True)

    if raw_df.empty:
        raise ValueError("El archivo no contiene datos válidos.")

    header_row_index = detect_header_row(raw_df)

    reset_excel_pointer(excel_file)
    df = pd.read_excel(excel_file, header=header_row_index, dtype=str)
    df = df.dropna(how="all").reset_index(drop=True)

    if df.empty:
        raise ValueError("El archivo no contiene datos válidos después de la fila de encabezados.")

    columns, warnings = resolve_source_columns(df)
    warnings.insert(0, f"Fila de encabezados detectada automáticamente: {header_row_index + 1}.")
    return df, columns, warnings


def extract_package_size(state_value: str) -> int | None:
    match = PACKAGE_PATTERN.search(state_value)
    if not match:
        return None
    return int(match.group(1))


def parse_orders(df: pd.DataFrame, columns: ResolvedColumns) -> tuple[list[OrderGroup], list[str]]:
    """
    Soporta ambos casos:
    1) Venta normal individual.
    2) Venta marcada como "Paquete de N productos" con filas hijas debajo.
    3) Venta marcada como paquete pero sin filas hijas completas -> se exporta como individual.
    """
    orders: list[OrderGroup] = []
    warnings: list[str] = []
    skipped_artifacts = 0
    i = 0

    while i < len(df):
        row = df.iloc[i]

        if looks_like_header_artifact(row, columns):
            skipped_artifacts += 1
            i += 1
            continue

        main_sale = clean_value(row[columns.sale]) or f"SIN-VENTA-{i + 1}"
        state = clean_value(row[columns.state])
        package_size = extract_package_size(state)

        # ---------------------------------------------------------------------
        # CASO PAQUETE
        # ---------------------------------------------------------------------
        if package_size:
            items: list[LineItem] = []
            next_index = i + 1

            while next_index < len(df) and len(items) < package_size:
                candidate = df.iloc[next_index]

                if looks_like_header_artifact(candidate, columns):
                    next_index += 1
                    continue

                # Si aparece otra venta-paquete, se corta el bloque
                if looks_like_package_summary_row(candidate, columns):
                    break

                # Si trae un # de venta distinto y además parece una venta normal, ya no es hija
                candidate_sale = clean_value(candidate[columns.sale])
                candidate_state = clean_value(candidate[columns.state])

                starts_new_normal_order = (
                    candidate_sale != ""
                    and candidate_sale != main_sale
                    and candidate_state != ""
                    and extract_package_size(candidate_state) is None
                )

                if starts_new_normal_order and not looks_like_child_row(candidate, columns):
                    break

                if row_has_meaningful_item_data(candidate, columns):
                    items.append(
                        LineItem(
                            sale_id=clean_value(candidate[columns.sale]) or main_sale,
                            units=clean_value(candidate[columns.units]) or "1",
                            sku=normalize_sku(candidate[columns.sku]),
                            title=clean_value(candidate[columns.title]),
                        )
                    )
                    next_index += 1
                    continue

                # Si la fila no aporta nada útil, dejamos de consumir
                break

            if items:
                orders.append(
                    OrderGroup(
                        main_sale=main_sale,
                        delivery_type=f"JUNTO ({len(items)} productos)",
                        items=items,
                    )
                )

                if len(items) < package_size:
                    warnings.append(
                        f"La venta '{main_sale}' indica 'Paquete de {package_size}', "
                        f"pero solo se pudieron identificar {len(items)} producto(s) hijo."
                    )

                i = next_index
                continue

            # Fallback seguro: si no hay hijos, tratarlo como individual
            warnings.append(
                f"La venta '{main_sale}' está marcada como '{state}', "
                "pero no se encontraron renglones hijos válidos. Se exportó como individual."
            )
            orders.append(
                OrderGroup(
                    main_sale=main_sale,
                    delivery_type="INDIVIDUAL",
                    items=[
                        LineItem(
                            sale_id=main_sale,
                            units=clean_value(row[columns.units]) or "1",
                            sku=normalize_sku(row[columns.sku]),
                            title=clean_value(row[columns.title]),
                        )
                    ],
                )
            )
            i += 1
            continue

        # ---------------------------------------------------------------------
        # CASO INDIVIDUAL
        # ---------------------------------------------------------------------
        orders.append(
            OrderGroup(
                main_sale=main_sale,
                delivery_type="INDIVIDUAL",
                items=[
                    LineItem(
                        sale_id=main_sale,
                        units=clean_value(row[columns.units]) or "1",
                        sku=normalize_sku(row[columns.sku]),
                        title=clean_value(row[columns.title]),
                    )
                ],
            )
        )
        i += 1

    if skipped_artifacts:
        warnings.append(
            f"Se omitieron {skipped_artifacts} fila(s) que parecían encabezados o texto auxiliar dentro del Excel."
        )

    return orders, warnings


# -----------------------------------------------------------------------------
# AGREGACIONES
# -----------------------------------------------------------------------------


def build_preparation_items(orders: Iterable[OrderGroup]) -> tuple[list[PreparationItem], list[str]]:
    grouped: dict[str, dict[str, object]] = {}
    warnings: list[str] = []

    for order in orders:
        seen_skus_in_order: set[str] = set()

        for item in order.items:
            sku = normalize_sku(item.sku)
            qty = item.qty_number

            if not sku:
                warnings.append(
                    f"La venta '{order.main_sale}' contiene un producto sin SKU; no se incluyó en la preparación consolidada por SKU."
                )
                continue

            if qty is None:
                warnings.append(
                    f"El SKU '{sku}' en la venta '{order.main_sale}' tiene una cantidad no numérica ('{item.units}'); no se incluyó en la suma consolidada."
                )
                continue

            if sku not in grouped:
                grouped[sku] = {
                    "names": [],
                    "total_units": 0.0,
                    "order_count": 0,
                    "line_count": 0,
                }

            grouped[sku]["names"].append(item.title)
            grouped[sku]["total_units"] += qty
            grouped[sku]["line_count"] += 1

            if sku not in seen_skus_in_order:
                grouped[sku]["order_count"] += 1
                seen_skus_in_order.add(sku)

    preparation_items = [
        PreparationItem(
            sku=sku,
            product_name=pick_canonical_product_name(data["names"]),
            total_units=float(data["total_units"]),
            order_count=int(data["order_count"]),
            line_count=int(data["line_count"]),
        )
        for sku, data in grouped.items()
    ]

    preparation_items.sort(key=lambda x: (-x.total_units, x.sku))
    return preparation_items, warnings


# -----------------------------------------------------------------------------
# DATAFRAMES
# -----------------------------------------------------------------------------


def build_control_dataframe(orders: Iterable[OrderGroup]) -> pd.DataFrame:
    rows = []
    for order in orders:
        for idx, item in enumerate(order.items, start=1):
            rows.append(
                {
                    "OK": "",
                    "Venta principal": order.main_sale,
                    "Tipo": order.short_type,
                    "Renglón": idx,
                    "Cantidad": item.units or "-",
                    "SKU": item.sku or "SIN SKU",
                    "Producto": item.title or "Sin descripción",
                }
            )
    return pd.DataFrame(rows)


def build_preparation_dataframe(preparation_items: Iterable[PreparationItem]) -> pd.DataFrame:
    rows = []
    for idx, item in enumerate(preparation_items, start=1):
        rows.append(
            {
                "Prioridad": idx,
                "SKU": item.sku,
                "Producto": item.product_name,
                "Total a preparar": format_number(item.total_units),
            }
        )
    return pd.DataFrame(rows)


def build_summary_dataframe(
    orders: Iterable[OrderGroup],
    preparation_items: Iterable[PreparationItem],
) -> pd.DataFrame:
    order_list = list(orders)
    prep_list = list(preparation_items)
    grouped_count = sum(1 for order in order_list if order.is_package)
    individual_count = len(order_list) - grouped_count
    product_lines = sum(order.item_count for order in order_list)
    total_detected_units = sum(item.total_units for item in prep_list)

    return pd.DataFrame(
        [
            {"Métrica": "Total de entregas", "Valor": len(order_list)},
            {"Métrica": "Individuales", "Valor": individual_count},
            {"Métrica": "Paquetes", "Valor": grouped_count},
            {"Métrica": "Líneas de producto", "Valor": product_lines},
            {"Métrica": "SKUs únicos para preparar", "Valor": len(prep_list)},
            {"Métrica": "Unidades consolidadas para preparación", "Valor": format_number(total_detected_units)},
        ]
    )


# -----------------------------------------------------------------------------
# PDF ESTILOS
# -----------------------------------------------------------------------------


def build_styles() -> dict[str, ParagraphStyle]:
    sample = getSampleStyleSheet()
    return {
        "title": ParagraphStyle(
            "TitleMain",
            parent=sample["Title"],
            fontName="Helvetica-Bold",
            fontSize=15.5,
            leading=18,
            textColor=COLOR_BLACK,
            spaceAfter=2,
        ),
        "subtitle": ParagraphStyle(
            "Subtitle",
            parent=sample["BodyText"],
            fontName="Helvetica",
            fontSize=8.0,
            leading=10,
            textColor=COLOR_TEXT_SOFT,
        ),
        "meta": ParagraphStyle(
            "Meta",
            parent=sample["BodyText"],
            fontName="Helvetica",
            fontSize=7.5,
            leading=9,
            textColor=COLOR_BLACK,
            alignment=2,
        ),
        "table_header": ParagraphStyle(
            "TableHeader",
            parent=sample["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=7.4,
            leading=8.7,
            textColor=COLOR_WHITE,
            alignment=1,
        ),
        "cell_center": ParagraphStyle(
            "CellCenter",
            parent=sample["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=7.2,
            leading=8.4,
            alignment=1,
        ),
        "cell_sale": ParagraphStyle(
            "CellSale",
            parent=sample["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=7.15,
            leading=8.4,
        ),
        "cell_qty": ParagraphStyle(
            "CellQty",
            parent=sample["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=7.2,
            leading=8.4,
            alignment=1,
        ),
        "cell_type": ParagraphStyle(
            "CellType",
            parent=sample["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=7.05,
            leading=8.2,
            alignment=1,
        ),
        "cell_product": ParagraphStyle(
            "CellProduct",
            parent=sample["BodyText"],
            fontName="Helvetica",
            fontSize=7.0,
            leading=8.5,
            splitLongWords=True,
        ),
        "metric_label": ParagraphStyle(
            "MetricLabel",
            parent=sample["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=6.3,
            leading=7.4,
        ),
        "metric_value": ParagraphStyle(
            "MetricValue",
            parent=sample["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=12,
            leading=13.2,
        ),
        "instruction": ParagraphStyle(
            "Instruction",
            parent=sample["BodyText"],
            fontName="Helvetica",
            fontSize=7.1,
            leading=8.4,
            textColor=COLOR_BLACK,
        ),
    }


def make_metric_box(label: str, value: str, width: float, styles: dict[str, ParagraphStyle]) -> Table:
    table = Table(
        [
            [Paragraph(escape(label.upper()), styles["metric_label"])],
            [Paragraph(escape(value), styles["metric_value"])],
        ],
        colWidths=[width],
    )
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), COLOR_WHITE),
                ("BOX", (0, 0), (-1, -1), 0.75, COLOR_BLACK),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    return table


def build_intro_block(
    document_title: str,
    subtitle: str,
    metrics: list[tuple[str, str]],
    page_label: str,
    doc_width: float,
    styles: dict[str, ParagraphStyle],
) -> list[object]:
    generated_at = datetime.now().strftime("%d/%m/%Y %H:%M")

    title_table = Table(
        [[
            Paragraph(escape(document_title), styles["title"]),
            Paragraph(
                f"<b>Generado:</b> {generated_at}<br/><b>Formato:</b> {escape(page_label)}<br/><b>Diseño:</b> profesional horizontal",
                styles["meta"],
            ),
        ]],
        colWidths=[doc_width * 0.67, doc_width * 0.33],
    )
    title_table.setStyle(
        TableStyle(
            [
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ]
        )
    )

    metric_width = doc_width / max(len(metrics), 1)
    metric_row = [make_metric_box(label, value, metric_width - 4, styles) for label, value in metrics]
    metric_table = Table([metric_row], colWidths=[metric_width] * len(metrics))
    metric_table.setStyle(
        TableStyle(
            [
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ]
        )
    )

    instruction_table = Table(
        [[Paragraph(escape(subtitle), styles["instruction"])]],
        colWidths=[doc_width],
    )
    instruction_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), COLOR_SUBHEADER),
                ("BOX", (0, 0), (-1, -1), 0.75, COLOR_BLACK),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )

    return [
        title_table,
        Paragraph("Documento optimizado para operación, lectura rápida y máxima claridad visual.", styles["subtitle"]),
        Spacer(1, 5),
        metric_table,
        Spacer(1, 5),
        instruction_table,
        Spacer(1, 6),
    ]


# -----------------------------------------------------------------------------
# PDF PRINCIPAL
# -----------------------------------------------------------------------------


def build_orders_table(orders: list[OrderGroup], width: float, styles: dict[str, ParagraphStyle]) -> LongTable:
    fixed_widths = {
        "no": 0.38 * inch,
        "ok": 0.46 * inch,
        "sale": 1.45 * inch,
        "type": 0.86 * inch,
        "qty": 0.70 * inch,
    }
    product_width = width - sum(fixed_widths.values())
    col_widths = [
        fixed_widths["no"],
        fixed_widths["ok"],
        fixed_widths["sale"],
        fixed_widths["type"],
        fixed_widths["qty"],
        product_width,
    ]

    data: list[list[object]] = [[
        Paragraph("No.", styles["table_header"]),
        Paragraph("OK", styles["table_header"]),
        Paragraph("Venta", styles["table_header"]),
        Paragraph("Tipo", styles["table_header"]),
        Paragraph("Cant.", styles["table_header"]),
        Paragraph("Producto / SKU", styles["table_header"]),
    ]]

    style_commands: list[tuple] = [
        ("BACKGROUND", (0, 0), (-1, 0), COLOR_HEADER),
        ("TEXTCOLOR", (0, 0), (-1, 0), COLOR_WHITE),
        ("BOX", (0, 0), (-1, -1), 0.75, COLOR_BLACK),
        ("INNERGRID", (0, 0), (-1, -1), 0.40, COLOR_BLACK),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [COLOR_ROW, COLOR_ROW_ALT]),
    ]

    row_index = 1
    order_number = 1

    for order in orders:
        start_row = row_index

        for item in order.items:
            data.append(
                [
                    Paragraph(str(order_number), styles["cell_center"]),
                    Checkbox(9.0),
                    Paragraph(escape(order.main_sale), styles["cell_sale"]),
                    Paragraph(escape(order.short_type), styles["cell_type"]),
                    Paragraph(escape(item.units or "-"), styles["cell_qty"]),
                    Paragraph(build_product_name_for_pdf(item), styles["cell_product"]),
                ]
            )
            row_index += 1

        end_row = row_index - 1

        if end_row > start_row:
            style_commands.extend(
                [
                    ("SPAN", (0, start_row), (0, end_row)),
                    ("SPAN", (1, start_row), (1, end_row)),
                    ("SPAN", (2, start_row), (2, end_row)),
                    ("SPAN", (3, start_row), (3, end_row)),
                    ("VALIGN", (0, start_row), (3, end_row), "MIDDLE"),
                ]
            )

        if order.is_package:
            style_commands.append(("BACKGROUND", (0, start_row), (-1, end_row), COLOR_PACKAGE))

        order_number += 1

    table = LongTable(data, colWidths=col_widths, repeatRows=1, splitByRow=1)
    table.setStyle(TableStyle(style_commands))
    return table


def build_main_pdf_buffer(orders: list[OrderGroup], page_label: str) -> io.BytesIO:
    styles = build_styles()
    page_size = PAGE_OPTIONS[page_label]
    buffer = io.BytesIO()

    total_lines = sum(order.item_count for order in orders)
    grouped_count = sum(1 for order in orders if order.is_package)
    individual_count = len(orders) - grouped_count

    doc = SimpleDocTemplate(
        buffer,
        pagesize=page_size,
        leftMargin=PAGE_MARGINS["left"],
        rightMargin=PAGE_MARGINS["right"],
        topMargin=PAGE_MARGINS["top"],
        bottomMargin=PAGE_MARGINS["bottom"],
        title="Lista de empaque profesional",
        author="OpenAI",
    )

    elements: list[object] = []
    elements.extend(
        build_intro_block(
            document_title="Lista de empaque",
            subtitle=(
                "Se separó la cantidad del producto para eliminar confusión visual. "
                "Cuando el Excel incluye hijos de un paquete, se mantienen agrupados; "
                "si no vienen, la venta se exporta como individual para no perderla."
            ),
            metrics=[
                ("Entregas", str(len(orders))),
                ("Individuales", str(individual_count)),
                ("Paquetes", str(grouped_count)),
                ("Líneas", str(total_lines)),
            ],
            page_label=page_label,
            doc_width=doc.width,
            styles=styles,
        )
    )
    elements.append(build_orders_table(orders, doc.width, styles))

    doc.build(
        elements,
        onFirstPage=lambda canvas, current_doc: draw_page_decoration(canvas, current_doc, "LISTA DE EMPAQUE"),
        onLaterPages=lambda canvas, current_doc: draw_page_decoration(canvas, current_doc, "LISTA DE EMPAQUE"),
    )
    buffer.seek(0)
    return buffer


# -----------------------------------------------------------------------------
# PDF DE PREPARACIÓN
# -----------------------------------------------------------------------------


def build_preparation_table(items: list[PreparationItem], width: float, styles: dict[str, ParagraphStyle]) -> LongTable:
    fixed_widths = {
        "no": 0.42 * inch,
        "qty": 0.95 * inch,
        "sku": 1.65 * inch,
    }
    product_width = width - sum(fixed_widths.values())
    col_widths = [
        fixed_widths["no"],
        fixed_widths["qty"],
        fixed_widths["sku"],
        product_width,
    ]

    data: list[list[object]] = [[
        Paragraph("#", styles["table_header"]),
        Paragraph("Total", styles["table_header"]),
        Paragraph("SKU", styles["table_header"]),
        Paragraph("Producto", styles["table_header"]),
    ]]

    for idx, item in enumerate(items, start=1):
        data.append(
            [
                Paragraph(str(idx), styles["cell_center"]),
                Paragraph(format_number(item.total_units), styles["cell_qty"]),
                Paragraph(escape(item.sku), styles["cell_sale"]),
                Paragraph(safe_paragraph_text(item.product_name), styles["cell_product"]),
            ]
        )

    table = LongTable(data, colWidths=col_widths, repeatRows=1, splitByRow=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), COLOR_HEADER),
                ("TEXTCOLOR", (0, 0), (-1, 0), COLOR_WHITE),
                ("BOX", (0, 0), (-1, -1), 0.75, COLOR_BLACK),
                ("INNERGRID", (0, 0), (-1, -1), 0.40, COLOR_BLACK),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [COLOR_ROW, COLOR_ROW_ALT]),
            ]
        )
    )
    return table


def build_preparation_pdf_buffer(preparation_items: list[PreparationItem], page_label: str) -> io.BytesIO:
    styles = build_styles()
    page_size = PAGE_OPTIONS[page_label]
    buffer = io.BytesIO()

    total_units = sum(item.total_units for item in preparation_items)
    doc = SimpleDocTemplate(
        buffer,
        pagesize=page_size,
        leftMargin=PAGE_MARGINS["left"],
        rightMargin=PAGE_MARGINS["right"],
        topMargin=PAGE_MARGINS["top"],
        bottomMargin=PAGE_MARGINS["bottom"],
        title="Lista de preparación por SKU",
        author="OpenAI",
    )

    elements: list[object] = []
    elements.extend(
        build_intro_block(
            document_title="Lista de preparación previa por SKU",
            subtitle=(
                "Este documento consolida cantidades por SKU antes de empacar. "
                "Se usa SKU como llave principal para evitar errores cuando coinciden nombres de producto."
            ),
            metrics=[
                ("SKUs", str(len(preparation_items))),
                ("Unidades", format_number(total_units)),
                ("Top SKU", preparation_items[0].sku if preparation_items else "-"),
                ("Mayor total", format_number(preparation_items[0].total_units) if preparation_items else "0"),
            ],
            page_label=page_label,
            doc_width=doc.width,
            styles=styles,
        )
    )
    elements.append(build_preparation_table(preparation_items, doc.width, styles))

    doc.build(
        elements,
        onFirstPage=lambda canvas, current_doc: draw_page_decoration(canvas, current_doc, "PREPARACIÓN POR SKU"),
        onLaterPages=lambda canvas, current_doc: draw_page_decoration(canvas, current_doc, "PREPARACIÓN POR SKU"),
    )
    buffer.seek(0)
    return buffer


def draw_page_decoration(canvas, doc, title: str) -> None:
    page_width, page_height = doc.pagesize
    canvas.saveState()
    canvas.setStrokeColor(COLOR_BLACK)
    canvas.setLineWidth(0.7)
    canvas.line(doc.leftMargin, page_height - 0.27 * inch, page_width - doc.rightMargin, page_height - 0.27 * inch)

    canvas.setFont("Helvetica-Bold", 8.4)
    canvas.drawString(doc.leftMargin, page_height - 0.20 * inch, title)

    canvas.setFont("Helvetica", 7.8)
    canvas.drawRightString(page_width - doc.rightMargin, page_height - 0.20 * inch, f"Página {canvas.getPageNumber()}")

    canvas.setLineWidth(0.5)
    canvas.line(doc.leftMargin, 0.29 * inch, page_width - doc.rightMargin, 0.29 * inch)
    canvas.setFont("Helvetica", 7.1)
    canvas.drawString(doc.leftMargin, 0.15 * inch, "Formato operativo optimizado para impresión y control de almacén")
    canvas.drawRightString(page_width - doc.rightMargin, 0.15 * inch, datetime.now().strftime("%d/%m/%Y %H:%M"))
    canvas.restoreState()


# -----------------------------------------------------------------------------
# GENERACIÓN CENTRAL
# -----------------------------------------------------------------------------


def generate_files(
    excel_file: BinaryIO,
    page_label: str,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, io.BytesIO, io.BytesIO, list[str], list[OrderGroup], list[PreparationItem]]:
    source_df, columns, warnings = load_source_dataframe(excel_file)
    orders, parse_warnings = parse_orders(source_df, columns)
    warnings.extend(parse_warnings)

    if not orders:
        raise ValueError("No se encontraron ventas válidas para exportar.")

    preparation_items, preparation_warnings = build_preparation_items(orders)
    warnings.extend(preparation_warnings)

    summary_df = build_summary_dataframe(orders, preparation_items)
    control_df = build_control_dataframe(orders)
    preparation_df = build_preparation_dataframe(preparation_items)
    main_pdf_buffer = build_main_pdf_buffer(orders, page_label)
    prep_pdf_buffer = build_preparation_pdf_buffer(preparation_items, page_label)

    return (
        summary_df,
        control_df,
        preparation_df,
        main_pdf_buffer,
        prep_pdf_buffer,
        warnings,
        orders,
        preparation_items,
    )


# -----------------------------------------------------------------------------
# STREAMLIT UI
# -----------------------------------------------------------------------------


def render_metrics(orders: list[OrderGroup], preparation_items: list[PreparationItem]) -> None:
    total_orders = len(orders)
    grouped_orders = sum(1 for order in orders if order.is_package)
    total_lines = sum(order.item_count for order in orders)
    total_prep_units = sum(item.total_units for item in preparation_items)

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Entregas", total_orders)
    col2.metric("Paquetes", grouped_orders)
    col3.metric("Líneas", total_lines)
    col4.metric("Unidades a preparar", format_number(total_prep_units))


def render_downloads(main_pdf_buffer: io.BytesIO, prep_pdf_buffer: io.BytesIO) -> None:
    col1, col2 = st.columns(2)

    with col1:
        st.download_button(
            label="Descargar PDF de empaque",
            data=main_pdf_buffer.getvalue(),
            file_name="lista_empaque_profesional.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

    with col2:
        st.download_button(
            label="Descargar PDF de preparación por SKU",
            data=prep_pdf_buffer.getvalue(),
            file_name="lista_preparacion_por_sku.pdf",
            mime="application/pdf",
            use_container_width=True,
        )


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)
    st.caption(
        "Procesa tu Excel y genera dos PDFs profesionales: uno para empaque con columnas claras y otro para preparación previa consolidada por SKU."
    )

    with st.container(border=True):
        st.markdown(
            """
            **Mejoras implementadas**

            - La **cantidad** ahora va en su propia columna, separada del producto.
            - Se eliminó la columna **Iniciales / Hora** del PDF principal.
            - Los **paquetes** se agrupan cuando el Excel trae renglones hijo válidos.
            - Si un paquete no trae renglones hijo completos, **no falla**: se exporta como individual.
            - Se genera un segundo PDF con la **preparación total por SKU** antes del empacado.
            - La consolidación usa **SKU como llave principal**, no el nombre del producto.
            """
        )

    col_a, col_b = st.columns([1.3, 1])

    with col_a:
        page_label = st.radio(
            "Formato PDF",
            list(PAGE_OPTIONS.keys()),
            index=0,
            horizontal=True,
        )

    with col_b:
        st.markdown("**Recomendación:** Carta horizontal para impresoras estándar.")

    uploaded_file = st.file_uploader("Sube el archivo Excel", type=["xlsx", "xls"])
    if uploaded_file is None:
        st.info("Esperando archivo Excel para generar los PDFs.")
        return

    try:
        with st.spinner("Procesando archivo y construyendo PDFs profesionales..."):
            (
                summary_df,
                control_df,
                preparation_df,
                main_pdf_buffer,
                prep_pdf_buffer,
                warnings,
                orders,
                preparation_items,
            ) = generate_files(uploaded_file, page_label)

        st.success("Proceso completado correctamente.")
        render_metrics(orders, preparation_items)

        if warnings:
            with st.expander("Ver advertencias detectadas"):
                for warning in warnings:
                    st.warning(warning)

        tab1, tab2, tab3 = st.tabs(["Lista de empaque", "Preparación por SKU", "Descargas"])

        with tab1:
            st.dataframe(control_df, use_container_width=True, height=460)

        with tab2:
            st.dataframe(preparation_df, use_container_width=True, height=460)
            st.dataframe(summary_df, use_container_width=True, height=240)

        with tab3:
            render_downloads(main_pdf_buffer, prep_pdf_buffer)

    except Exception as error:
        st.error(f"Error al procesar el archivo: {error}")


if __name__ == "__main__":
    main()
