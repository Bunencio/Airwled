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
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.pagesizes import legal, letter, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Flowable, LongTable, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


APP_TITLE = "Generador profesional de lista de empaque"
HEADER_ROW_INDEX = 4  # fila 5 del Excel
PACKAGE_PATTERN = re.compile(r"paquete\s+de\s+(\d+)", re.IGNORECASE)

PAGE_OPTIONS: dict[str, tuple[float, float]] = {
    "Carta horizontal": landscape(letter),
    "Oficio horizontal": landscape(legal),
}

PAGE_MARGINS = {
    "left": 0.34 * inch,
    "right": 0.34 * inch,
    "top": 0.52 * inch,
    "bottom": 0.42 * inch,
}

COLOR_BLACK = colors.black
COLOR_WHITE = colors.white
COLOR_GRAY_02 = colors.HexColor("#FAFAFA")
COLOR_GRAY_05 = colors.HexColor("#F3F3F3")
COLOR_GRAY_08 = colors.HexColor("#ECECEC")
COLOR_GRAY_10 = colors.HexColor("#E2E2E2")
COLOR_GRAY_15 = colors.HexColor("#D4D4D4")
COLOR_GRAY_75 = colors.HexColor("#3A3A3A")
COLOR_GRAY_85 = colors.HexColor("#222222")
COLOR_MID = colors.HexColor("#666666")


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
            number = to_number(item.units)
            if number is None:
                if clean_value(item.units):
                    has_non_numeric = True
                continue
            values.append(number)

        if values and not has_non_numeric:
            total = sum(values)
            return str(int(total)) if float(total).is_integer() else f"{total:.2f}"
        return "-"

    @property
    def short_type(self) -> str:
        if self.delivery_type == "INDIVIDUAL":
            return "INDIV."
        return f"PAQ x{self.item_count}"

    @property
    def is_package(self) -> bool:
        return self.delivery_type != "INDIVIDUAL"


COLUMN_RULES: tuple[ColumnRule, ...] = (
    ColumnRule("sale", 0, ("venta", "pedido", "order", "folio")),
    ColumnRule("state", 2, ("estado", "status")),
    ColumnRule("units", 6, ("unidades", "cantidad", "cant", "qty")),
    ColumnRule("sku", 16, ("sku", "codigo", "código", "asin", "referencia")),
    ColumnRule(
        "title",
        20,
        ("titulo", "título", "producto", "descripcion", "descripción", "articulo", "artículo"),
    ),
)


class Checkbox(Flowable):
    """Casilla vacía para marcado manual."""

    def __init__(self, size: float = 9.2, stroke_width: float = 1.05) -> None:
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



def slugify_filename(name: str) -> str:
    text = normalize_text(name)
    text = re.sub(r"[^a-z0-9]+", "_", text).strip("_")
    return text or "archivo"



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



def looks_like_header_artifact(row: pd.Series, columns: ResolvedColumns) -> bool:
    sale = normalize_text(row[columns.sale])
    title = normalize_text(row[columns.title])
    sku = normalize_text(row[columns.sku])
    units = normalize_text(row[columns.units])

    joined = " ".join(part for part in (sale, title, sku, units) if part)
    if not joined:
        return True

    suspicious_markers = (
        "# de venta",
        "venta principal",
        "titulo de la publicacion",
        "unidades x sku",
        "contenido verificado",
    )
    return any(marker in joined for marker in suspicious_markers)



def compact_type_label(order: OrderGroup) -> str:
    return order.short_type



def build_item_line_html(item: LineItem, index: int | None = None) -> str:
    qty = escape(item.units) if clean_value(item.units) else "-"
    sku = escape(item.sku) if clean_value(item.sku) else "SIN SKU"
    title = escape(item.title) if clean_value(item.title) else "Sin descripción"
    prefix = f"{index}. " if index is not None else ""
    return f"{prefix}{qty} x <b>{sku}</b> - {title}"



def build_order_detail_html(order: OrderGroup) -> str:
    lines = []
    for idx, item in enumerate(order.items, start=1):
        lines.append(build_item_line_html(item, idx if order.item_count > 1 else None))
    return "<br/>".join(lines)


# -----------------------------------------------------------------------------
# LECTURA Y PARSEO DE EXCEL
# -----------------------------------------------------------------------------


def resolve_source_columns(df: pd.DataFrame) -> tuple[ResolvedColumns, list[str]]:
    normalized_headers = {column: normalize_text(column) for column in df.columns}
    warnings: list[str] = []
    resolved: dict[str, str] = {}

    for rule in COLUMN_RULES:
        found_column = None

        for column_name, normalized_header in normalized_headers.items():
            if any(alias in normalized_header for alias in rule.aliases):
                found_column = column_name
                break

        if found_column is None:
            if rule.fallback_index >= len(df.columns):
                raise ValueError(
                    f"No se pudo resolver la columna '{rule.key}' y el índice de respaldo {rule.fallback_index + 1} no existe."
                )
            found_column = df.columns[rule.fallback_index]
            warnings.append(
                f"La columna '{rule.key}' no se encontró por nombre; se usó la columna en posición {rule.fallback_index + 1}."
            )

        resolved[rule.key] = found_column

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



def load_source_dataframe(excel_file: BinaryIO) -> tuple[pd.DataFrame, ResolvedColumns, list[str]]:
    df = pd.read_excel(excel_file, header=HEADER_ROW_INDEX)
    df = df.dropna(how="all").reset_index(drop=True)

    if df.empty:
        raise ValueError("El archivo no contiene datos válidos después de la fila de encabezados.")

    columns, warnings = resolve_source_columns(df)
    return df, columns, warnings



def extract_package_size(state_value: str) -> int | None:
    match = PACKAGE_PATTERN.search(state_value)
    if not match:
        return None
    return int(match.group(1))



def parse_orders(df: pd.DataFrame, columns: ResolvedColumns) -> tuple[list[OrderGroup], list[str]]:
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

        state = clean_value(row[columns.state])
        main_sale = clean_value(row[columns.sale]) or f"SIN-VENTA-{i + 1}"
        package_size = extract_package_size(state)

        if package_size:
            available_rows = len(df) - (i + 1)
            real_size = min(package_size, available_rows)
            items: list[LineItem] = []

            if available_rows < package_size:
                warnings.append(
                    f"La venta '{main_sale}' indica 'Paquete de {package_size}', pero solo hay {available_rows} filas posteriores disponibles."
                )

            for offset in range(1, real_size + 1):
                child_row = df.iloc[i + offset]
                if looks_like_header_artifact(child_row, columns):
                    continue
                items.append(
                    LineItem(
                        sale_id=clean_value(child_row[columns.sale]),
                        units=clean_value(child_row[columns.units]),
                        sku=clean_value(child_row[columns.sku]),
                        title=clean_value(child_row[columns.title]),
                    )
                )

            if not items:
                warnings.append(
                    f"La venta '{main_sale}' estaba marcada como paquete, pero no se encontraron artículos hijos. Se exportó como individual."
                )
                items = [
                    LineItem(
                        sale_id=main_sale,
                        units=clean_value(row[columns.units]),
                        sku=clean_value(row[columns.sku]),
                        title=clean_value(row[columns.title]),
                    )
                ]
                delivery_type = "INDIVIDUAL"
                i += 1
            else:
                delivery_type = f"JUNTO ({len(items)} productos)"
                i += package_size + 1

            orders.append(OrderGroup(main_sale=main_sale, delivery_type=delivery_type, items=items))
            continue

        orders.append(
            OrderGroup(
                main_sale=main_sale,
                delivery_type="INDIVIDUAL",
                items=[
                    LineItem(
                        sale_id=main_sale,
                        units=clean_value(row[columns.units]),
                        sku=clean_value(row[columns.sku]),
                        title=clean_value(row[columns.title]),
                    )
                ],
            )
        )
        i += 1

    if skipped_artifacts:
        warnings.append(f"Se omitieron {skipped_artifacts} fila(s) que parecían encabezados o texto auxiliar dentro del Excel.")

    return orders, warnings


# -----------------------------------------------------------------------------
# DATAFRAMES DE SALIDA
# -----------------------------------------------------------------------------


def build_control_dataframe(orders: Iterable[OrderGroup]) -> pd.DataFrame:
    rows = []
    for order in orders:
        rows.append(
            {
                "OK": "",
                "Iniciales / Hora": "",
                "Venta principal": order.main_sale,
                "Tipo": order.delivery_type,
                "Artículos": order.item_count,
                "Unidades totales": order.total_units,
                "Detalle": "\n".join(
                    f"{item.units or '-'} x {item.sku or 'SIN SKU'} - {item.title or 'Sin descripción'}"
                    for item in order.items
                ),
            }
        )
    return pd.DataFrame(rows)



def build_detail_dataframe(orders: Iterable[OrderGroup]) -> pd.DataFrame:
    rows = []
    for order in orders:
        for idx, item in enumerate(order.items, start=1):
            rows.append(
                {
                    "Venta principal": order.main_sale,
                    "Tipo": order.delivery_type,
                    "Renglón": idx,
                    "Venta hijo": item.sale_id,
                    "Unidades": item.units,
                    "SKU": item.sku,
                    "Producto": item.title,
                }
            )
    return pd.DataFrame(rows)



def build_summary_dataframe(orders: Iterable[OrderGroup]) -> pd.DataFrame:
    order_list = list(orders)
    grouped_count = sum(1 for order in order_list if order.is_package)
    individual_count = len(order_list) - grouped_count
    product_lines = sum(order.item_count for order in order_list)
    numeric_unit_values = [to_number(order.total_units) for order in order_list]
    total_units = sum(value for value in numeric_unit_values if value is not None)

    return pd.DataFrame(
        [
            {"Métrica": "Total de entregas", "Valor": len(order_list)},
            {"Métrica": "Individuales", "Valor": individual_count},
            {"Métrica": "Paquetes", "Valor": grouped_count},
            {"Métrica": "Líneas de producto", "Valor": product_lines},
            {"Métrica": "Unidades numéricas detectadas", "Valor": int(total_units) if float(total_units).is_integer() else total_units},
        ]
    )


# -----------------------------------------------------------------------------
# EXPORTACIÓN EXCEL
# -----------------------------------------------------------------------------


def apply_sheet_formatting(ws, preferred_widths: dict[str, int], freeze_panes: str) -> None:
    ws.freeze_panes = freeze_panes
    header_fill = PatternFill(fill_type="solid", start_color="D9D9D9", end_color="D9D9D9")
    border = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000"),
    )

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    for row in ws.iter_rows(min_row=2):
        max_lines = 1
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = border
            line_count = str(cell.value or "").count("\n") + 1
            max_lines = max(max_lines, line_count)
        ws.row_dimensions[row[0].row].height = max(20, min(16 * max_lines, 92))

    for idx in range(1, ws.max_column + 1):
        letter_code = get_column_letter(idx)
        ws.column_dimensions[letter_code].width = preferred_widths.get(letter_code, 18)



def build_excel_buffer(summary_df: pd.DataFrame, control_df: pd.DataFrame, detail_df: pd.DataFrame) -> io.BytesIO:
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        summary_df.to_excel(writer, index=False, sheet_name="Resumen")
        control_df.to_excel(writer, index=False, sheet_name="Control")
        detail_df.to_excel(writer, index=False, sheet_name="Detalle")

        apply_sheet_formatting(writer.sheets["Resumen"], {"A": 34, "B": 18}, "A2")
        apply_sheet_formatting(
            writer.sheets["Control"],
            {"A": 10, "B": 16, "C": 22, "D": 18, "E": 10, "F": 14, "G": 72},
            "A2",
        )
        apply_sheet_formatting(
            writer.sheets["Detalle"],
            {"A": 22, "B": 18, "C": 10, "D": 18, "E": 12, "F": 22, "G": 72},
            "A2",
        )

    buffer.seek(0)
    return buffer


# -----------------------------------------------------------------------------
# ESTILOS PDF
# -----------------------------------------------------------------------------


def build_styles() -> dict[str, ParagraphStyle]:
    sample = getSampleStyleSheet()

    return {
        "title": ParagraphStyle(
            "TitleMain",
            parent=sample["Title"],
            fontName="Helvetica-Bold",
            fontSize=16,
            leading=18,
            textColor=COLOR_BLACK,
            spaceAfter=2,
        ),
        "subtitle": ParagraphStyle(
            "Subtitle",
            parent=sample["BodyText"],
            fontName="Helvetica",
            fontSize=8.2,
            leading=10,
            textColor=COLOR_MID,
            spaceAfter=0,
        ),
        "meta": ParagraphStyle(
            "Meta",
            parent=sample["BodyText"],
            fontName="Helvetica",
            fontSize=7.7,
            leading=9.2,
            textColor=COLOR_BLACK,
            alignment=2,
        ),
        "metric_label": ParagraphStyle(
            "MetricLabel",
            parent=sample["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=6.5,
            leading=8,
            textColor=COLOR_BLACK,
        ),
        "metric_value": ParagraphStyle(
            "MetricValue",
            parent=sample["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=12.6,
            leading=14,
            textColor=COLOR_BLACK,
        ),
        "instruction": ParagraphStyle(
            "Instruction",
            parent=sample["BodyText"],
            fontName="Helvetica",
            fontSize=7.2,
            leading=8.6,
            textColor=COLOR_BLACK,
        ),
        "table_header": ParagraphStyle(
            "TableHeader",
            parent=sample["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=7.9,
            leading=9.2,
            textColor=COLOR_WHITE,
            alignment=1,
        ),
        "cell_center": ParagraphStyle(
            "CellCenter",
            parent=sample["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=7.6,
            leading=8.8,
            textColor=COLOR_BLACK,
            alignment=1,
        ),
        "cell_sale": ParagraphStyle(
            "CellSale",
            parent=sample["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=7.6,
            leading=8.8,
            textColor=COLOR_BLACK,
        ),
        "cell_type": ParagraphStyle(
            "CellType",
            parent=sample["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=7.5,
            leading=8.8,
            textColor=COLOR_BLACK,
            alignment=1,
        ),
        "cell_hint": ParagraphStyle(
            "CellHint",
            parent=sample["BodyText"],
            fontName="Helvetica",
            fontSize=5.7,
            leading=6.8,
            textColor=COLOR_MID,
            alignment=1,
        ),
        "cell_detail": ParagraphStyle(
            "CellDetail",
            parent=sample["BodyText"],
            fontName="Helvetica",
            fontSize=7.45,
            leading=8.95,
            textColor=COLOR_BLACK,
            splitLongWords=True,
            allowWidows=1,
            allowOrphans=1,
        ),
    }


# -----------------------------------------------------------------------------
# CONSTRUCCIÓN PDF
# -----------------------------------------------------------------------------


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
                ("BOX", (0, 0), (-1, -1), 0.9, COLOR_BLACK),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )
    )
    return table



def build_intro_block(
    orders: list[OrderGroup],
    styles: dict[str, ParagraphStyle],
    doc_width: float,
    page_label: str,
) -> list:
    grouped_count = sum(1 for order in orders if order.is_package)
    individual_count = len(orders) - grouped_count
    product_lines = sum(order.item_count for order in orders)
    generated_at = datetime.now().strftime("%d/%m/%Y %H:%M")

    title_table = Table(
        [
            [
                Paragraph("Lista de empaque", styles["title"]),
                Paragraph(
                    f"<b>Generado:</b> {generated_at}<br/><b>Formato:</b> {escape(page_label)}<br/><b>Diseño:</b> horizontal compacto",
                    styles["meta"],
                ),
            ]
        ],
        colWidths=[doc_width * 0.68, doc_width * 0.32],
    )
    title_table.setStyle(
        TableStyle(
            [
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ]
        )
    )

    summary_widths = [doc_width * 0.13, doc_width * 0.13, doc_width * 0.13, doc_width * 0.13, doc_width * 0.48]
    summary_table = Table(
        [
            [
                make_metric_box("Entregas", str(len(orders)), summary_widths[0], styles),
                make_metric_box("Individuales", str(individual_count), summary_widths[1], styles),
                make_metric_box("Paquetes", str(grouped_count), summary_widths[2], styles),
                make_metric_box("Líneas", str(product_lines), summary_widths[3], styles),
                make_metric_box("Control", "1 casilla + iniciales/hora por venta", summary_widths[4], styles),
            ]
        ],
        colWidths=summary_widths,
    )
    summary_table.setStyle(
        TableStyle(
            [
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ]
        )
    )

    instruction_table = Table(
        [[
            Paragraph(
                "<b>Uso:</b> validar venta y contenido, marcar la casilla al terminar y escribir iniciales u hora. "
                "<b>Clave:</b> INDIV. = individual, PAQ xN = entrega conjunta.",
                styles["instruction"],
            )
        ]],
        colWidths=[doc_width],
    )
    instruction_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), COLOR_GRAY_08),
                ("BOX", (0, 0), (-1, -1), 0.8, COLOR_BLACK),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )
    )

    return [
        title_table,
        Paragraph(
            "Documento preparado para impresión en blanco y negro con densidad alta, control manual claro y mejor aprovechamiento de hoja.",
            styles["subtitle"],
        ),
        Spacer(1, 5),
        summary_table,
        Spacer(1, 5),
        instruction_table,
        Spacer(1, 6),
    ]



def build_orders_table(orders: list[OrderGroup], width: float, styles: dict[str, ParagraphStyle]) -> LongTable:
    fixed_widths = {
        "no": 0.40 * inch,
        "ok": 0.48 * inch,
        "mark": 1.02 * inch,
        "sale": 1.52 * inch,
        "type": 0.92 * inch,
        "items": 0.54 * inch,
    }
    detail_width = width - sum(fixed_widths.values())
    col_widths = [
        fixed_widths["no"],
        fixed_widths["ok"],
        fixed_widths["mark"],
        fixed_widths["sale"],
        fixed_widths["type"],
        fixed_widths["items"],
        detail_width,
    ]

    data: list[list[object]] = [
        [
            Paragraph("No.", styles["table_header"]),
            Paragraph("OK", styles["table_header"]),
            Paragraph("Iniciales / hora", styles["table_header"]),
            Paragraph("Venta principal", styles["table_header"]),
            Paragraph("Tipo", styles["table_header"]),
            Paragraph("Art.", styles["table_header"]),
            Paragraph("Contenido verificado", styles["table_header"]),
        ]
    ]

    style_commands: list[tuple] = [
        ("BACKGROUND", (0, 0), (-1, 0), COLOR_GRAY_85),
        ("TEXTCOLOR", (0, 0), (-1, 0), COLOR_WHITE),
        ("BOX", (0, 0), (-1, -1), 0.75, COLOR_BLACK),
        ("INNERGRID", (0, 0), (-1, -1), 0.45, COLOR_BLACK),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("ALIGN", (0, 0), (5, -1), "CENTER"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [COLOR_WHITE, COLOR_GRAY_05]),
    ]

    for row_index, order in enumerate(orders, start=1):
        initials_cell = Paragraph("__________<br/><font size='5.6'>inic. / hora</font>", styles["cell_hint"])
        data.append(
            [
                Paragraph(str(row_index), styles["cell_center"]),
                Checkbox(9.2),
                initials_cell,
                Paragraph(escape(order.main_sale), styles["cell_sale"]),
                Paragraph(escape(compact_type_label(order)), styles["cell_type"]),
                Paragraph(str(order.item_count), styles["cell_center"]),
                Paragraph(build_order_detail_html(order), styles["cell_detail"]),
            ]
        )

        if order.is_package:
            style_commands.extend(
                [
                    ("BACKGROUND", (4, row_index), (5, row_index), COLOR_GRAY_10),
                    ("FONTNAME", (4, row_index), (5, row_index), "Helvetica-Bold"),
                ]
            )

    table = LongTable(data, colWidths=col_widths, repeatRows=1, splitByRow=1)
    table.setStyle(TableStyle(style_commands))
    return table



def draw_page_decoration(canvas, doc) -> None:
    page_width, page_height = doc.pagesize
    canvas.saveState()
    canvas.setStrokeColor(COLOR_BLACK)
    canvas.setLineWidth(0.7)
    canvas.line(doc.leftMargin, page_height - 0.28 * inch, page_width - doc.rightMargin, page_height - 0.28 * inch)

    canvas.setFont("Helvetica-Bold", 8.6)
    canvas.drawString(doc.leftMargin, page_height - 0.21 * inch, "LISTA DE EMPAQUE")

    canvas.setFont("Helvetica", 8)
    canvas.drawRightString(page_width - doc.rightMargin, page_height - 0.21 * inch, f"Página {canvas.getPageNumber()}")

    canvas.setLineWidth(0.5)
    canvas.line(doc.leftMargin, 0.30 * inch, page_width - doc.rightMargin, 0.30 * inch)
    canvas.setFont("Helvetica", 7.2)
    canvas.drawString(doc.leftMargin, 0.16 * inch, "Formato horizontal compacto para control operativo")
    canvas.drawRightString(page_width - doc.rightMargin, 0.16 * inch, datetime.now().strftime("%d/%m/%Y %H:%M"))
    canvas.restoreState()



def build_pdf_buffer(orders: list[OrderGroup], page_label: str) -> io.BytesIO:
    buffer = io.BytesIO()
    styles = build_styles()
    page_size = PAGE_OPTIONS[page_label]

    doc = SimpleDocTemplate(
        buffer,
        pagesize=page_size,
        leftMargin=PAGE_MARGINS["left"],
        rightMargin=PAGE_MARGINS["right"],
        topMargin=PAGE_MARGINS["top"],
        bottomMargin=PAGE_MARGINS["bottom"],
        title="Lista de empaque horizontal compacta",
        author="OpenAI",
    )

    elements: list[object] = []
    elements.extend(build_intro_block(orders, styles, doc.width, page_label))
    elements.append(build_orders_table(orders, doc.width, styles))

    doc.build(elements, onFirstPage=draw_page_decoration, onLaterPages=draw_page_decoration)
    buffer.seek(0)
    return buffer


# -----------------------------------------------------------------------------
# API PRINCIPAL
# -----------------------------------------------------------------------------


def generate_files(
    excel_file: BinaryIO,
    page_label: str,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, io.BytesIO, io.BytesIO, list[str], list[OrderGroup]]:
    source_df, columns, warnings = load_source_dataframe(excel_file)
    orders, parse_warnings = parse_orders(source_df, columns)
    warnings.extend(parse_warnings)

    if not orders:
        raise ValueError("No se encontraron ventas válidas para exportar.")

    summary_df = build_summary_dataframe(orders)
    control_df = build_control_dataframe(orders)
    detail_df = build_detail_dataframe(orders)
    excel_buffer = build_excel_buffer(summary_df, control_df, detail_df)
    pdf_buffer = build_pdf_buffer(orders, page_label)

    return summary_df, control_df, detail_df, excel_buffer, pdf_buffer, warnings, orders


# -----------------------------------------------------------------------------
# INTERFAZ STREAMLIT
# -----------------------------------------------------------------------------


def render_metrics(orders: list[OrderGroup]) -> None:
    total_orders = len(orders)
    grouped_orders = sum(1 for order in orders if order.is_package)
    individual_orders = total_orders - grouped_orders
    total_lines = sum(order.item_count for order in orders)

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Entregas", total_orders)
    col2.metric("Individuales", individual_orders)
    col3.metric("Paquetes", grouped_orders)
    col4.metric("Líneas de producto", total_lines)



def render_downloads(excel_buffer: io.BytesIO, pdf_buffer: io.BytesIO) -> None:
    col1, col2 = st.columns(2)

    with col1:
        st.download_button(
            label="Descargar Excel",
            data=excel_buffer.getvalue(),
            file_name="lista_empaque_horizontal.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with col2:
        st.download_button(
            label="Descargar PDF",
            data=pdf_buffer.getvalue(),
            file_name="lista_empaque_horizontal.pdf",
            mime="application/pdf",
            use_container_width=True,
        )



def main() -> None:
    st.set_page_config(page_title="Lista de empaque horizontal", layout="wide")

    st.title(APP_TITLE)
    st.caption(
        "Sube tu Excel y genera una lista de empaque horizontal, más compacta y pensada para meter más entregas por página sin sacrificar legibilidad."
    )

    with st.container(border=True):
        st.markdown(
            """
            **Qué cambia en esta versión horizontal**

            - Aprovecha mejor el ancho de la hoja para reducir páginas.
            - Consolida el control manual en **1 casilla OK + Iniciales / hora** por venta.
            - Mantiene legibilidad alta en blanco y negro.
            - Repite encabezados en cada página y resalta paquetes sin saturar el diseño.
            """
        )

    col_a, col_b = st.columns([1.3, 1])
    with col_a:
        page_label = st.radio(
            "Formato PDF",
            list(PAGE_OPTIONS.keys()),
            index=0,
            horizontal=True,
            help="Carta horizontal es la opción más compatible. Oficio horizontal mete aún más entregas por hoja si tu impresora lo soporta.",
        )
    with col_b:
        st.markdown("**Recomendación:** usa *Carta horizontal* si imprimes en equipos estándar.")

    uploaded_file = st.file_uploader("Sube el archivo Excel", type=["xlsx", "xls"])

    if uploaded_file is None:
        st.info("Esperando archivo Excel para procesar.")
        return

    try:
        with st.spinner("Procesando archivo y generando PDF horizontal compacto..."):
            summary_df, control_df, detail_df, excel_buffer, pdf_buffer, warnings, orders = generate_files(
                uploaded_file,
                page_label,
            )

        st.success("Archivo procesado correctamente.")
        render_metrics(orders)

        if warnings:
            with st.expander("Ver advertencias de lectura"):
                for warning in warnings:
                    st.warning(warning)

        tab1, tab2, tab3 = st.tabs(["Control", "Detalle", "Descargas"])

        with tab1:
            st.dataframe(control_df, use_container_width=True, height=430)

        with tab2:
            st.dataframe(detail_df, use_container_width=True, height=430)

        with tab3:
            st.dataframe(summary_df, use_container_width=True, height=220)
            render_downloads(excel_buffer, pdf_buffer)

    except Exception as error:
        st.error(f"Error al procesar el archivo: {error}")


if __name__ == "__main__":
    main()
