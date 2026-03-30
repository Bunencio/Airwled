
from __future__ import annotations

import io
import re
import unicodedata
import zipfile
from collections import Counter
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import BinaryIO, Iterable, Sequence
from xml.sax.saxutils import escape

import pandas as pd
try:
    import streamlit as st
except ModuleNotFoundError:  # pragma: no cover - permite probar la lógica sin Streamlit
    class _StreamlitStub:
        def __getattr__(self, name):
            raise ModuleNotFoundError(
                "streamlit no está instalado. Instálalo con 'pip install streamlit' para usar la interfaz."
            )

    st = _StreamlitStub()
from reportlab.lib import colors
from reportlab.lib.pagesizes import legal, letter, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Flowable, LongTable, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


APP_TITLE = "Centro profesional de empaque"
DEFAULT_HEADER_ROW = 5  # índice humano (1-based)
PACKAGE_PATTERN = re.compile(r"(?:paquete|pack|paq)\s*(?:de|x)?\s*(\d+)", re.IGNORECASE)

PAGE_OPTIONS: dict[str, tuple[float, float]] = {
    "Carta horizontal": landscape(letter),
    "Oficio horizontal": landscape(legal),
}

PAGE_MARGINS = {
    "left": 0.36 * inch,
    "right": 0.36 * inch,
    "top": 0.58 * inch,
    "bottom": 0.48 * inch,
}

COLOR_TEXT = colors.HexColor("#0F172A")
COLOR_HEADER = colors.HexColor("#111827")
COLOR_MUTED = colors.HexColor("#64748B")
COLOR_BORDER = colors.HexColor("#CBD5E1")
COLOR_BORDER_DARK = colors.HexColor("#94A3B8")
COLOR_SURFACE = colors.HexColor("#FFFFFF")
COLOR_SURFACE_ALT = colors.HexColor("#F8FAFC")
COLOR_PACKAGE_HEAD = colors.HexColor("#E2E8F0")
COLOR_PACKAGE_BODY = colors.HexColor("#F8FAFC")
COLOR_NOTE_BG = colors.HexColor("#EFF6FF")
COLOR_WARNING_BG = colors.HexColor("#FFF7ED")
WARNING_TEXT_HEX = "#9A3412"

APP_CSS = """
<style>
    .stApp {
        background: linear-gradient(180deg, #f8fafc 0%, #ffffff 62%);
    }
    .main .block-container {
        max-width: 1380px;
        padding-top: 1.2rem;
        padding-bottom: 2rem;
    }
    .hero-card {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
        border-radius: 24px;
        padding: 1.35rem 1.5rem 1.25rem 1.5rem;
        color: #ffffff;
        box-shadow: 0 20px 40px rgba(15, 23, 42, 0.14);
        margin-bottom: 0.9rem;
    }
    .hero-card h1 {
        margin: 0 0 0.35rem 0;
        font-size: 2rem;
        line-height: 1.1;
        color: #ffffff;
    }
    .hero-card p {
        margin: 0;
        color: #dbeafe;
        font-size: 0.98rem;
        line-height: 1.5;
    }
    .feature-grid {
        display: grid;
        grid-template-columns: repeat(3, minmax(0, 1fr));
        gap: 12px;
        margin: 0.35rem 0 1rem 0;
    }
    .feature-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 18px;
        padding: 0.95rem 1rem 0.9rem 1rem;
        box-shadow: 0 10px 26px rgba(15, 23, 42, 0.05);
    }
    .feature-card h3 {
        margin: 0 0 0.28rem 0;
        color: #0f172a;
        font-size: 1rem;
    }
    .feature-card p {
        margin: 0;
        color: #475569;
        font-size: 0.92rem;
        line-height: 1.45;
    }
    .subtle-note {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 18px;
        padding: 0.9rem 1rem;
        margin-bottom: 0.9rem;
        color: #334155;
    }
    .subtle-note strong {
        color: #0f172a;
    }
    div[data-testid="stDownloadButton"] > button {
        width: 100%;
        border-radius: 999px;
        border: 1px solid #cbd5e1;
        font-weight: 700;
        padding: 0.7rem 1rem;
    }
    div[data-testid="stDataFrame"] {
        border: 1px solid #e2e8f0;
        border-radius: 18px;
        overflow: hidden;
    }
    [data-testid="stMetric"] {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 18px;
        padding: 0.65rem 0.9rem;
        box-shadow: 0 8px 22px rgba(15, 23, 42, 0.04);
    }
    @media (max-width: 980px) {
        .feature-grid {
            grid-template-columns: 1fr;
        }
    }
</style>
"""


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
    units_raw: str
    sku: str
    title: str
    quantity: Decimal | None

    @property
    def sku_display(self) -> str:
        return self.sku or "SIN SKU"

    @property
    def title_display(self) -> str:
        return self.title or "Sin descripción"

    @property
    def quantity_display(self) -> str:
        if self.quantity is not None:
            return format_decimal(self.quantity)
        return self.units_raw or "-"


@dataclass(frozen=True)
class OrderGroup:
    main_sale: str
    items: tuple[LineItem, ...]
    is_package_group: bool
    raw_state: str = ""

    @property
    def item_count(self) -> int:
        return len(self.items)

    @property
    def short_type(self) -> str:
        return f"PAQ x{self.item_count}" if self.is_package_group else "INDIV."

    @property
    def delivery_type_display(self) -> str:
        return f"PAQUETE ({self.item_count} productos)" if self.is_package_group else "INDIVIDUAL"

    @property
    def total_numeric_quantity(self) -> Decimal:
        return sum((item.quantity or Decimal("0") for item in self.items), Decimal("0"))


@dataclass(frozen=True)
class PreparationEntry:
    sku_key: str
    sku_display: str
    titles: tuple[str, ...]
    total_quantity: Decimal | None
    line_count: int
    order_count: int
    orders: tuple[str, ...]
    omitted_quantity_count: int = 0

    @property
    def total_display(self) -> str:
        if self.total_quantity is None:
            return "-"
        return format_decimal(self.total_quantity)

    @property
    def needs_attention(self) -> bool:
        return self.sku_display == "SIN SKU" or self.omitted_quantity_count > 0 or self.total_quantity is None


@dataclass
class GeneratedOutputs:
    summary_df: pd.DataFrame
    packing_preview_df: pd.DataFrame
    preparation_preview_df: pd.DataFrame
    packing_pdf_buffer: io.BytesIO
    preparation_pdf_buffer: io.BytesIO
    zip_buffer: io.BytesIO
    warnings: list[str]
    orders: list[OrderGroup]
    preparations: list[PreparationEntry]
    file_names: dict[str, str]
    generated_at_display: str


COLUMN_RULES: tuple[ColumnRule, ...] = (
    ColumnRule("sale", 0, ("venta", "pedido", "order", "folio", "venta principal")),
    ColumnRule("state", 2, ("estado", "status", "entrega", "envio", "envío")),
    ColumnRule("units", 6, ("unidades", "cantidad", "cant", "qty", "units")),
    ColumnRule("sku", 16, ("sku", "seller sku", "codigo", "código", "asin", "referencia")),
    ColumnRule(
        "title",
        20,
        (
            "titulo",
            "título",
            "producto",
            "descripcion",
            "descripción",
            "articulo",
            "artículo",
            "titulo de la publicacion",
        ),
    ),
)


class Checkbox(Flowable):
    """Casilla vacía para marcado manual."""

    def __init__(self, size: float = 9.6, stroke_width: float = 1.0) -> None:
        super().__init__()
        self.size = size
        self.stroke_width = stroke_width
        self.width = size
        self.height = size

    def draw(self) -> None:
        self.canv.setLineWidth(self.stroke_width)
        self.canv.rect(0, 0, self.size, self.size)


# -----------------------------------------------------------------------------
# UTILIDADES GENERALES
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


def dedupe_preserve_order(values: Iterable[str]) -> list[str]:
    seen: set[str] = set()
    ordered: list[str] = []
    for value in values:
        if value not in seen:
            seen.add(value)
            ordered.append(value)
    return ordered


def slugify_filename(name: str) -> str:
    text = normalize_text(name)
    text = re.sub(r"[^a-z0-9]+", "_", text).strip("_")
    return text or "archivo"


def normalize_sku_key(value: str) -> str:
    text = clean_value(value).upper()
    text = re.sub(r"\s+", "", text)
    return text or "SIN SKU"


def parse_quantity(value: object) -> Decimal | None:
    text = clean_value(value)
    if not text:
        return None

    candidate = re.sub(r"[^0-9,\.\-]", "", text.replace(" ", ""))
    if not candidate or candidate in {"-", ".", ",", "-.", "-,"}:
        return None

    if "," in candidate and "." in candidate:
        if candidate.rfind(",") > candidate.rfind("."):
            candidate = candidate.replace(".", "").replace(",", ".")
        else:
            candidate = candidate.replace(",", "")
    elif candidate.count(",") > 1 and "." not in candidate:
        candidate = candidate.replace(",", "")
    elif candidate.count(".") > 1 and "," not in candidate:
        candidate = candidate.replace(".", "")
    elif "," in candidate:
        left, right = candidate.split(",", 1)
        if len(right) == 3 and left not in {"", "-"}:
            candidate = left + right
        else:
            candidate = left + "." + right
    elif "." in candidate:
        left, right = candidate.split(".", 1)
        if len(right) == 3 and left not in {"", "-"}:
            candidate = left + right

    try:
        return Decimal(candidate)
    except InvalidOperation:
        return None


def format_decimal(value: Decimal | None) -> str:
    if value is None:
        return "-"
    text = format(value.normalize(), "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text or "0"


def build_line_item(sale_id: object, units_raw: object, sku: object, title: object) -> LineItem:
    units_clean = clean_value(units_raw)
    return LineItem(
        sale_id=clean_value(sale_id),
        units_raw=units_clean,
        sku=clean_value(sku),
        title=clean_value(title),
        quantity=parse_quantity(units_clean),
    )


def has_meaningful_item_data(row: pd.Series, columns: ResolvedColumns) -> bool:
    return any(clean_value(row[column]) for column in (columns.sku, columns.title, columns.units))


def looks_like_header_artifact(row: pd.Series, columns: ResolvedColumns) -> bool:
    parts = [
        normalize_text(row[columns.sale]),
        normalize_text(row[columns.state]),
        normalize_text(row[columns.units]),
        normalize_text(row[columns.sku]),
        normalize_text(row[columns.title]),
    ]
    joined = " ".join(part for part in parts if part)
    if not joined:
        return True

    suspicious_markers = (
        "# de venta",
        "venta principal",
        "titulo de la publicacion",
        "contenido verificado",
        "iniciales / hora",
        "unidades x sku",
        "preparacion previa por sku",
        "productos asociados",
        "lista de empaque",
    )
    return any(marker in joined for marker in suspicious_markers)


def extract_package_size_from_row(row: pd.Series, columns: ResolvedColumns) -> int | None:
    for candidate in (clean_value(row[columns.state]), clean_value(row[columns.title])):
        match = PACKAGE_PATTERN.search(candidate)
        if match:
            return int(match.group(1))
    return None




def row_starts_new_order(row: pd.Series, columns: ResolvedColumns) -> bool:
    return bool(clean_value(row[columns.sale]))


def row_has_same_sale(row: pd.Series, columns: ResolvedColumns, main_sale: str) -> bool:
    sale = clean_value(row[columns.sale])
    return bool(sale) and sale == main_sale


def row_has_package_signal(row: pd.Series, columns: ResolvedColumns) -> bool:
    return extract_package_size_from_row(row, columns) is not None


def collect_package_items(
    df: pd.DataFrame,
    start_index: int,
    columns: ResolvedColumns,
    main_sale: str,
    package_size: int,
) -> tuple[list[LineItem], int]:
    """
    Consolida los productos del mismo paquete.

    Reglas:
    - la fila que dice "Paquete de N" puede traer el primer producto y debe incluirse;
    - filas siguientes sin número de venta suelen ser continuación del mismo paquete;
    - si aparece otra venta distinta, termina el paquete;
    - si aparece otra fila con nueva señal de paquete, también termina el paquete.
    """
    items: list[LineItem] = []
    cursor = start_index

    while cursor < len(df):
        row = df.iloc[cursor]

        if looks_like_header_artifact(row, columns):
            cursor += 1
            continue

        if cursor > start_index and row_has_package_signal(row, columns):
            break

        if cursor > start_index and row_starts_new_order(row, columns) and not row_has_same_sale(row, columns, main_sale):
            break

        if has_meaningful_item_data(row, columns):
            items.append(
                build_line_item(
                    sale_id=clean_value(row[columns.sale]) or main_sale,
                    units_raw=row[columns.units],
                    sku=row[columns.sku],
                    title=row[columns.title],
                )
            )
            cursor += 1
            if len(items) >= package_size:
                # Seguimos absorbiendo renglones vacíos de venta solo si claramente son continuación.
                while cursor < len(df):
                    next_row = df.iloc[cursor]
                    if looks_like_header_artifact(next_row, columns):
                        cursor += 1
                        continue
                    if row_has_package_signal(next_row, columns):
                        break
                    if row_starts_new_order(next_row, columns) and not row_has_same_sale(next_row, columns, main_sale):
                        break
                    if not has_meaningful_item_data(next_row, columns):
                        cursor += 1
                        continue
                    # permite capturar paquetes mal etiquetados donde realmente vienen más de N productos
                    items.append(
                        build_line_item(
                            sale_id=clean_value(next_row[columns.sale]) or main_sale,
                            units_raw=next_row[columns.units],
                            sku=next_row[columns.sku],
                            title=next_row[columns.title],
                        )
                    )
                    cursor += 1
                break
            continue

        if cursor > start_index:
            break
        cursor += 1

    return items, cursor

def build_output_filenames(source_name: str, generated_at: datetime) -> dict[str, str]:
    source_slug = slugify_filename(Path(source_name).stem)
    stamp = generated_at.strftime("%Y%m%d_%H%M%S")
    prefix = f"{source_slug}_{stamp}"
    return {
        "packing_pdf": f"{prefix}_lista_empaque.pdf",
        "preparation_pdf": f"{prefix}_preparacion_sku.pdf",
        "zip": f"{prefix}_pdfs_empaque.zip",
    }


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


def load_source_dataframe(excel_file: BinaryIO, header_row_number: int) -> tuple[pd.DataFrame, ResolvedColumns, list[str]]:
    if header_row_number < 1:
        raise ValueError("La fila del encabezado debe ser 1 o mayor.")

    df = pd.read_excel(
        excel_file,
        header=header_row_number - 1,
        dtype=object,
        keep_default_na=False,
    )
    df = df.dropna(how="all").reset_index(drop=True)

    if df.empty:
        raise ValueError("El archivo no contiene datos válidos después de la fila de encabezados.")

    columns, warnings = resolve_source_columns(df)
    return df, columns, warnings


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

        main_sale = clean_value(row[columns.sale]) or f"SIN-VENTA-{i + 1}"
        raw_state = clean_value(row[columns.state])
        package_size = extract_package_size_from_row(row, columns)

        if package_size:
            package_items, next_index = collect_package_items(
                df=df,
                start_index=i,
                columns=columns,
                main_sale=main_sale,
                package_size=package_size,
            )

            if package_items:
                if len(package_items) < package_size:
                    warnings.append(
                        f"La venta '{main_sale}' indica paquete de {package_size}, pero solo se pudieron consolidar {len(package_items)} artículo(s) del mismo paquete."
                    )
                elif len(package_items) > package_size:
                    warnings.append(
                        f"La venta '{main_sale}' indica paquete de {package_size}, pero se detectaron {len(package_items)} productos asociados en la misma caja. Se respetó el agrupado real."
                    )

                orders.append(
                    OrderGroup(
                        main_sale=main_sale,
                        items=tuple(package_items),
                        is_package_group=len(package_items) > 1,
                        raw_state=raw_state,
                    )
                )
                i = next_index
                continue

            warnings.append(
                f"La venta '{main_sale}' estaba marcada como paquete, pero no se encontraron artículos válidos. Se exportó como individual."
            )

        if not has_meaningful_item_data(row, columns) and not clean_value(row[columns.sale]):
            i += 1
            continue

        orders.append(
            OrderGroup(
                main_sale=main_sale,
                items=(
                    build_line_item(
                        sale_id=main_sale,
                        units_raw=row[columns.units],
                        sku=row[columns.sku],
                        title=row[columns.title],
                    ),
                ),
                is_package_group=False,
                raw_state=raw_state,
            )
        )
        i += 1

    if skipped_artifacts:
        warnings.append(
            f"Se omitieron {skipped_artifacts} fila(s) que parecían encabezados, texto auxiliar o filas completamente vacías dentro del Excel."
        )

    return orders, warnings


# -----------------------------------------------------------------------------
# CONSOLIDACIÓN POR SKU Y VISTAS PREVIAS
# -----------------------------------------------------------------------------


def build_preparation_entries(orders: Iterable[OrderGroup]) -> tuple[list[PreparationEntry], list[str]]:
    grouped: dict[str, dict[str, object]] = {}
    omitted_examples: list[str] = []
    missing_sku_lines = 0

    for order in orders:
        for item in order.items:
            sku_key = normalize_sku_key(item.sku)
            if sku_key == "SIN SKU":
                missing_sku_lines += 1

            if sku_key not in grouped:
                grouped[sku_key] = {
                    "sku_display": item.sku_display,
                    "title_counter": Counter(),
                    "line_count": 0,
                    "orders": [],
                    "orders_seen": set(),
                    "total_quantity": Decimal("0"),
                    "has_numeric": False,
                    "omitted_quantity_count": 0,
                }

            entry = grouped[sku_key]
            title_counter = entry["title_counter"]
            assert isinstance(title_counter, Counter)
            title_counter[item.title_display] += 1

            entry["line_count"] = int(entry["line_count"]) + 1

            orders_seen = entry["orders_seen"]
            assert isinstance(orders_seen, set)
            if order.main_sale not in orders_seen:
                orders_seen.add(order.main_sale)
                orders_list = entry["orders"]
                assert isinstance(orders_list, list)
                orders_list.append(order.main_sale)

            if item.quantity is None:
                entry["omitted_quantity_count"] = int(entry["omitted_quantity_count"]) + 1
                if item.units_raw:
                    omitted_examples.append(
                        f"SKU {item.sku_display} en venta {order.main_sale}: cantidad '{item.units_raw}' no se pudo sumar al consolidado."
                    )
                continue

            entry["has_numeric"] = True
            entry["total_quantity"] = Decimal(entry["total_quantity"]) + item.quantity

    preparation_entries: list[PreparationEntry] = []
    for sku_key, raw in grouped.items():
        title_counter = raw["title_counter"]
        assert isinstance(title_counter, Counter)
        titles = tuple(title for title, _ in title_counter.most_common())
        orders_list = raw["orders"]
        assert isinstance(orders_list, list)
        has_numeric = bool(raw["has_numeric"])
        total_quantity = Decimal(raw["total_quantity"]) if has_numeric else None

        preparation_entries.append(
            PreparationEntry(
                sku_key=sku_key,
                sku_display=str(raw["sku_display"]),
                titles=titles,
                total_quantity=total_quantity,
                line_count=int(raw["line_count"]),
                order_count=len(orders_list),
                orders=tuple(orders_list),
                omitted_quantity_count=int(raw["omitted_quantity_count"]),
            )
        )

    preparation_entries.sort(
        key=lambda entry: (
            entry.sku_display == "SIN SKU",
            entry.total_quantity is None,
            -(entry.total_quantity or Decimal("0")),
            entry.sku_display,
        )
    )

    warnings: list[str] = []
    omitted_count = sum(entry.omitted_quantity_count for entry in preparation_entries)
    if omitted_count:
        warnings.append(
            f"Se omitieron {omitted_count} cantidad(es) no numéricas al consolidar el PDF de preparación por SKU."
        )
        warnings.extend(omitted_examples[:12])
        if len(omitted_examples) > 12:
            warnings.append(f"Hay {len(omitted_examples) - 12} caso(s) adicional(es) con cantidades no numéricas omitidas.")

    if missing_sku_lines:
        warnings.append(
            f"Se detectaron {missing_sku_lines} línea(s) sin SKU. Se agruparon como 'SIN SKU' en el PDF de preparación; conviene revisarlas manualmente."
        )

    return preparation_entries, warnings


def build_summary_dataframe(orders: list[OrderGroup], preparations: list[PreparationEntry]) -> pd.DataFrame:
    package_count = sum(1 for order in orders if order.is_package_group)
    individual_count = len(orders) - package_count
    total_lines = sum(order.item_count for order in orders)
    total_numeric_units = sum((order.total_numeric_quantity for order in orders), Decimal("0"))
    omitted_in_preparation = sum(entry.omitted_quantity_count for entry in preparations)
    lines_without_sku = sum(1 for order in orders for item in order.items if item.sku_display == "SIN SKU")
    preparation_total = sum((entry.total_quantity or Decimal("0") for entry in preparations), Decimal("0"))

    return pd.DataFrame(
        [
            {"Métrica": "Total de entregas", "Valor": str(len(orders))},
            {"Métrica": "Pedidos individuales", "Valor": str(individual_count)},
            {"Métrica": "Paquetes", "Valor": str(package_count)},
            {"Métrica": "Líneas de producto", "Valor": str(total_lines)},
            {"Métrica": "SKU únicos para preparación", "Valor": str(len(preparations))},
            {"Métrica": "Unidades numéricas detectadas", "Valor": format_decimal(total_numeric_units)},
            {"Métrica": "Unidades totales a preparar", "Valor": format_decimal(preparation_total)},
            {"Métrica": "Líneas sin SKU", "Valor": str(lines_without_sku)},
            {"Métrica": "Cantidades omitidas del consolidado", "Valor": str(omitted_in_preparation)},
        ]
    )


def build_packing_preview_dataframe(orders: Iterable[OrderGroup]) -> pd.DataFrame:
    rows: list[dict[str, object]] = []
    for order_number, order in enumerate(orders, start=1):
        for item_index, item in enumerate(order.items, start=1):
            rows.append(
                {
                    "No.": order_number if item_index == 1 else "",
                    "OK": "☐" if item_index == 1 else "",
                    "Venta principal": order.main_sale if item_index == 1 else "",
                    "Tipo": order.short_type if item_index == 1 else "",
                    "SKU": item.sku_display,
                    "Producto": f"{item_index}. {item.title_display}" if order.item_count > 1 else item.title_display,
                    "Cantidad": item.quantity_display,
                }
            )
    return pd.DataFrame(rows)


def build_preparation_preview_dataframe(entries: Iterable[PreparationEntry]) -> pd.DataFrame:
    rows: list[dict[str, object]] = []
    for entry in entries:
        rows.append(
            {
                "SKU": entry.sku_display,
                "Productos asociados": "\n".join(entry.titles) if entry.titles else "Sin descripción",
                "Total unidades": entry.total_display,
                "Líneas": entry.line_count,
                "Ventas": entry.order_count,
                "Ventas principales": ", ".join(entry.orders),
                "Cantidades omitidas": entry.omitted_quantity_count,
            }
        )
    return pd.DataFrame(rows)


# -----------------------------------------------------------------------------
# ESTILOS Y BLOQUES PDF
# -----------------------------------------------------------------------------


def build_styles() -> dict[str, ParagraphStyle]:
    sample = getSampleStyleSheet()
    return {
        "title": ParagraphStyle(
            "TitleMain",
            parent=sample["Title"],
            fontName="Helvetica-Bold",
            fontSize=16.4,
            leading=18.6,
            textColor=COLOR_TEXT,
            spaceAfter=1,
        ),
        "subtitle": ParagraphStyle(
            "Subtitle",
            parent=sample["BodyText"],
            fontName="Helvetica",
            fontSize=8.4,
            leading=10.4,
            textColor=COLOR_MUTED,
        ),
        "meta": ParagraphStyle(
            "Meta",
            parent=sample["BodyText"],
            fontName="Helvetica",
            fontSize=7.5,
            leading=9.1,
            textColor=COLOR_TEXT,
            alignment=2,
        ),
        "metric_label": ParagraphStyle(
            "MetricLabel",
            parent=sample["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=6.5,
            leading=7.8,
            textColor=COLOR_MUTED,
        ),
        "metric_value": ParagraphStyle(
            "MetricValue",
            parent=sample["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=12.8,
            leading=14.4,
            textColor=COLOR_TEXT,
        ),
        "note": ParagraphStyle(
            "Note",
            parent=sample["BodyText"],
            fontName="Helvetica",
            fontSize=7.25,
            leading=8.9,
            textColor=COLOR_TEXT,
        ),
        "table_header": ParagraphStyle(
            "TableHeader",
            parent=sample["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=7.8,
            leading=9.3,
            textColor=colors.white,
            alignment=1,
        ),
        "cell_center": ParagraphStyle(
            "CellCenter",
            parent=sample["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=7.2,
            leading=8.8,
            textColor=COLOR_TEXT,
            alignment=1,
        ),
        "cell_sale": ParagraphStyle(
            "CellSale",
            parent=sample["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=7.2,
            leading=8.8,
            textColor=COLOR_TEXT,
        ),
        "cell_type": ParagraphStyle(
            "CellType",
            parent=sample["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=7.0,
            leading=8.6,
            textColor=COLOR_TEXT,
            alignment=1,
        ),
        "cell_sku": ParagraphStyle(
            "CellSKU",
            parent=sample["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=7.1,
            leading=8.7,
            textColor=COLOR_TEXT,
        ),
        "cell_product": ParagraphStyle(
            "CellProduct",
            parent=sample["BodyText"],
            fontName="Helvetica",
            fontSize=7.2,
            leading=9.0,
            textColor=COLOR_TEXT,
            splitLongWords=True,
        ),
        "cell_qty": ParagraphStyle(
            "CellQty",
            parent=sample["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=7.3,
            leading=8.8,
            textColor=COLOR_TEXT,
            alignment=1,
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
                ("BACKGROUND", (0, 0), (-1, -1), COLOR_SURFACE),
                ("BOX", (0, 0), (-1, -1), 0.9, COLOR_BORDER_DARK),
                ("LEFTPADDING", (0, 0), (-1, -1), 7),
                ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )
    )
    return table


def build_intro_block(
    *,
    title: str,
    subtitle: str,
    source_name: str,
    page_label: str,
    generated_at_display: str,
    metric_pairs: list[tuple[str, str]],
    note_html: str,
    styles: dict[str, ParagraphStyle],
    doc_width: float,
    note_bg: colors.Color = COLOR_NOTE_BG,
) -> list[object]:
    top_table = Table(
        [
            [
                Paragraph(escape(title), styles["title"]),
                Paragraph(
                    f"<b>Generado:</b> {escape(generated_at_display)}<br/>"
                    f"<b>Fuente:</b> {escape(source_name)}<br/>"
                    f"<b>Formato:</b> {escape(page_label)}",
                    styles["meta"],
                ),
            ]
        ],
        colWidths=[doc_width * 0.68, doc_width * 0.32],
    )
    top_table.setStyle(
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

    metric_width = doc_width / max(len(metric_pairs), 1)
    metrics_table = Table(
        [[make_metric_box(label, value, metric_width - 6, styles) for label, value in metric_pairs]],
        colWidths=[metric_width] * len(metric_pairs),
    )
    metrics_table.setStyle(
        TableStyle(
            [
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ]
        )
    )

    note_table = Table([[Paragraph(note_html, styles["note"])]], colWidths=[doc_width])
    note_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), note_bg),
                ("BOX", (0, 0), (-1, -1), 0.85, COLOR_BORDER),
                ("LEFTPADDING", (0, 0), (-1, -1), 7),
                ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )

    return [
        top_table,
        Paragraph(escape(subtitle), styles["subtitle"]),
        Spacer(1, 6),
        metrics_table,
        Spacer(1, 6),
        note_table,
        Spacer(1, 8),
    ]


def make_page_decorator(header_label: str, footer_label: str, generated_at_display: str):
    def decorator(canvas, doc) -> None:
        page_width, page_height = doc.pagesize
        canvas.saveState()
        canvas.setStrokeColor(COLOR_BORDER_DARK)
        canvas.setFillColor(COLOR_TEXT)
        canvas.setLineWidth(0.75)
        canvas.line(doc.leftMargin, page_height - 0.28 * inch, page_width - doc.rightMargin, page_height - 0.28 * inch)

        canvas.setFont("Helvetica-Bold", 8.7)
        canvas.drawString(doc.leftMargin, page_height - 0.21 * inch, header_label.upper())

        canvas.setFont("Helvetica", 7.8)
        canvas.drawRightString(page_width - doc.rightMargin, page_height - 0.21 * inch, f"Pagina {canvas.getPageNumber()}")

        canvas.setLineWidth(0.55)
        canvas.line(doc.leftMargin, 0.30 * inch, page_width - doc.rightMargin, 0.30 * inch)
        canvas.setFont("Helvetica", 7.0)
        canvas.drawString(doc.leftMargin, 0.16 * inch, footer_label)
        canvas.drawRightString(page_width - doc.rightMargin, 0.16 * inch, generated_at_display)
        canvas.restoreState()

    return decorator


def build_packing_table(orders: list[OrderGroup], width: float, styles: dict[str, ParagraphStyle]) -> LongTable:
    fixed_widths = {
        "number": 0.40 * inch,
        "ok": 0.48 * inch,
        "sale": 1.46 * inch,
        "type": 0.90 * inch,
        "sku": 1.42 * inch,
        "qty": 0.74 * inch,
    }
    product_width = width - sum(fixed_widths.values())
    col_widths = [
        fixed_widths["number"],
        fixed_widths["ok"],
        fixed_widths["sale"],
        fixed_widths["type"],
        fixed_widths["sku"],
        product_width,
        fixed_widths["qty"],
    ]

    data: list[list[object]] = [
        [
            Paragraph("No.", styles["table_header"]),
            Paragraph("OK", styles["table_header"]),
            Paragraph("Venta principal", styles["table_header"]),
            Paragraph("Tipo", styles["table_header"]),
            Paragraph("SKU", styles["table_header"]),
            Paragraph("Producto", styles["table_header"]),
            Paragraph("Cant.", styles["table_header"]),
        ]
    ]

    style_commands: list[tuple] = [
        ("BACKGROUND", (0, 0), (-1, 0), COLOR_HEADER),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("BOX", (0, 0), (-1, -1), 0.75, COLOR_BORDER_DARK),
        ("INNERGRID", (0, 0), (-1, -1), 0.35, COLOR_BORDER),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("ALIGN", (0, 0), (1, -1), "CENTER"),
        ("ALIGN", (3, 1), (3, -1), "CENTER"),
        ("ALIGN", (6, 1), (6, -1), "CENTER"),
    ]

    for order_number, order in enumerate(orders, start=1):
        for item_index, item in enumerate(order.items, start=1):
            product_text = f"{item_index}. {item.title_display}" if order.item_count > 1 else item.title_display
            data.append(
                [
                    Paragraph(str(order_number), styles["cell_center"]) if item_index == 1 else "",
                    Checkbox(9.6) if item_index == 1 else "",
                    Paragraph(escape(order.main_sale), styles["cell_sale"]) if item_index == 1 else "",
                    Paragraph(escape(order.short_type), styles["cell_type"]) if item_index == 1 else "",
                    Paragraph(escape(item.sku_display), styles["cell_sku"]),
                    Paragraph(escape(product_text), styles["cell_product"]),
                    Paragraph(escape(item.quantity_display), styles["cell_qty"]),
                ]
            )
            row_idx = len(data) - 1

            if order.is_package_group:
                background = COLOR_PACKAGE_HEAD if item_index == 1 else COLOR_PACKAGE_BODY
            else:
                background = COLOR_SURFACE if order_number % 2 else COLOR_SURFACE_ALT

            style_commands.append(("BACKGROUND", (0, row_idx), (-1, row_idx), background))

            if item_index == 1:
                style_commands.append(("LINEABOVE", (0, row_idx), (-1, row_idx), 0.65, COLOR_BORDER_DARK))
            if item_index == order.item_count:
                style_commands.append(("LINEBELOW", (0, row_idx), (-1, row_idx), 0.55, COLOR_BORDER))

    table = LongTable(data, colWidths=col_widths, repeatRows=1, splitByRow=1)
    table.setStyle(TableStyle(style_commands))
    return table


def build_preparation_product_html(entry: PreparationEntry) -> str:
    if entry.titles:
        visible_titles = list(entry.titles[:3])
        html = "<br/>".join(escape(title) for title in visible_titles)
    else:
        html = "Sin descripción"

    notes: list[str] = []
    extra_titles = max(len(entry.titles) - 3, 0)
    if extra_titles:
        notes.append(f"+ {extra_titles} nombre(s) adicional(es)")
    if entry.omitted_quantity_count:
        notes.append(f"{entry.omitted_quantity_count} cantidad(es) omitida(s) del total")
    if entry.sku_display == "SIN SKU":
        notes.append("Revisar: línea(s) sin SKU")
    if entry.total_quantity is None:
        notes.append("Sin cantidad numérica utilizable")

    if notes:
        html += f"<br/><font size='6.1' color='{WARNING_TEXT_HEX}'>{escape(' - '.join(notes))}</font>"
    return html


def build_preparation_table(entries: list[PreparationEntry], width: float, styles: dict[str, ParagraphStyle]) -> LongTable:
    fixed_widths = {
        "number": 0.42 * inch,
        "sku": 1.72 * inch,
        "total": 0.96 * inch,
        "lines": 0.68 * inch,
        "sales": 0.78 * inch,
    }
    products_width = width - sum(fixed_widths.values())
    col_widths = [
        fixed_widths["number"],
        fixed_widths["sku"],
        products_width,
        fixed_widths["total"],
        fixed_widths["lines"],
        fixed_widths["sales"],
    ]

    data: list[list[object]] = [
        [
            Paragraph("No.", styles["table_header"]),
            Paragraph("SKU", styles["table_header"]),
            Paragraph("Producto(s) asociados", styles["table_header"]),
            Paragraph("Total unid.", styles["table_header"]),
            Paragraph("Líneas", styles["table_header"]),
            Paragraph("Ventas", styles["table_header"]),
        ]
    ]

    style_commands: list[tuple] = [
        ("BACKGROUND", (0, 0), (-1, 0), COLOR_HEADER),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("BOX", (0, 0), (-1, -1), 0.75, COLOR_BORDER_DARK),
        ("INNERGRID", (0, 0), (-1, -1), 0.35, COLOR_BORDER),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("ALIGN", (0, 1), (0, -1), "CENTER"),
        ("ALIGN", (3, 1), (5, -1), "CENTER"),
    ]

    for row_number, entry in enumerate(entries, start=1):
        data.append(
            [
                Paragraph(str(row_number), styles["cell_center"]),
                Paragraph(escape(entry.sku_display), styles["cell_sku"]),
                Paragraph(build_preparation_product_html(entry), styles["cell_product"]),
                Paragraph(escape(entry.total_display), styles["cell_qty"]),
                Paragraph(str(entry.line_count), styles["cell_center"]),
                Paragraph(str(entry.order_count), styles["cell_center"]),
            ]
        )
        row_idx = len(data) - 1
        background = COLOR_WARNING_BG if entry.needs_attention else (COLOR_SURFACE if row_number % 2 else COLOR_SURFACE_ALT)
        style_commands.append(("BACKGROUND", (0, row_idx), (-1, row_idx), background))

    table = LongTable(data, colWidths=col_widths, repeatRows=1, splitByRow=1)
    table.setStyle(TableStyle(style_commands))
    return table


# -----------------------------------------------------------------------------
# GENERACIÓN DE PDFS
# -----------------------------------------------------------------------------


def build_packing_pdf_buffer(
    orders: list[OrderGroup],
    source_name: str,
    page_label: str,
    generated_at_display: str,
) -> io.BytesIO:
    styles = build_styles()
    page_size = PAGE_OPTIONS[page_label]
    package_count = sum(1 for order in orders if order.is_package_group)
    total_lines = sum(order.item_count for order in orders)
    total_units = sum((order.total_numeric_quantity for order in orders), Decimal("0"))

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=page_size,
        leftMargin=PAGE_MARGINS["left"],
        rightMargin=PAGE_MARGINS["right"],
        topMargin=PAGE_MARGINS["top"],
        bottomMargin=PAGE_MARGINS["bottom"],
        title="Lista de empaque operativa",
        author="OpenAI",
    )

    elements: list[object] = []
    elements.extend(
        build_intro_block(
            title="Lista de empaque operativa",
            subtitle="Producto y cantidad separados por columna, con paquetes desplegados en renglones consecutivos.",
            source_name=source_name,
            page_label=page_label,
            generated_at_display=generated_at_display,
            metric_pairs=[
                ("Entregas", str(len(orders))),
                ("Paquetes", str(package_count)),
                ("Líneas", str(total_lines)),
                ("Unidades", format_decimal(total_units)),
            ],
            note_html=(
                "<b>Lectura rápida:</b> cada línea muestra un SKU específico con su <b>producto</b> y <b>cantidad</b> en columnas separadas. "
                "En los paquetes se deja una sola casilla OK por venta y los productos quedan debajo de la misma venta para que no se confundan. "
                "Se eliminó completamente la columna de iniciales / hora."
            ),
            styles=styles,
            doc_width=doc.width,
            note_bg=COLOR_NOTE_BG,
        )
    )
    elements.append(build_packing_table(orders, doc.width, styles))

    decorator = make_page_decorator(
        header_label="Lista de empaque operativa",
        footer_label="Control manual por venta con SKU, producto y cantidad separados",
        generated_at_display=generated_at_display,
    )
    doc.build(elements, onFirstPage=decorator, onLaterPages=decorator)
    buffer.seek(0)
    return buffer


def build_preparation_pdf_buffer(
    preparations: list[PreparationEntry],
    orders: list[OrderGroup],
    source_name: str,
    page_label: str,
    generated_at_display: str,
) -> io.BytesIO:
    styles = build_styles()
    page_size = PAGE_OPTIONS[page_label]
    preparation_total = sum((entry.total_quantity or Decimal("0") for entry in preparations), Decimal("0"))
    total_lines = sum(entry.line_count for entry in preparations)
    order_coverage = len({order.main_sale for order in orders})
    omitted_count = sum(entry.omitted_quantity_count for entry in preparations)
    missing_sku_entries = sum(1 for entry in preparations if entry.sku_display == "SIN SKU")

    note_color = COLOR_WARNING_BG if (omitted_count or missing_sku_entries) else COLOR_NOTE_BG
    note_text = (
        "<b>Objetivo:</b> este documento le dice al empleado qué debe preparar antes de empezar a empacar. "
        "Todo está consolidado estrictamente por <b>SKU</b>, así que si el mismo SKU aparece repetido en varios paquetes o ventas, "
        "aquí sale sumado en un solo renglón. El nombre del producto es solo una referencia visual."
    )
    if omitted_count or missing_sku_entries:
        note_text += (
            f"<br/><font color='{WARNING_TEXT_HEX}'><b>Atención:</b> cantidades omitidas del consolidado: {omitted_count}. "
            f"SKU vacíos agrupados como 'SIN SKU': {missing_sku_entries}.</font>"
        )

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=page_size,
        leftMargin=PAGE_MARGINS["left"],
        rightMargin=PAGE_MARGINS["right"],
        topMargin=PAGE_MARGINS["top"],
        bottomMargin=PAGE_MARGINS["bottom"],
        title="Preparación previa por SKU",
        author="OpenAI",
    )

    elements: list[object] = []
    elements.extend(
        build_intro_block(
            title="Preparación previa por SKU",
            subtitle="Totales consolidados para surtir antes de empezar a empacar.",
            source_name=source_name,
            page_label=page_label,
            generated_at_display=generated_at_display,
            metric_pairs=[
                ("SKU únicos", str(len(preparations))),
                ("Unidades", format_decimal(preparation_total)),
                ("Líneas", str(total_lines)),
                ("Ventas", str(order_coverage)),
            ],
            note_html=note_text,
            styles=styles,
            doc_width=doc.width,
            note_bg=note_color,
        )
    )
    elements.append(build_preparation_table(preparations, doc.width, styles))

    decorator = make_page_decorator(
        header_label="Preparación previa por SKU",
        footer_label="Totales consolidados por SKU para surtido previo al empaque",
        generated_at_display=generated_at_display,
    )
    doc.build(elements, onFirstPage=decorator, onLaterPages=decorator)
    buffer.seek(0)
    return buffer


def build_zip_buffer(
    *,
    packing_pdf_buffer: io.BytesIO,
    preparation_pdf_buffer: io.BytesIO,
    file_names: dict[str, str],
) -> io.BytesIO:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(file_names["packing_pdf"], packing_pdf_buffer.getvalue())
        archive.writestr(file_names["preparation_pdf"], preparation_pdf_buffer.getvalue())
    buffer.seek(0)
    return buffer


# -----------------------------------------------------------------------------
# API PRINCIPAL
# -----------------------------------------------------------------------------


def generate_outputs(
    *,
    excel_bytes: bytes,
    source_name: str,
    page_label: str,
    header_row_number: int,
) -> GeneratedOutputs:
    generated_at = datetime.now()
    generated_at_display = generated_at.strftime("%d/%m/%Y %H:%M")
    file_names = build_output_filenames(source_name, generated_at)

    source_df, columns, warnings = load_source_dataframe(io.BytesIO(excel_bytes), header_row_number)
    orders, parse_warnings = parse_orders(source_df, columns)
    warnings.extend(parse_warnings)

    if not orders:
        raise ValueError("No se encontraron ventas válidas para exportar.")

    preparations, preparation_warnings = build_preparation_entries(orders)
    warnings.extend(preparation_warnings)
    warnings = dedupe_preserve_order(warnings)

    summary_df = build_summary_dataframe(orders, preparations)
    packing_preview_df = build_packing_preview_dataframe(orders)
    preparation_preview_df = build_preparation_preview_dataframe(preparations)

    packing_pdf_buffer = build_packing_pdf_buffer(
        orders=orders,
        source_name=source_name,
        page_label=page_label,
        generated_at_display=generated_at_display,
    )
    preparation_pdf_buffer = build_preparation_pdf_buffer(
        preparations=preparations,
        orders=orders,
        source_name=source_name,
        page_label=page_label,
        generated_at_display=generated_at_display,
    )
    zip_buffer = build_zip_buffer(
        packing_pdf_buffer=packing_pdf_buffer,
        preparation_pdf_buffer=preparation_pdf_buffer,
        file_names=file_names,
    )

    return GeneratedOutputs(
        summary_df=summary_df,
        packing_preview_df=packing_preview_df,
        preparation_preview_df=preparation_preview_df,
        packing_pdf_buffer=packing_pdf_buffer,
        preparation_pdf_buffer=preparation_pdf_buffer,
        zip_buffer=zip_buffer,
        warnings=warnings,
        orders=orders,
        preparations=preparations,
        file_names=file_names,
        generated_at_display=generated_at_display,
    )


# -----------------------------------------------------------------------------
# INTERFAZ STREAMLIT
# -----------------------------------------------------------------------------


def inject_css() -> None:
    st.markdown(APP_CSS, unsafe_allow_html=True)


def render_header() -> None:
    st.markdown(
        f"""
        <div class="hero-card">
            <h1>{APP_TITLE}</h1>
            <p>
                Convierte tu Excel en dos PDFs operativos: una <strong>lista de empaque limpia</strong> con SKU, producto y cantidad separados,
                y una <strong>preparación previa por SKU</strong> para que el empleado deje surtido todo antes de empezar a empacar.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class="feature-grid">
            <div class="feature-card">
                <h3>PDF 1 · Empaque claro</h3>
                <p>Quita la columna de iniciales / hora y separa producto y cantidad para que no se vea raro ni se preste a confusión.</p>
            </div>
            <div class="feature-card">
                <h3>PDF 2 · Preparación por SKU</h3>
                <p>Suma automáticamente los SKU repetidos aunque estén en distintos paquetes para que el empleado prepare todo antes de empacar.</p>
            </div>
            <div class="feature-card">
                <h3>Parser robusto</h3>
                <p>Respeta paquetes, detecta columnas por nombre o por posición de respaldo, preserva SKU con ceros a la izquierda y muestra advertencias útiles.</p>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class="subtle-note">
            <strong>Tip:</strong> deja la fila de encabezado en <strong>5</strong> si tu archivo sigue el mismo formato del sistema actual.
            Si cambia la plantilla, ajusta la fila desde las opciones avanzadas y el generador seguirá funcionando.
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_metrics(outputs: GeneratedOutputs) -> None:
    package_count = sum(1 for order in outputs.orders if order.is_package_group)
    total_lines = sum(order.item_count for order in outputs.orders)
    total_units = sum((entry.total_quantity or Decimal("0") for entry in outputs.preparations), Decimal("0"))

    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Entregas", len(outputs.orders))
    col2.metric("Paquetes", package_count)
    col3.metric("Líneas", total_lines)
    col4.metric("SKU únicos", len(outputs.preparations))
    col5.metric("Unidades a preparar", format_decimal(total_units))


def render_downloads(outputs: GeneratedOutputs) -> None:
    col1, col2, col3 = st.columns(3)

    with col1:
        st.download_button(
            label="Descargar PDF de empaque",
            data=outputs.packing_pdf_buffer.getvalue(),
            file_name=outputs.file_names["packing_pdf"],
            mime="application/pdf",
            use_container_width=True,
        )

    with col2:
        st.download_button(
            label="Descargar PDF de preparación SKU",
            data=outputs.preparation_pdf_buffer.getvalue(),
            file_name=outputs.file_names["preparation_pdf"],
            mime="application/pdf",
            use_container_width=True,
        )

    with col3:
        st.download_button(
            label="Descargar ZIP con ambos PDFs",
            data=outputs.zip_buffer.getvalue(),
            file_name=outputs.file_names["zip"],
            mime="application/zip",
            use_container_width=True,
        )


def main() -> None:
    st.set_page_config(page_title="Empaque profesional", layout="wide")
    inject_css()
    render_header()

    controls_left, controls_right = st.columns([1.3, 1])
    with controls_left:
        page_label = st.radio(
            "Formato PDF",
            list(PAGE_OPTIONS.keys()),
            index=0,
            horizontal=True,
            help="Carta horizontal suele ser la opción más compatible. Oficio horizontal aprovecha aún más ancho si tu impresora lo soporta.",
        )
    with controls_right:
        with st.expander("Opciones avanzadas", expanded=False):
            header_row_number = st.number_input(
                "Fila del encabezado en Excel",
                min_value=1,
                max_value=50,
                value=DEFAULT_HEADER_ROW,
                step=1,
                help="Usa 5 si tus encabezados empiezan en la fila 5. El valor es humano: 1 = primera fila del archivo.",
            )

    uploaded_file = st.file_uploader("Sube el archivo Excel", type=["xlsx", "xls"])

    if uploaded_file is None:
        st.info("Sube tu archivo Excel para generar los dos PDFs profesionales.")
        return

    try:
        with st.spinner("Procesando archivo, separando producto/cantidad y construyendo los dos PDFs..."):
            outputs = generate_outputs(
                excel_bytes=uploaded_file.getvalue(),
                source_name=uploaded_file.name,
                page_label=page_label,
                header_row_number=int(header_row_number),
            )

        st.success("Listo. Se generaron la lista de empaque y el PDF de preparación por SKU.")
        render_metrics(outputs)

        if outputs.warnings:
            with st.expander("Ver advertencias de lectura y consolidación", expanded=False):
                for warning in outputs.warnings:
                    st.warning(warning)

        tabs = st.tabs(["Vista empaque", "Vista preparación SKU", "Resumen y descargas"])

        with tabs[0]:
            st.caption("Producto y cantidad ya aparecen en columnas separadas. Los paquetes quedan desplegados por renglón bajo la misma venta.")
            st.dataframe(outputs.packing_preview_df, use_container_width=True, height=540, hide_index=True)

        with tabs[1]:
            st.caption("Consolidado por SKU para preparar mercancía antes de empacar. Si un SKU aparece en varios paquetes, aquí ya está sumado.")
            st.dataframe(outputs.preparation_preview_df, use_container_width=True, height=540, hide_index=True)

        with tabs[2]:
            st.dataframe(outputs.summary_df, use_container_width=True, height=320, hide_index=True)
            render_downloads(outputs)

    except Exception as error:
        st.error(f"Error al procesar el archivo: {error}")


if __name__ == "__main__":
    main()
