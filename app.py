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
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (
    Flowable,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

# =============================================================================
# CONFIGURACION GENERAL
# =============================================================================

APP_TITLE = "2.0 Generador profesional de lista de empaque"
HEADER_ROW_INDEX = 4  # Fila 5 de Excel
PACKAGE_PATTERN = re.compile(r"paquete\s+de\s+(\d+)", re.IGNORECASE)

PAGE_MARGINS = {
    "left": 0.48 * inch,
    "right": 0.48 * inch,
    "top": 0.58 * inch,
    "bottom": 0.55 * inch,
}

COLOR_BLACK = colors.black
COLOR_WHITE = colors.white
COLOR_GRAY_05 = colors.HexColor("#F3F3F3")
COLOR_GRAY_10 = colors.HexColor("#E4E4E4")
COLOR_GRAY_20 = colors.HexColor("#D0D0D0")

# =============================================================================
# MODELOS DE DATOS
# =============================================================================


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
        numeric_values: list[float] = []
        found_non_numeric = False

        for item in self.items:
            if not item.units:
                continue
            value = to_number(item.units)
            if value is None:
                found_non_numeric = True
                continue
            numeric_values.append(value)

        if numeric_values and not found_non_numeric:
            total = sum(numeric_values)
            return str(int(total)) if float(total).is_integer() else f"{total:.2f}"
        return "-"


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

# =============================================================================
# COMPONENTES PDF
# =============================================================================


class Checkbox(Flowable):
    """Casilla vacía para marcar manualmente en la impresión."""

    def __init__(self, size: float = 11, stroke_width: float = 1.2) -> None:
        super().__init__()
        self.size = size
        self.stroke_width = stroke_width
        self.width = size
        self.height = size

    def draw(self) -> None:
        self.canv.setLineWidth(self.stroke_width)
        self.canv.rect(0, 0, self.size, self.size)


# =============================================================================
# UTILIDADES
# =============================================================================


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



def slugify_filename(name: str) -> str:
    text = normalize_text(name)
    text = re.sub(r"[^a-z0-9]+", "_", text).strip("_")
    return text or "archivo"


# =============================================================================
# LECTURA Y PARSEO DE EXCEL
# =============================================================================


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
    i = 0

    while i < len(df):
        row = df.iloc[i]
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
                sub_row = df.iloc[i + offset]
                items.append(
                    LineItem(
                        sale_id=clean_value(sub_row[columns.sale]),
                        units=clean_value(sub_row[columns.units]),
                        sku=clean_value(sub_row[columns.sku]),
                        title=clean_value(sub_row[columns.title]),
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
                delivery_type = f"JUNTO ({package_size} productos)"
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

    return orders, warnings



def build_output_dataframe(orders: Iterable[OrderGroup]) -> pd.DataFrame:
    rows = []
    for order in orders:
        rows.append(
            {
                "Estatus": "",
                "Venta principal": order.main_sale,
                "Tipo": order.delivery_type,
                "Productos": order.item_count,
                "Unidades totales": order.total_units,
                "SKU": "\n".join(item.sku or "-" for item in order.items),
                "Detalle": "\n".join(item.title or "-" for item in order.items),
            }
        )
    return pd.DataFrame(rows)


# =============================================================================
# EXPORTACION EXCEL
# =============================================================================


def build_excel_buffer(df_output: pd.DataFrame) -> io.BytesIO:
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_output.to_excel(writer, index=False, sheet_name="Lista")
        ws = writer.sheets["Lista"]
        ws.freeze_panes = "A2"

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
            ws.row_dimensions[row[0].row].height = max(20, min(16 * max_lines, 90))

        preferred_widths = {
            "A": 12,
            "B": 22,
            "C": 22,
            "D": 12,
            "E": 16,
            "F": 22,
            "G": 60,
        }

        for idx, column_name in enumerate(df_output.columns, start=1):
            column_letter = get_column_letter(idx)
            if column_letter in preferred_widths:
                ws.column_dimensions[column_letter].width = preferred_widths[column_letter]
                continue

            max_length = max(
                len(str(column_name)),
                *(len(str(value)) for value in df_output[column_name].fillna("")),
            )
            ws.column_dimensions[column_letter].width = min(max(max_length * 0.9, 12), 50)

    buffer.seek(0)
    return buffer


# =============================================================================
# ESTILOS PDF
# =============================================================================


def build_styles() -> dict[str, ParagraphStyle]:
    sample = getSampleStyleSheet()

    return {
        "title": ParagraphStyle(
            "TitleMain",
            parent=sample["Title"],
            fontName="Helvetica-Bold",
            fontSize=18,
            leading=22,
            textColor=COLOR_BLACK,
            spaceAfter=3,
        ),
        "subtitle": ParagraphStyle(
            "Subtitle",
            parent=sample["BodyText"],
            fontName="Helvetica",
            fontSize=9,
            leading=11,
            textColor=COLOR_BLACK,
            spaceAfter=0,
        ),
        "summary_label": ParagraphStyle(
            "SummaryLabel",
            parent=sample["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=8,
            leading=10,
            textColor=COLOR_BLACK,
        ),
        "summary_value": ParagraphStyle(
            "SummaryValue",
            parent=sample["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=13,
            leading=15,
            textColor=COLOR_BLACK,
        ),
        "card_label": ParagraphStyle(
            "CardLabel",
            parent=sample["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=8,
            leading=10,
            textColor=COLOR_BLACK,
        ),
        "card_value_large": ParagraphStyle(
            "CardValueLarge",
            parent=sample["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=14,
            leading=17,
            textColor=COLOR_BLACK,
        ),
        "card_value": ParagraphStyle(
            "CardValue",
            parent=sample["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=10,
            leading=12,
            textColor=COLOR_BLACK,
        ),
        "table_header": ParagraphStyle(
            "TableHeader",
            parent=sample["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=8.5,
            leading=10,
            textColor=COLOR_BLACK,
        ),
        "table_cell": ParagraphStyle(
            "TableCell",
            parent=sample["BodyText"],
            fontName="Helvetica",
            fontSize=10,
            leading=12,
            textColor=COLOR_BLACK,
        ),
        "table_cell_center": ParagraphStyle(
            "TableCellCenter",
            parent=sample["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=10,
            leading=12,
            alignment=1,
            textColor=COLOR_BLACK,
        ),
        "status": ParagraphStyle(
            "Status",
            parent=sample["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=8.5,
            leading=10,
            textColor=COLOR_BLACK,
        ),
    }


# =============================================================================
# CONSTRUCCION PDF
# =============================================================================


def make_summary_box(label: str, value: str, width: float, styles: dict[str, ParagraphStyle]) -> Table:
    table = Table(
        [
            [Paragraph(escape(label.upper()), styles["summary_label"])],
            [Paragraph(escape(value), styles["summary_value"])],
        ],
        colWidths=[width],
    )
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), COLOR_WHITE),
                ("BOX", (0, 0), (-1, -1), 1, COLOR_BLACK),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )
    )
    return table



def build_header_block(orders: list[OrderGroup], styles: dict[str, ParagraphStyle], doc_width: float) -> list:
    grouped_count = sum(1 for order in orders if order.delivery_type.startswith("JUNTO"))
    individual_count = sum(1 for order in orders if order.delivery_type == "INDIVIDUAL")
    generated_at = datetime.now().strftime("%d/%m/%Y %H:%M")

    summary_widths = [doc_width * 0.20, doc_width * 0.20, doc_width * 0.20, doc_width * 0.40]

    summary_table = Table(
        [[
            make_summary_box("Ventas", str(len(orders)), summary_widths[0], styles),
            make_summary_box("Individuales", str(individual_count), summary_widths[1], styles),
            make_summary_box("Juntos", str(grouped_count), summary_widths[2], styles),
            make_summary_box("Instrucción", "Marcar, revisar y firmar cada venta", summary_widths[3], styles),
        ]],
        colWidths=summary_widths,
    )
    summary_table.setStyle(
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

    return [
        Paragraph("Lista de empaque operativa", styles["title"]),
        Paragraph(
            (
                f"Formato optimizado para impresión en blanco y negro. "
                f"Generado el {generated_at}. "
                "La tipografía y las áreas de control fueron ajustadas para lectura rápida y marcado manual."
            ),
            styles["subtitle"],
        ),
        Spacer(1, 10),
        summary_table,
        Spacer(1, 12),
    ]



def build_order_header(order: OrderGroup, width: float, styles: dict[str, ParagraphStyle]) -> Table:
    col_widths = [width * 0.44, width * 0.28, width * 0.28]

    data = [[
        Paragraph(
            f"<b>VENTA PRINCIPAL</b><br/>{escape(order.main_sale)}",
            styles["card_value_large"],
        ),
        Paragraph(
            f"<b>TIPO DE ENTREGA</b><br/>{escape(order.delivery_type)}",
            styles["card_value"],
        ),
        Paragraph(
            f"<b>PRODUCTOS</b>: {order.item_count}<br/><b>UNIDADES</b>: {escape(order.total_units)}",
            styles["card_value"],
        ),
    ]]

    table = Table(data, colWidths=col_widths)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), COLOR_GRAY_10),
                ("BOX", (0, 0), (-1, -1), 1.0, COLOR_BLACK),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )
    )
    return table



def build_items_table(order: OrderGroup, width: float, styles: dict[str, ParagraphStyle]) -> Table:
    qty_width = 0.76 * inch
    sku_width = 1.75 * inch
    product_width = width - qty_width - sku_width

    data = [[
        Paragraph("CANT.", styles["table_header"]),
        Paragraph("SKU", styles["table_header"]),
        Paragraph("PRODUCTO", styles["table_header"]),
    ]]

    for item in order.items:
        data.append(
            [
                Paragraph(safe_paragraph_text(item.units), styles["table_cell_center"]),
                Paragraph(safe_paragraph_text(item.sku), styles["table_cell"]),
                Paragraph(safe_paragraph_text(item.title), styles["table_cell"]),
            ]
        )

    table = Table(data, colWidths=[qty_width, sku_width, product_width], repeatRows=1, splitByRow=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), COLOR_GRAY_20),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [COLOR_WHITE, COLOR_GRAY_05]),
                ("BOX", (0, 0), (-1, -1), 0.9, COLOR_BLACK),
                ("INNERGRID", (0, 0), (-1, -1), 0.55, COLOR_BLACK),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (0, -1), "CENTER"),
            ]
        )
    )
    return table



def build_status_table(width: float, styles: dict[str, ParagraphStyle]) -> Table:
    fixed_width = (0.20 + 0.94 + 0.20 + 1.08 + 0.20 + 1.12 + 1.40) * inch
    col_widths = [
        0.20 * inch,
        0.94 * inch,
        0.20 * inch,
        1.08 * inch,
        0.20 * inch,
        1.12 * inch,
        1.40 * inch,
        width - fixed_width,
    ]

    data = [[
        Checkbox(11),
        Paragraph("SURTIDO", styles["status"]),
        Checkbox(11),
        Paragraph("REVISADO", styles["status"]),
        Checkbox(11),
        Paragraph("ENTREGADO", styles["status"]),
        Paragraph("INICIALES: ________", styles["status"]),
        Paragraph("HORA: ________", styles["status"]),
    ]]

    table = Table(data, colWidths=col_widths)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), COLOR_WHITE),
                ("BOX", (0, 0), (-1, -1), 0.9, COLOR_BLACK),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )
    )
    return table



def draw_page_footer(canvas, doc) -> None:
    page_number = canvas.getPageNumber()
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(COLOR_BLACK)
    canvas.drawCentredString(letter[0] / 2, 0.28 * inch, f"Lista de empaque operativa - Página {page_number}")
    canvas.restoreState()



def build_pdf_buffer(orders: list[OrderGroup]) -> io.BytesIO:
    buffer = io.BytesIO()
    styles = build_styles()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        leftMargin=PAGE_MARGINS["left"],
        rightMargin=PAGE_MARGINS["right"],
        topMargin=PAGE_MARGINS["top"],
        bottomMargin=PAGE_MARGINS["bottom"],
        title="Lista de empaque operativa",
        author="OpenAI",
    )

    elements: list = []
    elements.extend(build_header_block(orders, styles, doc.width))

    for index, order in enumerate(orders, start=1):
        elements.append(build_order_header(order, doc.width, styles))
        elements.append(Spacer(1, 5))
        elements.append(build_items_table(order, doc.width, styles))
        elements.append(Spacer(1, 6))
        elements.append(build_status_table(doc.width, styles))
        if index != len(orders):
            elements.append(Spacer(1, 12))

    doc.build(elements, onFirstPage=draw_page_footer, onLaterPages=draw_page_footer)
    buffer.seek(0)
    return buffer


# =============================================================================
# API PRINCIPAL
# =============================================================================


def generate_files(excel_file: BinaryIO) -> tuple[pd.DataFrame, io.BytesIO, io.BytesIO, list[str], list[OrderGroup]]:
    source_df, columns, warnings = load_source_dataframe(excel_file)
    orders, parse_warnings = parse_orders(source_df, columns)
    warnings.extend(parse_warnings)

    if not orders:
        raise ValueError("No se encontraron ventas válidas para exportar.")

    output_df = build_output_dataframe(orders)
    excel_buffer = build_excel_buffer(output_df)
    pdf_buffer = build_pdf_buffer(orders)

    return output_df, excel_buffer, pdf_buffer, warnings, orders


# =============================================================================
# INTERFAZ STREAMLIT
# =============================================================================


def render_metrics(orders: list[OrderGroup]) -> None:
    total_orders = len(orders)
    grouped_orders = sum(1 for order in orders if order.delivery_type.startswith("JUNTO"))
    individual_orders = total_orders - grouped_orders
    total_items = sum(order.item_count for order in orders)

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Ventas", total_orders)
    col2.metric("Individuales", individual_orders)
    col3.metric("Juntos", grouped_orders)
    col4.metric("Productos listados", total_items)



def render_downloads(excel_buffer: io.BytesIO, pdf_buffer: io.BytesIO) -> None:
    col1, col2 = st.columns(2)

    with col1:
        st.download_button(
            label="Descargar Excel",
            data=excel_buffer.getvalue(),
            file_name="lista_empaque_profesional.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with col2:
        st.download_button(
            label="Descargar PDF",
            data=pdf_buffer.getvalue(),
            file_name="lista_empaque_operativa.pdf",
            mime="application/pdf",
            use_container_width=True,
        )



def main() -> None:
    st.set_page_config(page_title="Lista de empaque profesional", layout="wide")

    st.title(APP_TITLE)
    st.caption(
        "Sube tu Excel y genera una lista de empaque mucho más clara para impresión en blanco y negro, con checkboxes y espacio para validación manual."
    )

    with st.container(border=True):
        st.markdown(
            """
            **Qué mejora este formato**

            - Letra más grande y jerarquía visual clara.
            - Diseño limpio para impresión en blanco y negro.
            - Bloques por venta para reducir errores operativos.
            - Casillas para surtido, revisión, entrega, iniciales y hora.
            """
        )

    uploaded_file = st.file_uploader("Sube el archivo Excel", type=["xlsx", "xls"])

    if uploaded_file is None:
        st.info("Esperando archivo Excel para procesar.")
        return

    try:
        with st.spinner("Procesando archivo y construyendo PDF profesional..."):
            df_output, excel_buffer, pdf_buffer, warnings, orders = generate_files(uploaded_file)

        st.success("Archivo procesado correctamente.")
        render_metrics(orders)

        if warnings:
            with st.expander("Ver advertencias de lectura"):
                for warning in warnings:
                    st.warning(warning)

        tab1, tab2 = st.tabs(["Vista previa", "Descargas"])

        with tab1:
            st.dataframe(df_output, use_container_width=True, height=460)

        with tab2:
            render_downloads(excel_buffer, pdf_buffer)

    except Exception as error:
        st.error(f"Error al procesar el archivo: {error}")


if __name__ == "__main__":
    main()
