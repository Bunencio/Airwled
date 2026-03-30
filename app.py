from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from datetime import datetime
from html import escape
from typing import BinaryIO, Sequence

import pandas as pd

try:
    import streamlit as st
except ImportError:  # permite probar la lógica sin Streamlit
    st = None

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (
    BaseDocTemplate,
    Flowable,
    Frame,
    LongTable,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)

# -----------------------------------------------------------------------------
# Configuración general
# -----------------------------------------------------------------------------

APP_TITLE = "Generador profesional de lista de empaque"
HEADER_ROW_INDEX = 4  # fila 5 en Excel
DEFAULT_COLUMN_INDEXES = {
    "venta": 0,
    "estado": 2,
    "unidades": 6,
    "sku": 16,
    "titulo": 20,
}

# Escala neutra para impresión a blanco y negro
BLACK = colors.HexColor("#111111")
DARK_GRAY = colors.HexColor("#2B2B2B")
MID_GRAY = colors.HexColor("#666666")
LIGHT_GRAY = colors.HexColor("#D7D7D7")
SOFT_GRAY = colors.HexColor("#ECECEC")
VERY_LIGHT_GRAY = colors.HexColor("#F7F7F7")
WHITE = colors.white

PAGE_SIZE = landscape(letter)
PAGE_MARGINS = {
    "left": 0.42 * inch,
    "right": 0.42 * inch,
    "top": 0.52 * inch,
    "bottom": 0.48 * inch,
}


# -----------------------------------------------------------------------------
# Modelos
# -----------------------------------------------------------------------------

@dataclass(slots=True)
class ProductLine:
    venta: str
    unidades: str
    sku: str
    titulo: str

    @property
    def contenido_linea(self) -> str:
        cantidad = self.unidades or "-"
        sku = self.sku or "SIN SKU"
        titulo = self.titulo or "SIN TÍTULO"
        return f"{cantidad} x {sku} - {titulo}"


@dataclass(slots=True)
class PackingTask:
    numero: int
    venta_principal: str
    tipo: str
    grupo: list[ProductLine] = field(default_factory=list)

    @property
    def es_paquete(self) -> bool:
        return self.tipo.startswith("PAQUETE")

    @property
    def cantidad_lineas(self) -> int:
        return len(self.grupo)

    @property
    def skus(self) -> str:
        return "\n".join(item.sku for item in self.grupo)

    @property
    def unidades(self) -> str:
        return "\n".join(item.unidades for item in self.grupo)

    @property
    def productos(self) -> str:
        return "\n".join(item.titulo for item in self.grupo)

    @property
    def contenido_formateado(self) -> str:
        return "\n".join(item.contenido_linea for item in self.grupo)


@dataclass(slots=True)
class ParseResult:
    tasks: list[PackingTask]
    warnings: list[str]


# -----------------------------------------------------------------------------
# Utilidades
# -----------------------------------------------------------------------------

def clean_value(value: object) -> str:
    if pd.isna(value):
        return ""

    if isinstance(value, float) and value.is_integer():
        value = int(value)

    text = str(value).strip()
    text = re.sub(r"\s+", " ", text)
    return text


def html_lines(text: str) -> str:
    return escape(text).replace("\n", "<br/>")


def validate_required_columns(df: pd.DataFrame, indexes: dict[str, int]) -> None:
    required_max_index = max(indexes.values())
    if len(df.columns) <= required_max_index:
        raise ValueError(
            "El archivo no tiene la estructura esperada. "
            f"Se necesitan al menos {required_max_index + 1} columnas y se detectaron {len(df.columns)}."
        )


def read_source_dataframe(excel_file: BinaryIO) -> pd.DataFrame:
    df = pd.read_excel(excel_file, header=HEADER_ROW_INDEX)
    df = df.dropna(how="all").reset_index(drop=True)
    validate_required_columns(df, DEFAULT_COLUMN_INDEXES)
    return df


def line_from_row(row: pd.Series, columns: dict[str, str]) -> ProductLine:
    return ProductLine(
        venta=clean_value(row[columns["venta"]]),
        unidades=clean_value(row[columns["unidades"]]),
        sku=clean_value(row[columns["sku"]]),
        titulo=clean_value(row[columns["titulo"]]),
    )


def parse_packing_tasks(df: pd.DataFrame) -> ParseResult:
    cols = {key: df.columns[idx] for key, idx in DEFAULT_COLUMN_INDEXES.items()}

    tasks: list[PackingTask] = []
    warnings: list[str] = []

    i = 0
    task_number = 1

    while i < len(df):
        row = df.iloc[i]
        estado = clean_value(row[cols["estado"]])
        venta_actual = clean_value(row[cols["venta"]])

        match = re.search(r"Paquete de (\d+)", estado, flags=re.IGNORECASE)

        if match:
            expected_count = int(match.group(1))
            group: list[ProductLine] = []

            for offset in range(1, expected_count + 1):
                next_index = i + offset
                if next_index >= len(df):
                    warnings.append(
                        f"La venta '{venta_actual or 'SIN ID'}' indica un paquete de {expected_count}, "
                        f"pero el archivo termina antes de completar el grupo."
                    )
                    break

                next_row = df.iloc[next_index]
                group.append(line_from_row(next_row, cols))

            tasks.append(
                PackingTask(
                    numero=task_number,
                    venta_principal=venta_actual or "SIN ID",
                    tipo=f"PAQUETE ({len(group)} producto{'s' if len(group) != 1 else ''})",
                    grupo=group,
                )
            )

            i += expected_count + 1
            task_number += 1
            continue

        tasks.append(
            PackingTask(
                numero=task_number,
                venta_principal=venta_actual or "SIN ID",
                tipo="INDIVIDUAL",
                grupo=[line_from_row(row, cols)],
            )
        )
        i += 1
        task_number += 1

    return ParseResult(tasks=tasks, warnings=warnings)


def tasks_to_dataframe(tasks: Sequence[PackingTask]) -> pd.DataFrame:
    rows: list[dict[str, str]] = []

    for task in tasks:
        rows.append(
            {
                "No.": task.numero,
                "Hecho": "",
                "Terminado": "",
                "Iniciales / Hora": "",
                "Venta principal": task.venta_principal,
                "Tipo": task.tipo,
                "SKU(s)": task.skus,
                "Unidades": task.unidades,
                "Productos": task.productos,
                "Contenido verificado": task.contenido_formateado,
            }
        )

    return pd.DataFrame(rows)


def build_excel_bytes(df_final: pd.DataFrame) -> io.BytesIO:
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_final.to_excel(writer, index=False, sheet_name="Lista empaque")
        workbook = writer.book
        worksheet = writer.sheets["Lista empaque"]

        header_fill = PatternFill(fill_type="solid", fgColor="202020")
        header_font = Font(color="FFFFFF", bold=True, size=12)
        thin = Side(style="thin", color="000000")
        medium = Side(style="medium", color="000000")

        column_widths = {
            "A": 8,
            "B": 12,
            "C": 14,
            "D": 18,
            "E": 20,
            "F": 18,
            "G": 24,
            "H": 12,
            "I": 48,
            "J": 60,
        }

        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = Border(left=medium, right=medium, top=medium, bottom=medium)

        for column_letter, width in column_widths.items():
            worksheet.column_dimensions[column_letter].width = width

        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions

        for row_idx, row in enumerate(worksheet.iter_rows(min_row=2), start=2):
            fill_color = "FFFFFF" if row_idx % 2 == 0 else "F4F4F4"
            for cell in row:
                cell.fill = PatternFill(fill_type="solid", fgColor=fill_color)
                cell.alignment = Alignment(vertical="center", wrap_text=True)
                cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
                cell.font = Font(name="Calibri", size=12)

            # SKU(s) en negritas -> columna G
            worksheet[f"G{row_idx}"].font = Font(name="Calibri", size=12, bold=True)

            # Columnas de control más centradas
            for control_col in ("A", "B", "C", "D", "H"):
                worksheet[f"{control_col}{row_idx}"].alignment = Alignment(
                    horizontal="center",
                    vertical="center",
                    wrap_text=True,
                )

        for row_idx in range(2, worksheet.max_row + 1):
            worksheet.row_dimensions[row_idx].height = 40

        workbook.properties.creator = "OpenAI - Generador profesional de lista de empaque"
        workbook.properties.title = "Lista de empaque"

    buffer.seek(0)
    return buffer


# -----------------------------------------------------------------------------
# Componentes PDF
# -----------------------------------------------------------------------------

class CheckBox(Flowable):
    def __init__(self, size: float = 14):
        super().__init__()
        self.size = size
        self.width = size
        self.height = size

    def wrap(self, available_width: float, available_height: float) -> tuple[float, float]:
        return self.width, self.height

    def draw(self) -> None:
        self.canv.setStrokeColor(BLACK)
        self.canv.setLineWidth(1.2)
        self.canv.roundRect(0, 0, self.size, self.size, 1.8, stroke=1, fill=0)


class SignatureLine(Flowable):
    def __init__(self, width: float = 54, label: str = "Iniciales / hora"):
        super().__init__()
        self.width = width
        self.height = 18
        self.label = label

    def wrap(self, available_width: float, available_height: float) -> tuple[float, float]:
        return self.width, self.height

    def draw(self) -> None:
        self.canv.setStrokeColor(BLACK)
        self.canv.setLineWidth(0.9)
        self.canv.line(0, 12, self.width, 12)
        self.canv.setFont("Helvetica", 6.6)
        self.canv.setFillColor(MID_GRAY)
        self.canv.drawCentredString(self.width / 2, 2, self.label)



def build_styles() -> dict[str, ParagraphStyle]:
    base_styles = getSampleStyleSheet()

    return {
        "title": ParagraphStyle(
            "TitlePacking",
            parent=base_styles["Title"],
            fontName="Helvetica-Bold",
            fontSize=20,
            leading=23,
            textColor=BLACK,
            alignment=TA_LEFT,
            spaceAfter=4,
        ),
        "subtitle": ParagraphStyle(
            "SubtitlePacking",
            parent=base_styles["Normal"],
            fontName="Helvetica",
            fontSize=10,
            leading=12,
            textColor=MID_GRAY,
            alignment=TA_LEFT,
        ),
        "metric_label": ParagraphStyle(
            "MetricLabel",
            parent=base_styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=8.4,
            leading=10,
            textColor=DARK_GRAY,
            alignment=TA_CENTER,
        ),
        "metric_value": ParagraphStyle(
            "MetricValue",
            parent=base_styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=16.5,
            leading=18,
            textColor=BLACK,
            alignment=TA_CENTER,
        ),
        "helper": ParagraphStyle(
            "Helper",
            parent=base_styles["Normal"],
            fontName="Helvetica",
            fontSize=9.4,
            leading=11.8,
            textColor=BLACK,
            alignment=TA_LEFT,
        ),
        "header": ParagraphStyle(
            "HeaderCell",
            parent=base_styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=10.4,
            leading=12.2,
            textColor=WHITE,
            alignment=TA_CENTER,
        ),
        "body": ParagraphStyle(
            "BodyCell",
            parent=base_styles["Normal"],
            fontName="Helvetica",
            fontSize=10.4,
            leading=13.0,
            textColor=BLACK,
            alignment=TA_LEFT,
        ),
        "body_center": ParagraphStyle(
            "BodyCenter",
            parent=base_styles["Normal"],
            fontName="Helvetica",
            fontSize=10.2,
            leading=12.6,
            textColor=BLACK,
            alignment=TA_CENTER,
        ),
        "body_bold": ParagraphStyle(
            "BodyBold",
            parent=base_styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=10.4,
            leading=13.0,
            textColor=BLACK,
            alignment=TA_LEFT,
        ),
    }



def metric_card(label: str, value: str, styles: dict[str, ParagraphStyle]) -> Table:
    table = Table(
        [
            [Paragraph(label, styles["metric_label"])],
            [Paragraph(value, styles["metric_value"])],
        ],
        colWidths=[1.58 * inch],
        rowHeights=[0.30 * inch, 0.45 * inch],
    )
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), LIGHT_GRAY),
                ("BACKGROUND", (0, 1), (-1, -1), WHITE),
                ("BOX", (0, 0), (-1, -1), 1.0, BLACK),
                ("LINEBELOW", (0, 0), (-1, 0), 0.8, BLACK),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    return table



def build_summary_table(tasks: Sequence[PackingTask], styles: dict[str, ParagraphStyle]) -> Table:
    total = len(tasks)
    paquetes = sum(task.es_paquete for task in tasks)
    individuales = total - paquetes

    summary = Table(
        [[
            metric_card("TOTAL DE ENTREGAS", str(total), styles),
            metric_card("INDIVIDUALES", str(individuales), styles),
            metric_card("PAQUETES", str(paquetes), styles),
        ]],
        colWidths=[1.66 * inch, 1.66 * inch, 1.66 * inch],
    )
    summary.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ]
        )
    )
    return summary



def build_help_box(styles: dict[str, ParagraphStyle], content_width: float) -> Table:
    text = (
        "<b>Uso sugerido:</b> 1) validar venta y contenido, 2) marcar <b>HECHO</b> al preparar, "
        "3) marcar <b>TERMINADO</b> al cerrar la tarea, 4) escribir iniciales u hora, "
        "5) revisar cantidades y SKU antes de entregar."
    )
    table = Table([[Paragraph(text, styles["helper"])]], colWidths=[content_width])
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), VERY_LIGHT_GRAY),
                ("BOX", (0, 0), (-1, -1), 0.9, BLACK),
                ("LEFTPADDING", (0, 0), (-1, -1), 9),
                ("RIGHTPADDING", (0, 0), (-1, -1), 9),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ]
        )
    )
    return table



def build_table_rows(tasks: Sequence[PackingTask], styles: dict[str, ParagraphStyle]) -> list[list[object]]:
    rows: list[list[object]] = [
        [
            Paragraph("No.", styles["header"]),
            Paragraph("Hecho", styles["header"]),
            Paragraph("Terminado", styles["header"]),
            Paragraph("Iniciales / Hora", styles["header"]),
            Paragraph("Venta principal", styles["header"]),
            Paragraph("Tipo", styles["header"]),
            Paragraph("Contenido verificado", styles["header"]),
        ]
    ]

    for task in tasks:
        body_lines: list[str] = []
        for line in task.grupo:
            line_html = (
                f"{escape(line.unidades or '-')} x <b>{escape(line.sku or 'SIN SKU')}</b> - "
                f"{escape(line.titulo or 'SIN TÍTULO')}"
            )
            body_lines.append(line_html)

        rows.append(
            [
                Paragraph(str(task.numero), styles["body_center"]),
                CheckBox(size=14),
                CheckBox(size=14),
                SignatureLine(width=60),
                Paragraph(html_lines(task.venta_principal), styles["body"]),
                Paragraph(html_lines(task.tipo), styles["body_bold"]),
                Paragraph("<br/>".join(body_lines), styles["body"]),
            ]
        )
    return rows



def draw_page_chrome(canvas, doc) -> None:
    page_width, page_height = PAGE_SIZE
    canvas.saveState()

    canvas.setStrokeColor(BLACK)
    canvas.setLineWidth(1.0)
    canvas.line(doc.leftMargin, page_height - 18, page_width - doc.rightMargin, page_height - 18)

    canvas.setFont("Helvetica-Bold", 9.5)
    canvas.setFillColor(BLACK)
    canvas.drawString(doc.leftMargin, page_height - 13, "LISTA DE EMPAQUE")

    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(MID_GRAY)
    canvas.drawRightString(
        page_width - doc.rightMargin,
        page_height - 13,
        f"Página {canvas.getPageNumber()}",
    )

    canvas.setLineWidth(0.6)
    canvas.setStrokeColor(LIGHT_GRAY)
    canvas.line(doc.leftMargin, 22, page_width - doc.rightMargin, 22)
    canvas.setFont("Helvetica", 7.2)
    canvas.setFillColor(MID_GRAY)
    canvas.drawString(doc.leftMargin, 10, "Formato optimizado para lectura rápida e impresión en blanco y negro")
    canvas.drawRightString(page_width - doc.rightMargin, 10, getattr(doc, "generated_at", ""))

    canvas.restoreState()



def build_pdf_bytes(tasks: Sequence[PackingTask]) -> io.BytesIO:
    buffer = io.BytesIO()
    styles = build_styles()

    doc = BaseDocTemplate(
        buffer,
        pagesize=PAGE_SIZE,
        leftMargin=PAGE_MARGINS["left"],
        rightMargin=PAGE_MARGINS["right"],
        topMargin=PAGE_MARGINS["top"],
        bottomMargin=PAGE_MARGINS["bottom"],
        title="Lista de empaque",
        author="OpenAI",
    )

    frame = Frame(
        doc.leftMargin,
        doc.bottomMargin,
        doc.width,
        doc.height,
        id="normal",
    )
    doc.addPageTemplates([PageTemplate(id="packing", frames=[frame], onPage=draw_page_chrome)])

    generated_at = datetime.now().strftime("%d/%m/%Y %H:%M")
    doc.generated_at = generated_at
    story: list[object] = [
        Paragraph("Lista de empaque", styles["title"]),
        Paragraph(
            f"Documento listo para impresión. Generado el {generated_at}. "
            "Diseño optimizado para lectura rápida, validación manual y control visual por el equipo.",
            styles["subtitle"],
        ),
        Spacer(1, 0.16 * inch),
        build_summary_table(tasks, styles),
        Spacer(1, 0.14 * inch),
        build_help_box(styles, doc.width),
        Spacer(1, 0.18 * inch),
    ]

    table_rows = build_table_rows(tasks, styles)
    col_widths = [
        0.42 * inch,  # No.
        0.68 * inch,  # Hecho
        0.78 * inch,  # Terminado
        1.05 * inch,  # Iniciales / Hora
        1.42 * inch,  # Venta principal
        1.38 * inch,  # Tipo
        4.95 * inch,  # Contenido verificado
    ]

    main_table = LongTable(table_rows, colWidths=col_widths, repeatRows=1)

    table_style_commands: list[tuple] = [
        ("BACKGROUND", (0, 0), (-1, 0), DARK_GRAY),
        ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
        ("BOX", (0, 0), (-1, -1), 0.9, BLACK),
        ("INNERGRID", (0, 0), (-1, -1), 0.45, BLACK),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (3, -1), "CENTER"),
        ("ALIGN", (4, 1), (5, -1), "LEFT"),
        ("ALIGN", (6, 1), (6, -1), "LEFT"),
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 7),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]

    for row_idx, task in enumerate(tasks, start=1):
        base_background = SOFT_GRAY if row_idx % 2 == 0 else WHITE
        table_style_commands.append(("BACKGROUND", (0, row_idx), (-1, row_idx), base_background))

        if task.es_paquete:
            table_style_commands.extend(
                [
                    ("BACKGROUND", (4, row_idx), (5, row_idx), LIGHT_GRAY),
                    ("LINEBEFORE", (4, row_idx), (4, row_idx), 1.0, BLACK),
                    ("LINEABOVE", (0, row_idx), (-1, row_idx), 0.85, BLACK),
                ]
            )

    main_table.setStyle(TableStyle(table_style_commands))
    story.append(main_table)

    observations = Table(
        [[
            Paragraph("<b>Observaciones generales:</b>", styles["body_bold"]),
            Paragraph("________________________________________________________________________________", styles["body"]),
        ]],
        colWidths=[1.95 * inch, 7.95 * inch],
    )
    observations.setStyle(
        TableStyle(
            [
                ("BOX", (0, 0), (-1, -1), 0.8, BLACK),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 9),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 9),
                ("BACKGROUND", (0, 0), (-1, -1), VERY_LIGHT_GRAY),
            ]
        )
    )
    story.extend([Spacer(1, 0.18 * inch), observations])

    doc.build(story)
    buffer.seek(0)
    return buffer


# -----------------------------------------------------------------------------
# Pipeline principal
# -----------------------------------------------------------------------------

def generate_outputs(excel_file: BinaryIO) -> tuple[pd.DataFrame, io.BytesIO, io.BytesIO, list[str]]:
    df_source = read_source_dataframe(excel_file)
    parse_result = parse_packing_tasks(df_source)
    df_final = tasks_to_dataframe(parse_result.tasks)
    excel_buffer = build_excel_bytes(df_final)
    pdf_buffer = build_pdf_bytes(parse_result.tasks)
    return df_final, excel_buffer, pdf_buffer, parse_result.warnings


# -----------------------------------------------------------------------------
# Streamlit UI
# -----------------------------------------------------------------------------

def main() -> None:
    if st is None:
        raise RuntimeError("Streamlit no está instalado en este entorno.")

    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)
    st.caption(
        "Versión mejorada con tipografía más legible, doble control visual (Hecho y Terminado), "
        "SKU resaltado y estructura profesional para reducir errores en operación."
    )

    uploaded_file = st.file_uploader("Sube el archivo Excel", type=["xlsx", "xls"])

    if uploaded_file is None:
        st.info("Carga un archivo para generar la lista de empaque mejorada.")
        return

    try:
        df_resultado, excel_out, pdf_out, warnings = generate_outputs(uploaded_file)

        st.success("Archivo procesado correctamente.")

        if warnings:
            for warning in warnings:
                st.warning(warning)

        st.dataframe(df_resultado, use_container_width=True, height=520)

        left, right = st.columns(2)

        with left:
            st.download_button(
                label="Descargar Excel profesional",
                data=excel_out,
                file_name="lista_empaque_profesional.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with right:
            st.download_button(
                label="Descargar PDF profesional",
                data=pdf_out,
                file_name="lista_empaque_profesional.pdf",
                mime="application/pdf",
                use_container_width=True,
            )

    except Exception as exc:
        st.error(f"Error al procesar el archivo: {exc}")


if __name__ == "__main__":
    main()
