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
except ImportError:
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
# Configuracion general
# -----------------------------------------------------------------------------

APP_TITLE = "Generador profesional de lista de empaque"
HEADER_ROW_INDEX = 4
DEFAULT_COLUMN_INDEXES = {
    "venta": 0,
    "estado": 2,
    "unidades": 6,
    "sku": 16,
    "titulo": 20,
}

BLACK        = colors.HexColor("#111111")
DARK_GRAY    = colors.HexColor("#2D2D2D")
MID_GRAY     = colors.HexColor("#6B6B6B")
LIGHT_GRAY   = colors.HexColor("#D9D9D9")
VERY_LIGHT_GRAY = colors.HexColor("#F3F3F3")
ACCENT_DARK  = colors.HexColor("#1A1A2E")   # azul muy oscuro para cabeceras
ACCENT_MID   = colors.HexColor("#16213E")
STAMP_BG     = colors.HexColor("#F7F7F7")
WHITE        = colors.white

PAGE_SIZE    = landscape(letter)
PAGE_MARGINS = {
    "left":   0.45 * inch,
    "right":  0.45 * inch,
    "top":    0.55 * inch,
    "bottom": 0.50 * inch,
}


# -----------------------------------------------------------------------------
# Modelos
# -----------------------------------------------------------------------------

@dataclass(slots=True)
class ProductLine:
    venta:    str
    unidades: str
    sku:      str
    titulo:   str

    @property
    def contenido_linea(self) -> str:
        cantidad = self.unidades or "-"
        sku      = self.sku      or "SIN SKU"
        titulo   = self.titulo   or "SIN TITULO"
        return f"{cantidad} x {sku} - {titulo}"


@dataclass(slots=True)
class PackingTask:
    numero:         int
    venta_principal: str
    tipo:           str
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
    tasks:    list[PackingTask]
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
            f"Se necesitan al menos {required_max_index + 1} columnas "
            f"y se detectaron {len(df.columns)}."
        )


def read_source_dataframe(excel_file: BinaryIO) -> pd.DataFrame:
    df = pd.read_excel(excel_file, header=HEADER_ROW_INDEX)

    # ❌ eliminar la fila fake tipo "# de venta"
    df = df.iloc[1:].reset_index(drop=True)

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

    tasks:    list[PackingTask] = []
    warnings: list[str]        = []
    i            = 0
    task_number  = 1

    while i < len(df):
        row         = df.iloc[i]
        estado      = clean_value(row[cols["estado"]])
        venta_actual = clean_value(row[cols["venta"]])

        match = re.search(r"Paquete de (\d+)", estado, flags=re.IGNORECASE)

        if match:
            expected_count = int(match.group(1))
            group: list[ProductLine] = []

            for offset in range(1, expected_count + 1):
                next_index = i + offset
                if next_index >= len(df):
                    warnings.append(
                        f"La venta '{venta_actual or 'SIN ID'}' indica un paquete de "
                        f"{expected_count}, pero el archivo termina antes de completar el grupo."
                    )
                    break
                group.append(line_from_row(df.iloc[next_index], cols))

            tasks.append(PackingTask(
                numero=task_number,
                venta_principal=venta_actual or "SIN ID",
                tipo=f"PAQUETE ({len(group)} producto{'s' if len(group) != 1 else ''})",
                grupo=group,
            ))
            i += expected_count + 1
            task_number += 1
            continue

        tasks.append(PackingTask(
            numero=task_number,
            venta_principal=venta_actual or "SIN ID",
            tipo="INDIVIDUAL",
            grupo=[line_from_row(row, cols)],
        ))
        i += 1
        task_number += 1

    return ParseResult(tasks=tasks, warnings=warnings)


def tasks_to_dataframe(tasks: Sequence[PackingTask]) -> pd.DataFrame:
    rows: list[dict[str, str]] = []
    for task in tasks:
        rows.append({
            "No.":              task.numero,
            "Hecho":            "",
            "Revisado":         "",
            "Iniciales / Hora": "",
            "Venta principal":  task.venta_principal,
            "Tipo":             task.tipo,
            "SKU(s)":           task.skus,
            "Unidades":         task.unidades,
            "Productos":        task.productos,
            "Contenido verificado": task.contenido_formateado,
        })
    return pd.DataFrame(rows)


def build_excel_bytes(df_final: pd.DataFrame) -> io.BytesIO:
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_final.to_excel(writer, index=False, sheet_name="Lista empaque")
        workbook  = writer.book
        worksheet = writer.sheets["Lista empaque"]

        header_fill = PatternFill(fill_type="solid", fgColor="1F1F1F")
        header_font = Font(color="FFFFFF", bold=True)
        thin   = Side(style="thin",   color="000000")
        medium = Side(style="medium", color="000000")

        column_widths = {
            "A":  8, "B": 10, "C": 10, "D": 18,
            "E": 20, "F": 18, "G": 24, "H": 12,
            "I": 48, "J": 60,
        }

        for cell in worksheet[1]:
            cell.fill      = header_fill
            cell.font      = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border    = Border(left=medium, right=medium, top=medium, bottom=medium)

        for col_letter, width in column_widths.items():
            worksheet.column_dimensions[col_letter].width = width

        worksheet.freeze_panes      = "A2"
        worksheet.auto_filter.ref   = worksheet.dimensions

        for row_idx, row in enumerate(worksheet.iter_rows(min_row=2), start=2):
            fill_color = "FFFFFF" if row_idx % 2 == 0 else "F2F2F2"
            for cell in row:
                cell.fill      = PatternFill(fill_type="solid", fgColor=fill_color)
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                cell.border    = Border(left=thin, right=thin, top=thin, bottom=thin)

        for row_idx in range(2, worksheet.max_row + 1):
            worksheet.row_dimensions[row_idx].height = 34

        workbook.properties.creator = "Generador profesional de lista de empaque"
        workbook.properties.title   = "Lista de empaque"

    buffer.seek(0)
    return buffer


# -----------------------------------------------------------------------------
# Componentes PDF custom
# -----------------------------------------------------------------------------

class DualStampCell(Flowable):
    """
    Dibuja dos recuadros horizontales HECHO | REVISADO dentro de una celda,
    cada uno con su propio checkbox y linea de firma.
    """

    LABEL_FONT_SIZE  = 6.2
    SIGN_FONT_SIZE   = 5.4
    BOX_SIZE         = 10.5
    OUTER_PAD        = 3.0
    INNER_GAP        = 5.0

    def __init__(self, available_width: float = 1.60 * inch):
        super().__init__()
        self._avail_w = available_width
        # Cada bloque ocupa la mitad del ancho disponible menos el gap
        self._block_w = (available_width - self.INNER_GAP - 2 * self.OUTER_PAD) / 2
        self.width    = available_width
        self.height   = 38

    def wrap(self, aw: float, ah: float):
        return self.width, self.height

    def _draw_stamp_block(self, x: float, y_base: float, label: str) -> None:
        """Dibuja un bloque individual (etiqueta + checkbox + linea firma)."""
        c = self.canv
        bw = self._block_w

        # Borde exterior del bloque
        c.setStrokeColor(BLACK)
        c.setLineWidth(0.9)
        c.roundRect(x, y_base, bw, self.height - 2 * self.OUTER_PAD, 2, stroke=1, fill=0)

        # Etiqueta centrada en la parte superior
        label_y = y_base + self.height - 2 * self.OUTER_PAD - 9
        c.setFont("Helvetica-Bold", self.LABEL_FONT_SIZE)
        c.setFillColor(ACCENT_DARK)
        c.drawCentredString(x + bw / 2, label_y, label)

        # Separador bajo la etiqueta
        sep_y = label_y - 3
        c.setStrokeColor(LIGHT_GRAY)
        c.setLineWidth(0.5)
        c.line(x + 4, sep_y, x + bw - 4, sep_y)

        # Checkbox
        cb_size = self.BOX_SIZE
        cb_x    = x + bw / 2 - cb_size / 2
        cb_y    = y_base + self.height - 2 * self.OUTER_PAD - 9 - 4 - cb_size - 2
        c.setStrokeColor(BLACK)
        c.setLineWidth(1.0)
        c.setFillColor(WHITE)
        c.rect(cb_x, cb_y, cb_size, cb_size, stroke=1, fill=1)

        # Linea de firma debajo del checkbox
        line_y  = cb_y - 7
        line_x0 = x + 5
        line_x1 = x + bw - 5
        c.setStrokeColor(BLACK)
        c.setLineWidth(0.7)
        c.line(line_x0, line_y, line_x1, line_y)

        # Mini-etiqueta "Firma / Hora"
        c.setFont("Helvetica", self.SIGN_FONT_SIZE)
        c.setFillColor(MID_GRAY)
        c.drawCentredString(x + bw / 2, line_y - 6, "Firma / Hora")

    def draw(self) -> None:
        c = self.canv
        c.saveState()

        x0_hecho    = self.OUTER_PAD
        x0_revisado = self.OUTER_PAD + self._block_w + self.INNER_GAP
        y_base      = self.OUTER_PAD

        self._draw_stamp_block(x0_hecho,    y_base, "HECHO")
        self._draw_stamp_block(x0_revisado, y_base, "REVISADO")

        c.restoreState()


class NumberBadge(Flowable):
    """Circulo con el numero de tarea."""

    def __init__(self, number: int, size: float = 22):
        super().__init__()
        self.number = number
        self.size   = size
        self.width  = size
        self.height = size

    def wrap(self, aw, ah):
        return self.width, self.height

    def draw(self):
        c = self.canv
        c.saveState()
        radius = self.size / 2
        c.setFillColor(DARK_GRAY)
        c.setStrokeColor(BLACK)
        c.setLineWidth(0.6)
        c.circle(radius, radius, radius, stroke=1, fill=1)
        c.setFont("Helvetica-Bold", 8.5)
        c.setFillColor(WHITE)
        c.drawCentredString(radius, radius - 3, str(self.number))
        c.restoreState()


# -----------------------------------------------------------------------------
# Estilos
# -----------------------------------------------------------------------------

def build_styles() -> dict[str, ParagraphStyle]:
    base = getSampleStyleSheet()

    return {
        "title": ParagraphStyle(
            "TitlePacking", parent=base["Title"],
            fontName="Helvetica-Bold", fontSize=18, leading=21,
            textColor=BLACK, alignment=TA_LEFT, spaceAfter=4,
        ),
        "subtitle": ParagraphStyle(
            "SubtitlePacking", parent=base["Normal"],
            fontName="Helvetica", fontSize=8.6, leading=10.5,
            textColor=MID_GRAY, alignment=TA_LEFT,
        ),
        "metric_label": ParagraphStyle(
            "MetricLabel", parent=base["Normal"],
            fontName="Helvetica-Bold", fontSize=7.5, leading=9,
            textColor=DARK_GRAY, alignment=TA_CENTER,
        ),
        "metric_value": ParagraphStyle(
            "MetricValue", parent=base["Normal"],
            fontName="Helvetica-Bold", fontSize=15, leading=16,
            textColor=BLACK, alignment=TA_CENTER,
        ),
        "helper": ParagraphStyle(
            "Helper", parent=base["Normal"],
            fontName="Helvetica", fontSize=8.2, leading=10,
            textColor=BLACK, alignment=TA_LEFT,
        ),
        "header": ParagraphStyle(
            "HeaderCell", parent=base["Normal"],
            fontName="Helvetica-Bold", fontSize=8.8, leading=10.2,
            textColor=WHITE, alignment=TA_CENTER,
        ),
        "body": ParagraphStyle(
            "BodyCell", parent=base["Normal"],
            fontName="Helvetica", fontSize=8.8, leading=11.0,
            textColor=BLACK, alignment=TA_LEFT,
        ),
        "body_center": ParagraphStyle(
            "BodyCenter", parent=base["Normal"],
            fontName="Helvetica", fontSize=8.8, leading=11.0,
            textColor=BLACK, alignment=TA_CENTER,
        ),
        "body_bold": ParagraphStyle(
            "BodyBold", parent=base["Normal"],
            fontName="Helvetica-Bold", fontSize=9.0, leading=11.2,
            textColor=BLACK, alignment=TA_LEFT,
        ),
        "tipo_paquete": ParagraphStyle(
            "TipoPaquete", parent=base["Normal"],
            fontName="Helvetica-Bold", fontSize=8.8, leading=11.0,
            textColor=ACCENT_DARK, alignment=TA_LEFT,
        ),
    }


# -----------------------------------------------------------------------------
# Bloques de resumen y ayuda
# -----------------------------------------------------------------------------

def metric_card(label: str, value: str, styles: dict[str, ParagraphStyle]) -> Table:
    table = Table(
        [[Paragraph(label, styles["metric_label"])],
         [Paragraph(value, styles["metric_value"])]],
        colWidths=[1.48 * inch],
        rowHeights=[0.28 * inch, 0.42 * inch],
    )
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), LIGHT_GRAY),
        ("BACKGROUND", (0, 1), (-1, -1), WHITE),
        ("BOX", (0, 0), (-1, -1), 1.0, BLACK),
        ("LINEBELOW", (0, 0), (-1, 0), 0.8, BLACK),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING",   (0, 0), (-1, -1), 6),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 6),
        ("TOPPADDING",    (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    return table


def build_summary_table(tasks: Sequence[PackingTask], styles: dict[str, ParagraphStyle]) -> Table:
    total        = len(tasks)
    paquetes     = sum(task.es_paquete for task in tasks)
    individuales = total - paquetes

    summary = Table(
        [[
            metric_card("TOTAL DE ENTREGAS", str(total), styles),
            metric_card("INDIVIDUALES",       str(individuales), styles),
            metric_card("PAQUETES",           str(paquetes), styles),
        ]],
        colWidths=[1.58 * inch, 1.58 * inch, 1.58 * inch],
    )
    summary.setStyle(TableStyle([
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING",   (0, 0), (-1, -1), 0),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 10),
        ("TOPPADDING",    (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    return summary


def build_help_box(styles: dict[str, ParagraphStyle], content_width: float) -> Table:
    text = (
        "<b>Uso sugerido:</b> 1) validar venta y contenido del pedido, "
        "2) marcar <b>HECHO</b> al empacar y firmar, "
        "3) marcar <b>REVISADO</b> al verificar el paquete antes de entregar, "
        "4) revisar cantidades y SKUs contra el sistema antes de despachar."
    )
    table = Table([[Paragraph(text, styles["helper"])]], colWidths=[content_width])
    table.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), VERY_LIGHT_GRAY),
        ("BOX",           (0, 0), (-1, -1), 0.9, BLACK),
        ("LEFTPADDING",   (0, 0), (-1, -1), 8),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 8),
        ("TOPPADDING",    (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    return table


# -----------------------------------------------------------------------------
# Filas de la tabla principal
# -----------------------------------------------------------------------------

def build_table_rows(
    tasks: Sequence[PackingTask],
    styles: dict[str, ParagraphStyle],
    stamp_col_width: float,
) -> list[list[object]]:

    # Cabecera
    rows: list[list[object]] = [[
        Paragraph("No.",              styles["header"]),
        Paragraph("Hecho / Revisado", styles["header"]),
        Paragraph("Venta principal",  styles["header"]),
        Paragraph("Tipo",             styles["header"]),
        Paragraph("Contenido verificado", styles["header"]),
    ]]

    for task in tasks:
        body_lines: list[str] = []
        for line in task.grupo:
            line_html = (
                f"<b>{escape(line.unidades or '-')} x</b> "
                f"{escape(line.sku or 'SIN SKU')} &ndash; "
                f"{escape(line.titulo or 'SIN TITULO')}"
            )
            body_lines.append(line_html)

        tipo_style = styles["tipo_paquete"] if task.es_paquete else styles["body_bold"]

        rows.append([
            NumberBadge(task.numero),
            DualStampCell(available_width=stamp_col_width),
            Paragraph(html_lines(task.venta_principal), styles["body"]),
            Paragraph(html_lines(task.tipo),            tipo_style),
            Paragraph("<br/>".join(body_lines),          styles["body"]),
        ])

    return rows


# -----------------------------------------------------------------------------
# Chrome de pagina
# -----------------------------------------------------------------------------

def draw_page_chrome(canvas, doc) -> None:
    page_width, page_height = PAGE_SIZE
    canvas.saveState()

    # Barra superior
    canvas.setFillColor(ACCENT_DARK)
    canvas.rect(
        doc.leftMargin, page_height - 20,
        page_width - doc.leftMargin - doc.rightMargin, 18,
        stroke=0, fill=1,
    )
    canvas.setFont("Helvetica-Bold", 9)
    canvas.setFillColor(WHITE)
    canvas.drawString(doc.leftMargin + 6, page_height - 14, "LISTA DE EMPAQUE")
    canvas.setFont("Helvetica", 8)
    canvas.drawRightString(
        page_width - doc.rightMargin - 6,
        page_height - 14,
        f"Pagina {canvas.getPageNumber()}",
    )

    # Pie de pagina
    canvas.setLineWidth(0.6)
    canvas.setStrokeColor(LIGHT_GRAY)
    canvas.line(doc.leftMargin, 22, page_width - doc.rightMargin, 22)
    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(MID_GRAY)
    canvas.drawString(doc.leftMargin, 10, "Formato optimizado para impresion en blanco y negro")
    canvas.drawRightString(page_width - doc.rightMargin, 10, getattr(doc, "generated_at", ""))

    canvas.restoreState()


# -----------------------------------------------------------------------------
# Construccion del PDF
# -----------------------------------------------------------------------------

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
        author="Generador profesional de lista de empaque",
    )

    frame = Frame(
        doc.leftMargin, doc.bottomMargin,
        doc.width, doc.height,
        id="normal",
    )
    doc.addPageTemplates([
        PageTemplate(id="packing", frames=[frame], onPage=draw_page_chrome)
    ])

    generated_at    = datetime.now().strftime("%d/%m/%Y %H:%M")
    doc.generated_at = generated_at

    story: list[object] = [
        Paragraph("Lista de empaque", styles["title"]),
        Paragraph(
            f"Documento listo para impresion. Generado el {generated_at}. "
            "Diseno optimizado para lectura rapida, validacion manual y doble control visual.",
            styles["subtitle"],
        ),
        Spacer(1, 0.16 * inch),
        build_summary_table(tasks, styles),
        Spacer(1, 0.14 * inch),
        build_help_box(styles, doc.width),
        Spacer(1, 0.18 * inch),
    ]

    # Anchos de columna
    # No. | Hecho/Revisado | Venta | Tipo | Contenido
    stamp_col_width = 1.90 * inch
    col_widths = [
        0.38 * inch,   # No.
        stamp_col_width,  # Hecho / Revisado (doble sello)
        1.40 * inch,   # Venta principal
        1.32 * inch,   # Tipo
        5.02 * inch,   # Contenido verificado
    ]

    table_rows = build_table_rows(tasks, styles, stamp_col_width)

    main_table = LongTable(table_rows, colWidths=col_widths, repeatRows=1)

    table_style_commands: list[tuple] = [
        # Cabecera
        ("BACKGROUND", (0, 0), (-1, 0), ACCENT_DARK),
        ("TEXTCOLOR",  (0, 0), (-1, 0), WHITE),
        # Bordes generales
        ("BOX",        (0, 0), (-1, -1), 0.9, BLACK),
        ("INNERGRID",  (0, 0), (-1, -1), 0.40, colors.HexColor("#CCCCCC")),
        # Alineacion
        ("VALIGN",  (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN",   (0, 0), (0, -1),  "CENTER"),
        ("ALIGN",   (1, 0), (1, -1),  "CENTER"),
        ("ALIGN",   (2, 1), (4, -1),  "LEFT"),
        # Padding
        ("LEFTPADDING",   (0, 0), (-1, -1), 5),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 5),
        ("TOPPADDING",    (0, 0), (-1, 0),  7),
        ("BOTTOMPADDING", (0, 0), (-1, 0),  7),
        ("TOPPADDING",    (0, 1), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 1), (-1, -1), 5),
        # Cabecera altura
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [WHITE, VERY_LIGHT_GRAY]),
    ]

    # Resalte de paquetes y banding
    for row_idx, task in enumerate(tasks, start=1):
        if task.es_paquete:
            table_style_commands.extend([
                ("BACKGROUND",  (2, row_idx), (3, row_idx), LIGHT_GRAY),
                ("LINEABOVE",   (0, row_idx), (-1, row_idx), 0.9, DARK_GRAY),
                ("LINEBELOW",   (0, row_idx), (-1, row_idx), 0.9, DARK_GRAY),
                ("LINEBEFORE",  (2, row_idx), (2, row_idx), 1.0, BLACK),
            ])

    main_table.setStyle(TableStyle(table_style_commands))
    story.append(main_table)

    # Pie de observaciones
    observations = Table(
        [[
            Paragraph("<b>Observaciones generales:</b>", styles["body_bold"]),
            Paragraph(
                "_" * 96,
                styles["body"],
            ),
        ]],
        colWidths=[1.8 * inch, 8.1 * inch],
    )
    observations.setStyle(TableStyle([
        ("BOX",           (0, 0), (-1, -1), 0.8, BLACK),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING",   (0, 0), (-1, -1), 8),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 8),
        ("TOPPADDING",    (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("BACKGROUND",    (0, 0), (-1, -1), VERY_LIGHT_GRAY),
    ]))
    story.extend([Spacer(1, 0.18 * inch), observations])

    doc.build(story)
    buffer.seek(0)
    return buffer


# -----------------------------------------------------------------------------
# Pipeline principal
# -----------------------------------------------------------------------------

def generate_outputs(
    excel_file: BinaryIO,
) -> tuple[pd.DataFrame, io.BytesIO, io.BytesIO, list[str]]:
    df_source    = read_source_dataframe(excel_file)
    parse_result = parse_packing_tasks(df_source)
    df_final     = tasks_to_dataframe(parse_result.tasks)
    excel_buffer = build_excel_bytes(df_final)
    pdf_buffer   = build_pdf_bytes(parse_result.tasks)
    return df_final, excel_buffer, pdf_buffer, parse_result.warnings


# -----------------------------------------------------------------------------
# Streamlit UI
# -----------------------------------------------------------------------------

def main() -> None:
    if st is None:
        raise RuntimeError("Streamlit no esta instalado en este entorno.")

    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)
    st.caption(
        "Version enfocada en impresion a blanco y negro con doble control visual: "
        "recuadros HECHO y REVISADO en cada fila para validacion de dos pasos."
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
