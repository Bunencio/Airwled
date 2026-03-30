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

# ─────────────────────────────────────────────────────────────────────────────
# PALETA  —  Alto contraste, pensada para impresion B&W en papel de oficina
# ─────────────────────────────────────────────────────────────────────────────
INK         = colors.HexColor("#0A0A0A")    # negro casi puro
CHARCOAL    = colors.HexColor("#1E1E1E")    # cabeceras oscuras
GRAPHITE    = colors.HexColor("#3D3D3D")    # elementos secundarios
SLATE       = colors.HexColor("#6A6A6A")    # texto auxiliar
SILVER      = colors.HexColor("#AEAEB2")    # bordes suaves
SMOKE       = colors.HexColor("#D8D8D8")    # fondo alterno filas
GHOST       = colors.HexColor("#F0F0F0")    # fondo muy claro
SNOW        = colors.HexColor("#F8F8F8")    # fondo casi blanco
WHITE       = colors.HexColor("#FFFFFF")
PKG_ACCENT  = colors.HexColor("#C8C8C8")    # fondo columnas en paquetes

# ─────────────────────────────────────────────────────────────────────────────
# LAYOUT
# ─────────────────────────────────────────────────────────────────────────────
PAGE_SIZE = landscape(letter)   # 11 x 8.5 in
L_MARGIN  = 0.48 * inch
R_MARGIN  = 0.48 * inch
T_MARGIN  = 0.75 * inch
B_MARGIN  = 0.55 * inch

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURACION DEL EXCEL DE ORIGEN
# ─────────────────────────────────────────────────────────────────────────────
HEADER_ROW_INDEX = 4
DEFAULT_COLUMN_INDEXES = {
    "venta":    0,
    "estado":   2,
    "unidades": 6,
    "sku":      16,
    "titulo":   20,
}
APP_TITLE = "Generador profesional de lista de empaque"


# ─────────────────────────────────────────────────────────────────────────────
# MODELOS
# ─────────────────────────────────────────────────────────────────────────────

@dataclass(slots=True)
class ProductLine:
    venta:    str
    unidades: str
    sku:      str
    titulo:   str

    @property
    def contenido_linea(self) -> str:
        return (
            f"{self.unidades or '-'} x "
            f"{self.sku or 'SIN SKU'} - "
            f"{self.titulo or 'SIN TITULO'}"
        )


@dataclass(slots=True)
class PackingTask:
    numero:          int
    venta_principal: str
    tipo:            str
    grupo: list[ProductLine] = field(default_factory=list)

    @property
    def es_paquete(self) -> bool:
        return self.tipo.startswith("PAQUETE")

    @property
    def skus(self) -> str:
        return "\n".join(i.sku for i in self.grupo)

    @property
    def unidades(self) -> str:
        return "\n".join(i.unidades for i in self.grupo)

    @property
    def productos(self) -> str:
        return "\n".join(i.titulo for i in self.grupo)

    @property
    def contenido_formateado(self) -> str:
        return "\n".join(i.contenido_linea for i in self.grupo)


@dataclass(slots=True)
class ParseResult:
    tasks:    list[PackingTask]
    warnings: list[str]


# ─────────────────────────────────────────────────────────────────────────────
# UTILIDADES DE PARSEO
# ─────────────────────────────────────────────────────────────────────────────

def clean_value(v: object) -> str:
    if pd.isna(v):
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return re.sub(r"\s+", " ", str(v).strip())


def html_lines(text: str) -> str:
    return escape(text).replace("\n", "<br/>")


def validate_required_columns(df: pd.DataFrame, indexes: dict[str, int]) -> None:
    req = max(indexes.values())
    if len(df.columns) <= req:
        raise ValueError(
            f"El archivo necesita al menos {req + 1} columnas; "
            f"se detectaron {len(df.columns)}."
        )


def read_source_dataframe(f: BinaryIO) -> pd.DataFrame:
    df = pd.read_excel(f, header=HEADER_ROW_INDEX)
    df = df.dropna(how="all").reset_index(drop=True)
    validate_required_columns(df, DEFAULT_COLUMN_INDEXES)
    return df


def line_from_row(row: pd.Series, cols: dict[str, str]) -> ProductLine:
    return ProductLine(
        venta=clean_value(row[cols["venta"]]),
        unidades=clean_value(row[cols["unidades"]]),
        sku=clean_value(row[cols["sku"]]),
        titulo=clean_value(row[cols["titulo"]]),
    )


def parse_packing_tasks(df: pd.DataFrame) -> ParseResult:
    cols     = {k: df.columns[v] for k, v in DEFAULT_COLUMN_INDEXES.items()}
    tasks:    list[PackingTask] = []
    warnings: list[str]        = []
    i = task_n = 0

    while i < len(df):
        row    = df.iloc[i]
        estado = clean_value(row[cols["estado"]])
        venta  = clean_value(row[cols["venta"]])
        match  = re.search(r"Paquete de (\d+)", estado, flags=re.IGNORECASE)

        if match:
            n     = int(match.group(1))
            group = []
            for off in range(1, n + 1):
                ni = i + off
                if ni >= len(df):
                    warnings.append(
                        f"Venta '{venta or 'SIN ID'}': paquete de {n} "
                        f"incompleto al final del archivo."
                    )
                    break
                group.append(line_from_row(df.iloc[ni], cols))
            tasks.append(PackingTask(
                numero=task_n + 1,
                venta_principal=venta or "SIN ID",
                tipo=f"PAQUETE ({len(group)} prod.)",
                grupo=group,
            ))
            i += n + 1
        else:
            tasks.append(PackingTask(
                numero=task_n + 1,
                venta_principal=venta or "SIN ID",
                tipo="INDIVIDUAL",
                grupo=[line_from_row(row, cols)],
            ))
            i += 1
        task_n += 1

    return ParseResult(tasks=tasks, warnings=warnings)


def tasks_to_dataframe(tasks: Sequence[PackingTask]) -> pd.DataFrame:
    return pd.DataFrame([{
        "No.":             t.numero,
        "Hecho":           "",
        "Revisado":        "",
        "Venta principal": t.venta_principal,
        "Tipo":            t.tipo,
        "SKU(s)":          t.skus,
        "Unidades":        t.unidades,
        "Productos":       t.productos,
        "Contenido":       t.contenido_formateado,
    } for t in tasks])


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL
# ─────────────────────────────────────────────────────────────────────────────

def build_excel_bytes(df: pd.DataFrame) -> io.BytesIO:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Lista empaque")
        wb  = writer.book
        ws  = writer.sheets["Lista empaque"]
        hf  = PatternFill(fill_type="solid", fgColor="1C1C1C")
        hfn = Font(color="FFFFFF", bold=True)
        thin   = Side(style="thin",   color="000000")
        medium = Side(style="medium", color="000000")
        widths = {"A": 8, "B": 10, "C": 10, "D": 22, "E": 18,
                  "F": 26, "G": 12, "H": 50, "I": 62}
        for cell in ws[1]:
            cell.fill      = hf
            cell.font      = hfn
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border    = Border(left=medium, right=medium, top=medium, bottom=medium)
        for col, w in widths.items():
            ws.column_dimensions[col].width = w
        ws.freeze_panes    = "A2"
        ws.auto_filter.ref = ws.dimensions
        for ri, row in enumerate(ws.iter_rows(min_row=2), 2):
            bg = "FFFFFF" if ri % 2 == 0 else "F4F4F4"
            for cell in row:
                cell.fill      = PatternFill(fill_type="solid", fgColor=bg)
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                cell.border    = Border(left=thin, right=thin, top=thin, bottom=thin)
        for ri in range(2, ws.max_row + 1):
            ws.row_dimensions[ri].height = 36
        wb.properties.title   = "Lista de empaque"
        wb.properties.creator = APP_TITLE
    buf.seek(0)
    return buf


# ─────────────────────────────────────────────────────────────────────────────
# FLOWABLES CUSTOM
# ─────────────────────────────────────────────────────────────────────────────

class RowNumber(Flowable):
    """
    Numero de fila en cuadro negro solido con cifra blanca.
    Grande y visible — el empleado lo localiza de un vistazo.
    """
    SIZE = 34

    def __init__(self, number: int):
        super().__init__()
        self.number = number
        self.width  = self.SIZE
        self.height = self.SIZE

    def wrap(self, aw, ah):
        return self.width, self.height

    def draw(self):
        c = self.canv
        s = self.SIZE
        c.setFillColor(INK)
        c.rect(0, 0, s, s, stroke=0, fill=1)
        fs = 15 if self.number < 100 else 11
        c.setFont("Helvetica-Bold", fs)
        c.setFillColor(WHITE)
        c.drawCentredString(s / 2, s / 2 - fs * 0.36, str(self.number))


class StampBlock(Flowable):
    """
    Dos casillas de verificacion HECHO | REVISADO en horizontal.
    Casillas grandes (18 pt) para marcar con lapiz o pluma sin esfuerzo.
    Cada casilla tiene:
      - Banda de etiqueta oscura con texto blanco en negrita
      - Checkbox XL con guia visual tenue
      - Linea de firma con etiqueta 'Firma / Hora'
    """
    CHECK  = 18.0   # tamanio del checkbox en puntos
    LBL_H  = 16.0   # altura de la banda de etiqueta
    PAD    = 5.0    # padding interno
    GAP    = 6.0    # espacio entre los dos bloques
    RADIUS = 3.0    # esquinas redondeadas del bloque exterior

    def __init__(self, total_width: float):
        super().__init__()
        self._tw  = total_width
        self._bw  = (total_width - self.GAP) / 2
        self.width  = total_width
        self.height = self.LBL_H + self.PAD + self.CHECK + 4 + 10 + 8 + self.PAD

    def wrap(self, aw, ah):
        return self.width, self.height

    def _draw_block(self, x0: float, label: str) -> None:
        c  = self.canv
        bw = self._bw
        h  = self.height

        # Borde exterior redondeado
        c.setStrokeColor(INK)
        c.setLineWidth(1.8)
        c.roundRect(x0, 0, bw, h, self.RADIUS, stroke=1, fill=0)

        # Banda superior de etiqueta oscura
        lh = self.LBL_H
        c.setFillColor(CHARCOAL)
        c.rect(x0, h - lh, bw, lh, stroke=0, fill=1)
        c.setFont("Helvetica-Bold", 9.5)
        c.setFillColor(WHITE)
        c.drawCentredString(x0 + bw / 2, h - lh + 4.5, label)

        # Checkbox XL centrado
        cs   = self.CHECK
        cb_x = x0 + bw / 2 - cs / 2
        cb_y = h - lh - self.PAD - cs
        c.setFillColor(GHOST)
        c.setStrokeColor(INK)
        c.setLineWidth(2.0)
        c.rect(cb_x, cb_y, cs, cs, stroke=1, fill=1)

        # Guia de palomita muy tenue
        c.setStrokeColor(SMOKE)
        c.setLineWidth(1.2)
        mid_x = cb_x + cs * 0.38
        c.line(cb_x + 2.5,  cb_y + cs * 0.46, mid_x,         cb_y + 2.5)
        c.line(mid_x,        cb_y + 2.5,        cb_x + cs - 2, cb_y + cs * 0.72)

        # Divisor interno
        sep_y = cb_y - 6
        c.setStrokeColor(SMOKE)
        c.setLineWidth(0.6)
        c.line(x0 + 5, sep_y, x0 + bw - 5, sep_y)

        # Linea de firma
        line_y = sep_y - 10
        c.setStrokeColor(INK)
        c.setLineWidth(1.1)
        c.line(x0 + 7, line_y, x0 + bw - 7, line_y)

        # Mini-etiqueta bajo la linea
        c.setFont("Helvetica", 6.2)
        c.setFillColor(SLATE)
        c.drawCentredString(x0 + bw / 2, line_y - 7.5, "Firma / Hora")

    def draw(self):
        self.canv.saveState()
        self._draw_block(0,                   "HECHO")
        self._draw_block(self._bw + self.GAP, "REVISADO")
        self.canv.restoreState()


class TypeBadge(Flowable):
    """
    Insignia de tipo de envio.
    PAQUETE -> fondo gris oscuro, texto blanco, borde solido.
    INDIVIDUAL -> fondo blanco, texto oscuro, borde fino.
    Diferenciacion visual en menos de 1 segundo.
    """
    def __init__(self, text: str, is_package: bool, w: float, h: float = 42):
        super().__init__()
        self._text = text
        self._pkg  = is_package
        self.width  = w
        self.height = h

    def wrap(self, aw, ah):
        return self.width, self.height

    def draw(self):
        c   = self.canv
        w   = self.width
        h   = self.height
        pad = 5
        rr  = 5

        if self._pkg:
            c.setFillColor(GRAPHITE)
            c.setStrokeColor(INK)
            c.setLineWidth(1.5)
            c.roundRect(pad, pad, w - 2*pad, h - 2*pad, rr, stroke=1, fill=1)
            text_color = WHITE
        else:
            c.setFillColor(WHITE)
            c.setStrokeColor(GRAPHITE)
            c.setLineWidth(1.2)
            c.roundRect(pad, pad, w - 2*pad, h - 2*pad, rr, stroke=1, fill=1)
            text_color = INK

        c.setFillColor(text_color)
        parts = self._text.split("(")
        if len(parts) == 2:
            c.setFont("Helvetica-Bold", 8.5)
            c.drawCentredString(w / 2, h / 2 + 3.5, parts[0].strip())
            c.setFont("Helvetica", 7.5)
            c.drawCentredString(w / 2, h / 2 - 6.5, "(" + parts[1].strip())
        else:
            c.setFont("Helvetica-Bold", 9.2)
            c.drawCentredString(w / 2, h / 2 - 3.5, self._text)


# ─────────────────────────────────────────────────────────────────────────────
# ESTILOS DE PARRAFO
# ─────────────────────────────────────────────────────────────────────────────

def build_styles() -> dict[str, ParagraphStyle]:
    base = getSampleStyleSheet()
    return {
        "doc_title": ParagraphStyle(
            "DocTitle", parent=base["Normal"],
            fontName="Helvetica-Bold", fontSize=22, leading=26,
            textColor=INK, alignment=TA_LEFT,
        ),
        "doc_sub": ParagraphStyle(
            "DocSub", parent=base["Normal"],
            fontName="Helvetica", fontSize=8.8, leading=11,
            textColor=SLATE, alignment=TA_LEFT, spaceBefore=2,
        ),
        "metric_lbl": ParagraphStyle(
            "MetLbl", parent=base["Normal"],
            fontName="Helvetica-Bold", fontSize=7.5, leading=9,
            textColor=GRAPHITE, alignment=TA_CENTER,
        ),
        "metric_val": ParagraphStyle(
            "MetVal", parent=base["Normal"],
            fontName="Helvetica-Bold", fontSize=26, leading=30,
            textColor=INK, alignment=TA_CENTER,
        ),
        "help": ParagraphStyle(
            "Help", parent=base["Normal"],
            fontName="Helvetica", fontSize=8.5, leading=12,
            textColor=INK, alignment=TA_LEFT,
        ),
        "col_hdr": ParagraphStyle(
            "ColHdr", parent=base["Normal"],
            fontName="Helvetica-Bold", fontSize=8.8, leading=10.5,
            textColor=WHITE, alignment=TA_CENTER,
        ),
        "venta": ParagraphStyle(
            "Venta", parent=base["Normal"],
            fontName="Helvetica-Bold", fontSize=10.5, leading=13,
            textColor=INK, alignment=TA_LEFT,
        ),
        "content_line": ParagraphStyle(
            "ContentLine", parent=base["Normal"],
            fontName="Helvetica", fontSize=9.5, leading=13.5,
            textColor=INK, alignment=TA_LEFT,
        ),
        "obs_label": ParagraphStyle(
            "ObsLabel", parent=base["Normal"],
            fontName="Helvetica-Bold", fontSize=9.0, leading=11,
            textColor=INK, alignment=TA_LEFT,
        ),
    }


# ─────────────────────────────────────────────────────────────────────────────
# ENCABEZADO DEL DOCUMENTO
# ─────────────────────────────────────────────────────────────────────────────

def build_doc_header(
    tasks: Sequence[PackingTask],
    styles: dict[str, ParagraphStyle],
    doc_width: float,
    generated_at: str,
) -> list[object]:

    total        = len(tasks)
    paquetes     = sum(t.es_paquete for t in tasks)
    individuales = total - paquetes

    def metric_card(label: str, value: str) -> Table:
        t = Table(
            [[Paragraph(value, styles["metric_val"])],
             [Paragraph(label, styles["metric_lbl"])]],
            colWidths=[1.35 * inch],
            rowHeights=[0.50 * inch, 0.24 * inch],
        )
        t.setStyle(TableStyle([
            ("BOX",           (0, 0), (-1, -1), 1.6, INK),
            ("LINEBELOW",     (0, 0), (-1, 0),  0.8, SMOKE),
            ("BACKGROUND",    (0, 0), (-1, 0),  WHITE),
            ("BACKGROUND",    (0, 1), (-1, 1),  GHOST),
            ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING",    (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING",   (0, 0), (-1, -1), 4),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 4),
        ]))
        return t

    metrics_row = Table(
        [[metric_card("TOTAL", str(total)),
          metric_card("INDIVIDUALES", str(individuales)),
          metric_card("PAQUETES", str(paquetes))]],
        colWidths=[1.45 * inch] * 3,
    )
    metrics_row.setStyle(TableStyle([
        ("LEFTPADDING",   (0, 0), (-1, -1), 0),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 7),
        ("TOPPADDING",    (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
    ]))

    instrucciones = (
        "<b>Instrucciones de uso:</b>"
        "&nbsp;&nbsp;"
        "<b>1.</b> Localiza el pedido por el numero de fila (#)."
        "&nbsp;&nbsp;"
        "<b>2.</b> Empaca TODOS los productos listados en la columna de contenido."
        "&nbsp;&nbsp;"
        "<b>3.</b> Marca la casilla <b>HECHO</b> y anota tu firma o la hora."
        "&nbsp;&nbsp;"
        "<b>4.</b> Un segundo empleado verifica el paquete y marca <b>REVISADO</b>."
        "&nbsp;&nbsp;"
        "<b>5.</b> Entrega el pedido al cliente."
    )
    help_box = Table(
        [[Paragraph(instrucciones, styles["help"])]],
        colWidths=[doc_width],
    )
    help_box.setStyle(TableStyle([
        ("BOX",           (0, 0), (-1, -1), 1.5, INK),
        ("BACKGROUND",    (0, 0), (-1, -1), GHOST),
        ("LEFTPADDING",   (0, 0), (-1, -1), 11),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 11),
        ("TOPPADDING",    (0, 0), (-1, -1), 9),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 9),
    ]))

    title_block = Table(
        [[Paragraph("Lista de Empaque", styles["doc_title"])],
         [Paragraph(f"Generado: {generated_at}", styles["doc_sub"])]],
        colWidths=[4.4 * inch],
    )
    title_block.setStyle(TableStyle([
        ("TOPPADDING",    (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ("LEFTPADDING",   (0, 0), (-1, -1), 0),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 0),
    ]))

    top_row = Table(
        [[title_block, metrics_row]],
        colWidths=[4.5 * inch, doc_width - 4.5 * inch],
    )
    top_row.setStyle(TableStyle([
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",    (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ("LEFTPADDING",   (0, 0), (-1, -1), 0),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 0),
    ]))

    return [top_row, Spacer(1, 7), help_box, Spacer(1, 10)]


# ─────────────────────────────────────────────────────────────────────────────
# FILAS DE LA TABLA PRINCIPAL
# ─────────────────────────────────────────────────────────────────────────────

def build_table_rows(
    tasks: Sequence[PackingTask],
    styles: dict[str, ParagraphStyle],
    stamp_w: float,
    tipo_w: float,
    row_h: float,
) -> list[list[object]]:

    header = [
        Paragraph("#",                    styles["col_hdr"]),
        Paragraph("HECHO  /  REVISADO",   styles["col_hdr"]),
        Paragraph("VENTA / PEDIDO",        styles["col_hdr"]),
        Paragraph("TIPO",                  styles["col_hdr"]),
        Paragraph("CONTENIDO DEL PEDIDO",  styles["col_hdr"]),
    ]
    rows: list[list[object]] = [header]

    for task in tasks:
        lines_html: list[str] = []
        for line in task.grupo:
            qty   = escape(line.unidades or "-")
            sku   = escape(line.sku      or "SIN SKU")
            title = escape(line.titulo   or "SIN TITULO")
            lines_html.append(
                f"<b>{qty}&nbsp;x&nbsp;&nbsp;{sku}</b>"
                f"&nbsp;&nbsp;-&nbsp;&nbsp;{title}"
            )

        rows.append([
            RowNumber(task.numero),
            StampBlock(total_width=stamp_w),
            Paragraph(html_lines(task.venta_principal), styles["venta"]),
            TypeBadge(task.tipo, task.es_paquete, w=tipo_w, h=row_h - 8),
            Paragraph("<br/>".join(lines_html), styles["content_line"]),
        ])

    return rows


# ─────────────────────────────────────────────────────────────────────────────
# CHROME DE PAGINA  (barra superior + pie)
# ─────────────────────────────────────────────────────────────────────────────

def draw_page_chrome(canvas, doc) -> None:
    pw, ph = PAGE_SIZE
    canvas.saveState()

    # Barra superior negra
    BAR_H = 24
    canvas.setFillColor(INK)
    canvas.rect(0, ph - BAR_H, pw, BAR_H, stroke=0, fill=1)

    # Linea de acento gris debajo de la barra
    canvas.setStrokeColor(SILVER)
    canvas.setLineWidth(1.5)
    canvas.line(0, ph - BAR_H - 1, pw, ph - BAR_H - 1)

    canvas.setFont("Helvetica-Bold", 10.5)
    canvas.setFillColor(WHITE)
    canvas.drawString(L_MARGIN, ph - BAR_H + 7, "LISTA DE EMPAQUE")

    company = getattr(doc, "company", "")
    if company:
        canvas.setFont("Helvetica", 9)
        canvas.setFillColor(SILVER)
        canvas.drawCentredString(pw / 2, ph - BAR_H + 7, company)

    canvas.setFont("Helvetica", 8.5)
    canvas.setFillColor(SMOKE)
    canvas.drawRightString(
        pw - R_MARGIN, ph - BAR_H + 7,
        f"Pagina {canvas.getPageNumber()}",
    )

    # Pie de pagina
    canvas.setStrokeColor(SMOKE)
    canvas.setLineWidth(0.7)
    canvas.line(L_MARGIN, B_MARGIN - 4, pw - R_MARGIN, B_MARGIN - 4)
    canvas.setFont("Helvetica", 6.8)
    canvas.setFillColor(SILVER)
    canvas.drawString(L_MARGIN,         B_MARGIN - 15, "Documento de uso interno - no distribuir")
    canvas.drawRightString(pw - R_MARGIN, B_MARGIN - 15, getattr(doc, "generated_at", ""))

    canvas.restoreState()


# ─────────────────────────────────────────────────────────────────────────────
# CONSTRUCCION DEL PDF
# ─────────────────────────────────────────────────────────────────────────────

def build_pdf_bytes(tasks: Sequence[PackingTask]) -> io.BytesIO:
    buffer       = io.BytesIO()
    styles       = build_styles()
    generated_at = datetime.now().strftime("%d/%m/%Y  %H:%M")

    doc = BaseDocTemplate(
        buffer,
        pagesize=PAGE_SIZE,
        leftMargin=L_MARGIN,
        rightMargin=R_MARGIN,
        topMargin=T_MARGIN,
        bottomMargin=B_MARGIN,
        title="Lista de empaque",
        author=APP_TITLE,
    )
    doc.generated_at = generated_at
    doc.company      = ""   # <-- pon el nombre de tu empresa aqui si lo deseas

    frame = Frame(
        doc.leftMargin, doc.bottomMargin,
        doc.width, doc.height,
        id="main", showBoundary=0,
    )
    doc.addPageTemplates([
        PageTemplate(id="packing", frames=[frame], onPage=draw_page_chrome)
    ])

    # Anchos de columna
    num_w     = 0.42 * inch   # columna '#'
    stamp_w   = 2.10 * inch   # columna HECHO / REVISADO
    venta_w   = 1.45 * inch   # columna venta/pedido
    tipo_w    = 1.30 * inch   # columna tipo
    content_w = doc.width - num_w - stamp_w - venta_w - tipo_w
    col_widths = [num_w, stamp_w, venta_w, tipo_w, content_w]

    # Altura de fila de datos — amplia para comodidad visual
    DATA_ROW_H = 58

    story: list[object] = []
    story += build_doc_header(tasks, styles, doc.width, generated_at)

    table_rows = build_table_rows(tasks, styles, stamp_w, tipo_w, DATA_ROW_H)
    row_heights = [28] + [DATA_ROW_H] * len(tasks)

    main_table = LongTable(
        table_rows,
        colWidths=col_widths,
        rowHeights=row_heights,
        repeatRows=1,
    )

    cmds: list[tuple] = [
        # Cabecera
        ("BACKGROUND",     (0, 0), (-1, 0),  CHARCOAL),
        ("TEXTCOLOR",      (0, 0), (-1, 0),  WHITE),
        ("TOPPADDING",     (0, 0), (-1, 0),  6),
        ("BOTTOMPADDING",  (0, 0), (-1, 0),  6),
        ("LEFTPADDING",    (0, 0), (-1, 0),  6),
        ("RIGHTPADDING",   (0, 0), (-1, 0),  6),
        # Bordes globales
        ("BOX",            (0, 0), (-1, -1), 1.6, INK),
        ("LINEBELOW",      (0, 0), (-1, 0),  2.5, INK),
        ("INNERGRID",      (0, 0), (-1, -1), 0.5, SMOKE),
        ("LINEAFTER",      (0, 0), (0, -1),  1.0, SILVER),
        ("LINEAFTER",      (1, 0), (1, -1),  1.6, INK),
        ("LINEAFTER",      (2, 0), (2, -1),  0.8, SILVER),
        ("LINEAFTER",      (3, 0), (3, -1),  1.2, GRAPHITE),
        # Alineacion y valign
        ("VALIGN",         (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN",          (0, 0), (0, -1),  "CENTER"),
        ("ALIGN",          (1, 0), (1, -1),  "CENTER"),
        ("ALIGN",          (2, 0), (4, -1),  "LEFT"),
        # Padding filas de datos
        ("LEFTPADDING",    (0, 1), (-1, -1), 7),
        ("RIGHTPADDING",   (0, 1), (-1, -1), 7),
        ("TOPPADDING",     (0, 1), (-1, -1), 0),
        ("BOTTOMPADDING",  (0, 1), (-1, -1), 0),
        # Banding alterno blanco / gris muy claro
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [WHITE, GHOST]),
        # Linea divisoria entre filas de datos
        ("LINEBELOW",      (0, 1), (-1, -1), 0.8, SMOKE),
    ]

    # Resalte especial para paquetes
    for ri, task in enumerate(tasks, start=1):
        if task.es_paquete:
            cmds += [
                ("BACKGROUND", (2, ri), (3, ri), PKG_ACCENT),
                ("LINEABOVE",  (0, ri), (-1, ri), 2.0, GRAPHITE),
                ("LINEBELOW",  (0, ri), (-1, ri), 2.0, GRAPHITE),
            ]

    main_table.setStyle(TableStyle(cmds))
    story.append(main_table)

    # Bloque de observaciones al final
    obs = Table(
        [[
            Paragraph("<b>Observaciones:</b>", styles["obs_label"]),
            Paragraph("_" * 115,               styles["content_line"]),
        ]],
        colWidths=[1.55 * inch, doc.width - 1.55 * inch],
        rowHeights=[0.50 * inch],
    )
    obs.setStyle(TableStyle([
        ("BOX",           (0, 0), (-1, -1), 1.4, INK),
        ("LINEBEFORE",    (1, 0), (1, 0),   1.4, INK),
        ("BACKGROUND",    (0, 0), (0, 0),   GHOST),
        ("BACKGROUND",    (1, 0), (1, 0),   WHITE),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING",   (0, 0), (-1, -1), 9),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 9),
        ("TOPPADDING",    (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    story += [Spacer(1, 11), obs]

    doc.build(story)
    buffer.seek(0)
    return buffer


# ─────────────────────────────────────────────────────────────────────────────
# PIPELINE PRINCIPAL
# ─────────────────────────────────────────────────────────────────────────────

def generate_outputs(
    excel_file: BinaryIO,
) -> tuple[pd.DataFrame, io.BytesIO, io.BytesIO, list[str]]:
    df_source    = read_source_dataframe(excel_file)
    parse_result = parse_packing_tasks(df_source)
    df_final     = tasks_to_dataframe(parse_result.tasks)
    excel_buf    = build_excel_bytes(df_final)
    pdf_buf      = build_pdf_bytes(parse_result.tasks)
    return df_final, excel_buf, pdf_buf, parse_result.warnings


# ─────────────────────────────────────────────────────────────────────────────
# STREAMLIT UI
# ─────────────────────────────────────────────────────────────────────────────

def main() -> None:
    if st is None:
        raise RuntimeError("Streamlit no esta instalado en este entorno.")

    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)
    st.caption(
        "PDF con doble control HECHO / REVISADO · "
        "Alta legibilidad para empleados · Optimizado para impresion en blanco y negro."
    )

    uploaded_file = st.file_uploader(
        "Sube el archivo Excel de ventas", type=["xlsx", "xls"]
    )

    if uploaded_file is None:
        st.info("Carga un archivo Excel para generar la lista de empaque.")
        return

    try:
        df_res, xl_out, pdf_out, warns = generate_outputs(uploaded_file)
        st.success(f"Archivo procesado: {len(df_res)} pedidos encontrados.")

        for w in warns:
            st.warning(w)

        st.dataframe(df_res, use_container_width=True, height=480)

        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                "Descargar Excel",
                data=xl_out,
                file_name="lista_empaque.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with col2:
            st.download_button(
                "Descargar PDF para imprimir",
                data=pdf_out,
                file_name="lista_empaque.pdf",
                mime="application/pdf",
                use_container_width=True,
            )

    except Exception as exc:
        st.error(f"Error al procesar: {exc}")
        raise


if __name__ == "__main__":
    main()
