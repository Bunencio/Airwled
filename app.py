from __future__ import annotations

import io
import re
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from typing import BinaryIO
from xml.sax.saxutils import escape

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

# =============================================================================
# CONFIG
# =============================================================================

APP_TITLE = "Generador de lista de empaque"
HEADER_ROW_INDEX = 4  # fila 5 en Excel
PACKAGE_PATTERN = re.compile(r"paquete\s+de\s+(\d+)", re.IGNORECASE)

PAGE_SIZE = landscape(letter)
LEFT_MARGIN = 0.22 * inch
RIGHT_MARGIN = 0.22 * inch
TOP_MARGIN = 0.34 * inch
BOTTOM_MARGIN = 0.34 * inch

COLOR_BLACK = colors.black
COLOR_WHITE = colors.white
COLOR_HEADER = colors.HexColor("#222222")
COLOR_GRID = colors.HexColor("#4A4A4A")
COLOR_ROW_A = colors.HexColor("#FFFFFF")
COLOR_ROW_B = colors.HexColor("#F2F2F2")
COLOR_PANEL = colors.HexColor("#EAEAEA")
COLOR_PANEL_DARK = colors.HexColor("#D8D8D8")


# =============================================================================
# MODELOS
# =============================================================================

@dataclass(frozen=True)
class SourceColumns:
    sale: str
    state: str
    units: str
    sku: str
    title: str


# =============================================================================
# HELPERS
# =============================================================================


def clean_value(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()



def normalize_text(value: object) -> str:
    text = clean_value(value)
    text = unicodedata.normalize("NFKD", text)
    text = text.encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"\s+", " ", text).strip().lower()
    return text



def safe_paragraph(value: object, default: str = "-") -> str:
    text = clean_value(value) or default
    return escape(text).replace("\n", "<br/>")



def detect_columns(df: pd.DataFrame) -> tuple[SourceColumns, list[str]]:
    warnings: list[str] = []
    normalized = {col: normalize_text(col) for col in df.columns}

    def find_column(aliases: tuple[str, ...], fallback_index: int, label: str) -> str:
        for col, norm in normalized.items():
            if any(alias in norm for alias in aliases):
                return col
        if fallback_index >= len(df.columns):
            raise ValueError(
                f"No encontré la columna '{label}' y el índice de respaldo {fallback_index + 1} no existe."
            )
        warnings.append(
            f"No encontré por nombre la columna '{label}', así que usé la posición {fallback_index + 1}."
        )
        return df.columns[fallback_index]

    columns = SourceColumns(
        sale=find_column(("venta", "pedido", "order", "folio"), 0, "venta"),
        state=find_column(("estado", "status"), 2, "estado"),
        units=find_column(("unidades", "cantidad", "cant", "qty"), 6, "unidades"),
        sku=find_column(("sku", "codigo", "código", "asin", "referencia"), 16, "sku"),
        title=find_column(("titulo", "título", "producto", "descripcion", "descripción", "articulo", "artículo"), 20, "titulo"),
    )
    return columns, warnings



def is_garbage_row(sale: str, state: str, sku: str, title: str) -> bool:
    combined = normalize_text(" ".join([sale, state, sku, title]))
    if not combined:
        return True

    garbage_markers = (
        "# de venta",
        "venta principal",
        "titulo de la publicacion",
        "título de la publicación",
        "unidades x sku",
        "contenido verificado",
        "iniciales / hora",
        "hecho",
        "no.",
    )

    return any(marker in combined for marker in garbage_markers)



def load_dataframe(excel_file: BinaryIO) -> tuple[pd.DataFrame, SourceColumns, list[str]]:
    df = pd.read_excel(excel_file, header=HEADER_ROW_INDEX)
    df = df.dropna(how="all").reset_index(drop=True)

    if df.empty:
        raise ValueError("El archivo no contiene datos válidos después de la fila de encabezados.")

    columns, warnings = detect_columns(df)
    return df, columns, warnings


# =============================================================================
# PARSEO DE PEDIDOS
# =============================================================================


def parse_orders(df: pd.DataFrame, columns: SourceColumns) -> tuple[pd.DataFrame, list[str]]:
    resultado = []
    warnings: list[str] = []
    i = 0

    while i < len(df):
        fila = df.iloc[i]
        venta_actual = clean_value(fila[columns.sale])
        estado = clean_value(fila[columns.state])
        sku_actual = clean_value(fila[columns.sku])
        titulo_actual = clean_value(fila[columns.title])

        if is_garbage_row(venta_actual, estado, sku_actual, titulo_actual):
            i += 1
            continue

        match = PACKAGE_PATTERN.search(estado)

        if match:
            cantidad = int(match.group(1))
            grupo = []
            disponibles = len(df) - (i + 1)
            real_size = min(cantidad, disponibles)

            if disponibles < cantidad:
                warnings.append(
                    f"La venta '{venta_actual}' indica Paquete de {cantidad}, pero solo encontré {disponibles} filas posteriores."
                )

            for j in range(1, real_size + 1):
                subfila = df.iloc[i + j]
                venta_hija = clean_value(subfila[columns.sale])
                estado_hijo = clean_value(subfila[columns.state])
                sku_hijo = clean_value(subfila[columns.sku])
                titulo_hijo = clean_value(subfila[columns.title])

                if is_garbage_row(venta_hija, estado_hijo, sku_hijo, titulo_hijo):
                    continue

                grupo.append(
                    {
                        "Venta": venta_hija,
                        "Unidades": clean_value(subfila[columns.units]),
                        "SKU": sku_hijo,
                        "Titulo": titulo_hijo,
                    }
                )

            if grupo:
                resultado.append(
                    {
                        "OK": "☐",
                        "Iniciales / Hora": "__________",
                        "Venta principal": venta_actual,
                        "Tipo": f"PAQUETE ({len(grupo)} productos)",
                        "Contenido": "\n".join(
                            f"{item['Unidades'] or '-'} x {item['SKU'] or '-'} - {item['Titulo'] or '-'}"
                            for item in grupo
                        ),
                    }
                )
                i += cantidad + 1
                continue

        resultado.append(
            {
                "OK": "☐",
                "Iniciales / Hora": "__________",
                "Venta principal": venta_actual,
                "Tipo": "INDIVIDUAL",
                "Contenido": f"{clean_value(fila[columns.units]) or '-'} x {sku_actual or '-'} - {titulo_actual or '-'}",
            }
        )
        i += 1

    if not resultado:
        raise ValueError("No se encontraron ventas válidas para exportar.")

    df_final = pd.DataFrame(resultado)
    df_final.insert(0, "No.", range(1, len(df_final) + 1))
    return df_final, warnings


# =============================================================================
# EXCEL
# =============================================================================


def build_excel_buffer(df_final: pd.DataFrame) -> io.BytesIO:
    excel_buffer = io.BytesIO()

    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        df_final.to_excel(writer, index=False, sheet_name="Lista")
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
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border

        for row in ws.iter_rows(min_row=2):
            max_lines = 1
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
                cell.border = border
                max_lines = max(max_lines, str(cell.value or "").count("\n") + 1)
            ws.row_dimensions[row[0].row].height = max(22, min(15 * max_lines, 80))

        widths = {
            "A": 8,
            "B": 8,
            "C": 18,
            "D": 22,
            "E": 18,
            "F": 95,
        }
        for idx in range(1, len(df_final.columns) + 1):
            col_letter = get_column_letter(idx)
            ws.column_dimensions[col_letter].width = widths.get(col_letter, 18)

    excel_buffer.seek(0)
    return excel_buffer


# =============================================================================
# PDF
# =============================================================================


def build_styles() -> dict[str, ParagraphStyle]:
    base = getSampleStyleSheet()
    return {
        "title": ParagraphStyle(
            "Title",
            parent=base["Title"],
            fontName="Helvetica-Bold",
            fontSize=18,
            leading=21,
            textColor=COLOR_BLACK,
            spaceAfter=2,
        ),
        "meta": ParagraphStyle(
            "Meta",
            parent=base["BodyText"],
            fontName="Helvetica",
            fontSize=8.2,
            leading=9.5,
            textColor=COLOR_BLACK,
        ),
        "summary_label": ParagraphStyle(
            "SummaryLabel",
            parent=base["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=7.2,
            leading=8.2,
            alignment=1,
        ),
        "summary_value": ParagraphStyle(
            "SummaryValue",
            parent=base["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=16,
            leading=18,
            alignment=1,
        ),
        "header": ParagraphStyle(
            "Header",
            parent=base["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=8.1,
            leading=9.2,
            alignment=1,
            textColor=COLOR_WHITE,
        ),
        "cell": ParagraphStyle(
            "Cell",
            parent=base["BodyText"],
            fontName="Helvetica",
            fontSize=8.2,
            leading=9.4,
            wordWrap="CJK",
            textColor=COLOR_BLACK,
        ),
        "cell_bold": ParagraphStyle(
            "CellBold",
            parent=base["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=8.2,
            leading=9.4,
            wordWrap="CJK",
            textColor=COLOR_BLACK,
        ),
        "cell_center": ParagraphStyle(
            "CellCenter",
            parent=base["BodyText"],
            fontName="Helvetica",
            fontSize=8.2,
            leading=9.4,
            alignment=1,
            textColor=COLOR_BLACK,
        ),
        "small": ParagraphStyle(
            "Small",
            parent=base["BodyText"],
            fontName="Helvetica",
            fontSize=7,
            leading=8,
            alignment=1,
            textColor=COLOR_BLACK,
        ),
        "foot": ParagraphStyle(
            "Foot",
            parent=base["BodyText"],
            fontName="Helvetica",
            fontSize=7.5,
            leading=8.5,
            textColor=COLOR_BLACK,
        ),
    }



def summary_box(label: str, value: str, width: float, styles: dict[str, ParagraphStyle]) -> Table:
    t = Table(
        [
            [Paragraph(escape(label), styles["summary_label"])],
            [Paragraph(escape(value), styles["summary_value"])],
        ],
        colWidths=[width],
    )
    t.setStyle(
        TableStyle(
            [
                ("BOX", (0, 0), (-1, -1), 0.9, COLOR_BLACK),
                ("BACKGROUND", (0, 0), (-1, -1), COLOR_WHITE),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    return t



def draw_page(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica-Bold", 8.4)
    canvas.drawString(LEFT_MARGIN, PAGE_SIZE[1] - 0.18 * inch, "LISTA DE EMPAQUE")
    canvas.setFont("Helvetica", 8)
    canvas.drawRightString(PAGE_SIZE[0] - RIGHT_MARGIN, PAGE_SIZE[1] - 0.18 * inch, f"Página {canvas.getPageNumber()}")
    canvas.line(LEFT_MARGIN, PAGE_SIZE[1] - 0.21 * inch, PAGE_SIZE[0] - RIGHT_MARGIN, PAGE_SIZE[1] - 0.21 * inch)

    canvas.setFont("Helvetica", 7.5)
    canvas.drawString(LEFT_MARGIN, 0.14 * inch, "Formato optimizado para impresión en blanco y negro")
    canvas.drawRightString(PAGE_SIZE[0] - RIGHT_MARGIN, 0.14 * inch, datetime.now().strftime("%d/%m/%Y %H:%M"))
    canvas.restoreState()



def build_pdf_buffer(df_final: pd.DataFrame) -> io.BytesIO:
    pdf_buffer = io.BytesIO()
    styles = build_styles()

    doc = SimpleDocTemplate(
        pdf_buffer,
        pagesize=PAGE_SIZE,
        leftMargin=LEFT_MARGIN,
        rightMargin=RIGHT_MARGIN,
        topMargin=TOP_MARGIN,
        bottomMargin=BOTTOM_MARGIN,
        title="Lista de empaque",
        author="OpenAI",
    )

    total = len(df_final)
    individuales = int((df_final["Tipo"] == "INDIVIDUAL").sum())
    paquetes = total - individuales
    generated_at = datetime.now().strftime("%d/%m/%Y %H:%M")

    summary_col_widths = [1.58 * inch, 1.58 * inch, 1.58 * inch]
    summary = Table(
        [[
            summary_box("TOTAL DE ENTREGAS", str(total), summary_col_widths[0], styles),
            summary_box("INDIVIDUALES", str(individuales), summary_col_widths[1], styles),
            summary_box("PAQUETES", str(paquetes), summary_col_widths[2], styles),
        ]],
        colWidths=summary_col_widths,
    )
    summary.setStyle(
        TableStyle(
            [
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ]
        )
    )

    instruction = Table(
        [[Paragraph(
            "<b>Uso sugerido:</b> 1) validar venta y contenido, 2) marcar la casilla al completar, 3) escribir iniciales u hora para dejar evidencia, 4) revisar cantidades antes de entregar.",
            styles["meta"],
        )]],
        colWidths=[doc.width],
    )
    instruction.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), COLOR_PANEL),
                ("BOX", (0, 0), (-1, -1), 0.8, COLOR_BLACK),
                ("LEFTPADDING", (0, 0), (-1, -1), 7),
                ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )

    table_data = [[
        Paragraph("<b>No.</b>", styles["header"]),
        Paragraph("<b>Hecho</b>", styles["header"]),
        Paragraph("<b>Iniciales / Hora</b>", styles["header"]),
        Paragraph("<b>Venta principal</b>", styles["header"]),
        Paragraph("<b>Tipo</b>", styles["header"]),
        Paragraph("<b>Contenido verificado</b>", styles["header"]),
    ]]

    for _, row in df_final.iterrows():
        table_data.append([
            Paragraph(str(row["No."]), styles["cell_center"]),
            Paragraph("☐", styles["cell_center"]),
            Paragraph("__________<br/><font size='6'>Iniciales / hora</font>", styles["small"]),
            Paragraph(safe_paragraph(row["Venta principal"]), styles["cell"]),
            Paragraph(safe_paragraph(row["Tipo"]), styles["cell_bold"]),
            Paragraph(safe_paragraph(row["Contenido"]), styles["cell"]),
        ])

    col_widths = [0.42 * inch, 0.58 * inch, 0.90 * inch, 1.52 * inch, 1.36 * inch, doc.width - (0.42 + 0.58 + 0.90 + 1.52 + 1.36) * inch]

    table = Table(table_data, colWidths=col_widths, repeatRows=1, splitByRow=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), COLOR_HEADER),
                ("TEXTCOLOR", (0, 0), (-1, 0), COLOR_WHITE),
                ("BOX", (0, 0), (-1, -1), 0.8, COLOR_GRID),
                ("INNERGRID", (0, 0), (-1, -1), 0.45, COLOR_GRID),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [COLOR_ROW_A, COLOR_ROW_B]),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ("ALIGN", (0, 0), (2, -1), "CENTER"),
            ]
        )
    )

    observations = Table(
        [[
            Paragraph("<b>Observaciones generales:</b> " + "_" * 110, styles["foot"])
        ]],
        colWidths=[doc.width],
    )
    observations.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), COLOR_PANEL_DARK),
                ("BOX", (0, 0), (-1, -1), 0.8, COLOR_BLACK),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )

    elements = [
        Spacer(1, 0.12 * inch),
        Paragraph("Lista de empaque", styles["title"]),
        Paragraph(
            f"Documento listo para impresión. Generado el {generated_at}. Diseño optimizado para lectura rápida, validación manual y control visual.",
            styles["meta"],
        ),
        Spacer(1, 0.10 * inch),
        summary,
        Spacer(1, 0.08 * inch),
        instruction,
        Spacer(1, 0.10 * inch),
        table,
        Spacer(1, 0.10 * inch),
        observations,
    ]

    doc.build(elements, onFirstPage=draw_page, onLaterPages=draw_page)
    pdf_buffer.seek(0)
    return pdf_buffer


# =============================================================================
# PIPELINE
# =============================================================================


def generar_archivos(excel_file: BinaryIO):
    df, columns, warnings = load_dataframe(excel_file)
    df_final, parse_warnings = parse_orders(df, columns)
    warnings.extend(parse_warnings)

    excel_buffer = build_excel_buffer(df_final)
    pdf_buffer = build_pdf_buffer(df_final)
    return df_final, excel_buffer, pdf_buffer, warnings


# =============================================================================
# STREAMLIT
# =============================================================================


def main() -> None:
    st.set_page_config(page_title="Lista de empaque", layout="wide")
    st.title(APP_TITLE)
    st.caption(
        "Versión horizontal compacta para meter más entregas por página, mantener buena lectura y reducir hojas impresas."
    )

    uploaded_file = st.file_uploader("Sube el archivo Excel", type=["xlsx", "xls"])

    if uploaded_file is None:
        st.info("Esperando archivo Excel para procesar.")
        return

    try:
        with st.spinner("Procesando archivo y construyendo PDF horizontal..."):
            df_resultado, excel_out, pdf_out, warnings = generar_archivos(uploaded_file)

        st.success("Archivo procesado correctamente")

        c1, c2, c3 = st.columns(3)
        c1.metric("Entregas", len(df_resultado))
        c2.metric("Individuales", int((df_resultado["Tipo"] == "INDIVIDUAL").sum()))
        c3.metric("Paquetes", int((df_resultado["Tipo"] != "INDIVIDUAL").sum()))

        if warnings:
            with st.expander("Advertencias de lectura"):
                for warning in warnings:
                    st.warning(warning)

        st.dataframe(df_resultado, use_container_width=True, height=480)

        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                "Descargar Excel",
                data=excel_out.getvalue(),
                file_name="lista_empaque_horizontal.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with col2:
            st.download_button(
                "Descargar PDF horizontal",
                data=pdf_out.getvalue(),
                file_name="lista_empaque_horizontal.pdf",
                mime="application/pdf",
                use_container_width=True,
            )

    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")


if __name__ == "__main__":
    main()
