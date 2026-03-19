import io
import re
import pandas as pd
import streamlit as st

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.units import inch

st.set_page_config(page_title="Lista de empaque", layout="wide")
st.title("Generador de lista de empaque")

uploaded_file = st.file_uploader("Sube el archivo Excel", type=["xlsx", "xls"])

def limpiar_valor(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip()

def generar_archivos(excel_file):
    # Encabezados en fila 5
    df = pd.read_excel(excel_file, header=4)
    df = df.dropna(how="all").reset_index(drop=True)

    # Columnas según tu estructura
    col_venta = df.columns[0]
    col_estado = df.columns[2]
    col_unidades = df.columns[6]
    col_sku = df.columns[16]
    col_titulo = df.columns[20]

    resultado = []
    i = 0

    while i < len(df):
        fila = df.iloc[i]
        estado = limpiar_valor(fila[col_estado])
        venta_actual = limpiar_valor(fila[col_venta])

        match = re.search(r"Paquete de (\d+)", estado, re.IGNORECASE)

        if match:
            cantidad = int(match.group(1))
            grupo = []

            for j in range(1, cantidad + 1):
                if i + j < len(df):
                    subfila = df.iloc[i + j]
                    grupo.append({
                        "Venta": limpiar_valor(subfila[col_venta]),
                        "Unidades": limpiar_valor(subfila[col_unidades]),
                        "SKU": limpiar_valor(subfila[col_sku]),
                        "Titulo": limpiar_valor(subfila[col_titulo]),
                    })

            resultado.append({
                "Check": "☐",
                "Venta principal": venta_actual,  # ID del paquete
                "Tipo": f"JUNTO ({cantidad} productos)",
                "Grupo": grupo
            })

            i += cantidad + 1

        else:
            resultado.append({
                "Check": "☐",
                "Venta principal": venta_actual,
                "Tipo": "INDIVIDUAL",
                "Grupo": [{
                    "Venta": venta_actual,
                    "Unidades": limpiar_valor(fila[col_unidades]),
                    "SKU": limpiar_valor(fila[col_sku]),
                    "Titulo": limpiar_valor(fila[col_titulo]),
                }]
            })

            i += 1

    filas_finales = []

    for item in resultado:
        
        unidades = "\n".join([x["Unidades"] for x in item["Grupo"]])
        skus = "\n".join([x["SKU"] for x in item["Grupo"]])
        titulos = "\n".join([x["Titulo"] for x in item["Grupo"]])

        filas_finales.append({
            "✔": item["Check"],
            "Venta principal": item["Venta principal"],
            "Tipo": item["Tipo"],
            "SKU": skus,
            "Unidades": unidades,
            "Productos": titulos
        })

    df_final = pd.DataFrame(filas_finales)

    # Excel en memoria
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        df_final.to_excel(writer, index=False, sheet_name="Lista")
    excel_buffer.seek(0)

    # PDF en memoria
    pdf_buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        pdf_buffer,
        pagesize=landscape(letter),
        rightMargin=15,
        leftMargin=15,
        topMargin=20,
        bottomMargin=20
    )

    styles = getSampleStyleSheet()

    style_header = ParagraphStyle(
        name="header",
        parent=styles["Normal"],
        alignment=TA_CENTER,
        fontName="Helvetica-Bold",
        fontSize=8,
        leading=10,
        textColor=colors.white
    )

    style_cell = ParagraphStyle(
        name="cell",
        parent=styles["Normal"],
        alignment=TA_LEFT,
        fontName="Helvetica",
        fontSize=7,
        leading=9,
        wordWrap="CJK"
    )

    style_center = ParagraphStyle(
        name="center",
        parent=styles["Normal"],
        alignment=TA_CENTER,
        fontName="Helvetica",
        fontSize=7,
        leading=9
    )

    headers = [
        Paragraph("<b>✔</b>", style_header),
        Paragraph("<b>Venta principal</b>", style_header),
        Paragraph("<b>Tipo</b>", style_header),
        Paragraph("<b>SKU</b>", style_header),
        
        Paragraph("<b>Unidades</b>", style_header),
        Paragraph("<b>Productos</b>", style_header),
    ]

    table_data = [headers]

    for _, row in df_final.iterrows():
        table_data.append([
            Paragraph(str(row["✔"]), style_center),
            Paragraph(str(row["Venta principal"]).replace("\n", "<br/>"), style_cell),
            Paragraph(str(row["Tipo"]).replace("\n", "<br/>"), style_cell),
            Paragraph(str(row["SKU"]).replace("\n", "<br/>"), style_cell),
            
            Paragraph(str(row["Unidades"]).replace("\n", "<br/>"), style_center),
            Paragraph(str(row["Productos"]).replace("\n", "<br/>"), style_cell),
        ])

    col_widths = [
        0.35 * inch,  # check
        1.45 * inch,  # venta principal
        1.40 * inch,  # tipo
        1.70 * inch,  # sku
        
        0.80 * inch,  # unidades
        3.80 * inch   # productos
    ]

    table = Table(table_data, colWidths=col_widths, repeatRows=1)

    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#4F81BD")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('ALIGN', (0, 0), (0, -1), 'CENTER'),
        ('ALIGN', (5, 1), (5, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.whitesmoke, colors.lightgrey]),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))

    elements = [
        Paragraph("<b>Lista de empaque</b>", styles["Title"]),
        Spacer(1, 8),
        table
    ]

    doc.build(elements)
    pdf_buffer.seek(0)

    return df_final, excel_buffer, pdf_buffer


if uploaded_file is not None:
    try:
        df_resultado, excel_out, pdf_out = generar_archivos(uploaded_file)

        st.success("Archivo procesado correctamente")
        st.dataframe(df_resultado, use_container_width=True)

        col1, col2 = st.columns(2)

        with col1:
            st.download_button(
                "Descargar Excel",
                data=excel_out,
                file_name="lista_empaque.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with col2:
            st.download_button(
                "Descargar PDF",
                data=pdf_out,
                file_name="lista_empaque.pdf",
                mime="application/pdf"
            )

    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")
