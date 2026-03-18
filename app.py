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

def generar_archivos(excel_file):
    df = pd.read_excel(excel_file, header=4)
    df = df.dropna(how="all").reset_index(drop=True)

    col_venta = df.columns[0]
    col_unidades = df.columns[6]
    col_estado = df.columns[2]
    col_titulo = df.columns[20]

    resultado = []
    i = 0

    while i < len(df):
        fila = df.iloc[i]
        estado = str(fila[col_estado]).strip() if pd.notna(fila[col_estado]) else ""
        venta_actual = "" if pd.isna(fila[col_venta]) else str(fila[col_venta])

        match = re.search(r'Paquete de (\d+)', estado, re.IGNORECASE)

        if match:
            cantidad = int(match.group(1))
            grupo = []

            for j in range(1, cantidad + 1):
                if i + j < len(df):
                    subfila = df.iloc[i + j]
                    grupo.append({
                        "Venta": "" if pd.isna(subfila[col_venta]) else str(subfila[col_venta]),
                        "Unidades": "" if pd.isna(subfila[col_unidades]) else str(subfila[col_unidades]),
                        "Titulo": "" if pd.isna(subfila[col_titulo]) else str(subfila[col_titulo])
                    })

            resultado.append({
                "Check": "☐",
                "Venta principal": venta_actual,
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
                    "Unidades": "" if pd.isna(fila[col_unidades]) else str(fila[col_unidades]),
                    "Titulo": "" if pd.isna(fila[col_titulo]) else str(fila[col_titulo])
                }]
            })
            i += 1

    filas_finales = []
    for item in resultado:
        if item["Tipo"].startswith("JUNTO"):
            ventas_detalle = "\n".join([x["Venta"] for x in item["Grupo"]])
            unidades = "\n".join([x["Unidades"] for x in item["Grupo"]])
            titulos = "\n".join([x["Titulo"] for x in item["Grupo"]])
        else:
            ventas_detalle = item["Grupo"][0]["Venta"]
            unidades = item["Grupo"][0]["Unidades"]
            titulos = item["Grupo"][0]["Titulo"]

        filas_finales.append({
            "✔": item["Check"],
            "Venta principal": item["Venta principal"],
            "Tipo": item["Tipo"],
            "Ventas detalle": ventas_detalle,
            "Unidades": unidades,
            "Productos": titulos
        })

    df_final = pd.DataFrame(filas_finales)

    excel_buffer = io.BytesIO()
    df_final.to_excel(excel_buffer, index=False)
    excel_buffer.seek(0)

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
        Paragraph("<b>Ventas detalle</b>", style_header),
        Paragraph("<b>Unidades</b>", style_header),
        Paragraph("<b>Productos</b>", style_header),
    ]

    table_data = [headers]
    for _, row in df_final.iterrows():
        table_data.append([
            Paragraph(str(row["✔"]), style_center),
            Paragraph(str(row["Venta principal"]).replace("\n", "<br/>"), style_cell),
            Paragraph(str(row["Tipo"]).replace("\n", "<br/>"), style_cell),
            Paragraph(str(row["Ventas detalle"]).replace("\n", "<br/>"), style_cell),
            Paragraph(str(row["Unidades"]).replace("\n", "<br/>"), style_center),
            Paragraph(str(row["Productos"]).replace("\n", "<br/>"), style_cell),
        ])

    col_widths = [0.35 * inch, 1.55 * inch, 1.45 * inch, 1.75 * inch, 0.85 * inch, 4.55 * inch]
    table = Table(table_data, colWidths=col_widths, repeatRows=1)

    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#4F81BD")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('ALIGN', (0, 0), (0, -1), 'CENTER'),
        ('ALIGN', (4, 1), (4, -1), 'CENTER'),
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

        st.download_button(
            "Descargar Excel",
            data=excel_out,
            file_name="lista_empaque.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.download_button(
            "Descargar PDF",
            data=pdf_out,
            file_name="lista_empaque.pdf",
            mime="application/pdf"
        )
    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")
