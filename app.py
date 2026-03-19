!pip install pandas openpyxl reportlab -q

import pandas as pd
import re
from google.colab import files

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.units import inch

print("⬆️ Sube tu archivo Excel")
uploaded = files.upload()

# Leer archivo (encabezados en fila 5)
file_name = list(uploaded.keys())[0]
df = pd.read_excel(file_name, header=4)

# ==============================
# 🎯 COLUMNAS FIJAS
# ==============================

col_venta = df.columns[0]
col_unidades = df.columns[6]
col_estado = df.columns[2]
col_titulo = df.columns[20]
col_sku = df.columns[16]  # <-- AJUSTA SI CAMBIA, pero aquí suele estar SKU

df = df.dropna(how="all").reset_index(drop=True)

# ==============================
# 🧠 PROCESAMIENTO
# ==============================

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
                    "Venta": str(subfila[col_venta]),
                    "Unidades": str(subfila[col_unidades]),
                    "Titulo": str(subfila[col_titulo]),
                    "SKU": str(subfila[col_sku])
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
                "Unidades": str(fila[col_unidades]),
                "Titulo": str(fila[col_titulo]),
                "SKU": str(fila[col_sku])
            }]
        })

        i += 1

# ==============================
# 📋 FORMATO FINAL
# ==============================

filas_finales = []

for item in resultado:
    ventas = "\n".join([x["Venta"] for x in item["Grupo"]])
    unidades = "\n".join([x["Unidades"] for x in item["Grupo"]])
    titulos = "\n".join([x["Titulo"] for x in item["Grupo"]])
    skus = "\n".join([x["SKU"] for x in item["Grupo"]])

    filas_finales.append({
        "✔": item["Check"],
        "Venta principal": item["Venta principal"],
        "Tipo": item["Tipo"],
        "SKU": skus,
        "Ventas detalle": ventas,
        "Unidades": unidades,
        "Productos": titulos
    })

df_final = pd.DataFrame(filas_finales)

# ==============================
# 💾 EXCEL
# ==============================

excel_name = "lista_empaque.xlsx"
df_final.to_excel(excel_name, index=False)

# ==============================
# 📄 PDF HORIZONTAL
# ==============================

pdf_name = "lista_empaque.pdf"

doc = SimpleDocTemplate(
    pdf_name,
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
    textColor=colors.white
)

style_cell = ParagraphStyle(
    name="cell",
    parent=styles["Normal"],
    alignment=TA_LEFT,
    fontSize=7,
    wordWrap='CJK'
)

style_center = ParagraphStyle(
    name="center",
    parent=styles["Normal"],
    alignment=TA_CENTER,
    fontSize=7
)

headers = [
    Paragraph("<b>✔</b>", style_header),
    Paragraph("<b>Venta</b>", style_header),
    Paragraph("<b>Tipo</b>", style_header),
    Paragraph("<b>SKU</b>", style_header),
    Paragraph("<b>Unidades</b>", style_header),
    Paragraph("<b>Productos</b>", style_header),
]

table_data = [headers]

for _, row in df_final.iterrows():
    table_data.append([
        Paragraph(str(row["✔"]), style_center),
        Paragraph(str(row["Venta principal"]), style_cell),
        Paragraph(str(row["Tipo"]), style_cell),
        Paragraph(str(row["SKU"]).replace("\n", "<br/>"), style_cell),
        Paragraph(str(row["Unidades"]).replace("\n", "<br/>"), style_center),
        Paragraph(str(row["Productos"]).replace("\n", "<br/>"), style_cell),
    ])

col_widths = [
    0.3 * inch,
    1.5 * inch,
    1.3 * inch,
    2.0 * inch,   # SKU
    0.8 * inch,
    4.5 * inch
]

table = Table(table_data, colWidths=col_widths, repeatRows=1)

table.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#4F81BD")),
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
    ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
]))

doc.build([table])

# ==============================
# 📥 DESCARGA
# ==============================

files.download(excel_name)
files.download(pdf_name)

print("✅ LISTO con SKU incluido")
