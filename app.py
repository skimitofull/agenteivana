import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
from io import BytesIO
from PIL import Image, ImageDraw, ImageFont
import numpy as np

st.title("Agente IA: Modificador avanzado de Estado de Cuenta PDF")

st.write("Sube tu PDF original y el Excel con los movimientos nuevos.")

pdf_file = st.file_uploader("Sube el estado de cuenta PDF", type=["pdf"])
excel_file = st.file_uploader("Sube el archivo Excel de movimientos", type=["xlsx"])

def format_currency(amount):
    if pd.isnull(amount) or str(amount).strip() == '' or str(amount).lower() == 'nan':
        return ''
    try:
        if isinstance(amount, str):
            amount = float(str(amount).replace('$', '').replace(',', ''))
        return '${:,.2f}'.format(amount) if amount != 0 else ''
    except:
        return ''

def format_date(date):
    if pd.isnull(date):
        return ''
    try:
        if isinstance(date, str) and ('NOV' in date or 'DIC' in date):
            return date.strip()
        return pd.to_datetime(date).strftime('%d %b').upper()
    except:
        return str(date).upper()

def create_table_image(df):
    # Configuración visual
    width = 2480
    row_height = 40
    header_height = 60
    margin = 50
    headers = ['Fecha', 'Concepto', 'Origen / Referencia', 'Depósito', 'Retiro', 'Saldo']
    col_widths = [180, 900, 400, 250, 250, 300]
    x_positions = np.cumsum([margin] + col_widths[:-1])
    total_rows = len(df) + 1
    total_height = (total_rows * row_height) + header_height + 2 * margin

    # Crear imagen
    img = Image.new('RGB', (width, total_height), 'white')
    draw = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype("arial.ttf", 22)
        font_bold = ImageFont.truetype("arialbd.ttf", 24)
    except:
        font = ImageFont.load_default()
        font_bold = ImageFont.load_default()

    # Encabezado
    y = margin
    for i, header in enumerate(headers):
        x = x_positions[i]
        draw.rectangle([x, y, x + col_widths[i], y + header_height], fill='#E8E8E8')
        draw.text((x + 10, y + 15), header, fill='black', font=font_bold)

    # Filas
    for idx, row in df.iterrows():
        y = margin + header_height + (idx * row_height)
        fill_color = '#F5F5F5' if idx % 2 == 0 else 'white'
        draw.rectangle([margin, y, width - margin, y + row_height], fill=fill_color)
        for i, col in enumerate(headers):
            x = x_positions[i]
            value = str(row[col]) if col in row else ''
            if col in ['Depósito', 'Retiro', 'Saldo']:
                text_width = draw.textlength(value, font=font)
                draw.text((x + col_widths[i] - text_width - 10, y + 8), value, fill='black', font=font)
            else:
                draw.text((x + 10, y + 8), value, fill='black', font=font)
    return img

if pdf_file and excel_file:
    # Leer movimientos
    df = pd.read_excel(excel_file)
    # Limpiar y formatear
    df = df.dropna(subset=['Fecha', 'Concepto'], how='all')
    df['Fecha'] = df['Fecha'].apply(format_date)
    for col in ['Depósito', 'Retiro', 'Saldo']:
        df[col] = df[col].apply(format_currency)
    st.write("Vista previa de movimientos:", df.head())

   # Crear imagen de la tabla
img = create_table_image(df)
img_bytes = BytesIO()
img.save(img_bytes, format='PNG')
st.image(img_bytes.getvalue(), caption="Vista previa de la tabla de movimientos")

# Convertir la imagen a PDF usando PIL
pdf_bytes_table = BytesIO()
img_rgb = img.convert('RGB')
img_rgb.save(pdf_bytes_table, format='PDF')
pdf_bytes_table.seek(0)

# Procesar PDF original
pdf_bytes = pdf_file.read()
doc = fitz.open(stream=pdf_bytes, filetype="pdf")

# Abrir el PDF de la tabla como documento PDF
img_pdf = fitz.open(stream=pdf_bytes_table.getvalue(), filetype="pdf")

# Insertar la tabla como nueva página al final
doc.insert_pdf(img_pdf)

    # Guardar PDF modificado
    output = BytesIO()
    doc.save(output)
    st.success("¡PDF generado con la tabla de movimientos avanzada!")
    st.download_button("Descargar PDF modificado", output.getvalue(), file_name="estado_cuenta_modificado.pdf", mime="application/pdf")
