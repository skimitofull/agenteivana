import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import base64

# --- CONSTANTES GLOBALES ---
PAGE_WIDTH, PAGE_HEIGHT = letter
MARGIN = 25 * mm
HEADER_HEIGHT = 15 * mm
ROW_HEIGHT = 15 * mm
LINE_SPACING = 5 * mm
ROWS_PER_PAGE = 15

# Registrar fuente
try:
    pdfmetrics.registerFont(TTFont('Arial', 'Arial.ttf'))
    FONT = 'Arial'
except:
    FONT = 'Helvetica'

def clean_amount(amount):
    if pd.isnull(amount) or str(amount).strip() == '':
        return ''
    try:
        if isinstance(amount, str):
            amount = float(amount.replace('$', '').replace(',', ''))
        return '${:,.2f}'.format(amount) if amount != 0 else ''
    except ValueError:
        return ''

def clean_date(date):
    if pd.isnull(date):
        return ''
    try:
        if isinstance(date, str) and ('NOV' in date.upper() or 'DIC' in date.upper()):
            return date.upper().strip()
        return pd.to_datetime(date).strftime('%d %b').upper()
    except:
        return str(date).upper()

def calculate_row_height(concept_parts):
    return max(len(concept_parts) * LINE_SPACING, ROW_HEIGHT)

def split_concept(concept):
    if pd.isnull(concept):
        return ['']
    concept = str(concept).strip()
    parts = []
    max_line_length = 50

    if "TRANSF INTERBANCARIA SPEI" in concept:
        parts = ['TRANSF INTERBANCARIA SPEI']
        remaining = concept.replace('TRANSF INTERBANCARIA SPEI', '').strip()

        # Extraer fecha si existe
        for word in remaining.split():
            if any(month in word.upper() for month in ['NOV', 'DIC']):
                parts.append(word)
                remaining = remaining.replace(word, '').strip()
                break

        # Extraer referencia si existe
        for word in remaining.split():
            if word.startswith('//'):
                parts.append(word)
                remaining = remaining.replace(word, '').strip()
                break

        # Procesar el resto del texto
        if remaining:
            words = remaining.split()
            current_line = []
            for word in words:
                if len(' '.join(current_line + [word])) <= max_line_length:
                    current_line.append(word)
                else:
                    if current_line:
                        parts.append(' '.join(current_line))
                    current_line = [word]
            if current_line:
                parts.append(' '.join(current_line))

    elif "SCOTIALINE" in concept:
        parts = ["SWEB PAGO A SCOTIALINE"]
        # Extraer número de referencia si existe
        for word in concept.split():
            if word.isdigit() and len(word) > 10:
                parts.append(word)
                break

    else:
        words = concept.split()
        current_line = []
        for word in words:
            if len(' '.join(current_line + [word])) <= max_line_length:
                current_line.append(word)
            else:
                if current_line:
                    parts.append(' '.join(current_line))
                current_line = [word]
        if current_line:
            parts.append(' '.join(current_line))

    return parts

def create_pdf(df):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    width, height = PAGE_WIDTH, PAGE_HEIGHT

    # Definición de columnas
    headers = ['Fecha', 'Concepto', 'Origen / Referencia', 'Depósito', 'Retiro', 'Saldo']
    col_widths = [
        0.10,  # Fecha
        0.35,  # Concepto
        0.25,  # Origen / Referencia
        0.10,  # Depósito
        0.10,  # Retiro
        0.10   # Saldo
    ]

    # Calcular posiciones X
    usable_width = width - (2 * MARGIN)
    col_widths = [w * usable_width for w in col_widths]
    col_x = [MARGIN]
    for w in col_widths[:-1]:
        col_x.append(col_x[-1] + w)

    total_rows = len(df)
    row_idx = 0
    page_number = 1

    while row_idx < total_rows:
        y = height - MARGIN
        header_y = y - HEADER_HEIGHT

        # Encabezado
        c.setFillColor(colors.black)
        c.setStrokeColor(colors.black)
        c.setFont(FONT, 10)
        c.rect(MARGIN, header_y, usable_width, HEADER_HEIGHT, fill=1)

        # Texto del encabezado
        c.setFillColor(colors.white)
        for i, header in enumerate(headers):
            x_pos = col_x[i] + 2 * mm
            if header in ['Depósito', 'Retiro', 'Saldo']:
                text_width = c.stringWidth(header, FONT, 10)
                x_pos = col_x[i] + col_widths[i] - text_width - 2 * mm
            c.drawString(x_pos, header_y + 4 * mm, header)

        c.setFillColor(colors.black)
        y = header_y - ROW_HEIGHT
        rows_on_page = 0

        while row_idx < total_rows and rows_on_page < ROWS_PER_PAGE:
            row = df.iloc[row_idx]
            concept_parts = split_concept(row['Concepto'])
            row_height = calculate_row_height(concept_parts)

            if y - row_height < MARGIN:
                break

            # Color de fondo alternado
            if rows_on_page % 2 == 0:
                c.setFillColor(colors.HexColor("#F0F0F0"))
                c.rect(MARGIN, y - row_height, usable_width, row_height, fill=1, stroke=0)
                c.setFillColor(colors.black)

            # Contenido
            for i, col in enumerate(headers):
                value = str(row[col]) if pd.notnull(row[col]) else ''
                x_pos = col_x[i] + 2 * mm

                if col == 'Concepto':
                    for j, line in enumerate(concept_parts):
                        text_y = y - (j + 1) * LINE_SPACING
                        c.drawString(x_pos, text_y, line)
                elif col in ['Depósito', 'Retiro', 'Saldo']:
                    text_width = c.stringWidth(value, FONT, 10)
                    x_pos = col_x[i] + col_widths[i] - text_width - 2 * mm
                    c.drawString(x_pos, y - LINE_SPACING, value)
                else:
                    c.drawString(x_pos, y - LINE_SPACING, value)

            # Líneas verticales
            c.setStrokeColor(colors.HexColor("#E5E5E5"))
            for x in col_x:
                c.line(x, header_y, x, y - row_height)
            c.line(MARGIN + usable_width, header_y, MARGIN + usable_width, y - row_height)

            y -= row_height
            row_idx += 1
            rows_on_page += 1

        # Número de página
        c.setFont(FONT, 9)
        page_text = f"Página {page_number}"
        c.drawRightString(width - MARGIN, MARGIN / 2, page_text)

        c.showPage()
        page_number += 1

    c.save()
    buffer.seek(0)
    return buffer

def create_preview_html(pdf_bytes):
    b64 = base64.b64encode(pdf_bytes).decode('utf-8')
    html = f'''
        <iframe
            src="data:application/pdf;base64,{b64}"
            width="800"
            height="1000"
            type="application/pdf"
        >
        </iframe>
    '''
    return html

# --- INTERFAZ STREAMLIT ---
st.set_page_config(page_title="Generador de Estado de Cuenta Scotiabank", layout="wide")

st.title("Generador de Estado de Cuenta Scotiabank")
st.write("Sube tu archivo Excel con movimientos y descarga el PDF generado.")

uploaded_file = st.file_uploader("Selecciona tu archivo Excel (.xlsx)", type=['xlsx'])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl")

        st.write("Vista previa de los movimientos:")
        st.dataframe(df)

        df['Fecha'] = df['Fecha'].apply(clean_date)
        df['Depósito'] = df['Depósito'].apply(clean_amount)
        df['Retiro'] = df['Retiro'].apply(clean_amount)
        df['Saldo'] = df['Saldo'].apply(clean_amount)

        if st.button("Generar PDF"):
            with st.spinner('Generando PDF...'):
                pdf_buffer = create_pdf(df)
                st.success('¡PDF generado correctamente!')

                col1, col2 = st.columns([1, 1])

                with col1:
                    st.download_button(
                        label="Descargar PDF",
                        data=pdf_buffer.getvalue(),
                        file_name="estado_cuenta_modificado.pdf",
                        mime="application/pdf"
                    )

                st.subheader("Vista previa del PDF:")
                pdf_preview = create_preview_html(pdf_buffer.getvalue())
                st.markdown(pdf_preview, unsafe_allow_html=True)

    except Exception as e:
        st.error(f"Error al procesar el archivo: {str(e)}")
else:
    st.info("Por favor, sube un archivo Excel para comenzar.")
