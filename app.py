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

# --- CONSTANTES GLOBALES AJUSTADAS ---
PAGE_WIDTH, PAGE_HEIGHT = letter
MARGIN = 30 * mm
HEADER_HEIGHT = 12 * mm
ROW_HEIGHT = 12 * mm
LINE_SPACING = 4 * mm
ROWS_PER_PAGE = 25
FONT_SIZE = 8
HEADER_FONT_SIZE = 9

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

def split_concept(concept, max_chars=45):
    if pd.isnull(concept):
        return ['']

    concept = str(concept).strip()
    parts = []

    # Manejo especial para SPEI
    if "TRANSF INTERBANCARIA SPEI" in concept:
        # Extraer componentes específicos
        date_part = None
        ref_part = None
        name_part = None

        # Buscar fecha
        words = concept.split()
        for i, word in enumerate(words):
            if any(month in word.upper() for month in ['NOV', 'DIC']):
                date_part = f"{words[i-1]} {word}" if i > 0 else word
                break

        # Buscar referencia (comienza con //)
        for word in words:
            if word.startswith('//'):
                ref_part = word
                break

        # Buscar nombre específico
        if "JOSE TOMAS COLSA CHALITA" in concept:
            name_part = "JOSE TOMAS COLSA CHALITA"

        # Construir partes en orden específico
        parts.append('TRANSF INTERBANCARIA SPEI')
        if date_part:
            parts.append(date_part)
        if name_part:
            parts.append(name_part)
        if ref_part:
            parts.append(ref_part)

    elif "SCOTIALINE" in concept:
        parts = ["SWEB PAGO A SCOTIALINE"]
        # Extraer número de referencia
        for word in concept.split():
            if word.isdigit() and len(word) > 10:
                parts.append(word)
                break
    else:
        # División general de texto
        current_line = []
        words = concept.split()

        for word in words:
            if len(' '.join(current_line + [word])) <= max_chars:
                current_line.append(word)
            else:
                if current_line:
                    parts.append(' '.join(current_line))
                current_line = [word]
        if current_line:
            parts.append(' '.join(current_line))

    return parts

def calculate_row_height(concept_parts):
    return max(len(concept_parts) * LINE_SPACING, ROW_HEIGHT)

def create_pdf(df):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    width, height = PAGE_WIDTH, PAGE_HEIGHT

    # Definición de columnas ajustada
    headers = ['Fecha', 'Concepto', 'Origen / Referencia', 'Depósito', 'Retiro', 'Saldo']
    col_widths = [
        0.08,  # Fecha
        0.42,  # Concepto
        0.20,  # Origen / Referencia
        0.10,  # Depósito
        0.10,  # Retiro
        0.10   # Saldo
    ]

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
        c.setFont(FONT, HEADER_FONT_SIZE)
        c.rect(MARGIN, header_y, usable_width, HEADER_HEIGHT, fill=1)

        # Texto del encabezado
        c.setFillColor(colors.white)
        for i, header in enumerate(headers):
            x_pos = col_x[i] + 2 * mm
            if header in ['Depósito', 'Retiro', 'Saldo']:
                text_width = c.stringWidth(header, FONT, HEADER_FONT_SIZE)
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

            # Fondo alternado
            if rows_on_page % 2 == 0:
                c.setFillColor(colors.HexColor("#F5F5F5"))
                c.rect(MARGIN, y - row_height, usable_width, row_height, fill=1, stroke=0)
                c.setFillColor(colors.black)

            # Contenido
            c.setFont(FONT, FONT_SIZE)
            for i, col in enumerate(headers):
                value = str(row[col]) if pd.notnull(row[col]) else ''
                x_pos = col_x[i] + 2 * mm

                if col == 'Concepto':
                    for j, line in enumerate(concept_parts):
                        text_y = y - ((j + 1) * LINE_SPACING)
                        c.drawString(x_pos, text_y, line)
                elif col in ['Depósito', 'Retiro', 'Saldo']:
                    text_width = c.stringWidth(value, FONT, FONT_SIZE)
                    x_pos = col_x[i] + col_widths[i] - text_width - 2 * mm
                    c.drawString(x_pos, y - LINE_SPACING, value)
                else:
                    c.drawString(x_pos, y - LINE_SPACING, value)

            # Líneas verticales
            c.setStrokeColor(colors.HexColor("#CCCCCC"))
            for x in col_x:
                c.line(x, header_y, x, y - row_height)
            c.line(MARGIN + usable_width, header_y, MARGIN + usable_width, y - row_height)

            y -= row_height
            row_idx += 1
            rows_on_page += 1

        # Número de página
        c.setFont(FONT, FONT_SIZE)
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
