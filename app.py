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
MARGIN = 20 * mm
HEADER_HEIGHT = 12 * mm
ROW_HEIGHT = 10 * mm
LINE_SPACING = 4 * mm
ROWS_PER_PAGE = 20

# Registrar fuente Arial si la tienes, si no, usa Helvetica
try:
    pdfmetrics.registerFont(TTFont('Arial', 'Arial.ttf'))
    FONT = 'Arial'
except:
    FONT = 'Helvetica'

def split_concept(concept):
    if pd.isnull(concept):
        return ['']
    concept = str(concept).strip()
    parts = []
    if "TRANSF INTERBANCARIA SPEI" in concept:
        parts = [
            'TRANSF INTERBANCARIA SPEI',
            'TRANSF INTERBANCARIA SPEI',
            '/TRANSFERENCIA A'
        ]
        for part in concept.split():
            if part.startswith('202'):
                parts.append(part)
                break
        if "DIC" in concept:
            parts.append("02 DIC")
        elif "NOV" in concept:
            parts.append("19 NOV")
        if "JOSE TOMAS COLSA CHALITA" in concept:
            parts.append("JOSE TOMAS COLSA CHALITA")
        for part in concept.split():
            if part.startswith('//'):
                parts.append(part)
                break
    elif "SCOTIALINE" in concept:
        parts = ["SWEB PAGO A SCOTIALINE"]
        for part in concept.split():
            if part.isdigit() and len(part) > 10:
                parts.append(part)
                break
    else:
        words = concept.split()
        current_line = []
        for word in words:
            current_line.append(word)
            if len(' '.join(current_line)) > 40:
                parts.append(' '.join(current_line[:-1]))
                current_line = [word]
        if current_line:
            parts.append(' '.join(current_line))
    return parts

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

def create_pdf(df):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    width, height = PAGE_WIDTH, PAGE_HEIGHT

    # Columnas
    headers = ['Fecha', 'Concepto', 'Origen / Referencia', 'Depósito', 'Retiro', 'Saldo']
    col_widths = [
        0.08,  # Fecha
        0.40,  # Concepto
        0.20,  # Origen / Referencia
        0.11,  # Depósito
        0.11,  # Retiro
        0.10   # Saldo
    ]
    col_widths = [w * (width - 2 * MARGIN) for w in col_widths]
    col_x = [MARGIN]
    for w in col_widths[:-1]:
        col_x.append(col_x[-1] + w)
    col_x.append(width - MARGIN)

    total_rows = len(df)
    row_idx = 0
    page_number = 1

    while row_idx < total_rows:
        y = height - MARGIN - HEADER_HEIGHT

        # Encabezado
        c.setFillColor(colors.black)
        c.setStrokeColor(colors.black)
        c.setFont(FONT, 10)
        c.rect(MARGIN, y, width - 2 * MARGIN, HEADER_HEIGHT, fill=1)
        c.setFillColor(colors.white)
        for i, header in enumerate(headers):
            c.drawString(col_x[i] + 2, y + 3, header)
        c.setFillColor(colors.black)

        y -= ROW_HEIGHT

        # Filas
        rows_on_page = 0
        while row_idx < total_rows and rows_on_page < ROWS_PER_PAGE:
            row = df.iloc[row_idx]
            concept_parts = split_concept(row['Concepto'])
            row_height = max(len(concept_parts) * (LINE_SPACING), ROW_HEIGHT)

            # Alternancia de color
            if rows_on_page % 2 == 0:
                c.setFillColor(colors.HexColor("#F0F0F0"))
                c.rect(MARGIN, y, width - 2 * MARGIN, row_height, fill=1, stroke=0)
                c.setFillColor(colors.black)

            # Columnas
            for i, col in enumerate(headers):
                x = col_x[i]
                value = str(row[col]) if pd.notnull(row[col]) else ''
                if col == 'Concepto':
                    for j, line in enumerate(concept_parts):
                        c.drawString(x + 2, y + row_height - (j + 1) * LINE_SPACING, line)
                elif col in ['Depósito', 'Retiro', 'Saldo']:
                    c.drawRightString(col_x[i+1] - 4, y + row_height - LINE_SPACING, value)
                else:
                    c.drawString(x + 2, y + row_height - LINE_SPACING, value)

            # Líneas verticales
            for x in col_x:
                c.setStrokeColor(colors.HexColor("#E5E5E5"))
                c.line(x, y, x, y + row_height)
            c.setStrokeColor(colors.black)

            y -= row_height
            row_idx += 1
            rows_on_page += 1

        # Línea vertical final
        c.setStrokeColor(colors.HexColor("#E5E5E5"))
        c.line(width - MARGIN, height - MARGIN - HEADER_HEIGHT, width - MARGIN, y + row_height)
        c.setStrokeColor(colors.black)

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
