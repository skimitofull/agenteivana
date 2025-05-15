import streamlit as st
import pandas as pd
import numpy as np
from PIL import Image, ImageDraw, ImageFont
from io import BytesIO
import math
import base64
from PyPDF2 import PdfReader
import tempfile

# --- CONSTANTES GLOBALES ---
LETTER_WIDTH = 800
LETTER_HEIGHT = 1000
MARGIN = 30
BASE_ROW_HEIGHT = 25
HEADER_HEIGHT = 25
LINE_SPACING = 15
ROWS_PER_PAGE = 20

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

def calculate_row_height(concept_parts):
    return max(len(concept_parts) * LINE_SPACING + 10, BASE_ROW_HEIGHT)

def create_page(df, start_idx, end_idx, page_number):
    img = Image.new('RGB', (LETTER_WIDTH, LETTER_HEIGHT), 'white')
    draw = ImageDraw.Draw(img)

    try:
        font = ImageFont.truetype("Arial", 9)
        font_bold = ImageFont.truetype("Arial Bold", 10)
    except:
        font = ImageFont.load_default()
        font_bold = ImageFont.load_default()

    width = LETTER_WIDTH - (2 * MARGIN)
    headers = ['Fecha', 'Concepto', 'Origen / Referencia', 'Depósito', 'Retiro', 'Saldo']
    col_widths = [
        int(width * 0.08),
        int(width * 0.40),
        int(width * 0.20),
        int(width * 0.11),
        int(width * 0.11),
        int(width * 0.10)
    ]
    col_widths[-1] = width - sum(col_widths[:-1])
    x_positions = np.cumsum([MARGIN] + col_widths[:-1])

    header_color = '#000000'
    header_text_color = '#FFFFFF'
    alternate_row_color = '#F0F0F0'
    border_color = '#E5E5E5'

    y = MARGIN
    for i, header in enumerate(headers):
        x = x_positions[i]
        draw.rectangle([x, y, x + col_widths[i], y + HEADER_HEIGHT],
                      fill=header_color, outline=border_color)
        text_width = draw.textlength(header, font=font_bold)
        text_x = x + (col_widths[i] - text_width) // 2
        draw.text((text_x, y + 5), header, fill=header_text_color, font=font_bold)

    current_y = MARGIN + HEADER_HEIGHT

    for idx in range(start_idx, min(end_idx, len(df))):
        row = df.iloc[idx]
        concept_parts = split_concept(row['Concepto'])
        row_height = calculate_row_height(concept_parts)

        if current_y + row_height > LETTER_HEIGHT - MARGIN:
            break

        if idx % 2 == 0:
            draw.rectangle([MARGIN, current_y, LETTER_WIDTH - MARGIN, current_y + row_height],
                         fill=alternate_row_color)

        for i, col in enumerate(headers):
            x = x_positions[i]
            value = str(row[col]) if pd.notnull(row[col]) else ''

            if col == 'Concepto':
                for line_idx, line in enumerate(concept_parts):
                    line_y = current_y + (line_idx * LINE_SPACING)
                    draw.text((x + 5, line_y + 2), line, fill='black', font=font)
            elif col in ['Depósito', 'Retiro', 'Saldo']:
                text_width = draw.textlength(value, font=font)
                draw.text((x + col_widths[i] - text_width - 5, current_y + 2),
                         value, fill='black', font=font)
            else:
                draw.text((x + 5, current_y + 2), value, fill='black', font=font)

            # Dibujar líneas verticales
            if i > 0:
                draw.line([(x, current_y), (x, current_y + row_height)], fill=border_color)

        current_y += row_height + 5

    # Dibujar línea vertical final
    draw.line([(LETTER_WIDTH - MARGIN, MARGIN + HEADER_HEIGHT), (LETTER_WIDTH - MARGIN, current_y)], fill=border_color)

    page_text = f"Página {page_number}"
    text_width = draw.textlength(page_text, font=font)
    draw.text((LETTER_WIDTH - MARGIN - text_width, LETTER_HEIGHT - MARGIN),
              page_text, fill='black', font=font)

    return img, end_idx

def create_pdf(df):
    total_rows = len(df)
    pages = []
    current_idx = 0
    page_num = 1

    while current_idx < total_rows:
        page, next_idx = create_page(df, current_idx, current_idx + ROWS_PER_PAGE, page_num)
        pages.append(page)
        if next_idx == current_idx:
            break
        current_idx = next_idx
        page_num += 1

    pdf_buffer = BytesIO()
    if pages:
        pages[0].convert('RGB').save(pdf_buffer, format='PDF', save_all=True,
                                   append_images=[page.convert('RGB') for page in pages[1:]])
        pdf_buffer.seek(0)
    return pdf_buffer

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
