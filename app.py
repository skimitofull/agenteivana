import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from PyPDF2 import PdfMerger
import base64

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
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=(LETTER_WIDTH, LETTER_HEIGHT))

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

    header_color = colors.HexColor('#000000')
    header_text_color = colors.HexColor('#FFFFFF')
    alternate_row_color = colors.HexColor('#F0F0F0')
    border_color = colors.HexColor('#E5E5E5')

    y = LETTER_HEIGHT - MARGIN
    header_y = y - HEADER_HEIGHT
    c.setFillColor(header_color)
    c.rect(MARGIN, header_y, width, HEADER_HEIGHT, fill=1)

    c.setFillColor(header_text_color)
    c.setFont('Helvetica-Bold', 10)
    for i, header in enumerate(headers):
        x = x_positions[i]
        text_width = c.stringWidth(header, 'Helvetica-Bold', 10)
        text_x = x + (col_widths[i] - text_width) // 2
        c.drawString(text_x, header_y + 7, header)

    current_y = header_y - BASE_ROW_HEIGHT

    c.setFont('Helvetica', 9)
    for idx in range(start_idx, min(end_idx, len(df))):
        row = df.iloc[idx]
        concept_parts = split_concept(row['Concepto'])
        row_height = calculate_row_height(concept_parts)

        if current_y - row_height < MARGIN:
            break

        if idx % 2 == 0:
            c.setFillColor(alternate_row_color)
            c.rect(MARGIN, current_y - row_height, width, row_height, fill=1, stroke=0)

        c.setFillColor(colors.black)
        for i, col in enumerate(headers):
            x = x_positions[i]
            value = str(row[col]) if pd.notnull(row[col]) else ''

            if col == 'Concepto':
                for line_idx, line in enumerate(concept_parts):
                    line_y = current_y - (line_idx * LINE_SPACING) - LINE_SPACING
                    c.drawString(x + 5, line_y, line)
            elif col in ['Depósito', 'Retiro', 'Saldo']:
                text_width = c.stringWidth(value, 'Helvetica', 9)
                c.drawString(x + col_widths[i] - text_width - 5, current_y - LINE_SPACING, value)
            else:
                c.drawString(x + 5, current_y - LINE_SPACING, value)

        c.setStrokeColor(border_color)
        for x in x_positions:
            c.line(x, header_y, x, current_y - row_height)
        c.line(MARGIN + width, header_y, MARGIN + width, current_y - row_height)

        current_y -= row_height + 5

    c.setFont('Helvetica', 9)
    page_text = f"Página {page_number}"
    text_width = c.stringWidth(page_text, 'Helvetica', 9)
    c.drawString(LETTER_WIDTH - MARGIN - text_width, MARGIN, page_text)

    c.save()
    buffer.seek(0)
    return buffer, end_idx

def create_pdf(df):
    total_rows = len(df)
    buffers = []
    current_idx = 0
    page_num = 1

    while current_idx < total_rows:
        buffer, next_idx = create_page(df, current_idx, current_idx + ROWS_PER_PAGE, page_num)
        buffers.append(buffer)
        if next_idx == current_idx:
            break
        current_idx = next_idx
        page_num += 1

    output = BytesIO()
    merger = PdfMerger()
    for buffer in buffers:
        merger.append(buffer)
    merger.write(output)
    merger.close()
    output.seek(0)
    return output

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
