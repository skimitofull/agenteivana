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
LETTER_WIDTH, LETTER_HEIGHT = letter
MARGIN = 30
BASE_ROW_HEIGHT = 25
HEADER_HEIGHT = 25
LINE_SPACING = 12
ROWS_PER_PAGE = 20
FONT_SIZE = 8.24
HEADER_FONT_SIZE = 8.24

def split_concept(concept):
    if pd.isnull(concept):
        return ['']

    concept = str(concept).strip().upper()
    parts = []

    if "TRANSF.INTERB SPEI" in concept:
        lines = concept.split('\n')
        parts = [line.strip().upper() for line in lines if line.strip()]
    else:
        words = concept.split()
        current_line = []
        for word in words:
            current_line.append(word)
            if len(' '.join(current_line)) > 40:
                parts.append(' '.join(current_line[:-1]).upper())
                current_line = [word]
        if current_line:
            parts.append(' '.join(current_line).upper())

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
    return len(concept_parts) * LINE_SPACING

def create_page(df, start_idx, end_idx, page_number):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)

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

    # Colores
    header_color = colors.HexColor('#000000')
    header_text_color = colors.HexColor('#FFFFFF')
    alternate_row_color = colors.HexColor('#F0F0F0')
    border_color = colors.HexColor('#E5E5E5')

    y = LETTER_HEIGHT - MARGIN
    header_y = y - HEADER_HEIGHT
    c.setFillColor(header_color)
    c.rect(MARGIN, header_y, width, HEADER_HEIGHT, fill=1)

    # Texto del encabezado
    c.setFillColor(header_text_color)
    c.setFont('Helvetica-Bold', HEADER_FONT_SIZE)
    for i, header in enumerate(headers):
        x = x_positions[i]
        text_width = c.stringWidth(header.upper(), 'Helvetica-Bold', HEADER_FONT_SIZE)
        text_x = x + (col_widths[i] - text_width) // 2
        c.drawString(text_x, header_y + 3, header.upper())

    current_y = header_y - BASE_ROW_HEIGHT

    # Contenido
    c.setFont('Helvetica', FONT_SIZE)
    prev_date = None

    for idx in range(start_idx, min(end_idx, len(df))):
        row = df.iloc[idx]
        concept_parts = split_concept(row['Concepto'])
        row_height = calculate_row_height(concept_parts)

        # Ajustar el espaciado solo cuando cambia la fecha
        current_date = row['Fecha']
        if prev_date is not None and current_date != prev_date:
            current_y -= 5  # Pequeño espacio adicional entre fechas diferentes
        prev_date = current_date

        if current_y - row_height < MARGIN:
            break

        # Fondo alternado
        if idx % 2 == 0:
            c.setFillColor(alternate_row_color)
            c.rect(MARGIN, current_y - row_height, width, row_height, fill=1, stroke=0)

        c.setFillColor(colors.black)
        for i, col in enumerate(headers):
            x = x_positions[i]
            value = str(row[col]) if pd.notnull(row[col]) else ''

            if col == 'Concepto':
                for j, line in enumerate(concept_parts):
                    text_y = current_y - ((j + 1) * LINE_SPACING) + (LINE_SPACING / 2)
                    c.drawString(x + 5, text_y, line)
            elif col in ['Depósito', 'Retiro', 'Saldo']:
                text_width = c.stringWidth(value, 'Helvetica', FONT_SIZE)
                c.drawString(x + col_widths[i] - text_width - 5, current_y - LINE_SPACING + (LINE_SPACING / 2), value)
            else:
                c.drawString(x + 5, current_y - LINE_SPACING + (LINE_SPACING / 2), value)

        # Líneas verticales
        c.setStrokeColor(border_color)
        for x in x_positions:
            c.line(x, header_y, x, current_y - row_height)
        c.line(MARGIN + width, header_y, MARGIN + width, current_y - row_height)

        current_y -= row_height

    # Número de página
    c.setFont('Helvetica', FONT_SIZE)
    page_text = f"Página {page_number}"
    text_width = c.stringWidth(page_text, 'Helvetica', FONT_SIZE)
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

        # Eliminar la primera columna si está vacía
        if df.iloc[:, 0].isnull().all():
            df = df.iloc[:, 1:]

        # Eliminar la primera fila si está vacía
        if df.iloc[0].isnull().all():
            df = df.iloc[1:]
            df = df.reset_index(drop=True)

        # Limpiar filas completamente vacías
        df = df.dropna(how='all')

        st.write("Vista previa de los movimientos:")
        st.dataframe(df)

        # Aplicar limpieza de datos
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
