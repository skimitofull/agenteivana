import streamlit as st
import pandas as pd
import numpy as np
from PIL import Image, ImageDraw, ImageFont
from io import BytesIO
import math

# --- CONSTANTES GLOBALES ---
LETTER_WIDTH = 800
LETTER_HEIGHT = 1000
MARGIN = 30
BASE_ROW_HEIGHT = 25  # Aumentado para más espacio
HEADER_HEIGHT = 25
LINE_SPACING = 15  # Aumentado para más espacio entre líneas
ROWS_PER_PAGE = 20  # Reducido para evitar sobrecarga

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
        # Dividir conceptos largos en múltiples líneas
        words = concept.split()
        current_line = []
        for word in words:
            current_line.append(word)
            if len(' '.join(current_line)) > 40:  # Ajustar según necesidad
                parts.append(' '.join(current_line[:-1]))
                current_line = [word]
        if current_line:
            parts.append(' '.join(current_line))
    return parts

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

    # Configuración de columnas y estilos
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
    header_color = '#000000'
    header_text_color = '#FFFFFF'
    alternate_row_color = '#F0F0F0'
    border_color = '#E5E5E5'

    # Dibujar encabezados
    y = MARGIN
    for i, header in enumerate(headers):
        x = x_positions[i]
        draw.rectangle([x, y, x + col_widths[i], y + HEADER_HEIGHT],
                      fill=header_color, outline=border_color)
        text_width = draw.textlength(header, font=font_bold)
        text_x = x + (col_widths[i] - text_width) // 2
        draw.text((text_x, y + 5), header, fill=header_text_color, font=font_bold)

    current_y = MARGIN + HEADER_HEIGHT

    # Verificar si hay espacio suficiente para cada fila
    for idx in range(start_idx, min(end_idx, len(df))):
        row = df.iloc[idx]
        concept_parts = split_concept(row['Concepto'])
        row_height = calculate_row_height(concept_parts)

        # Si la siguiente fila excedería el límite de la página, terminar esta página
        if current_y + row_height > LETTER_HEIGHT - MARGIN:
            break

        # Dibujar fila
        if idx % 2 == 0:
            draw.rectangle([MARGIN, current_y, LETTER_WIDTH - MARGIN, current_y + row_height],
                         fill=alternate_row_color)

        # Dibujar contenido
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

        # Línea separadora
        draw.line([(MARGIN, current_y + row_height),
                   (LETTER_WIDTH - MARGIN, current_y + row_height)], fill=border_color)

        current_y += row_height + 5  # Añadir espacio extra entre filas

    # Número de página
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
        if next_idx == current_idx:  # Si no se pudo agregar más filas
            break
        current_idx = next_idx
        page_num += 1

    pdf_buffer = BytesIO()
    if pages:
        pages[0].convert('RGB').save(pdf_buffer, format='PDF', save_all=True,
                                   append_images=[page.convert('RGB') for page in pages[1:]])
        pdf_buffer.seek(0)
    return pdf_buffer, pages

# --- INTERFAZ STREAMLIT ---
st.set_page_config(page_title="Generador de Estado de Cuenta Scotiabank", layout="wide")

st.title("Generador de Estado de Cuenta Scotiabank")
st.write("Sube tu archivo Excel con movimientos y descarga el PDF generado.")

uploaded_file = st.file_uploader("Selecciona tu archivo Excel (.xlsx)", type=['xlsx'])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl")

        # Mostrar vista previa de los datos
        st.write("Vista previa de los movimientos:")
        st.dataframe(df)

        # Procesar los datos
        df['Fecha'] = df['Fecha'].apply(clean_date)
        df['Depósito'] = df['Depósito'].apply(clean_amount)
        df['Retiro'] = df['Retiro'].apply(clean_amount)
        df['Saldo'] = df['Saldo'].apply(clean_amount)

        # Generar vista previa y PDF
        if st.button("Generar Vista Previa y PDF"):
            with st.spinner('Generando documento...'):
                pdf_buffer, preview_pages = create_pdf(df)
                st.success('¡Documento generado correctamente!')

                # Mostrar vista previa de todas las páginas
                st.subheader("Vista previa de todas las páginas:")
                for i, page in enumerate(preview_pages):
                    st.image(page, caption=f'Página {i+1}', use_column_width=True)
                    st.markdown("---")

                # Botón de descarga
                st.download_button(
                    label="Descargar PDF",
                    data=pdf_buffer,
                    file_name="estado_cuenta_modificado.pdf",
                    mime="application/pdf"
                )

    except Exception as e:
        st.error(f"Error al procesar el archivo: {str(e)}")
else:
    st.info("Por favor, sube un archivo Excel para comenzar.")
