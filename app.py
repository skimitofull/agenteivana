import streamlit as st
import pandas as pd
import numpy as np
from PIL import Image, ImageDraw, ImageFont
from io import BytesIO

st.set_page_config(page_title="Agente Modificador de Estados de Cuenta", layout="wide")

def clean_amount(amount):
    """Limpia y formatea montos al estilo del banco"""
    if pd.isnull(amount) or str(amount).strip() == '' or str(amount).lower() == 'nan':
        return ''
    try:
        if isinstance(amount, str):
            amount = float(str(amount).replace('$', '').replace(',', ''))
        return '${:,.2f}'.format(amount) if amount != 0 else ''
    except:
        return ''

def clean_date(date):
    """Limpia y formatea fechas al estilo del banco"""
    if pd.isnull(date):
        return ''
    try:
        if isinstance(date, str) and ('NOV' in date.upper() or 'DIC' in date.upper()):
            return date.upper().strip()
        return pd.to_datetime(date).strftime('%d %b').upper()
    except:
        return str(date).upper()

def create_table_image(df):
    """Crea la imagen de la tabla con el formato exacto del banco"""
    # Configuración A4 vertical (210mm × 297mm)
    width = 1240  # A4 a 150 DPI en vertical
    row_height = 25
    header_height = 35
    margin = 30

    # Configuración de columnas
    headers = ['Fecha', 'Concepto', 'Origen / Referencia', 'Depósito', 'Retiro', 'Saldo']
    col_widths = [100, 450, 220, 140, 140, 160]  # Proporciones ajustadas al formato vertical
    x_positions = np.cumsum([margin] + col_widths[:-1])

    # Calcular altura total
    total_rows = len(df) + 1
    total_height = (total_rows * row_height) + header_height + (2 * margin)

    # Crear imagen
    img = Image.new('RGB', (width, total_height), 'white')
    draw = ImageDraw.Draw(img)

    # Configurar fuentes
    try:
        font = ImageFont.truetype("Arial", 14)
        font_bold = ImageFont.truetype("Arial Bold", 14)
    except:
        font = ImageFont.load_default()
        font_bold = ImageFont.load_default()

    # Colores exactos del banco
    header_color = '#E6E6E6'
    alternate_row_color = '#F2F2F2'
    border_color = '#D9D9D9'

    # Dibujar encabezados
    y = margin
    for i, header in enumerate(headers):
        x = x_positions[i]
        # Fondo del encabezado
        draw.rectangle([x, y, x + col_widths[i], y + header_height],
                      fill=header_color, outline=border_color)
        # Texto centrado en el encabezado
        text_width = draw.textlength(header, font=font_bold)
        text_x = x + (col_widths[i] - text_width) // 2
        draw.text((text_x, y + 10), header, fill='black', font=font_bold)

    # Dibujar filas
    for idx, row in df.iterrows():
        y = margin + header_height + (idx * row_height)

        # Fondo alternado
        if idx % 2 == 0:
            draw.rectangle([margin, y, width - margin, y + row_height],
                         fill=alternate_row_color)

        # Dibujar bordes y datos
        for i, col in enumerate(headers):
            x = x_positions[i]
            value = str(row[col]) if col in row and pd.notnull(row[col]) else ''

            # Alineación según el tipo de dato
            if col in ['Depósito', 'Retiro', 'Saldo']:
                # Montos alineados a la derecha
                text_width = draw.textlength(value, font=font)
                draw.text((x + col_widths[i] - text_width - 5, y + 5),
                         value, fill='black', font=font)
            else:
                # Texto alineado a la izquierda
                draw.text((x + 5, y + 5), value, fill='black', font=font)

            # Líneas verticales
            draw.line([(x, y), (x, y + row_height)], fill=border_color)

        # Última línea vertical
        x_last = x_positions[-1] + col_widths[-1]
        draw.line([(x_last, y), (x_last, y + row_height)], fill=border_color)

        # Línea horizontal
        draw.line([(margin, y + row_height),
                   (width - margin, y + row_height)], fill=border_color)

    return img

def main():
    st.title("Agente Modificador de Estados de Cuenta")
    st.write("Sube el archivo Excel con los nuevos movimientos para generar la tabla en formato banco")

    excel_file = st.file_uploader("Excel con movimientos", type=['xlsx'])

    if excel_file:
        try:
            # Leer y limpiar datos
            df = pd.read_excel(excel_file)
            df = df.dropna(how='all')

            # Limpiar y formatear datos
            df['Fecha'] = df['Fecha'].apply(clean_date)
            df['Depósito'] = df['Depósito'].apply(clean_amount)
            df['Retiro'] = df['Retiro'].apply(clean_amount)
            df['Saldo'] = df['Saldo'].apply(clean_amount)

            # Mostrar preview
            st.write("Vista previa de los datos procesados:")
            st.dataframe(df)

            # Crear imagen y PDF
            img = create_table_image(df)

            # Convertir a PDF
            pdf_bytes = BytesIO()
            img_rgb = img.convert('RGB')
            img_rgb.save(pdf_bytes, format='PDF')
            pdf_bytes.seek(0)

            # Botón de descarga
            st.success("¡PDF generado exitosamente!")
            st.download_button(
                "📄 Descargar PDF",
                pdf_bytes.getvalue(),
                "estado_cuenta_modificado.pdf",
                "application/pdf"
            )

        except Exception as e:
            st.error(f"Error al procesar el archivo: {str(e)}")
            st.write("Por favor, verifica que el archivo Excel tenga el formato correcto.")

if __name__ == "__main__":
    main()
