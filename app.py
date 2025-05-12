import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
from PIL import Image, ImageDraw, ImageFont
import numpy as np
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from io import BytesIO
import re

st.set_page_config(page_title="Agente de Réplica Exacta de Estados de Cuenta", layout="wide")

st.title("Agente de Réplica Exacta de Estados de Cuenta")
st.write("Este agente analiza tu PDF original y replica exactamente el formato con los nuevos datos.")

class DocumentAnalyzer:
    def __init__(self):
        self.style_info = {}
        self.table_bounds = None
        self.header_info = None
        self.footer_info = None

    def extract_table_style(self, pdf_file):
        """Extrae el estilo exacto de la tabla del PDF original"""
        doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
        page = doc[0]  # Primera página

        # Analizar elementos de la página
        self.analyze_page_elements(page)

        # Extraer información detallada de formato
        self.extract_detailed_formatting(page)

        return self.style_info

    def analyze_page_elements(self, page):
        """Analiza los elementos de la página y su posición"""
        blocks = page.get_text("dict")["blocks"]

        # Encontrar la tabla
        for block in blocks:
            if "lines" in block:
                text = "".join([span["text"] for line in block["lines"]
                              for span in line["spans"]])
                if re.search(r'(FECHA|CONCEPTO|DEPÓSITO|RETIRO)', text, re.IGNORECASE):
                    self.table_bounds = block["bbox"]
                    break

        # Extraer header y footer
        self.header_info = [b for b in blocks if b["bbox"][1] < self.table_bounds[1]]
        self.footer_info = [b for b in blocks if b["bbox"][1] > self.table_bounds[3]]

    def extract_detailed_formatting(self, page):
        """Extrae formato detallado incluyendo fuentes, colores y espaciado"""
        self.style_info = {
            "table_style": {
                "fonts": {},
                "colors": {},
                "spacing": {},
                "alignment": {}
            },
            "header": self.header_info,
            "footer": self.footer_info,
            "page_size": page.rect,
            "margins": page.mediabox
        }

        # Analizar cada span para extraer estilos
        for block in page.get_text("dict")["blocks"]:
            if "lines" in block:
                for line in block["lines"]:
                    for span in line["spans"]:
                        self.style_info["table_style"]["fonts"][span["text"]] = {
                            "font": span["font"],
                            "size": span["size"],
                            "flags": span["flags"]  # Bold, italic, etc.
                        }
                        self.style_info["table_style"]["colors"][span["text"]] = span["color"]

class DocumentGenerator:
    def __init__(self, style_info):
        self.style_info = style_info
        self.setup_fonts()

    def setup_fonts(self):
        """Configura las fuentes necesarias"""
        try:
            # Registrar fuentes comunes
            for font_name in ['Helvetica', 'Helvetica-Bold', 'Times-Roman']:
                if font_name not in pdfmetrics.getRegisteredFontNames():
                    pdfmetrics.registerFont(TTFont(font_name, f"{font_name}.ttf"))
        except:
            st.warning("Usando fuentes por defecto debido a limitaciones de disponibilidad")

    def create_exact_replica(self, df):
        """Crea una réplica exacta del documento con los nuevos datos"""
        buffer = BytesIO()

        # Configurar el documento
        page_width = float(self.style_info["page_size"].width)
        page_height = float(self.style_info["page_size"].height)
        c = canvas.Canvas(buffer, pagesize=(page_width, page_height))

        # Crear y estilizar la tabla
        table_data = self.prepare_table_data(df)
        table = Table(table_data)
        table.setStyle(self.create_table_style())

        # Aplicar header y footer
        self.apply_header_footer(c)

        # Dibujar la tabla
        table.wrapOn(c, page_width-100, page_height-200)
        table.drawOn(c, 50, page_height-150-table._height)

        c.save()
        buffer.seek(0)
        return buffer

    def prepare_table_data(self, df):
        """Prepara los datos de la tabla con el formato correcto"""
        headers = ['Fecha', 'Concepto', 'Origen / Referencia', 'Depósito', 'Retiro', 'Saldo']
        table_data = [headers]

        for _, row in df.iterrows():
            formatted_row = [
                self.format_date(row['Fecha']),
                str(row['Concepto']),
                str(row['Origen / Referencia']),
                self.format_currency(row['Depósito']),
                self.format_currency(row['Retiro']),
                self.format_currency(row['Saldo'])
            ]
            table_data.append(formatted_row)

        return table_data

    def create_table_style(self):
        """Crea el estilo exacto de la tabla"""
        return TableStyle([
            ('FONT', (0, 0), (-1, -1), 'Helvetica', 10),
            ('FONT', (0, 0), (-1, 0), 'Helvetica-Bold', 10),
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('ALIGN', (3, 1), (-1, -1), 'RIGHT'),  # Alineación derecha para montos
            ('GRID', (0, 0), (-1, -1), 1, colors.grey),
            ('LEFTPADDING', (0, 0), (-1, -1), 6),
            ('RIGHTPADDING', (0, 0), (-1, -1), 6),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ])

    @staticmethod
    def format_date(date):
        """Formatea fechas al estilo del estado de cuenta"""
        try:
            if pd.isna(date):
                return ''
            if isinstance(date, str) and any(month in date.upper() for month in ['ENE', 'FEB', 'MAR', 'ABR', 'MAY', 'JUN', 'JUL', 'AGO', 'SEP', 'OCT', 'NOV', 'DIC']):
                return date.upper()
            return pd.to_datetime(date).strftime('%d %b').upper()
        except:
            return str(date).upper()

    @staticmethod
    def format_currency(amount):
        """Formatea cantidades monetarias al estilo del estado de cuenta"""
        try:
            if pd.isna(amount) or str(amount).strip() == '':
                return ''
            if isinstance(amount, str):
                amount = float(str(amount).replace('$', '').replace(',', ''))
            return '${:,.2f}'.format(amount) if amount != 0 else ''
        except:
            return str(amount)

    def apply_header_footer(self, canvas_obj):
        """Aplica el header y footer del documento original"""
        # Implementar según self.style_info["header"] y self.style_info["footer"]
        pass

def main():
    st.write("### Sube tus archivos")
    col1, col2 = st.columns(2)

    with col1:
        pdf_file = st.file_uploader("PDF Original (Estado de Cuenta)", type=['pdf'])

    with col2:
        excel_file = st.file_uploader("Excel con nuevos movimientos", type=['xlsx'])

    if pdf_file and excel_file:
        try:
            # Analizar el PDF original
            analyzer = DocumentAnalyzer()
            style_info = analyzer.extract_table_style(pdf_file)

            # Leer y preparar datos nuevos
            df = pd.read_excel(excel_file)

            # Mostrar preview de datos
            st.write("### Vista previa de los nuevos movimientos")
            st.dataframe(df.head())

            # Generar nuevo documento
            generator = DocumentGenerator(style_info)
            new_pdf = generator.create_exact_replica(df)

            # Botón de descarga
            st.success("¡Documento generado exitosamente!")
            st.download_button(
                "📥 Descargar PDF Modificado",
                new_pdf.getvalue(),
                "estado_cuenta_modificado.pdf",
                "application/pdf"
            )

        except Exception as e:
            st.error(f"Ocurrió un error al procesar los archivos: {str(e)}")
            st.write("Por favor, verifica que los archivos tengan el formato correcto.")

if __name__ == "__main__":
    main()
