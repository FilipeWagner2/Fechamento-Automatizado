from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
import os

# ─────────────────────────────────────────
# DADOS DA FICHA (extraídos da foto)
# ─────────────────────────────────────────
TITULO = "Fechamento Mensal"
DATA_PLANILHA = "08/05/2026"

# Colunas: Refer | Data | Produto | Cliente | PP | P | M | G | GG | XGG | XG | Qtde Peças | Preço | Total
LINHAS = [
    ["3154", "13/04/26", "US - Camiseta",  "Ultra Silk", "",  "1",  "3",  "",  "1",  "",  "11", "16", "3,50", "56,00"],
    ["3154", "13/04/26", "US - Camiseta",  "Ultra Silk", "",  "2",  "1",  "1", "",   "1", "7",  "12", "3,50", "42,00"],
    ["3154", "13/04/26", "US - Camiseta",  "Ultra Silk", "1", "10", "7",  "1", "",   "1", "4",  "24", "3,50", "84,00"],
    ["3154", "13/04/26", "US - Camiseta",  "Ultra Silk", "",  "11", "3",  "",  "1",  "",  "2",  "17", "3,50", "59,50"],
    ["3154", "13/04/26", "US - Bandeira",  "Ultra Silk", "",  "",   "4",  "",  "",   "",  "",   "4",  "3,50", "14,00"],
]

CABECALHO = ["Refer.", "Data", "Produto", "Cliente", "PP", "P", "M", "G", "GG", "XGG", "XG", "Qtde Peças", "Preço", "Total"]

# ─────────────────────────────────────────
# GERAÇÃO DO PDF
# ─────────────────────────────────────────
OUTPUT_PATH = os.path.join(os.path.dirname(__file__), "saida", "fechamento_3154.pdf")
os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)

doc = SimpleDocTemplate(
    OUTPUT_PATH,
    pagesize=landscape(A4),
    leftMargin=1*cm,
    rightMargin=1*cm,
    topMargin=1.5*cm,
    bottomMargin=1.5*cm,
)

styles = getSampleStyleSheet()
title_style = ParagraphStyle("title", fontSize=13, fontName="Helvetica-Bold", alignment=TA_LEFT)
date_style  = ParagraphStyle("date",  fontSize=11, fontName="Helvetica",      alignment=TA_LEFT)

elements = []

# Cabeçalho da planilha
elements.append(Paragraph(TITULO, title_style))
elements.append(Spacer(1, 0.2*cm))
elements.append(Paragraph(f"Data: {DATA_PLANILHA}", date_style))
elements.append(Spacer(1, 0.5*cm))

# Tabela
table_data = [CABECALHO] + LINHAS

# Larguras das colunas (total ~27cm em landscape A4)
col_widths = [
    1.5*cm,  # Refer.
    2.0*cm,  # Data
    4.5*cm,  # Produto
    2.5*cm,  # Cliente
    1.0*cm,  # PP
    1.0*cm,  # P
    1.0*cm,  # M
    1.0*cm,  # G
    1.0*cm,  # GG
    1.2*cm,  # XGG
    1.2*cm,  # XG
    2.0*cm,  # Qtde Peças
    1.5*cm,  # Preço
    2.0*cm,  # Total
]

table = Table(table_data, colWidths=col_widths, repeatRows=1)
table.setStyle(TableStyle([
    # Cabeçalho
    ("BACKGROUND",    (0, 0), (-1, 0),  colors.HexColor("#1F3864")),
    ("TEXTCOLOR",     (0, 0), (-1, 0),  colors.white),
    ("FONTNAME",      (0, 0), (-1, 0),  "Helvetica-Bold"),
    ("FONTSIZE",      (0, 0), (-1, 0),  8),
    ("ALIGN",         (0, 0), (-1, 0),  "CENTER"),
    ("VALIGN",        (0, 0), (-1, 0),  "MIDDLE"),
    ("ROWBACKGROUND", (0, 0), (-1, 0),  colors.HexColor("#1F3864")),

    # Dados
    ("FONTNAME",      (0, 1), (-1, -1), "Helvetica"),
    ("FONTSIZE",      (0, 1), (-1, -1), 8),
    ("ALIGN",         (0, 1), (-1, -1), "CENTER"),
    ("VALIGN",        (0, 1), (-1, -1), "MIDDLE"),
    ("ROWBACKGROUND", (0, 1), (-1, -1), colors.white),

    # Zebra
    *[("BACKGROUND", (0, i), (-1, i), colors.HexColor("#EBF0FA")) for i in range(2, len(table_data), 2)],

    # Grid
    ("GRID",          (0, 0), (-1, -1), 0.5, colors.HexColor("#AAAAAA")),
    ("BOX",           (0, 0), (-1, -1), 1,   colors.HexColor("#1F3864")),

    # Padding
    ("TOPPADDING",    (0, 0), (-1, -1), 4),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ("LEFTPADDING",   (0, 0), (-1, -1), 4),
    ("RIGHTPADDING",  (0, 0), (-1, -1), 4),
]))

elements.append(table)
doc.build(elements)
print(f"PDF gerado em: {OUTPUT_PATH}")
