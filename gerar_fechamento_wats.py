from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
import re

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

BASE_DIR = Path(__file__).resolve().parent
OUT_DIR = BASE_DIR / "saida"
OUT_DIR.mkdir(exist_ok=True)

XLSX_PATH = OUT_DIR / "fechamento_wats.xlsx"
PDF_PATH = OUT_DIR / "fechamento_wats.pdf"

CLIENTE = "WATS"
dados_fechamento = [
    (27102, "01/06/26", "MOLETOM LOGO BOTOM - VD. MARINA", 0, 4, 12, 14, 12, 4, 0, Decimal("14.00")),
    (27118, "01/06/26", "MOLETOM COLEGIAL - VD. MARINA", 0, 0, 10, 10, 10, 5, 0, Decimal("14.00")),
    (27345, "08/05/26", "CUECA BOXER - PRETO", 0, 40, 80, 60, 40, 0, 0, Decimal("4.00")),
    (27346, "08/05/26", "CUECA BOXER - BRANCO", 0, 20, 50, 50, 50, 0, 0, Decimal("4.00")),
    (27347, "08/05/26", "CUECA BOXER - CHUMBO", 0, 40, 65, 65, 55, 0, 0, Decimal("4.00")),
    ("V-27", "16/05/26", "SMALL BOX PRETO", 0, 0, 0, 8, 0, 0, 0, Decimal("5.00")),
    ("V-27", "16/05/26", "SPORTING PETROLEO", 0, 0, 0, 8, 0, 0, 0, Decimal("5.00")),
    ("V-27", "16/05/26", "WORLD BRANCO", 0, 0, 0, 8, 0, 0, 0, Decimal("5.00")),
    ("V-27", "19/05/26", "REGATA LISTRA COLORS - PRETO", 0, 0, 0, 9, 0, 0, 0, Decimal("5.00")),
    ("V-27", "16/05/26", "CAMISETA LOGO CENTRAL - PRETO/CHUMBO", 0, 0, 0, 8, 0, 0, 0, Decimal("5.00")),
    ("V-27", "16/05/26", "REGATA TAG DRY FIT - PRE/BRA", 0, 0, 0, 9, 0, 0, 0, Decimal("5.00")),
    ("V-27", "19/05/26", "REGATA LISTRA CLASSIC - PRE/BRA", 0, 0, 0, 9, 0, 0, 0, Decimal("5.00")),
    ("V-27", "19/05/26", "CAMISETA LISTRA CLASSIC - CAQUI/OFFWHITE", 0, 0, 0, 8, 0, 0, 0, Decimal("5.00")),
    ("V-27", "19/05/26", "CAMISETA LISTRA COLORS - PRE/AZUL", 0, 0, 0, 8, 0, 0, 0, Decimal("5.00")),
    ("V-27", "16/05/26", "CAMISETA LISTRA COLORS INFANTIL - PRE/AZU", 0, 0, 0, 9, 0, 0, 0, Decimal("5.00")),
    ("V-27", "19/05/26", "MANGA LONGA BEACH - PRETO", 0, 0, 0, 3, 0, 0, 0, Decimal("5.00")),
]

headers = [
    "Refer.", "Data", "Produto", "Cliente",
    "PP", "P", "M", "G", "GG", "XGG", "XG",
    "Qtde Peças", "Preço", "Total", "Status"
]


def format_brl(value: Decimal) -> str:
    return f"{value:.2f}".replace(".", ",")


def calc_line(pp: int, p: int, m: int, g: int, gg: int, xgg: int, xg: int, preco: Decimal):
    qtd = pp + p + m + g + gg + xgg + xg
    total = (Decimal(qtd) * preco).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return qtd, total


rows_out = []

# Processa dados do fechamento
for refer, data, produto, pp, p, m, g, gg, xgg, xg, preco in dados_fechamento:
    qtd, total = calc_line(pp, p, m, g, gg, xgg, xg, preco)
    rows_out.append([
        refer, data, produto, CLIENTE,
        pp if pp else "", p if p else "", m if m else "", g if g else "", gg if gg else "", xgg if xgg else "", xg if xg else "",
        qtd, format_brl(preco), format_brl(total), ""
    ])

# Calcula total geral de todos os itens (coluna Total, indice 13)
total_geral = Decimal("0.00")
for row in rows_out:
    valor_str = row[13].replace(",", ".")
    total_geral += Decimal(valor_str)

# Insere linha de total geral
rows_out.append([
    "", "", "TOTAL GERAL", CLIENTE,
    "", "", "", "", "", "", "",
    "", "", format_brl(total_geral), ""
])

# Gera XLSX
wb = Workbook()
ws = wb.active
ws.title = "Fechamento"

ws.append(headers)
for row in rows_out:
    ws.append(row)

header_fill = PatternFill("solid", fgColor="1F3864")
header_font = Font(color="FFFFFF", bold=True)
thin = Side(style="thin", color="CCCCCC")

for cell in ws[1]:
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=len(headers)):
    for cell in row:
        cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
        cell.alignment = Alignment(horizontal="center", vertical="center")

for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=3, max_col=3):
    for cell in row:
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

widths = {
    "A": 10, "B": 10, "C": 38, "D": 14,
    "E": 6, "F": 6, "G": 6, "H": 6, "I": 6, "J": 6, "K": 6,
    "L": 12, "M": 9, "N": 12, "O": 10
}
for col, width in widths.items():
    ws.column_dimensions[col].width = width

wb.save(XLSX_PATH)

# Gera PDF
doc = SimpleDocTemplate(
    str(PDF_PATH),
    pagesize=landscape(A4),
    leftMargin=0.8 * cm,
    rightMargin=0.8 * cm,
    topMargin=1.0 * cm,
    bottomMargin=1.0 * cm,
)

styles = getSampleStyleSheet()
title_style = ParagraphStyle("title", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=12)

elements = [
    Paragraph("Fechamento Mensal - WATS", title_style),
    Spacer(1, 0.25 * cm),
]

table_data = [headers] + rows_out
col_widths = [
    1.5 * cm, 1.8 * cm, 6.0 * cm, 2.5 * cm,
    0.8 * cm, 0.8 * cm, 0.8 * cm, 0.8 * cm, 0.8 * cm, 1.0 * cm, 1.0 * cm,
    1.8 * cm, 1.2 * cm, 1.6 * cm, 1.2 * cm,
]

table = Table(table_data, colWidths=col_widths, repeatRows=1)
table_style = [
    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F3864")),
    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
    ("FONTSIZE", (0, 0), (-1, 0), 8),
    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
    ("ALIGN", (2, 1), (2, -1), "LEFT"),
    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#B7B7B7")),
    ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
    ("FONTSIZE", (0, 1), (-1, -1), 7),
]
for i in range(1, len(table_data)):
    if i % 2 == 0:
        table_style.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor("#EDF2FB")))

table.setStyle(TableStyle(table_style))
elements.append(table)
doc.build(elements)

print(f"XLSX: {XLSX_PATH}")
print(f"PDF: {PDF_PATH}")
print(f"Linhas: {len(rows_out)}")
