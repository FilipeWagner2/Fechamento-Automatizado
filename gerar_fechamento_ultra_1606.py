from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

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

XLSX_PATH = OUT_DIR / "fechamento_ultra_silk_1606.xlsx"
PDF_PATH = OUT_DIR / "fechamento_ultra_silk_1606.pdf"

CLIENTE = "Ultra Silk"
DATA = "16/06/26"

P350 = Decimal("3.50")
P380 = Decimal("3.80")
P600 = Decimal("6.00")
P700 = Decimal("7.00")

# Ordem solicitada: 3315 (direita para esquerda) ate 3380.
dados = [
    (3315, DATA, "Camiseta", 0, 15, 20, 15, 10, 0, 0, P350),
    (3315, DATA, "Camiseta", 0, 15, 20, 15, 10, 0, 0, P350),
    (3315, DATA, "Camiseta", 0, 20, 20, 30, 15, 0, 0, P350),

    (3201, DATA, "Camiseta", 4, 11, 15, 16, 12, 0, 0, P350),
    (3201, DATA, "Camiseta", 3, 19, 46, 24, 9, 8, 1, P350),
    (3201, DATA, "Camiseta", 0, 0, 0, 0, 0, 0, 1, P350),

    (3130, DATA, "Baby Look", 0, 0, 1, 3, 0, 0, 0, P350),
    (3130, DATA, "Camiseta", 1, 0, 0, 0, 0, 0, 0, P350),
    (3130, DATA, "Regata Nadador FEM", 0, 0, 0, 0, 0, 0, 1, P350),

    (3131, DATA, "Baby Look", 0, 0, 2, 6, 4, 0, 0, P350),
    (3131, DATA, "Camiseta", 0, 0, 1, 1, 1, 0, 1, P350),
    (3131, DATA, "Camiseta FEM", 1, 2, 2, 0, 0, 0, 0, P350),
    (3131, DATA, "Regata Nadador FEM", 0, 0, 1, 0, 0, 0, 1, P350),

    (3132, DATA, "Regata Baby Look", 0, 0, 0, 0, 1, 0, 0, P350),
    (3132, DATA, "Camiseta", 1, 0, 0, 0, 0, 0, 0, P350),
    (3132, DATA, "Regata Baby Look", 1, 0, 0, 0, 0, 0, 1, P350),
    (3132, DATA, "Camiseta", 0, 1, 0, 0, 0, 0, 0, P350),

    (3133, DATA, "Baby Look", 0, 0, 1, 3, 0, 0, 0, P350),
    (3133, DATA, "Camiseta", 0, 0, 1, 0, 0, 0, 0, P350),
    (3133, DATA, "Camiseta FEM", 1, 1, 1, 0, 0, 0, 0, P350),
    (3133, DATA, "Regata Nadador FEM", 0, 0, 0, 0, 0, 0, 1, P350),

    (3495, DATA, "Camiseta FEM", 0, 0, 2, 1, 1, 0, 0, P350),
    (3495, DATA, "Camiseta", 0, 3, 7, 5, 5, 0, 2, P350),

    (3546, DATA, "Camiseta", 2, 8, 10, 3, 1, 0, 0, P350),
    (3546, DATA, "Camiseta", 2, 5, 1, 0, 0, 0, 0, P350),

    (3485, DATA, "Camiseta", 1, 12, 10, 0, 0, 0, 0, P350),
    (3485, DATA, "Camiseta", 0, 8, 2, 1, 1, 0, 0, P350),
    (3485, DATA, "Bandeira", 2, 0, 0, 0, 0, 0, 0, P350),

    (3428, DATA, "Camisa Raglan", 0, 4, 15, 11, 2, 0, 0, P350),
    (3428, DATA, "Machao", 0, 0, 2, 0, 0, 0, 0, P350),
    (3428, DATA, "Camiseta FEM Raglan", 1, 7, 8, 0, 0, 0, 0, P350),
    (3428, DATA, "Baby Look Raglan", 3, 1, 9, 4, 0, 0, 0, P350),

    (3429, DATA, "Camisa Raglan", 0, 4, 15, 12, 3, 0, 0, P350),
    (3429, DATA, "Machao", 0, 0, 3, 1, 0, 0, 0, P350),
    (3429, DATA, "Raglan Baby Look", 3, 1, 10, 6, 0, 0, 0, P350),
    (3429, DATA, "Camiseta Raglan FEM", 1, 8, 6, 0, 0, 1, 0, P350),

    (3430, DATA, "Calcao", 0, 2, 8, 4, 3, 0, 0, P700),
    (3430, DATA, "Calcao", 0, 0, 1, 0, 0, 0, 0, P700),
    (3430, DATA, "Raglan Baby Look", 0, 0, 0, 1, 0, 0, 0, P700),
    (3430, DATA, "Camisa Raglan", 0, 0, 1, 0, 0, 0, 0, P700),

    (3431, DATA, "Camisa Raglan", 0, 0, 2, 0, 0, 0, 0, P350),
    (3431, DATA, "Raglan Baby Look", 0, 0, 1, 0, 0, 0, 0, P350),

    (3432, DATA, "Camisa Raglan", 0, 0, 2, 0, 0, 0, 0, P700),
    (3432, DATA, "Raglan Baby Look", 0, 0, 1, 0, 0, 0, 0, P700),
    (3432, DATA, "Calcao", 0, 0, 1, 0, 0, 0, 0, P700),

    (3285, DATA, "Camiseta", 20, 20, 30, 25, 5, 0, 0, P350),

    (3513, DATA, "Camiseta", 0, 3, 10, 10, 10, 2, 6, P350),
    (3513, DATA, "Camiseta FEM", 0, 6, 8, 5, 0, 0, 0, P350),

    (3325, DATA, "Camiseta", 2, 15, 19, 14, 9, 0, 1, P350),

    (3434, DATA, "Camiseta FEM", 1, 2, 1, 0, 0, 0, 0, P350),
    (3434, DATA, "Baby Look", 0, 1, 1, 0, 0, 0, 0, P350),
    (3434, DATA, "Camiseta", 1, 1, 2, 0, 2, 0, 0, P350),
    (3434, DATA, "Machao", 0, 1, 0, 0, 1, 0, 0, P350),
    (3434, DATA, "Camiseta", 0, 0, 0, 0, 1, 0, 0, P350),

    (3433, DATA, "Camiseta", 1, 1, 1, 0, 0, 0, 0, P700),
    (3433, DATA, "Calcao", 1, 1, 1, 0, 0, 0, 0, P700),
    (3433, DATA, "Camiseta", 2, 1, 1, 0, 0, 0, 0, P700),
    (3433, DATA, "Short", 2, 1, 1, 1, 0, 0, 0, P700),

    (3336, DATA, "Raglan Manga Longa", 20, 20, 60, 60, 10, 3, 0, P380),
    (3336, DATA, "Camisa Raglan", 5, 8, 8, 0, 0, 0, 0, P380),

    (3327, DATA, "Raglan Manga Longa", 1, 8, 9, 19, 9, 4, 1, P380),
    (3327, DATA, "Raglan Manga Longa", 2, 3, 12, 9, 3, 2, 0, P380),
    (3327, DATA, "Raglan Manga Longa", 3, 6, 9, 2, 0, 0, 0, P380),
    (3327, DATA, "Raglan Manga Longa", 0, 0, 0, 0, 4, 2, 0, P380),

    (3383, DATA, "Raglan Manga Longa", 1, 4, 7, 8, 27, 2, 0, P380),

    (3380, DATA, "Moletom", 0, 2, 3, 8, 1, 0, 1, P600),

    # Novas fichas da pasta 16-06(2), inseridas ao final na ordem solicitada.
    (3401, DATA, "Camiseta", 7, 3, 2, 1, 6, 5, 1, P350),

    (3427, DATA, "Camiseta", 0, 20, 20, 20, 0, 3, 2, P350),

    (3411, DATA, "Camisa Raglan", 1, 1, 2, 11, 4, 5, 1, P350),
    (3411, DATA, "Camisa Raglan", 1, 0, 0, 0, 0, 0, 0, P350),
    (3411, DATA, "Camisa Raglan", 1, 1, 0, 2, 1, 0, 0, P350),

    (3444, DATA, "Camisa Raglan", 0, 10, 20, 20, 0, 0, 0, P350),

    (3260, DATA, "Camisa Raglan", 24, 42, 92, 61, 48, 57, 104, P350),
    (3260, DATA, "Camisa Raglan", 14, 124, 91, 36, 22, 3, 5, P350),
    (3260, DATA, "Camiseta Raglan FEM", 0, 0, 0, 0, 0, 0, 6, P350),
    (3260, DATA, "Camisa Raglan", 0, 2, 4, 2, 0, 2, 2, P350),

    (3443, DATA, "Camiseta", 1, 5, 6, 10, 10, 7, 7, P350),
    (3443, DATA, "Camiseta", 1, 1, 0, 0, 0, 0, 0, P350),
    (3443, DATA, "Regata Feminina", 0, 0, 0, 0, 1, 1, 0, P350),

    (3442, DATA, "Camiseta", 1, 0, 1, 3, 1, 0, 0, P350),

    (3390, DATA, "Camiseta", 0, 0, 4, 2, 1, 0, 1, P350),
    (3390, DATA, "Baby Look", 0, 1, 0, 1, 1, 0, 0, P350),
    (3390, DATA, "Camiseta FEM", 0, 2, 3, 0, 0, 0, 0, P350),

    # Segunda ficha 3390 (arquivo com sufixo (2)): inserir logo abaixo da primeira.
    (3390, DATA, "Camiseta FEM", 0, 2, 3, 0, 0, 0, 0, P350),

    (3450, DATA, "Camiseta", 0, 3, 4, 4, 1, 0, 0, P350),

    (3228, DATA, "Camiseta", 28, 4, 40, 3, 0, 0, 0, P350),
    (3228, DATA, "Camiseta", 0, 0, 5, 0, 3, 0, 0, P350),

    (3437, DATA, "Camiseta", 5, 7, 14, 26, 7, 1, 0, P350),

    (3494, DATA, "Camiseta", 5, 17, 17, 9, 2, 0, 0, P350),

    (3440, DATA, "Bandeira", 0, 0, 70, 0, 0, 0, 0, Decimal("2.00")),

    (3286, DATA, "Camiseta", 1, 0, 0, 2, 0, 0, 0, P350),
    (3286, DATA, "Camiseta FEM", 0, 1, 1, 0, 0, 0, 0, P350),

    (3229, DATA, "Abada", 0, 10, 35, 35, 15, 5, 0, P350),
    (3229, DATA, "Abada", 0, 0, 15, 20, 15, 0, 0, P350),

    (3227, DATA, "Raglan Manga Longa", 8, 9, 19, 5, 4, 6, 0, P380),
    (3227, DATA, "Raglan Manga Longa", 3, 12, 9, 7, 0, 0, 0, P380),
    (3227, DATA, "Raglan Manga Longa", 6, 9, 2, 3, 0, 0, 0, P380),
    (3227, DATA, "Raglan Manga Longa", 4, 0, 0, 0, 0, 2, 0, P380),

    ("", DATA, "Sem codigo de referencia", 1, 1, 1, 2, 0, 0, 0, P350),
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
for refer, data, produto, pp, p, m, g, gg, xgg, xg, preco in dados:
    qtd, total = calc_line(pp, p, m, g, gg, xgg, xg, preco)
    rows_out.append([
        refer, data, produto, CLIENTE,
        pp if pp else "", p if p else "", m if m else "", g if g else "", gg if gg else "", xgg if xgg else "", xg if xg else "",
        qtd, format_brl(preco), format_brl(total), ""
    ])

total_geral = Decimal("0.00")
for row in rows_out:
    total_geral += Decimal(row[13].replace(",", "."))

rows_out.append([
    "", "", "TOTAL GERAL", CLIENTE,
    "", "", "", "", "", "", "",
    "", "", format_brl(total_geral), ""
])

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

widths = {
    "A": 10, "B": 10, "C": 28, "D": 14,
    "E": 6, "F": 6, "G": 6, "H": 6, "I": 6, "J": 6, "K": 6,
    "L": 12, "M": 9, "N": 12, "O": 10
}
for col, width in widths.items():
    ws.column_dimensions[col].width = width

wb.save(XLSX_PATH)

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
    Paragraph("Fechamento Mensal - Ultra Silk (16-06)", title_style),
    Spacer(1, 0.25 * cm),
]

table_data = [headers] + rows_out
col_widths = [
    1.5 * cm, 1.8 * cm, 4.2 * cm, 2.5 * cm,
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
