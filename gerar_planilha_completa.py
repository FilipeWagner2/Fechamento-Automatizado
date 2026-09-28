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

XLSX_PATH = OUT_DIR / "fechamento_completo_conferencia.xlsx"
PDF_PATH = OUT_DIR / "fechamento_completo_conferencia.pdf"

PRECO = Decimal("3.50")

# Colunas: Refer, Data, Produto, Cliente, PP, P, M, G, GG, XGG, XG, Status
# Observacao: alguns itens vieram de leitura de foto com baixa nitidez e estao marcados para conferencia.
rows_raw = [
    # Ficha 1 (exemplo validado)
    (3154, "13/04/26", "US - Camiseta", "Ultra Silk", 0, 1, 3, 0, 1, 0, 11, ""),
    (3154, "13/04/26", "US - Camiseta", "Ultra Silk", 0, 2, 1, 1, 0, 1, 7, ""),
    (3154, "13/04/26", "US - Camiseta", "Ultra Silk", 1, 10, 7, 1, 0, 1, 4, ""),
    (3154, "13/04/26", "US - Camiseta", "Ultra Silk", 0, 11, 3, 0, 1, 0, 2, ""),
    (3154, "13/04/26", "US - Bandeira", "Ultra Silk", 0, 0, 4, 0, 0, 0, 0, ""),

    # Ficha 2
    (3129, "09/04/26", "US - Camiseta FEM", "Ultra Silk", 0, 0, 1, 0, 0, 0, 0, ""),

    # Ficha 3
    (3286, "06/05/26", "US - Camiseta", "Ultra Silk", 0, 1, 0, 0, 2, 0, 0, ""),
    (3286, "06/05/26", "US - Camiseta FEM", "Ultra Silk", 0, 1, 1, 0, 0, 0, 0, ""),

    # Ficha 4
    (3228, "28/04/26", "US - Camiseta", "Ultra Silk", 0, 28, 4, 40, 3, 0, 0, ""),
    (3228, "28/04/26", "US - Camiseta", "Ultra Silk", 0, 0, 5, 0, 3, 0, 0, ""),

    # Ficha 5
    (3229, "28/04/26", "US - Abada", "Ultra Silk", 0, 10, 35, 35, 15, 5, 0, ""),
    (3229, "28/04/26", "US - Abada", "Ultra Silk", 0, 15, 20, 15, 0, 0, 0, ""),

    # Ficha 6
    (3201, "23/04/26", "US - Camiseta", "Ultra Silk", 4, 11, 15, 16, 12, 0, 0, ""),
    (3201, "23/04/26", "US - Camiseta", "Ultra Silk", 3, 19, 46, 24, 9, 8, 1, ""),
    (3201, "23/04/26", "US - Camiseta", "Ultra Silk", 0, 0, 1, 0, 0, 0, 0, ""),

    # Ficha 7
    (3112, "07/04/26", "US - Camiseta", "Ultra Silk", 0, 1, 10, 4, 1, 0, 0, ""),
    (3112, "07/04/26", "US - Camiseta", "Ultra Silk", 0, 0, 1, 1, 0, 0, 0, ""),
    (3112, "07/04/26", "US - Camiseta FEM", "Ultra Silk", 0, 4, 4, 4, 3, 2, 0, ""),
    (3112, "07/04/26", "US - Camiseta FEM", "Ultra Silk", 0, 0, 1, 0, 1, 0, 0, ""),

    # Ficha 8
    (3111, "07/04/26", "US - Camiseta", "Ultra Silk", 0, 0, 4, 15, 11, 2, 0, ""),
    (3111, "07/04/26", "US - Camiseta", "Ultra Silk", 0, 0, 4, 15, 11, 2, 0, ""),

    # Ficha 9
    (2979, "18/03/26", "US - Camiseta FEM", "Ultra Silk", 2, 10, 7, 8, 7, 4, 0, ""),
    (2979, "18/03/26", "US - Camiseta FEM", "Ultra Silk", 2, 10, 7, 8, 7, 4, 0, ""),

    # Ficha 10
    (3128, "04/04/26", "US - Babby look", "Ultra Silk", 0, 1, 1, 1, 1, 0, 0, ""),
    (3128, "04/04/26", "US - Camiseta", "Ultra Silk", 0, 0, 1, 0, 0, 0, 0, ""),
    (3128, "04/04/26", "US - Camiseta FEM", "Ultra Silk", 0, 0, 1, 0, 0, 0, 0, ""),
    (3128, "04/04/26", "US - Camiseta", "Ultra Silk", 1, 0, 0, 0, 0, 0, 0, ""),

    # Ficha 11
    (3149, "13/04/26", "US - Camiseta", "Ultra Silk", 0, 0, 6, 8, 0, 0, 3, ""),
    (3149, "13/04/26", "US - Camiseta FEM", "Ultra Silk", 0, 3, 4, 2, 0, 0, 0, ""),
    (3149, "13/04/26", "US - Camiseta", "Ultra Silk", 0, 3, 10, 7, 0, 0, 0, ""),
    (3149, "13/04/26", "US - Camiseta FEM", "Ultra Silk", 0, 5, 3, 2, 0, 0, 0, ""),

    # Ficha 12
    (3043, "26/03/26", "US - Camiseta", "Ultra Silk", 0, 2, 18, 21, 17, 0, 0, ""),

    # Ficha 13 (conferencia)
    (3080, "31/03/26", "US - Camiseta", "Ultra Silk", 0, 3, 7, 14, 4, 0, 0, "Conferir"),

    # Ficha 14
    (3107, "07/04/26", "US - Camiseta", "Ultra Silk", 0, 10, 5, 11, 5, 0, 1, ""),
    (3107, "07/04/26", "US - Bandeira", "Ultra Silk", 0, 0, 1, 0, 0, 0, 0, ""),

    # Ficha 15 (conferencia)
    (3100, "03/04/26", "US - Babby look", "Ultra Silk", 3, 1, 1, 2, 1, 2, 0, "Conferir"),
    (3100, "03/04/26", "US - Camiseta FEM", "Ultra Silk", 1, 6, 15, 6, 1, 2, 1, "Conferir"),
    (3100, "03/04/26", "US - Camiseta", "Ultra Silk", 1, 1, 4, 4, 1, 0, 2, "Conferir"),
    (3100, "03/04/26", "US - Camiseta", "Ultra Silk", 1, 4, 4, 0, 1, 0, 0, "Conferir"),

    # Ficha 16
    (3105, "07/04/26", "US - Camiseta", "Ultra Silk", 2, 10, 18, 8, 2, 0, 0, ""),
    (3105, "07/04/26", "US - Regata", "Ultra Silk", 4, 10, 15, 8, 3, 0, 0, ""),

    # Ficha 17 (conferencia)
    (3228, "13/04/26", "US - Camiseta", "Ultra Silk", 28, 4, 40, 3, 0, 0, 0, "Conferir"),

    # Ficha 18 (conferencia)
    (2968, "17/03/26", "US - Camiseta FEM", "Ultra Silk", 1, 9, 4, 2, 1, 0, 0, "Conferir"),
]

headers = [
    "Refer.", "Data", "Produto", "Cliente",
    "PP", "P", "M", "G", "GG", "XGG", "XG",
    "Qtde Peças", "Preço", "Total", "Status"
]


def format_brl(value: Decimal) -> str:
    text = f"{value:.2f}".replace(".", ",")
    return text


def calc_line(pp: int, p: int, m: int, g: int, gg: int, xgg: int, xg: int):
    qtd = pp + p + m + g + gg + xgg + xg
    total = (Decimal(qtd) * PRECO).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return qtd, total


rows_out = []
for r in rows_raw:
    refer, data, produto, cliente, pp, p, m, g, gg, xgg, xg, status = r
    qtd, total = calc_line(pp, p, m, g, gg, xgg, xg)
    rows_out.append([
        refer, data, produto, cliente,
        pp if pp else "", p if p else "", m if m else "", g if g else "", gg if gg else "", xgg if xgg else "", xg if xg else "",
        qtd, format_brl(PRECO), format_brl(total), status
    ])

# XLSX
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

# Larguras
widths = {
    "A": 10, "B": 10, "C": 28, "D": 14,
    "E": 6, "F": 6, "G": 6, "H": 6, "I": 6, "J": 6, "K": 6,
    "L": 12, "M": 9, "N": 12, "O": 10
}
for col, width in widths.items():
    ws.column_dimensions[col].width = width

wb.save(XLSX_PATH)

# PDF
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
    Paragraph("Fechamento Mensal - Conferencia", title_style),
    Spacer(1, 0.25 * cm),
]

table_data = [headers] + rows_out
col_widths = [
    1.5 * cm, 1.8 * cm, 4.2 * cm, 2.5 * cm,
    0.8 * cm, 0.8 * cm, 0.8 * cm, 0.8 * cm, 0.8 * cm, 1.0 * cm, 1.0 * cm,
    1.8 * cm, 1.2 * cm, 1.6 * cm, 1.5 * cm,
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

# Destaca linhas com Status=Conferir
status_col = headers.index("Status")
for i, row in enumerate(rows_out, start=1):
    if row[status_col] == "Conferir":
        table_style.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor("#FFF3CD")))

table.setStyle(TableStyle(table_style))
elements.append(table)
doc.build(elements)

print(f"XLSX: {XLSX_PATH}")
print(f"PDF: {PDF_PATH}")
print(f"Linhas: {len(rows_out)}")
