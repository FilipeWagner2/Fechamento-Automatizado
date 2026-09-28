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

XLSX_PATH = OUT_DIR / "fechamento_novo.xlsx"
PDF_PATH = OUT_DIR / "fechamento_novo.pdf"

PRECO_PADRAO = Decimal("3.50")
PRECO_CONJUNTO = Decimal("7.00")
CLIENTE = "Ultra Silk"
FEITAS_DIR = BASE_DIR / "Feitas"

# Todos os dados com referência numérica
dados_com_ref = [
    (2968, "17/03/26", "US - Camiseta", 0, 0, 0, 1, 1, 1, 0, PRECO_PADRAO),
    (2968, "17/03/26", "US - Camiseta FEM", 1, 8, 4, 2, 1, 1, 0, PRECO_PADRAO),
    (2968, "17/03/26", "US - Machao", 0, 1, 0, 0, 0, 0, 0, PRECO_PADRAO),

    (2979, "18/03/26", "US - Camiseta FEM", 2, 10, 7, 8, 7, 4, 0, PRECO_PADRAO),
    (2979, "18/03/26", "US - Camiseta FEM", 2, 10, 7, 8, 7, 4, 0, PRECO_PADRAO),

    (3043, "26/03/26", "US - Camiseta", 0, 2, 18, 21, 17, 0, 0, PRECO_PADRAO),

    (3050, "27/03/26", "US - Babby look", 13, 4, 1, 6, 0, 0, 0, PRECO_PADRAO),
    (3050, "27/03/26", "US - Camiseta", 1, 2, 20, 4, 0, 0, 0, PRECO_PADRAO),

    (3077, "31/03/26", "US - Camiseta FEM", 0, 0, 2, 2, 0, 0, 0, PRECO_PADRAO),
    (3077, "31/03/26", "US - Camiseta", 1, 0, 8, 2, 0, 0, 1, PRECO_PADRAO),
    (3077, "31/03/26", "US - Camiseta", 0, 5, 10, 16, 0, 10, 4, PRECO_PADRAO),

    (3078, "31/03/26", "US - Camisa Raglan", 20, 30, 20, 0, 0, 5, 0, PRECO_PADRAO),
    (3078, "31/03/26", "US - Camiseta", 0, 0, 20, 20, 0, 10, 0, PRECO_PADRAO),
    (3078, "31/03/26", "US - Bandeira", 5, 7, 0, 0, 0, 2, 1, PRECO_PADRAO),
    (3078, "31/03/26", "US - Bandeira", 2, 0, 0, 0, 0, 0, 0, PRECO_PADRAO),

    (3080, "01/04/26", "US - Camiseta", 0, 1, 9, 14, 4, 0, 0, PRECO_PADRAO),
    (3080, "01/04/26", "US - Babby look", 0, 1, 2, 4, 1, 0, 2, PRECO_PADRAO),

    (3100, "06/04/26", "US - Babby look", 3, 2, 3, 1, 0, 1, 0, PRECO_PADRAO),
    (3100, "06/04/26", "US - Camiseta FEM", 1, 8, 15, 6, 1, 1, 0, PRECO_PADRAO),
    (3100, "06/04/26", "US - Camiseta", 0, 2, 4, 4, 2, 1, 0, PRECO_PADRAO),
    (3100, "06/04/26", "US - Camiseta", 1, 4, 4, 1, 0, 0, 0, PRECO_PADRAO),

    (3099, "03/04/26", "US - Sacochila", 0, 0, 3, 0, 0, 0, 0, PRECO_CONJUNTO),
    (3099, "03/04/26", "US - Camiseta", 0, 0, 108, 0, 0, 0, 0, PRECO_CONJUNTO),
    (3099, "03/04/26", "US - Calcao", 4, 7, 17, 18, 27, 6, 29, PRECO_CONJUNTO),
    (3099, "03/04/26", "US - Calcao", 4, 7, 17, 18, 27, 6, 29, PRECO_CONJUNTO),

    (3105, "07/04/26", "US - Camiseta", 2, 10, 18, 8, 2, 0, 0, PRECO_PADRAO),
    (3105, "07/04/26", "US - Regata", 4, 10, 15, 8, 3, 0, 0, PRECO_PADRAO),
    (3105, "07/04/26", "US - Camiseta", 2, 10, 18, 8, 2, 0, 0, PRECO_PADRAO),
    (3105, "07/04/26", "US - Regata", 4, 10, 15, 8, 3, 0, 0, PRECO_PADRAO),

    (3107, "07/04/26", "US - Camiseta", 10, 5, 11, 5, 0, 1, 0, PRECO_PADRAO),
    (3107, "07/04/26", "US - Bandeira", 0, 0, 1, 0, 0, 0, 0, PRECO_PADRAO),

    (3109, "07/04/26", "US - Camiseta", 30, 40, 40, 30, 20, 10, 0, PRECO_PADRAO),
    (3109, "07/04/26", "US - Camiseta", 15, 20, 35, 35, 0, 0, 0, PRECO_PADRAO),

    (3111, "07/04/26", "US - Camiseta", 0, 0, 4, 15, 11, 2, 0, PRECO_PADRAO),
    (3111, "07/04/26", "US - Camiseta", 0, 0, 4, 15, 11, 2, 0, PRECO_PADRAO),

    (3112, "07/04/26", "US - Camiseta", 0, 0, 1, 10, 4, 1, 0, PRECO_PADRAO),
    (3112, "07/04/26", "US - Camiseta", 0, 0, 0, 1, 1, 0, 0, PRECO_PADRAO),
    (3112, "07/04/26", "US - Camiseta FEM", 0, 4, 4, 4, 3, 2, 0, PRECO_PADRAO),
    (3112, "07/04/26", "US - Camiseta FEM", 0, 0, 1, 0, 0, 1, 0, PRECO_PADRAO),

    (3113, "07/04/26", "US - Camiseta", 0, 0, 0, 3, 0, 0, 0, PRECO_PADRAO),

    (3130, "09/04/26", "US - Babby look", 0, 0, 1, 3, 0, 0, 0, PRECO_PADRAO),
    (3130, "09/04/26", "US - Camiseta", 0, 1, 0, 0, 0, 0, 0, PRECO_PADRAO),
    (3130, "09/04/26", "US - Regata Nadador FEM", 0, 0, 0, 0, 0, 0, 1, PRECO_PADRAO),

    (3131, "09/04/26", "US - Babby look", 0, 0, 2, 6, 4, 0, 0, PRECO_PADRAO),
    (3131, "09/04/26", "US - Camiseta", 0, 0, 1, 1, 1, 0, 1, PRECO_PADRAO),
    (3131, "09/04/26", "US - Camiseta FEM", 1, 2, 2, 0, 0, 0, 0, PRECO_PADRAO),
    (3131, "09/04/26", "US - Regata Nadador FEM", 0, 0, 1, 0, 0, 0, 1, PRECO_PADRAO),

    (3132, "09/04/26", "US - Regata Baby look", 0, 0, 0, 0, 1, 0, 0, PRECO_PADRAO),
    (3132, "09/04/26", "US - Camiseta", 1, 0, 0, 0, 0, 0, 0, PRECO_PADRAO),
    (3132, "09/04/26", "US - Regata Baby look", 1, 0, 0, 0, 0, 0, 1, PRECO_PADRAO),
    (3132, "09/04/26", "US - Camiseta", 0, 1, 0, 0, 0, 0, 0, PRECO_PADRAO),

    (3133, "09/04/26", "US - Babby look", 0, 0, 1, 3, 0, 0, 0, PRECO_PADRAO),
    (3133, "09/04/26", "US - Camiseta", 0, 0, 1, 0, 0, 0, 0, PRECO_PADRAO),
    (3133, "09/04/26", "US - Camiseta FEM", 1, 1, 1, 0, 0, 0, 0, PRECO_PADRAO),
    (3133, "09/04/26", "US - Regata Nadador FEM", 0, 0, 0, 0, 0, 0, 1, PRECO_PADRAO),

    (3146, "13/04/26", "US - Camiseta", 0, 0, 10, 10, 0, 0, 0, PRECO_CONJUNTO),
    (3146, "13/04/26", "US - Camiseta FEM", 0, 10, 10, 0, 0, 0, 0, PRECO_CONJUNTO),
    (3146, "13/04/26", "US - Camiseta", 0, 0, 0, 3, 0, 0, 0, PRECO_CONJUNTO),
    (3146, "13/04/26", "US - Camiseta", 0, 0, 0, 1, 0, 0, 0, PRECO_CONJUNTO),
    (3146, "13/04/26", "US - Camisa M. Longa", 0, 0, 0, 0, 0, 0, 2, PRECO_CONJUNTO),

    (3149, "13/04/26", "US - Camiseta", 0, 0, 0, 6, 8, 0, 3, PRECO_PADRAO),
    (3149, "13/04/26", "US - Camiseta FEM", 0, 3, 4, 2, 0, 0, 0, PRECO_PADRAO),
    (3149, "13/04/26", "US - Camiseta FEM", 0, 0, 3, 10, 7, 0, 0, PRECO_PADRAO),
    (3149, "13/04/26", "US - Camiseta", 0, 5, 3, 2, 0, 0, 0, PRECO_PADRAO),

    (3154, "13/04/26", "US - Camiseta", 0, 1, 3, 0, 1, 0, 11, PRECO_PADRAO),
    (3154, "13/04/26", "US - Camiseta", 0, 2, 1, 1, 0, 1, 7, PRECO_PADRAO),
    (3154, "13/04/26", "US - Camiseta", 1, 10, 7, 1, 0, 1, 4, PRECO_PADRAO),
    (3154, "13/04/26", "US - Camiseta", 0, 11, 3, 0, 1, 0, 2, PRECO_PADRAO),
    (3154, "13/04/26", "US - Bandeira", 0, 0, 4, 0, 0, 0, 0, PRECO_PADRAO),

    (3197, "23/04/26", "US - Camiseta", 0, 0, 4, 4, 3, 0, 4, PRECO_PADRAO),
    (3197, "23/04/26", "US - Camiseta", 0, 2, 8, 5, 3, 0, 2, PRECO_PADRAO),

    (3201, "23/04/26", "US - Camiseta", 4, 11, 15, 16, 12, 0, 0, PRECO_PADRAO),
    (3201, "23/04/26", "US - Camiseta", 3, 19, 46, 24, 9, 8, 1, PRECO_PADRAO),
    (3201, "23/04/26", "US - Camiseta", 0, 0, 0, 0, 0, 0, 1, PRECO_PADRAO),

    (3228, "28/04/26", "US - Camiseta", 28, 4, 40, 3, 0, 0, 0, PRECO_PADRAO),
    (3228, "28/04/26", "US - Camiseta", 0, 0, 5, 0, 3, 0, 0, PRECO_PADRAO),
    (3228, "28/04/26", "US - Camiseta", 28, 4, 40, 3, 0, 0, 0, PRECO_PADRAO),
    (3228, "28/04/26", "US - Camiseta", 0, 0, 5, 0, 3, 0, 0, PRECO_PADRAO),

    (3229, "28/04/26", "US - Abada", 0, 10, 35, 35, 15, 5, 0, PRECO_PADRAO),
    (3229, "28/04/26", "US - Abada", 0, 15, 20, 15, 0, 0, 0, PRECO_PADRAO),

    (3286, "06/05/26", "US - Camiseta", 0, 1, 0, 2, 0, 0, 0, PRECO_PADRAO),
    (3286, "06/05/26", "US - Camiseta FEM", 0, 1, 1, 0, 0, 0, 0, PRECO_PADRAO),

    (3294, "07/05/26", "US - Camiseta", 30, 100, 100, 0, 0, 0, 0, PRECO_PADRAO),
]

# Dados sem referência (vão por último)
dados_sem_ref = [
    ("", "", "US - Camiseta", 1, 5, 5, 16, 14, 0, 5, PRECO_PADRAO),
    ("", "", "US - Camiseta FEM", 5, 7, 14, 13, 18, 0, 4, PRECO_PADRAO),
    ("", "", "US - Infantil", 1, 1, 2, 5, 3, 1, 14, PRECO_PADRAO),
    ("", "", "US - Camiseta Verde AABB", 0, 0, 920, 0, 0, 0, 0, PRECO_PADRAO),
    ("", "", "US - Camiseta Branca Algodao", 0, 0, 40, 0, 0, 0, 0, PRECO_PADRAO),
    ("", "", "US - Camiseta", 0, 0, 94, 0, 0, 0, 0, PRECO_PADRAO),
    ("", "", "US - Camiseta Algodao", 0, 0, 76, 0, 0, 0, 0, PRECO_PADRAO),
    ("", "", "US - Camisetas Guerreiros", 0, 0, 119, 0, 0, 0, 0, PRECO_PADRAO),
    ("", "", "US - Colete", 0, 0, 4180, 0, 0, 0, 0, Decimal("1.50")),
]

# Identifica referências de conjunto pelos nomes de arquivos na pasta Feitas.
# Ex.: "conjunto 2(3146).jpeg" -> 3146
refs_conjunto = set()
if FEITAS_DIR.exists():
    for f in FEITAS_DIR.iterdir():
        nome = f.name.lower()
        if "conjunto" in nome:
            match = re.search(r"\((\d+)\)", nome)
            if match:
                refs_conjunto.add(int(match.group(1)))

# Ordena por: nao-conjunto primeiro, depois conjunto; ambos em ordem crescente.
dados_com_ref.sort(key=lambda x: (x[0] in refs_conjunto, x[0]))

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

# Processa dados com referência (ordenados)
for refer, data, produto, pp, p, m, g, gg, xgg, xg, preco in dados_com_ref:
    qtd, total = calc_line(pp, p, m, g, gg, xgg, xg, preco)
    rows_out.append([
        refer, data, produto, CLIENTE,
        pp if pp else "", p if p else "", m if m else "", g if g else "", gg if gg else "", xgg if xgg else "", xg if xg else "",
        qtd, format_brl(preco), format_brl(total), ""
    ])

# Processa dados sem referência (vai por último)
for refer, data, produto, pp, p, m, g, gg, xgg, xg, preco in dados_sem_ref:
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

widths = {
    "A": 10, "B": 10, "C": 28, "D": 14,
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
    Paragraph("Fechamento Mensal - Organizado", title_style),
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
