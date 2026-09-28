from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet


BASE_DIR = Path(__file__).resolve().parent
OUT_DIR = BASE_DIR / "saida"
OUT_DIR.mkdir(exist_ok=True)

XLSX_PATH = OUT_DIR / "fechamento_ultra_silk_2907.xlsx"
PDF_PATH = OUT_DIR / "fechamento_ultra_silk_2907.pdf"

CLIENTE = "Ultra Silk"
DATA = "29/07/26"
P350 = Decimal("3.50")
P400 = Decimal("4.00")
P700 = Decimal("7.00")

SIZE_COLS = ["2", "4", "6", "8", "10", "12", "14", "16", "PP", "P", "M", "G", "GG", "XG", "EG", "EXG", "EXXG"]


def row(ref, produto, preco, **sizes):
    return {"ref": ref, "data": DATA, "produto": produto, "preco": preco, "sizes": sizes}


# Ordem: fichas na ordem de renomeacao/criacao. As fichas de R$ 7,00 ficam no fim.
dados = [
    row("3721", "Camiseta", P350, P=2, M=1, G=4, GG=3, XG=2),
    row("3606", "Colete", P350, **{"12": 100, "16": 100}, M=1, G=1, GG=1),
    row("3687", "Camiseta", P350, PP=4, P=22, M=19, G=9, GG=3),
    row("3601", "Segunda pele", P350, PP=15, P=20, M=20, G=20, GG=5, XG=5),
    row("3622", "Camisa Raglan", P350, P=5, M=5, G=20, GG=10),
    row("3622", "Camiseta FEM", P350, M=5, G=5, GG=5),
    row("3629", "Camisa Raglan", P350, **{"14": 2}, PP=2, P=18, M=4, G=14),
    row("3629", "Baby Look", P350, PP=2, P=4),
    row("3637", "Camiseta Raglan", P350, **{"12": 6, "14": 17}, PP=3, P=4),
    row("3565", "Camiseta - Cinza", P350, **{"12": 5, "14": 5}, PP=5, P=5, M=5, G=5),
    row("3565", "Camiseta - Verde", P350, **{"12": 5, "14": 5}, P=5, M=5, G=5),
    row("3565", "Camiseta - Laranja", P350, **{"12": 10, "14": 10, "16": 5}, PP=10, P=5, M=5),
    row("3571", "Camisa Raglan", P350, PP=7, P=2, M=11, G=4, GG=4),
    row("3596", "Camiseta", P350, PP=2, P=2, M=2, G=2, GG=2),
    row("3539", "Camiseta", P350, PP=2, P=1, M=5),
    row("3539", "Camiseta FEM", P350, PP=10),
    row("3614", "Camisa Raglan", P350, PP=9, P=14, M=8, G=1),
    row("3594", "Camiseta", P350, PP=1, P=2, M=4, G=3),
    row("3594", "Camiseta", P350, **{"12": 2, "14": 5, "16": 13}),
    row("3560", "Camiseta FEM - A", P350, M=1, G=4),
    row("3560", "Camiseta FEM - B", P350, PP=1, M=2, G=3, GG=1, XG=1),
    row("3572", "Camiseta FEM - Onca", P350, PP=5, P=2, G=1),
    row("3572", "Camiseta - Trem", P350, **{"14": 2}, P=1, M=2, G=3, GG=3),
    row("3577", "Camiseta - Rosa", P350, **{"2": 1}, P=3, M=4, G=5, GG=3),
    row("3577", "Camiseta - Azul", P350, P=5, G=5, GG=5),
    row("3588", "Camiseta", P350, PP=6, P=4),
    row("3587", "Camiseta - Amarela", P350, P=6, M=8, G=11, GG=10, XG=5),
    row("3587", "Camiseta - Azul", P350, P=6, M=8, G=11, GG=10, XG=5),
    row("3524", "Camiseta", P350, P=6, M=12, G=6, GG=1),
    row("3584", "Camiseta", P350, **{"2": 2, "4": 3, "6": 6, "10": 1, "12": 1}),
    row("", "Camiseta - Sem codigo de referencia", P350, **{"6": 5, "14": 5}, M=20, G=20),
    row("2908", "Camiseta", P350, **{"6": 1}, M=4, G=3),
    row("2908", "Baby Look", P350, M=2),
    row("3580", "Camiseta", P400, M=3, G=11, GG=4),
    row("3580", "Camiseta M. Longa", P400, P=1, G=3, GG=2),
    row("3609", "Camisa M. Longa", P350, M=10, G=11, GG=10, XG=10, EG=2, EXG=3),
    row("3609", "Camisa M. Longa", P350, **{"8": 8, "10": 5, "12": 5}),
    row("3609", "Camiseta", P350, P=6, M=16, G=16, GG=10, XG=10, EG=10, EXG=8),
    row("3609", "Baby Look", P350, PP=5, P=10, M=13, G=10, GG=10, XG=5, EG=5),
    row("3609", "Camiseta", P350, **{"2": 9, "4": 5, "6": 5, "8": 8, "10": 5, "12": 5, "14": 5}),
    row("3706", "Segunda pele", P400, P=14, M=32, G=24, XG=18, EXG=12),
    row("3433", "Camiseta", P700, **{"12": 1, "14": 1}, P=1),
    row("3433", "Calcao", P700, **{"12": 1, "14": 1}, P=1),
    row("3433", "Camiseta", P700, **{"12": 2, "16": 1}, PP=1),
    row("3433", "Short", P700, **{"12": 2, "16": 1}, PP=1, GG=1),
    row("3434", "Camiseta FEM", P700, **{"14": 1}, P=2, M=1),
    row("3434", "Baby Look", P700, P=1, M=1),
    row("3434", "Camiseta", P700, **{"12": 1}, M=1, GG=2, EG=2),
    row("3434", "Machao", P700, P=1, GG=1),
    row("3434", "Camiseta - Laranja", P700, GG=1),
    row("3540", "Camiseta", P700, **{"6": 2, "8": 5, "10": 4, "12": 4, "14": 5}, P=4, GG=1),
    row("3540", "Calcao", P700, **{"6": 2, "8": 3, "10": 4, "12": 3, "14": 4}, P=1, G=1),
    row("3540", "Camisa Goleiro", P700, **{"10": 1, "12": 1}),
    row("3540", "Camiseta Professor", P700, G=2),
]

headers = ["Refer.", "Data", "Produto", "Cliente", *SIZE_COLS, "Qtde Pecas", "Preco", "Total", "Status"]


def format_brl(value: Decimal) -> str:
    return f"{value:.2f}".replace(".", ",")


def calc_total(qtd: int, preco: Decimal) -> Decimal:
    return (Decimal(qtd) * preco).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


rows_out = []
for item in dados:
    qtd = sum(int(item["sizes"].get(size, 0) or 0) for size in SIZE_COLS)
    total = calc_total(qtd, item["preco"])
    rows_out.append([
        item["ref"], item["data"], item["produto"], CLIENTE,
        *[item["sizes"].get(size, "") for size in SIZE_COLS],
        qtd, format_brl(item["preco"]), format_brl(total), "",
    ])

total_geral = sum(Decimal(row[-2].replace(",", ".")) for row in rows_out)
rows_out.append([
    "", "", "TOTAL GERAL", CLIENTE,
    *["" for _ in SIZE_COLS],
    "", "", format_brl(total_geral), "",
])

wb = Workbook()
ws = wb.active
ws.title = "Fechamento"
ws.append(headers)
for row_data in rows_out:
    ws.append(row_data)

header_fill = PatternFill("solid", fgColor="1F3864")
header_font = Font(color="FFFFFF", bold=True)
thin = Side(style="thin", color="CCCCCC")

for cell in ws[1]:
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

for row_cells in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=len(headers)):
    for cell in row_cells:
        cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
        cell.alignment = Alignment(horizontal="center", vertical="center")

for row_cells in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=3, max_col=3):
    for cell in row_cells:
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

for cell in ws[ws.max_row]:
    cell.font = Font(bold=True)

widths = {"A": 9, "B": 10, "C": 24, "D": 13}
for idx, _size in enumerate(SIZE_COLS, start=5):
    widths[ws.cell(row=1, column=idx).column_letter] = 5
widths[ws.cell(row=1, column=len(headers) - 3).column_letter] = 11
widths[ws.cell(row=1, column=len(headers) - 2).column_letter] = 9
widths[ws.cell(row=1, column=len(headers) - 1).column_letter] = 11
widths[ws.cell(row=1, column=len(headers)).column_letter] = 9
for col, width in widths.items():
    ws.column_dimensions[col].width = width

wb.save(XLSX_PATH)

doc = SimpleDocTemplate(
    str(PDF_PATH),
    pagesize=landscape(A4),
    leftMargin=0.45 * cm,
    rightMargin=0.45 * cm,
    topMargin=0.8 * cm,
    bottomMargin=0.8 * cm,
)

styles = getSampleStyleSheet()
title_style = ParagraphStyle("title", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=11)
elements = [
    Paragraph("Fechamento - Ultra Silk | 29/07/26", title_style),
    Spacer(1, 0.2 * cm),
]

table_data = [headers] + rows_out
col_widths = [
    1.15 * cm, 1.45 * cm, 3.35 * cm, 1.75 * cm,
    *([0.45 * cm] * len(SIZE_COLS)),
    1.25 * cm, 1.0 * cm, 1.2 * cm, 0.85 * cm,
]

table = Table(table_data, colWidths=col_widths, repeatRows=1)
table_style = [
    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F3864")),
    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
    ("FONTSIZE", (0, 0), (-1, 0), 5.5),
    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
    ("ALIGN", (2, 1), (2, -1), "LEFT"),
    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#B7B7B7")),
    ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
    ("FONTSIZE", (0, 1), (-1, -1), 5.5),
    ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
    ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#D9E1F2")),
]
for i in range(1, len(table_data)):
    if i % 2 == 0:
        table_style.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor("#EDF2FB")))

table.setStyle(TableStyle(table_style))
elements.append(table)
doc.build(elements)

print(f"XLSX: {XLSX_PATH}")
print(f"PDF:  {PDF_PATH}")
print(f"Linhas de dados: {len(dados)}")
print(f"Total geral: R$ {format_brl(total_geral)}")
