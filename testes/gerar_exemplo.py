"""Gera uma planilha de faturamento de exemplo com o mesmo layout que a macro espera.

  Painel!D7          cupons do mês
  Diário!B6:B36      datas (fórmulas)
  Diário!E6:E36      Realizado
Uso: python3 gerar_exemplo.py saida.xlsx [--filtro "NOME A EXCLUIR"]
(depois passe pelo LibreOffice para gravar os valores das fórmulas)
"""
import sys, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import BarChart, Reference

saida = sys.argv[1]
filtro = sys.argv[sys.argv.index("--filtro") + 1] if "--filtro" in sys.argv else None
wb = Workbook()
painel = wb.active; painel.title = "Painel"
d = wb.create_sheet("Diário")
verde = PatternFill("solid", fgColor="1F6F5C"); branco = Font(bold=True, color="FFFFFF")

d["B2"] = "Faturamento diário — novembro/2026"; d["B2"].font = Font(bold=True, size=14)
for c, t in zip("BCDEF", ["Data", "Dia", "Meta", "Realizado", "Realizado − Meta"]):
    d[f"{c}5"] = t; d[f"{c}5"].font = branco; d[f"{c}5"].fill = verde; d[f"{c}5"].alignment = Alignment(horizontal="center")
d["B6"] = dt.date(2026, 11, 1)
for r in range(6, 36):
    if r > 6: d[f"B{r}"] = f"=B{r-1}+1"
    d[f"B{r}"].number_format = "dd/mm/yyyy"
    d[f"C{r}"] = f'=TEXT(B{r},"ddd")'
    d[f"D{r}"] = 4200 if r % 7 not in (4, 5) else 5600
    d[f"D{r}"].number_format = "#,##0.00"
    d[f"E{r}"].number_format = "#,##0.00"
    d[f"F{r}"] = f'=IF(E{r}="","",E{r}-D{r})'
    d[f"F{r}"].number_format = "#,##0.00;[Red]-#,##0.00"
# 15 primeiros dias já preenchidos (como se a macro tivesse rodado no trabalho)
for i, r in enumerate(range(6, 21)):
    d[f"E{r}"] = round(3900 + (i * 137) % 1900 + 0.37 * i, 2)
d["B37"] = "Total"; d["B37"].font = Font(bold=True)
for c in "DEF":
    d[f"{c}37"] = f"=SUM({c}6:{c}36)"; d[f"{c}37"].font = Font(bold=True); d[f"{c}37"].number_format = "#,##0.00"
d.column_dimensions["B"].width = 13
for c in "CDEF": d.column_dimensions[c].width = 16
d.freeze_panes = "B6"

painel["B2"] = "Painel do mês"; painel["B2"].font = Font(bold=True, size=14)
rotulos = [("B5", "Faturamento realizado", "D5", "=Diário!E37"),
           ("B6", "Meta do mês", "D6", "=Diário!D37"),
           ("B7", "Cupons no mês", "D7", 2150),
           ("B8", "Ticket médio", "D8", '=IF(D7=0,"",D5/D7)'),
           ("B9", "Atingimento", "D9", '=IF(D6=0,"",D5/D6)')]
for a, t, b, v in rotulos:
    painel[a] = t; painel[b] = v; painel[b].font = Font(bold=True)
painel["D5"].number_format = painel["D6"].number_format = painel["D8"].number_format = "#,##0.00"
painel["D7"].number_format = "#,##0"; painel["D9"].number_format = "0.0%"
painel.column_dimensions["B"].width = 26; painel.column_dimensions["D"].width = 16

g = BarChart(); g.title = "Realizado por dia"; g.height = 7; g.width = 16
g.add_data(Reference(d, min_col=5, min_row=5, max_row=35), titles_from_data=True)
g.set_categories(Reference(d, min_col=2, min_row=6, max_row=35))
painel.add_chart(g, "F2")
if filtro:
    f = wb.create_sheet("Filtros vendas")
    f["A1"] = "Máquinas / locais a NÃO somar no faturamento"; f["A2"] = filtro
wb.save(saida)
