"""Confere a planilha que o app devolveu.
Uso: python3 verificar_saida.py entrada.xlsm saida.xlsm esperado.json pasta_tmp"""
import sys, json, zipfile, subprocess, os, openpyxl
ent, sai, esp, tmp = sys.argv[1:5]
esperado = {int(k): v for k, v in json.load(open(esp)).items()}
def ok(c, m):
    if not c: raise SystemExit("FALHOU: " + m)
    print("  ✓", m)

zi, zo = zipfile.ZipFile(ent), zipfile.ZipFile(sai)
ok(zi.read("xl/vbaProject.bin") == zo.read("xl/vbaProject.bin"), "vbaProject.bin (macros) idêntico ao original")
ok(zi.read("xl/charts/chart1.xml") == zo.read("xl/charts/chart1.xml"), "gráfico intacto")
ok("macroEnabled" in zo.read("[Content_Types].xml").decode(), "continua sendo .xlsm")

a = openpyxl.load_workbook(ent); b = openpyxl.load_workbook(sai)
ok(b.sheetnames[:len(a.sheetnames)] == a.sheetnames and "Vendas API" in b.sheetnames and "Conferência vendas" in b.sheetnames,
   "abas: " + ", ".join(b.sheetnames))
da, db = a["Diário"], b["Diário"]
ok(all(da.cell(r, 5).value == db.cell(r, 5).value for r in range(6, 21)), "dias 01–15 mantidos")
ok(all(db.cell(r, 5).value == esperado[r - 5]["valor"] for r in range(21, 28)),
   "dias 16–22 = soma esperada (sem a loja filtrada): " + str([db.cell(r, 5).value for r in range(21, 28)]))
ok(all(db.cell(r, 5).value is None for r in range(28, 36)), "dias 23–30 continuam vazios")
ok(all(da.cell(r, c).value == db.cell(r, c).value for r in range(1, 40) for c in range(1, 8) if c != 5), "fórmulas e demais células iguais")
ok(db["B6"].number_format == da["B6"].number_format and db["E21"].number_format == da["E21"].number_format, "formatação preservada")
log = b["Vendas API"]
linhas = [[c.value for c in r] for r in log.iter_rows(min_row=2, max_col=7)]
ok(len(linhas) == 22, "log com 22 dias")
por_dia = {r[0].day: r for r in linhas}
ok(all(por_dia[d][3] == esperado[d]["cupons"] for d in range(16, 23)), "cupons estimados corretos")
ok(log["A2"].number_format != "General" and log["B2"].number_format != "General", "log com formato de data/número: %s / %s" % (log["A2"].number_format, log["B2"].number_format))
ok(b["Painel"]["D7"].value == sum(r[3] for r in linhas), "Painel!D7 = soma dos cupons do log (%s)" % b["Painel"]["D7"].value)

# Recalcula no LibreOffice (outro programa) para provar que o arquivo abre e as fórmulas funcionam
os.makedirs(tmp, exist_ok=True)
xlsx = os.path.join(tmp, "saida.xlsx"); open(xlsx, "wb").write(open(sai, "rb").read())
# perfil próprio do LibreOffice com "recalcular sempre ao abrir" (o padrão dele é usar os valores guardados)
perfil = os.path.join(tmp, "perfil-lo", "user")
os.makedirs(perfil, exist_ok=True)
open(os.path.join(perfil, "registrymodifications.xcu"), "w").write(
    '<?xml version="1.0" encoding="UTF-8"?><oor:items xmlns:oor="http://openoffice.org/2001/registry" '
    'xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
    '<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse">'
    '<value>0</value></prop></item></oor:items>')
subprocess.run(["soffice", "-env:UserInstallation=file://" + os.path.abspath(os.path.join(tmp, "perfil-lo")), "--headless", "--calc",
                "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", os.path.join(tmp, "lo"), xlsx],
               check=True, capture_output=True, timeout=180)
c = openpyxl.load_workbook(os.path.join(tmp, "lo", "saida.xlsx"), data_only=True)
soma = round(sum(c["Diário"].cell(r, 5).value or 0 for r in range(6, 37)), 2)
ok(round(c["Painel"]["D5"].value, 2) == soma, "LibreOffice abriu e recalculou: Painel!D5 = %s" % soma)
print("OK - arquivo de saída conferido")
