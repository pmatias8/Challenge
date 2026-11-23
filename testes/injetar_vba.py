"""Transforma um .xlsx em .xlsm com um vbaProject.bin FALSO (só para testar
que o app devolve o arquivo de macros intacto). Não abra o resultado no Excel.
Uso: python3 injetar_vba.py entrada.xlsx saida.xlsm"""
import sys, zipfile, os
ent, sai = sys.argv[1], sys.argv[2]
falso = bytes((i * 37 + 11) % 256 for i in range(20000))
with zipfile.ZipFile(ent) as zi, zipfile.ZipFile(sai, "w", zipfile.ZIP_DEFLATED) as zo:
    for it in zi.infolist():
        dados = zi.read(it.filename)
        if it.filename == "[Content_Types].xml":
            t = dados.decode()
            t = t.replace("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                          "application/vnd.ms-excel.sheet.macroEnabled.main+xml")
            t = t.replace("</Types>", '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>')
            dados = t.encode()
        if it.filename == "xl/_rels/workbook.xml.rels":
            dados = dados.decode().replace("</Relationships>",
                '<Relationship Id="rIdVBA" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>').encode()
        zo.writestr(it, dados)
    zo.writestr("xl/vbaProject.bin", falso)
print("ok", sai)
