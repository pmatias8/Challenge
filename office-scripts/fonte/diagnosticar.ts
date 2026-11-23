/**
 * ============================================================
 *  DIAGNÓSTICO: grava a resposta crua de 1 página de vendas
 *  (versão Office Scripts da macro VBA DiagnosticarVendasAPI)
 *
 *  Campo ao rodar: dia (dd/mm/aaaa). Vazio = ontem.
 *  Resultado na aba "DiagVendas" (recriada a cada vez).
 * ============================================================
 */

//@@COMUM_CONFIG@@

async function main(workbook: ExcelScript.Workbook, dia?: string, token?: string): Promise<string> {
  const tk = limparToken(token || TOKEN_VMPAY || "");
  if (!tk) throw new Error("Falta o token da VMpay: preencha o campo 'token' ou a constante TOKEN_VMPAY no script.");
  if (API_BASE.indexOf("SEU-USUARIO") >= 0) throw new Error("Configure API_BASE no topo do script com o endereço do seu proxy.");
  const d = dia && dia.trim() ? converterData(dia) : hojeSerial() - 1;
  if (d === null) throw new Error("Data inválida: " + dia);

  const r = await fetch(urlVends(tk, d, 1, 5));
  const corpo = await r.text();
  if (r.status !== 200) throw new Error(`HTTP ${r.status}: ${corpo.substring(0, 300)}`);

  const antiga = acharAba(workbook, "DiagVendas");
  if (antiga) antiga.delete();
  const w = workbook.addWorksheet("DiagVendas");
  w.getRange("A1").setValue(`Resposta crua de /vends (5 primeiras vendas de ${formatarData(d)})`);
  w.getRange("A1").getFormat().getFont().setBold(true);
  const texto = corpo.split("},{").join("},\n\n{").split(",\"").join(",\n\"");
  const pedacos: string[][] = [];
  for (let p = 0; p < texto.length; p += 30000) pedacos.push(["'" + texto.substring(p, p + 30000)]);
  if (pedacos.length) {
    const alvo = w.getRangeByIndexes(2, 0, pedacos.length, 1);
    alvo.setValues(pedacos);
    alvo.getFormat().setWrapText(true);
  }
  w.getRange("A:A").getFormat().setColumnWidth(640);
  w.activate();
  return "Resposta gravada na aba DiagVendas.";
}
