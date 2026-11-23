/**
 * ============================================================
 *  CONFERÊNCIA POR LOJA: total de um dia, loja a loja,
 *  incluindo as lojas que o filtro tira da conta
 *  (versão Office Scripts da macro VBA ConferirVendasPorLoja)
 *
 *  Campo ao rodar: dia (dd/mm/aaaa). Vazio = ontem.
 *  Resultado na aba "Conferência vendas" (recriada a cada vez).
 * ============================================================
 */

//@@COMUM_CONFIG@@
const ABA_CONF = "Conferência vendas";

async function main(workbook: ExcelScript.Workbook, dia?: string, token?: string): Promise<string> {
  const tk = (token || TOKEN_VMPAY || "").trim();
  if (!tk) throw new Error("Falta o token da VMpay: preencha o campo 'token' ou a constante TOKEN_VMPAY no script.");
  if (API_BASE.indexOf("SEU-USUARIO") >= 0) throw new Error("Configure API_BASE no topo do script com o endereço do seu proxy.");
  const d = dia && dia.trim() ? converterData(dia) : hojeSerial() - 1;
  if (d === null) throw new Error("Data inválida: " + dia);

  garantirFiltro(workbook);
  const acc = novoAcumulador(lerFiltro(workbook), true);
  await consultarDia(tk, d, acc);

  const antiga = acharAba(workbook, ABA_CONF);
  if (antiga) antiga.delete();
  const w = workbook.addWorksheet(ABA_CONF);

  const a1 = w.getRange("A1");
  a1.setValue(`Vendas de ${formatarData(d)} por loja (API /vends)`);
  a1.getFormat().getFont().setBold(true);
  a1.getFormat().getFont().setSize(13);
  const cab = w.getRange("A3:G3");
  cab.setValues([["Cliente", "Local", "Máquina", "Valor", "Itens", "Cupons", "Fora da conta (filtro)?"]]);
  cab.getFormat().getFont().setBold(true);
  cab.getFormat().getFont().setColor("#FFFFFF");
  cab.getFormat().getFill().setColor("#263238");

  const lojas = Array.from((acc.lojas as Map<string, Loja>).values()).sort((a, b) => b.valor - a.valor);
  let totExc = 0;
  const linhas: (string | number)[][] = lojas.map((l) => {
    if (l.excluida) totExc += l.valor;
    return [l.cliente, l.local, l.maquina, arred2(l.valor), l.itens, l.cupons.size, l.excluida ? "Sim" : "Não"];
  });
  let r = 4;
  if (linhas.length) {
    w.getRangeByIndexes(3, 0, linhas.length, 7).setValues(linhas);
    lojas.forEach((l, i) => { if (l.excluida) w.getRangeByIndexes(3 + i, 0, 1, 7).getFormat().getFill().setColor("#FDE2C4"); });
    r += linhas.length;
  }
  r++;
  const tot = w.getRangeByIndexes(r - 1, 2, 4, 2);
  tot.setValues([
    ["Total na planilha (sem as excluídas)", arred2(acc.total)],
    ["Lojas excluídas pelo filtro", arred2(totExc)],
    ["Total geral da API", arred2(acc.total + totExc)],
    ["Vendas com valor zero (brinde/troca)", acc.zeradas],
  ]);
  tot.getFormat().getFont().setBold(true);
  w.getRangeByIndexes(r - 1, 2, 4, 1).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);

  w.getRange("D:D").setNumberFormat("#,##0.00");
  w.getRangeByIndexes(r + 2, 3, 1, 1).setNumberFormat("0");
  w.getRange("E:F").setNumberFormat("#,##0");
  w.getRange("A:G").getFormat().autofitColumns();
  w.activate();

  const msg = `${formatarData(d)}: ${lojas.length} loja(s). Total na planilha ${arred2(acc.total)}; excluídas ${arred2(totExc)}.`;
  console.log(msg);
  return msg;
}
