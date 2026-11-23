/**
 * ============================================================
 *  ATUALIZAR VENDAS DA API VMPAY  ->  planilha Faturamento
 *  (versão Office Scripts da macro VBA AtualizarVendasAPI)
 *
 *  - Consulta GET /api/v1/vends dia a dia (dia de Brasília).
 *  - Soma o valor das vendas e escreve na coluna Realizado (E)
 *    da aba "Diário", na linha da data (coluna B).
 *  - Guarda tudo na aba "Vendas API" (data, faturamento, itens,
 *    cupons estimados, valor que estava na planilha, diferença).
 *  - Lojas a excluir: aba "Filtros vendas" (criada na 1a vez).
 *
 *  Diferenças em relação ao VBA:
 *  - Office Scripts não têm InputBox/MsgBox. As perguntas viram os
 *    campos que o Excel mostra ao rodar o script (todos opcionais):
 *      dataInicio / dataFim (dd/mm/aaaa)  -> vazio = mesma sugestão da macro
 *      substituirPreenchidos              -> "Sim" = substituir pelo valor da API
 *      atualizarCuponsPainel              -> "Sim" = gravar a estimativa em Painel!D7
 *      token                              -> só se não estiver em TOKEN_VMPAY
 *  - O resumo da execução fica em "Vendas API"!I1:I4 (e no painel de saída).
 *  - A API da VMpay não aceita chamadas de navegador (CORS), então o
 *    endereço abaixo deve ser o do seu proxy (veja proxy/README).
 * ============================================================
 */

//@@COMUM_CONFIG@@
const ABA_PROJ = "Diário";
const COL_DATA = 2;             // B
const COL_REAL = 5;             // E
const LIN_INI = 6;
const LIN_FIM = 36;
const ABA_TICKET = "Painel";
const CEL_TICKET = "D7";
const ABA_LOG = "Vendas API";

// ============================================================
//  MACRO PRINCIPAL
// ============================================================
async function main(
  workbook: ExcelScript.Workbook,
  dataInicio?: string,
  dataFim?: string,
  substituirPreenchidos: "Não" | "Sim" = "Não",
  atualizarCuponsPainel: "Não" | "Sim" = "Não",
  token?: string
): Promise<string> {
  const t0 = Date.now();
  const ws = acharAba(workbook, ABA_PROJ);
  if (!ws) throw new Error(`Abra a planilha de faturamento (aba '${ABA_PROJ}') antes de rodar.`);

  const tk = limparToken(token || TOKEN_VMPAY || "");
  if (!tk) throw new Error("Falta o token da VMpay: preencha o campo 'token' ou a constante TOKEN_VMPAY no script.");
  if (API_BASE.indexOf("SEU-USUARIO") >= 0) throw new Error("Configure API_BASE no topo do script com o endereço do seu proxy.");

  const hoje = hojeSerial();

  // --- Dias que a aba tem (coluna B) ---
  const vals = ws.getRangeByIndexes(LIN_INI - 1, 0, LIN_FIM - LIN_INI + 1, COL_REAL).getValues();
  const linhaDia = new Map<number, number>();
  let primDia = 0, ultDia = 0, primVazio = 0;
  for (let i = 0; i < vals.length; i++) {
    const d = lerData(vals[i][COL_DATA - 1]);
    if (d === null) continue;
    if (!linhaDia.has(d)) linhaDia.set(d, LIN_INI + i);
    if (!primDia) primDia = d;
    ultDia = d;
    if (!primVazio && vals[i][COL_REAL - 1] === "" && d < hoje) primVazio = d;
  }
  if (linhaDia.size === 0) throw new Error(`Não encontrei datas em ${ABA_PROJ}!B${LIN_INI}:B${LIN_FIM}.`);
  const realizado = (d: number) => vals[(linhaDia.get(d) as number) - LIN_INI][COL_REAL - 1];

  // --- Período ---
  let sugFim = Math.min(hoje - 1, ultDia);
  let sugIni = primVazio || sugFim;
  if (sugIni < primDia) sugIni = primDia;
  const dIni = dataInicio && dataInicio.trim() ? converterData(dataInicio) : sugIni;
  const dFim = dataFim && dataFim.trim() ? converterData(dataFim) : sugFim;
  if (dIni === null) throw new Error("Data de início inválida: " + dataInicio);
  if (dFim === null) throw new Error("Data de fim inválida: " + dataFim);
  if (dFim < dIni) throw new Error("A data de fim tem que ser igual ou depois do início.");
  const avisos: string[] = [];
  if (dFim >= hoje) avisos.push("O dia de hoje ainda não fechou e entrou parcial.");

  // --- Dias já preenchidos no período ---
  let jaTem = 0, foraAba = 0;
  for (let d = dIni; d <= dFim; d++) {
    if (linhaDia.has(d)) { if (realizado(d) !== "") jaTem++; } else foraAba++;
  }
  if (foraAba === dFim - dIni + 1) throw new Error(`Nenhum dia do período está na aba '${ABA_PROJ}' (coluna B).`);
  const sobrescrever = substituirPreenchidos === "Sim";

  garantirFiltro(workbook);
  const excluir = lerFiltro(workbook);

  // --- Consulta todos os dias primeiro (se der erro, nada é gravado) ---
  const dias: ResultadoDia[] = [];
  for (let d = dIni; d <= dFim; d++) {
    if (!linhaDia.has(d)) continue;
    console.log(`VMpay vendas: ${formatarData(d)}...`);
    const acc = novoAcumulador(excluir, false);
    await consultarDia(tk, d, acc);
    dias.push({ dia: d, valor: acc.total, itens: acc.itens, cupons: acc.chaves.size });
  }

  // --- Grava ---
  const wl = garantirLog(workbook);
  const agora = agoraSerial();
  let gravados = 0, mantidos = 0;
  for (const x of dias) {
    const linha = linhaDia.get(x.dia) as number;
    const antes = realizado(x.dia);
    if (antes === "" || sobrescrever) {
      ws.getCell(linha - 1, COL_REAL - 1).setValue(arred2(x.valor));
      gravados++;
    } else mantidos++;
    gravarLog(wl, x, antes, agora);
  }

  // --- Ticket (D7): soma dos cupons do mês no log ---
  const t = somarCuponsMes(wl, primDia, Math.min(hoje - 1, ultDia));
  let msgTicket = "";
  const painel = acharAba(workbook, ABA_TICKET);
  if (t.diasLog === t.diasMes && t.diasMes > 0 && painel) {
    const atual = painel.getRange(CEL_TICKET).getValue();
    if (atualizarCuponsPainel === "Sim") {
      painel.getRange(CEL_TICKET).setValue(t.cup);
      msgTicket = `Cupons do Painel atualizados: ${fmtInt(Number(atual) || 0)} -> ${fmtInt(t.cup)}.`;
    } else {
      msgTicket = `Cupons estimados no mês (até ontem): ${fmtInt(t.cup)} | no Painel: ${fmtInt(Number(atual) || 0)}. ` +
                  `Rode de novo com atualizarCuponsPainel = Sim para gravar.`;
    }
  }

  let msg = `${gravados} dia(s) gravado(s) em Realizado`;
  if (mantidos > 0) msg += `, ${mantidos} mantido(s) como estavam (veja a aba '${ABA_LOG}')`;
  if (jaTem > 0 && !sobrescrever) msg += ". Para substituir, rode com substituirPreenchidos = Sim";
  msg += `. Período ${formatarData(dIni)} a ${formatarData(dFim)}. Tempo: ${Math.round((Date.now() - t0) / 1000)}s.`;
  const linhas = [msg, msgTicket, avisos.join(" "), `Rodado em ${formatarDataHora(agora)}`];
  wl.getRange("I1:I4").setValues(linhas.map((l) => [l]));
  linhas.filter((l) => l).forEach((l) => console.log(l));
  return linhas.filter((l) => l).join("\n");
}

// ============================================================
//  ABA DE LOG
// ============================================================
function garantirLog(wb: ExcelScript.Workbook): ExcelScript.Worksheet {
  let w = acharAba(wb, ABA_LOG);
  if (w) return w;
  w = wb.addWorksheet(ABA_LOG);
  const cab = w.getRange("A1:G1");
  cab.setValues([["Data", "Faturamento API", "Itens", "Cupons (estimado)", "Valor na planilha antes", "Diferença (API - planilha)", "Atualizado em"]]);
  cab.getFormat().getFont().setBold(true);
  cab.getFormat().getFont().setColor("#FFFFFF");
  cab.getFormat().getFill().setColor("#263238");
  cab.getFormat().setWrapText(true);
  cab.getFormat().setRowHeight(30);
  w.getRange("A:A").getFormat().setColumnWidth(70);
  w.getRange("B:F").getFormat().setColumnWidth(90);
  w.getRange("G:G").getFormat().setColumnWidth(95);
  w.getRange("A:A").setNumberFormat("dd/mm/yyyy");
  w.getRange("B:B").setNumberFormat("#,##0.00");
  w.getRange("C:D").setNumberFormat("#,##0");
  w.getRange("E:F").setNumberFormat("#,##0.00");
  w.getRange("G:G").setNumberFormat("dd/mm/yyyy hh:mm");
  w.getRange("A1:G1").setNumberFormat("General");
  w.getFreezePanes().freezeRows(1);
  return w;
}

function gravarLog(w: ExcelScript.Worksheet, x: ResultadoDia, antes: string | number | boolean, agora: number): void {
  const ult = ultimaLinhaA(w);
  let alvo = 0;
  if (ult >= 2) {
    const a = w.getRangeByIndexes(1, 0, ult - 1, 1).getValues();
    for (let i = 0; i < a.length; i++) if (lerData(a[i][0]) === x.dia) { alvo = i + 2; break; }
  }
  if (!alvo) alvo = ult + 1;
  const temAntes = typeof antes === "number";
  w.getRangeByIndexes(alvo - 1, 0, 1, 7).setValues([[
    x.dia, arred2(x.valor), x.itens, x.cupons,
    temAntes ? (antes as number) : "", temAntes ? arred2(x.valor - (antes as number)) : "", agora,
  ]]);
}

function somarCuponsMes(w: ExcelScript.Worksheet, ini: number, fim: number): { cup: number; diasLog: number; diasMes: number } {
  let cup = 0, diasLog = 0;
  const diasMes = Math.max(0, fim - ini + 1);
  const ult = ultimaLinhaA(w);
  if (ult < 2) return { cup, diasLog, diasMes };
  const v = w.getRangeByIndexes(1, 0, ult - 1, 4).getValues();
  for (const linha of v) {
    const d = lerData(linha[0]);
    if (d !== null && d >= ini && d <= fim) { cup += Math.round(Number(linha[3]) || 0); diasLog++; }
  }
  return { cup, diasLog, diasMes };
}

