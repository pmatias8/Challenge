// ARQUIVO GERADO por office-scripts/montar.mjs a partir de office-scripts/fonte/ — copie e cole inteiro no Excel.
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

// ---------- CONFIGURE AQUI ----------
const API_BASE = "https://vmpay-proxy.SEU-USUARIO.workers.dev/api/v1";
// O script fica guardado no SEU OneDrive. Se você adicionar um botão deste
// script numa planilha compartilhada, o script (e este token) passa a ser
// compartilhado com quem abre a planilha. Nesse caso deixe vazio e digite o
// token no campo "token" ao rodar.
const TOKEN_VMPAY = "";
// ------------------------------------

const OFFSET_UTC = -3;
const PER_PAGE = 1000;
const LOTE = 8;                 // páginas pedidas em paralelo

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

  const tk = (token || TOKEN_VMPAY || "").trim();
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

const ABA_FILTRO = "Filtros vendas";

// ============================================================
//  CONSULTA DE UM DIA (paginação em lotes paralelos)
// ============================================================
interface Ref { name?: string | number | null; asset_number?: string | number | null; }
interface Venda {
  value?: number | string | null;
  quantity?: number | string | null;
  purchase_id?: string | number | null;
  order_id?: string | number | null;
  cashless_transaction_id?: string | number | null;
  machine_id?: string | number | null;
  occurred_at?: string | null;
  machine?: Ref | null;
  location?: Ref | null;
  client?: Ref | null;
}
interface Loja { cliente: string; local: string; maquina: string; valor: number; itens: number; cupons: Set<string>; excluida: boolean; }
interface Acumulador { excluir: Set<string>; total: number; itens: number; chaves: Set<string>; lojas: Map<string, Loja> | null; zeradas: number; }
interface ResultadoDia { dia: number; valor: number; itens: number; cupons: number; }
interface Pagina { lista: Venda[]; erro: string; }

function novoAcumulador(excluir: Set<string>, conferencia: boolean): Acumulador {
  return { excluir, total: 0, itens: 0, chaves: new Set<string>(), lojas: conferencia ? new Map<string, Loja>() : null, zeradas: 0 };
}

function urlVends(token: string, dia: number, pagina: number, porPagina: number): string {
  const ini = dia - OFFSET_UTC / 24;
  const fim = dia + 1 - OFFSET_UTC / 24 - 1 / 86400;
  return `${API_BASE}/vends?access_token=${encodeURIComponent(token)}` +
    `&start_date=${encodeURIComponent(formatarDataHora(ini))}` +
    `&end_date=${encodeURIComponent(formatarDataHora(fim))}` +
    `&page=${pagina}&per_page=${porPagina}`;
}

async function pedirPagina(url: string, dia: number, pagina: number): Promise<Pagina> {
  try {
    const r = await fetch(url);
    if (r.status === 401 || r.status === 403) return { lista: [], erro: `A VMpay recusou o token (HTTP ${r.status}).` };
    if (r.status !== 200) {
      const tx = (await r.text()).substring(0, 200);
      return { lista: [], erro: `HTTP ${r.status} em ${formatarData(dia)} (página ${pagina}): ${tx}` };
    }
    const lista: Venda[] = await r.json();
    if (!Array.isArray(lista)) return { lista: [], erro: "Resposta da API não é uma lista JSON." };
    return { lista, erro: "" };
  } catch (e) {
    return { lista: [], erro: `Não consegui falar com a API em ${formatarData(dia)} (${String(e)}). Confira API_BASE (proxy).` };
  }
}

async function consultarDia(token: string, dia: number, acc: Acumulador): Promise<void> {
  let pag0 = 0;
  for (;;) {
    const pedidos: Promise<Pagina>[] = [];
    for (let i = 1; i <= LOTE; i++) pedidos.push(pedirPagina(urlVends(token, dia, pag0 + i, PER_PAGE), dia, pag0 + i));
    for (let i = 0; i < LOTE; i++) {
      const p = await pedidos[i];
      if (p.erro) throw new Error(p.erro);
      for (const it of p.lista) somarVenda(acc, it);
      if (p.lista.length < PER_PAGE) return; // páginas seguintes do lote são ignoradas
    }
    pag0 += LOTE;
  }
}

function somarVenda(acc: Acumulador, it: Venda): void {
  if (!it || typeof it !== "object") return;
  const machine = it.machine || null, location = it.location || null, client = it.client || null;

  let excl = false;
  if (acc.excluir.size > 0) {
    excl = excluido(acc, txt(machine ? machine.name : null)) ||
           excluido(acc, txt(machine ? machine.asset_number : null)) ||
           excluido(acc, txt(location ? location.name : null));
  }
  const valor = typeof it.value === "number" ? it.value : 0;
  const qtd = typeof it.quantity === "number" ? it.quantity : 1;

  // Cupom: a API não manda id da compra; itens do mesmo carrinho saem com a
  // mesma máquina e o mesmo horário (até o segundo).
  let k = txt(it.purchase_id) || txt(it.order_id) || txt(it.cashless_transaction_id);
  if (!k) k = txt(it.machine_id) + "|" + txt(it.occurred_at).substring(0, 19);

  if (acc.lojas) {
    const cli = txt(client ? client.name : null), loc = txt(location ? location.name : null), maq = txt(machine ? machine.asset_number : null);
    const lj = `${cli} | ${loc} | ${maq}`;
    let l = acc.lojas.get(lj);
    if (!l) { l = { cliente: cli, local: loc, maquina: maq, valor: 0, itens: 0, cupons: new Set<string>(), excluida: false }; acc.lojas.set(lj, l); }
    l.valor += valor; l.itens += qtd; l.cupons.add(k); l.excluida = excl;
    if (valor <= 0) acc.zeradas++;
  }
  if (excl) return;
  acc.total += valor;
  acc.itens += qtd;
  acc.chaves.add(k);
}

function excluido(acc: Acumulador, s: string): boolean {
  return s !== "" && acc.excluir.has(normalizar(s));
}

// ============================================================
//  FILTRO DE LOJAS
// ============================================================
function garantirFiltro(wb: ExcelScript.Workbook): void {
  if (acharAba(wb, ABA_FILTRO)) return;
  const w = wb.addWorksheet(ABA_FILTRO);
  const a1 = w.getRange("A1");
  a1.setValue("Máquinas / locais a NÃO somar no faturamento");
  a1.getFormat().getFont().setBold(true);
  a1.getFormat().getFont().setColor("#FFFFFF");
  a1.getFormat().getFill().setColor("#263238");
  w.getRange("A:A").getFormat().setColumnWidth(255);
  // Começa vazia: o total bate com o relatório da VMpay.
  const c1 = w.getRange("C1");
  c1.setValue("Vazia = soma todas as lojas, igual ao relatório da VMpay. Coloque aqui (a partir da linha 2)" +
              " o nome de uma máquina ou local só se quiser tirar essa loja do faturamento.");
  c1.getFormat().setWrapText(true);
  w.getRange("C:C").getFormat().setColumnWidth(320);
}

function lerFiltro(wb: ExcelScript.Workbook): Set<string> {
  const s = new Set<string>();
  const w = acharAba(wb, ABA_FILTRO);
  if (!w) return s;
  const ult = ultimaLinhaA(w);
  if (ult < 2) return s;
  for (const linha of w.getRangeByIndexes(1, 0, ult - 1, 1).getValues()) {
    const v = normalizar(String(linha[0]));
    if (v) s.add(v);
  }
  return s;
}

// ============================================================
//  HELPERS
// ============================================================
function acharAba(wb: ExcelScript.Workbook, nome: string): ExcelScript.Worksheet | undefined {
  const alvo = nome.toLowerCase();
  return wb.getWorksheets().find((w) => w.getName().toLowerCase() === alvo);
}

function txt(v: string | number | boolean | null | undefined): string {
  return v === null || v === undefined ? "" : String(v);
}

function normalizar(s: string): string {
  const de = "áàãâäéêèëíìîóôõòúùüûçñ", para = "aaaaaeeeeiiioooouuuucn";
  let r = "";
  for (const c of s.trim().toLowerCase()) { const i = de.indexOf(c); r += i >= 0 ? para[i] : c; }
  return r;
}

const DIA_MS = 86400000;
const EPOCA = 25569; // serial do Excel de 01/01/1970

function hojeSerial(): number { return Math.floor(Date.now() / DIA_MS + EPOCA + OFFSET_UTC / 24); }
function agoraSerial(): number { return Date.now() / DIA_MS + EPOCA + OFFSET_UTC / 24; }

function lerData(v: string | number | boolean): number | null {
  if (typeof v === "number" && v > 0) return Math.round(v);
  if (typeof v === "string" && v.trim()) return converterData(v);
  return null;
}

function converterData(s: string): number | null {
  const p = s.trim().split("/");
  if (p.length !== 3) return null;
  const d = Number(p[0]), m = Number(p[1]), a = Number(p[2]);
  if (!Number.isInteger(d) || !Number.isInteger(m) || !Number.isInteger(a) || m < 1 || m > 12 || d < 1 || d > 31 || a < 1900) return null;
  return Date.UTC(a, m - 1, d) / DIA_MS + EPOCA;
}

function dois(n: number): string { return (n < 10 ? "0" : "") + n; }
function dataDe(serial: number): Date { return new Date(Math.round((serial - EPOCA) * 86400) * 1000); }
function formatarData(serial: number): string {
  const d = dataDe(serial);
  return `${dois(d.getUTCDate())}/${dois(d.getUTCMonth() + 1)}/${d.getUTCFullYear()}`;
}
function formatarDataHora(serial: number): string {
  const d = dataDe(serial);
  return `${formatarData(serial)} ${dois(d.getUTCHours())}:${dois(d.getUTCMinutes())}:${dois(d.getUTCSeconds())}`;
}
function arred2(x: number): number { return Math.round((x + Number.EPSILON) * 100) / 100; }
function fmtInt(x: number): string { return Math.round(x).toLocaleString("pt-BR"); }

function ultimaLinhaA(w: ExcelScript.Worksheet): number {
  const usado = w.getUsedRange(true);
  if (!usado) return 1;
  const fim = usado.getRowIndex() + usado.getRowCount();
  const a = w.getRangeByIndexes(0, 0, fim, 1).getValues();
  for (let r = a.length - 1; r >= 0; r--) if (a[r][0] !== "") return r + 1;
  return 1;
}
