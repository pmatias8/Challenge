// ============================================================
//  As 3 macros do modVendasAPI, operando sobre a planilha aberta
//  no aparelho. As perguntas (InputBox / MsgBox) ficam na tela do
//  app; aqui fica só a regra.
// ============================================================

import { consultarDia, novoAcumulador, normalizar, hojeSerial, agoraSerial, converterData, diagnosticar } from "./vendas.js";

export const ABA_PROJ = "Diário";
export const COL_DATA = 2;   // B
export const COL_REAL = 5;   // E
export const LIN_INI = 6;
export const LIN_FIM = 36;
export const ABA_TICKET = "Painel";
export const CEL_TICKET = "D7";
export const ABA_LOG = "Vendas API";
export const ABA_FILTRO = "Filtros vendas";
export const ABA_CONF = "Conferência vendas";

const COR_CAB = "263238", COR_EXCL = "FDE2C4", BRANCO = "FFFFFF";
const arred2 = (x) => Math.round((x + Number.EPSILON) * 100) / 100;

/** Data de uma célula -> serial (sistema 1900) ou null. */
function lerData(v, off) {
  if (typeof v === "number" && v > 0) return Math.round(v) + off;
  if (typeof v === "string") return converterData(v);
  return null;
}

// ------------------------------------------------------------
//  Leitura da aba Diário + sugestão de período
// ------------------------------------------------------------
export async function analisar(pasta, hoje = hojeSerial()) {
  if (!pasta.temAba(ABA_PROJ)) throw new Error(`A planilha não tem a aba '${ABA_PROJ}'. Abra a planilha de faturamento.`);
  const ws = await pasta.aba(ABA_PROJ);
  const off = pasta.data1904 ? 1462 : 0;
  const linhaDia = new Map();
  let primDia = 0, ultDia = 0, primVazio = 0;
  for (let r = LIN_INI; r <= LIN_FIM; r++) {
    const d = lerData(ws.valorEm(r, COL_DATA), off);
    if (d === null) continue;
    if (!linhaDia.has(d)) linhaDia.set(d, r);
    if (!primDia) primDia = d;
    ultDia = d;
    if (!primVazio && ws.valorEm(r, COL_REAL) === null && d < hoje) primVazio = d;
  }
  if (!linhaDia.size) throw new Error(`Não encontrei datas em ${ABA_PROJ}!B${LIN_INI}:B${LIN_FIM}.`);

  let sugFim = Math.min(hoje - 1, ultDia);
  let sugIni = primVazio || sugFim;
  if (sugIni < primDia) sugIni = primDia;

  const preenchidos = [...linhaDia].filter(([, r]) => ws.valorEm(r, COL_REAL) !== null).length;
  return { ws, off, linhaDia, primDia, ultDia, primVazio, sugIni, sugFim, hoje, preenchidos };
}

/** Quantos dias do período já têm valor e quantos estão fora da aba. */
export function verificarPeriodo(info, ini, fim) {
  let jaTem = 0, foraAba = 0;
  for (let d = ini; d <= fim; d++) {
    const r = info.linhaDia.get(d);
    if (r === undefined) foraAba++;
    else if (info.ws.valorEm(r, COL_REAL) !== null) jaTem++;
  }
  return { jaTem, foraAba, todosFora: foraAba === fim - ini + 1, incluiHoje: fim >= info.hoje };
}

// ------------------------------------------------------------
//  Abas auxiliares
// ------------------------------------------------------------
export async function garantirFiltro(pasta) {
  if (pasta.temAba(ABA_FILTRO)) return false;
  const w = await pasta.criarAba(ABA_FILTRO);
  const cab = await pasta.estilo({ negrito: true, corFonte: BRANCO, fundo: COR_CAB });
  const quebra = await pasta.estilo({ quebrar: true });
  w.definir("A1", "Máquinas / locais a NÃO somar no faturamento", cab);
  w.definir("C1", "Vazia = soma todas as lojas, igual ao relatório da VMpay. Coloque aqui (a partir da linha 2)" +
                  " o nome de uma máquina ou local só se quiser tirar essa loja do faturamento.", quebra);
  w.formatarColunas(1, 1, 48);
  w.formatarColunas(3, 3, 60);
  return true;
}

export async function lerFiltro(pasta) {
  const set = new Set();
  const w = await pasta.aba(ABA_FILTRO);
  if (!w) return set;
  const ult = w.ultimaLinha(1);
  for (let i = 2; i <= ult; i++) {
    const v = normalizar(w.valorEm(i, 1) ?? "");
    if (v) set.add(v);
  }
  return set;
}

async function garantirLog(pasta) {
  let w = await pasta.aba(ABA_LOG);
  if (w) return w;
  w = await pasta.criarAba(ABA_LOG);
  const cab = await pasta.estilo({ negrito: true, corFonte: BRANCO, fundo: COR_CAB, quebrar: true });
  ["Data", "Faturamento API", "Itens", "Cupons (estimado)", "Valor na planilha antes",
   "Diferença (API - planilha)", "Atualizado em"].forEach((t, i) => w.definirEm(1, i + 1, t, cab));
  w.alturaLinha(1, 30);
  w.formatarColunas(1, 1, 12, await pasta.estilo({ numFmt: 14 }));   // dd/mm/aaaa
  w.formatarColunas(2, 2, 16, await pasta.estilo({ numFmt: 4 }));    // #.##0,00
  w.formatarColunas(3, 4, 16, await pasta.estilo({ numFmt: 3 }));    // #.##0
  w.formatarColunas(5, 6, 16, await pasta.estilo({ numFmt: 4 }));
  w.formatarColunas(7, 7, 17, await pasta.estilo({ numFmt: 22 }));   // dd/mm/aaaa hh:mm
  w.congelarLinha1();
  return w;
}

function gravarLog(w, off, dia, valor, itens, cupons, antes, agora) {
  const ult = w.ultimaLinha(1);
  let alvo = 0;
  for (let r = 2; r <= ult; r++) {
    const d = lerData(w.valorEm(r, 1), off);
    if (d === dia) { alvo = r; break; }
  }
  if (!alvo) alvo = ult + 1;
  w.definirEm(alvo, 1, dia - off);
  w.definirEm(alvo, 2, arred2(valor));
  w.definirEm(alvo, 3, itens);
  w.definirEm(alvo, 4, cupons);
  if (typeof antes === "number") {
    w.definirEm(alvo, 5, antes);
    w.definirEm(alvo, 6, arred2(valor - antes));
  } else {
    w.definirEm(alvo, 5, null);
    w.definirEm(alvo, 6, null);
  }
  w.definirEm(alvo, 7, agora - off);
}

function somarCuponsMes(w, off, ini, fim) {
  let cup = 0, diasLog = 0;
  const diasMes = Math.max(0, fim - ini + 1);
  const ult = w.ultimaLinha(1);
  for (let r = 2; r <= ult; r++) {
    const d = lerData(w.valorEm(r, 1), off);
    if (d !== null && d >= ini && d <= fim) {
      cup += Math.round(Number(w.valorEm(r, 4)) || 0);
      diasLog++;
    }
  }
  return { cup, diasLog, diasMes };
}

// ------------------------------------------------------------
//  MACRO PRINCIPAL: AtualizarVendasAPI
// ------------------------------------------------------------
export async function atualizarVendas(pasta, info, { token, base, ini, fim, sobrescrever, fetchImpl, progresso = () => {} }) {
  await garantirFiltro(pasta);
  const excluir = await lerFiltro(pasta);

  // 1) consulta todos os dias primeiro (se der erro, nada é gravado)
  const dias = [];
  const total = fim - ini + 1;
  for (let d = ini, i = 1; d <= fim; d++, i++) {
    if (!info.linhaDia.has(d)) continue;
    progresso({ dia: d, i, n: total, pagina: 0 });
    const acc = novoAcumulador(excluir);
    await consultarDia({ base, token, dia: d, acc, fetchImpl, aoPaginar: (pg) => progresso({ dia: d, i, n: total, pagina: pg }) });
    dias.push({ dia: d, valor: acc.total, itens: acc.itens, cupons: acc.chaves.size });
  }

  // 2) grava
  const ws = info.ws, off = info.off;
  const wl = await garantirLog(pasta);
  const agora = agoraSerial();
  let gravados = 0, mantidos = 0;
  for (const x of dias) {
    const linha = info.linhaDia.get(x.dia);
    const antes = ws.valorEm(linha, COL_REAL);
    x.antes = antes;
    if (antes === null || sobrescrever) {
      ws.definirEm(linha, COL_REAL, arred2(x.valor));
      x.gravado = true; gravados++;
    } else { x.gravado = false; mantidos++; }
    gravarLog(wl, off, x.dia, x.valor, x.itens, x.cupons, antes, agora);
  }

  // 3) ticket do mês: soma dos cupons no log
  const fimMes = Math.min(info.hoje - 1, info.ultDia);
  const t = somarCuponsMes(wl, off, info.primDia, fimMes);
  let ticket = null;
  if (t.diasLog === t.diasMes && t.diasMes > 0 && pasta.temAba(ABA_TICKET)) {
    const painel = await pasta.aba(ABA_TICKET);
    ticket = { estimado: t.cup, atual: painel.valor(CEL_TICKET) };
  }
  return { dias, gravados, mantidos, ticket };
}

export async function aplicarTicket(pasta, valor) {
  (await pasta.aba(ABA_TICKET)).definir(CEL_TICKET, valor);
}

// ------------------------------------------------------------
//  ConferirVendasPorLoja
// ------------------------------------------------------------
export async function conferirPorLoja(pasta, { token, base, dia, fetchImpl, progresso = () => {} }) {
  await garantirFiltro(pasta);
  const acc = novoAcumulador(await lerFiltro(pasta), true);
  await consultarDia({ base, token, dia, acc, fetchImpl, aoPaginar: (pg) => progresso({ dia, pagina: pg }) });
  const lojas = [...acc.conf.lojas.values()].map((l) => ({ ...l, cupons: l.cupons.size }))
    .sort((a, b) => b.valor - a.valor);
  const totExc = lojas.filter((l) => l.excluida).reduce((s, l) => s + l.valor, 0);
  return { dia, lojas, total: acc.total, totExc, zeradas: acc.conf.zeradas };
}

/** Grava o resultado da conferência na aba "Conferência vendas" (como a macro faz). */
export async function gravarConferencia(pasta, res, formatarData) {
  let w = await pasta.aba(ABA_CONF);
  if (w) w.limparTudo(); else w = await pasta.criarAba(ABA_CONF);
  const titulo = await pasta.estilo({ negrito: true });
  const cab = await pasta.estilo({ negrito: true, corFonte: BRANCO, fundo: COR_CAB });
  const dinheiro = await pasta.estilo({ numFmt: 4 }), inteiro = await pasta.estilo({ numFmt: 3 });
  const dinheiroExc = await pasta.estilo({ numFmt: 4, fundo: COR_EXCL }), inteiroExc = await pasta.estilo({ numFmt: 3, fundo: COR_EXCL });
  const exc = await pasta.estilo({ fundo: COR_EXCL });
  const rotulo = await pasta.estilo({ negrito: true, alinhar: "right" });
  const totalFmt = await pasta.estilo({ negrito: true, numFmt: 4 }), totalInt = await pasta.estilo({ negrito: true, numFmt: 1 });

  w.definir("A1", `Vendas de ${formatarData(res.dia)} por loja (API /vends)`, titulo);
  ["Cliente", "Local", "Máquina", "Valor", "Itens", "Cupons", "Fora da conta (filtro)?"]
    .forEach((t, i) => w.definirEm(3, i + 1, t, cab));
  let r = 4;
  for (const l of res.lojas) {
    const e = l.excluida;
    w.definirEm(r, 1, l.cliente, e ? exc : undefined);
    w.definirEm(r, 2, l.local, e ? exc : undefined);
    w.definirEm(r, 3, l.maquina, e ? exc : undefined);
    w.definirEm(r, 4, arred2(l.valor), e ? dinheiroExc : dinheiro);
    w.definirEm(r, 5, l.itens, e ? inteiroExc : inteiro);
    w.definirEm(r, 6, l.cupons, e ? inteiroExc : inteiro);
    w.definirEm(r, 7, e ? "Sim" : "Não", e ? exc : undefined);
    r++;
  }
  r++;
  const linhas = [
    ["Total na planilha (sem as excluídas)", arred2(res.total), totalFmt],
    ["Lojas excluídas pelo filtro", arred2(res.totExc), totalFmt],
    ["Total geral da API", arred2(res.total + res.totExc), totalFmt],
    ["Vendas com valor zero (brinde/troca)", res.zeradas, totalInt],
  ];
  linhas.forEach(([t, v, s], i) => { w.definirEm(r + i, 3, t, rotulo); w.definirEm(r + i, 4, v, s); });
  w.formatarColunas(1, 2, 28); w.formatarColunas(3, 3, 36); w.formatarColunas(4, 6, 14); w.formatarColunas(7, 7, 22);
}

// ------------------------------------------------------------
//  DiagnosticarVendasAPI (mostra na tela; não grava na planilha)
// ------------------------------------------------------------
export { diagnosticar };
