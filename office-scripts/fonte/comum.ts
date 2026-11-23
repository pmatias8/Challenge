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

/** Aceita o token puro, "access_token=..." ou o endereço de teste inteiro. */
function limparToken(s: string): string {
  let t = s.trim();
  const i = t.indexOf("access_token=");
  if (i >= 0) t = t.substring(i + "access_token=".length);
  t = t.split(/[&\s]/)[0].replace(/^["'<]+|["'>]+$/g, "");
  if (/%[0-9a-fA-F]{2}/.test(t)) { try { t = decodeURIComponent(t); } catch (e) { /* mantém */ } }
  return t;
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

