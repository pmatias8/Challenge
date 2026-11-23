// ============================================================
//  Lógica da macro modVendasAPI (VBA) traduzida para JavaScript.
//  Mesmas regras: dia de Brasília, paginação em lotes paralelos,
//  filtro de lojas, estimativa de cupons e modo conferência.
//  Funciona no navegador e no Node (testes).
// ============================================================

export const API_VMPAY = "https://vmpay.vertitecnologia.com.br/api/v1";
export const OFFSET_UTC = -3;     // Brasília
export const PER_PAGE = 1000;
export const LOTE = 8;            // páginas pedidas em paralelo

const DIA_MS = 86400000;
const EPOCA_EXCEL = 25569;        // serial do Excel para 01/01/1970

// ---------- datas (serial do Excel, sistema 1900) ----------

/** Serial do Excel de "hoje" no fuso de Brasília (só a parte inteira). */
export function hojeSerial(agoraMs = Date.now()) {
  return Math.floor(agoraMs / DIA_MS + EPOCA_EXCEL + OFFSET_UTC / 24);
}

/** "Agora" em Brasília como serial com fração (para "Atualizado em"). */
export function agoraSerial(agoraMs = Date.now()) {
  return agoraMs / DIA_MS + EPOCA_EXCEL + OFFSET_UTC / 24;
}

/** Serial -> partes de data/hora (lido como relógio "de parede", sem fuso). */
function partes(serial) {
  const ms = Math.round((serial - EPOCA_EXCEL) * 86400) * 1000;
  const d = new Date(ms);
  return {
    a: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate(),
    h: d.getUTCHours(), n: d.getUTCMinutes(), s: d.getUTCSeconds(),
  };
}
const p2 = (x) => String(x).padStart(2, "0");

export function formatarData(serial) {
  const p = partes(serial);
  return `${p2(p.d)}/${p2(p.m)}/${p.a}`;
}
export function formatarDataHora(serial) {
  const p = partes(serial);
  return `${p2(p.d)}/${p2(p.m)}/${p.a} ${p2(p.h)}:${p2(p.n)}:${p2(p.s)}`;
}

/** "dd/mm/aaaa" -> serial (ou null). Igual ao TentarConverterData do VBA. */
export function converterData(texto) {
  const p = String(texto || "").trim().split("/");
  if (p.length !== 3) return null;
  const [d, m, a] = p.map((x) => Number(x));
  if (![d, m, a].every(Number.isInteger)) return null;
  if (m < 1 || m > 12 || d < 1 || d > 31 || a < 1900) return null;
  return Date.UTC(a, m - 1, d) / DIA_MS + EPOCA_EXCEL;
}

/** Serial <-> "aaaa-mm-dd" (campo <input type=date>). */
export function serialParaISO(serial) {
  const p = partes(serial);
  return `${p.a}-${p2(p.m)}-${p2(p.d)}`;
}
export function isoParaSerial(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || "");
  if (!m) return null;
  return Date.UTC(+m[1], +m[2] - 1, +m[3]) / DIA_MS + EPOCA_EXCEL;
}

/** Janela em UTC do dia de Brasília: 00:00 local = 03:00 UTC ... 23:59:59 local. */
export function janelaDiaUTC(dia) {
  const ini = dia - OFFSET_UTC / 24;
  const fim = dia + 1 - OFFSET_UTC / 24 - 1 / 86400;
  return { ini: formatarDataHora(ini), fim: formatarDataHora(fim) };
}

// ---------- token ----------

/**
 * Limpa o que foi colado como token: aceita o token puro, "access_token=..."
 * ou até o endereço inteiro de teste, e tira aspas, espaços e "&page=..." do fim.
 */
export function limparToken(s) {
  let t = String(s ?? "").trim();
  const i = t.indexOf("access_token=");
  if (i >= 0) t = t.slice(i + "access_token=".length);
  t = t.split(/[&\s]/)[0].replace(/^["'<]+|["'>]+$/g, "");
  try { if (/%[0-9a-f]{2}/i.test(t)) t = decodeURIComponent(t); } catch { /* deixa como está */ }
  return t;
}

/** "18 caracteres, termina em …f3a9" — para conferir sem mostrar o token. */
export function resumoToken(t) {
  return t ? `${t.length} caracteres, termina em …${t.slice(-4)}` : "nenhum token salvo";
}

// ---------- texto ----------

const ACENTOS = { "á": "a", "à": "a", "ã": "a", "â": "a", "ä": "a", "é": "e", "ê": "e", "è": "e", "ë": "e",
  "í": "i", "ì": "i", "î": "i", "ó": "o", "ô": "o", "õ": "o", "ò": "o", "ú": "u", "ù": "u", "ü": "u", "û": "u",
  "ç": "c", "ñ": "n" };

export function normalizar(s) {
  return String(s ?? "").trim().toLowerCase().replace(/[áàãâäéêèëíìîóôõòúùüûçñ]/g, (c) => ACENTOS[c]);
}

function getStr(obj, chave) {
  if (!obj || typeof obj !== "object") return "";
  const v = obj[chave];
  if (v === null || v === undefined || typeof v === "object") return "";
  return String(v);
}
function getNum(obj, chave) {
  if (!obj || typeof obj !== "object") return null;
  const v = obj[chave];
  return typeof v === "number" && Number.isFinite(v) ? v : null;
}
const sub = (it, k) => (it && it[k] && typeof it[k] === "object" ? it[k] : null);

// ---------- regra de negócio de uma venda ----------

/** Cria o acumulador de um dia. excluir = Set de nomes normalizados. */
export function novoAcumulador(excluir = new Set(), conferencia = false) {
  return {
    excluir, total: 0, itens: 0, chaves: new Set(),
    conf: conferencia ? { lojas: new Map(), zeradas: 0 } : null,
    cache: new Map(),
  };
}

function excluido(acc, nome) {
  if (!nome) return false;
  let n = acc.cache.get(nome);
  if (n === undefined) { n = normalizar(nome); acc.cache.set(nome, n); }
  return acc.excluir.has(n);
}

/** Igual ao SomarVenda do VBA. */
export function somarVenda(acc, it) {
  if (!it || typeof it !== "object") return;
  const machine = sub(it, "machine"), location = sub(it, "location"), client = sub(it, "client");

  let excl = false;
  if (acc.excluir.size > 0) {
    excl = excluido(acc, getStr(machine, "name")) ||
           excluido(acc, getStr(machine, "asset_number")) ||
           excluido(acc, getStr(location, "name"));
  }

  const valor = getNum(it, "value") ?? 0;
  const qtd = getNum(it, "quantity") ?? 1;

  // Cupom: a API não manda id da compra; itens do mesmo carrinho saem com a
  // mesma máquina e o mesmo horário (até o segundo).
  let k = getStr(it, "purchase_id") || getStr(it, "order_id") || getStr(it, "cashless_transaction_id");
  if (!k) k = getStr(it, "machine_id") + "|" + getStr(it, "occurred_at").slice(0, 19);

  if (acc.conf) {
    const lj = [getStr(client, "name"), getStr(location, "name"), getStr(machine, "asset_number")].join(" | ");
    let l = acc.conf.lojas.get(lj);
    if (!l) {
      l = { cliente: getStr(client, "name"), local: getStr(location, "name"), maquina: getStr(machine, "asset_number"),
            valor: 0, itens: 0, cupons: new Set(), excluida: false };
      acc.conf.lojas.set(lj, l);
    }
    l.valor += valor; l.itens += qtd; l.cupons.add(k); l.excluida = excl;
    if (valor <= 0) acc.conf.zeradas++;
  }

  if (excl) return;
  acc.total += valor;
  acc.itens += qtd;
  acc.chaves.add(k);
}

// ---------- API ----------

export class ErroApi extends Error {}

function montarUrl(base, token, dia, extra) {
  const j = janelaDiaUTC(dia);
  const q = new URLSearchParams({ access_token: token, start_date: j.ini, end_date: j.fim, ...extra });
  // espaços como %20 (igual ao UrlEncode do VBA); "+" literal já vira %2B
  return `${base.replace(/\/+$/, "")}/vends?${q.toString().replace(/\+/g, "%20")}`;
}

async function pedir(fetchImpl, url, signal, dia, pagina) {
  let r;
  try {
    r = await fetchImpl(url, { signal, headers: { Accept: "application/json" } });
  } catch (e) {
    if (e && e.name === "AbortError") throw e;
    throw new ErroApi(
      "Não consegui falar com a API (" + (e && e.message ? e.message : e) + "). " +
      "No navegador isso quase sempre é bloqueio de CORS: configure o endereço do proxy.");
  }
  if (r.status === 401 || r.status === 403) throw new ErroApi(`A VMpay recusou o token (HTTP ${r.status}).`);
  if (r.status !== 200) {
    const tx = (await r.text().catch(() => "")).slice(0, 200);
    throw new ErroApi(`HTTP ${r.status} em ${formatarData(dia)} (página ${pagina}): ${tx}`);
  }
  const dados = await r.json();
  if (!Array.isArray(dados)) throw new ErroApi("Resposta da API não é uma lista JSON.");
  return dados;
}

/**
 * Consulta um dia inteiro (igual ao ConsultarDia do VBA): pede LOTE páginas
 * em paralelo e para na primeira página com menos de PER_PAGE vendas.
 */
export async function consultarDia({ base = API_VMPAY, token, dia, acc, fetchImpl = fetch, aoPaginar }) {
  let pag0 = 0;
  for (;;) {
    const ctrl = new AbortController();
    const pedidos = [];
    for (let i = 1; i <= LOTE; i++) {
      const pg = pag0 + i;
      const p = pedir(fetchImpl, montarUrl(base, token, dia, { per_page: PER_PAGE, page: pg }), ctrl.signal, dia, pg);
      p.catch(() => {}); // evita "unhandled rejection" das páginas canceladas
      pedidos.push(p);
    }
    try {
      for (let i = 0; i < LOTE; i++) {
        const lista = await pedidos[i];
        for (const it of lista) somarVenda(acc, it);
        if (aoPaginar) aoPaginar(pag0 + i + 1);
        if (lista.length < PER_PAGE) { ctrl.abort(); return acc; }
      }
    } catch (e) {
      ctrl.abort();
      throw e;
    }
    pag0 += LOTE;
  }
}

/** Resposta crua (5 vendas) para diagnóstico. */
export async function diagnosticar({ base = API_VMPAY, token, dia, fetchImpl = fetch }) {
  return pedir(fetchImpl, montarUrl(base, token, dia, { page: 1, per_page: 5 }), undefined, dia, 1);
}
