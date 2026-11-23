// ============================================================
//  Modo demonstração: uma "API VMpay" falsa, dentro do aparelho,
//  para testar o app sem token e sem proxy. Os números são
//  inventados, mas o formato imita a resposta real de /vends.
// ============================================================

const MAQUINAS = [
  { id: 101, asset_number: "PM-0101", name: "Geladeira Academia Centro", local: "Academia Centro", cliente: "Rede Fit" },
  { id: 102, asset_number: "PM-0102", name: "Geladeira Academia Norte", local: "Academia Norte", cliente: "Rede Fit" },
  { id: 201, asset_number: "PM-0201", name: "Mercadinho Condomínio Sol", local: "Condomínio Sol", cliente: "Cond. Sol" },
  { id: 202, asset_number: "PM-0202", name: "Mercadinho Condomínio Lua", local: "Condomínio Lua", cliente: "Cond. Lua" },
  { id: 301, asset_number: "PM-0301", name: "Escritório Teste", local: "Escritório (teste)", cliente: "Interno" },
];

function rnd(seed) { // gerador determinístico (mulberry32)
  let a = seed >>> 0;
  return () => { a = (a + 0x6D2B79F5) >>> 0; let t = a; t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61); return ((t ^ (t >>> 14)) >>> 0) / 4294967296; };
}

/** Todas as vendas de um dia (UTC de início "dd/mm/aaaa hh:mm:ss"). */
export function vendasDoDia(inicioUTC) {
  const [d, m, a] = inicioUTC.slice(0, 10).split("/").map(Number);
  const r = rnd(a * 10000 + m * 100 + d);
  const n = 1800 + Math.floor(r() * 1400);       // 1.800 a 3.200 itens -> 2 a 4 páginas
  const base = Date.UTC(a, m - 1, d, 3, 0, 0);    // 00:00 em Brasília
  const out = [];
  let seg = 0, id = 1;
  while (out.length < n) {
    const mq = MAQUINAS[Math.floor(r() * MAQUINAS.length)];
    seg += 1 + Math.floor(r() * 40);
    const quando = new Date(base + Math.min(seg, 86399) * 1000).toISOString().replace(".000Z", "Z");
    const itensNoCarrinho = 1 + Math.floor(r() * 3);
    for (let k = 0; k < itensNoCarrinho && out.length < n; k++) {
      const brinde = r() < 0.01;
      out.push({
        id: id++, occurred_at: quando, machine_id: mq.id,
        value: brinde ? 0 : Math.round((3 + r() * 14) * 100) / 100,
        quantity: 1,
        machine: { id: mq.id, asset_number: mq.asset_number, name: mq.name },
        location: { name: mq.local }, client: { name: mq.cliente },
      });
    }
  }
  return out;
}

/** fetch falso: responde /vends com paginação, como a API real. */
export async function fetchDemo(url, opts = {}) {
  await new Promise((ok) => setTimeout(ok, 60 + Math.random() * 120));
  if (opts.signal && opts.signal.aborted) { const e = new Error("aborted"); e.name = "AbortError"; throw e; }
  const u = new URL(url);
  const todas = vendasDoDia(u.searchParams.get("start_date"));
  const por = Number(u.searchParams.get("per_page") || 1000), pg = Number(u.searchParams.get("page") || 1);
  const corpo = JSON.stringify(todas.slice((pg - 1) * por, pg * por));
  return new Response(corpo, { status: 200, headers: { "Content-Type": "application/json" } });
}
