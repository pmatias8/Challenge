// ============================================================
//  Proxy da API VMpay para navegador (Cloudflare Worker, plano grátis)
//
//  Por que existe: a API da VMpay não manda os cabeçalhos CORS, então
//  o navegador (Excel na Web / Office Scripts e o app no celular)
//  bloqueia a resposta. O Excel no Windows não tem essa regra, por isso
//  a macro VBA funciona direto.
//
//  O que faz: repassa SOMENTE pedidos GET para /api/v1/... da VMpay e
//  devolve a resposta com o cabeçalho que libera o navegador. Não guarda
//  nada. O token continua indo na própria chamada (access_token=...),
//  então sem o token este endereço não serve para nada.
// ============================================================

const VMPAY = "https://vmpay.vertitecnologia.com.br";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, OPTIONS",
  "Access-Control-Allow-Headers": "Accept, Content-Type",
  "Access-Control-Max-Age": "86400",
};

export default {
  async fetch(request) {
    if (request.method === "OPTIONS") return new Response(null, { status: 204, headers: CORS });
    if (request.method !== "GET") return new Response("Só GET", { status: 405, headers: CORS });

    const url = new URL(request.url);
    if (!url.pathname.startsWith("/api/v1/")) {
      return new Response("Proxy VMpay no ar. Use /api/v1/...", { status: 200, headers: { ...CORS, "Content-Type": "text/plain; charset=utf-8" } });
    }

    const resposta = await fetch(VMPAY + url.pathname + url.search, {
      headers: { Accept: "application/json" },
      cf: { cacheTtl: 0 },
    });
    const h = new Headers(CORS);
    h.set("Content-Type", resposta.headers.get("Content-Type") || "application/json");
    h.set("Cache-Control", "no-store");
    return new Response(resposta.body, { status: resposta.status, headers: h });
  },
};
