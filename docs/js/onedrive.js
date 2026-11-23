// ============================================================
//  Abrir e salvar a planilha ORIGINAL no OneDrive / SharePoint
//  (Microsoft Graph). O app baixa o arquivo, altera só as células
//  da macro e grava de volta no mesmo arquivo, com proteção contra
//  sobrescrever alterações feitas por outra pessoa nesse meio tempo.
// ============================================================

const GRAPH = "https://graph.microsoft.com/v1.0";
const ESCOPOS = ["Files.ReadWrite.All"];

let app = null, appChave = "";

/** Endereço de volta do login (sem "index.html", igual ao cadastrado no aplicativo). */
export const enderecoRetorno = () => location.origin + location.pathname.replace(/index\.html$/, "");

/** Configuração: { clientId, tenant } (tenant vazio = qualquer conta Microsoft). */
async function pca(cfg) {
  if (!cfg.clientId) throw new Error("Falta o ID do aplicativo Microsoft (⚙︎ → Conta Microsoft). Veja o guia.");
  const chave = cfg.clientId + "|" + (cfg.tenant || "");
  if (app && appChave === chave) return app;
  const M = globalThis.msal;
  if (!M) throw new Error("Biblioteca de login da Microsoft não carregou (sem internet?).");
  app = new M.PublicClientApplication({
    auth: {
      clientId: cfg.clientId,
      authority: "https://login.microsoftonline.com/" + (cfg.tenant || "common"),
      redirectUri: enderecoRetorno(),
    },
    cache: { cacheLocation: "localStorage" },
  });
  await app.initialize();
  appChave = chave;
  return app;
}

/** Chamar no início: termina um login que voltou da página da Microsoft. */
export async function concluirLogin(cfg) {
  if (!cfg.clientId) return null;
  const a = await pca(cfg);
  const r = await a.handleRedirectPromise();
  if (r && r.account) a.setActiveAccount(r.account);
  return conta(cfg);
}

export async function conta(cfg) {
  if (!cfg.clientId) return null;
  const a = await pca(cfg);
  const c = a.getActiveAccount() || a.getAllAccounts()[0] || null;
  if (c) a.setActiveAccount(c);
  return c;
}

export async function entrar(cfg) {
  const a = await pca(cfg);
  await a.loginRedirect({ scopes: ESCOPOS, prompt: "select_account" });
}

export async function sair(cfg) {
  const a = await pca(cfg);
  const c = await conta(cfg);
  if (c) await a.logoutRedirect({ account: c, postLogoutRedirectUri: enderecoRetorno() });
}

async function token(cfg) {
  if (globalThis.__TOKEN_TESTE__) return globalThis.__TOKEN_TESTE__; // testes automáticos
  const a = await pca(cfg);
  const c = await conta(cfg);
  if (!c) throw new Error("Entre com a sua conta Microsoft primeiro.");
  try {
    return (await a.acquireTokenSilent({ scopes: ESCOPOS, account: c })).accessToken;
  } catch (e) {
    // sessão expirou: volta para a página de login e retorna aqui
    await a.acquireTokenRedirect({ scopes: ESCOPOS, account: c });
    throw new Error("Redirecionando para o login…");
  }
}

async function graph(cfg, caminho, opcoes = {}) {
  const r = await fetch(caminho.startsWith("http") ? caminho : GRAPH + caminho, {
    ...opcoes,
    headers: { Authorization: "Bearer " + (await token(cfg)), ...(opcoes.headers || {}) },
  });
  if (r.ok) return r;
  let msg = "";
  try { const j = await r.json(); msg = (j.error && (j.error.message || j.error.code)) || ""; } catch { /* sem corpo */ }
  const e = new Error(`OneDrive respondeu HTTP ${r.status}${msg ? ": " + msg : ""}`);
  e.status = r.status;
  throw e;
}

/** Referência mínima que o app guarda para reabrir a planilha. */
function ref(item) {
  const alvo = item.remoteItem || item;
  return {
    driveId: (alvo.parentReference && alvo.parentReference.driveId) || (item.parentReference && item.parentReference.driveId),
    itemId: alvo.id,
    nome: alvo.name || item.name,
    local: (alvo.parentReference && alvo.parentReference.path ? alvo.parentReference.path.replace(/^.*root:/, "") : "") || "",
  };
}

const ehPlanilha = (nome) => /\.xls[xm]$/i.test(nome || "");

/** Planilhas usadas recentemente (OneDrive e SharePoint). */
export async function recentes(cfg) {
  const r = await (await graph(cfg, "/me/drive/recent?$top=100")).json();
  return (r.value || []).filter((it) => ehPlanilha((it.remoteItem || it).name) && !(it.remoteItem || it).folder).map(ref).slice(0, 15);
}

/** A partir de um link de compartilhamento ("Copiar link" no OneDrive/Excel). */
export async function porLink(cfg, link) {
  const b64 = btoa(unescape(encodeURIComponent(link.trim()))).replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
  const item = await (await graph(cfg, `/shares/u!${b64}/driveItem`)).json();
  if (!ehPlanilha(item.name)) throw new Error(`"${item.name}" não é uma planilha .xlsx/.xlsm.`);
  return ref(item);
}

/** Baixa a planilha. Devolve { bytes, eTag, nome, webUrl }. */
export async function baixar(cfg, r) {
  const item = await (await graph(cfg, `/drives/${r.driveId}/items/${r.itemId}`)).json();
  const url = item["@microsoft.graph.downloadUrl"];
  const resp = url ? await fetch(url) : await graph(cfg, `/drives/${r.driveId}/items/${r.itemId}/content`);
  if (!resp.ok) throw new Error(`Não consegui baixar a planilha (HTTP ${resp.status}).`);
  return { bytes: await resp.arrayBuffer(), eTag: item.eTag, nome: item.name, webUrl: item.webUrl };
}

/**
 * Grava por cima do arquivo original. Com If-Match: se alguém alterou a
 * planilha depois que o app a abriu, o OneDrive recusa (412) em vez de
 * apagar a alteração da outra pessoa.
 */
export async function salvar(cfg, r, bytes, eTag) {
  try {
    const resp = await graph(cfg, `/drives/${r.driveId}/items/${r.itemId}/content`, {
      method: "PUT", body: bytes,
      headers: { "Content-Type": "application/octet-stream", ...(eTag ? { "If-Match": eTag } : {}) },
    });
    const item = await resp.json();
    return { eTag: item.eTag, quando: item.lastModifiedDateTime };
  } catch (e) {
    if (e.status === 412) throw new Error("A planilha foi alterada no OneDrive depois que o app a abriu (por você em outro lugar ou por outra pessoa). " +
      "Para não apagar essa alteração, nada foi salvo. Abra a planilha de novo pelo app e rode outra vez.");
    if (e.status === 423) throw new Error("A planilha está bloqueada no OneDrive (aberta com bloqueio em algum lugar). Feche-a e tente de novo.");
    throw e;
  }
}
