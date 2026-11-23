import { Pasta } from "./xlsx.js";
import { API_VMPAY, formatarData, serialParaISO, isoParaSerial, hojeSerial, limparToken, resumoToken } from "./vendas.js";
import * as M from "./macros.js";
import { fetchDemo } from "./demo.js";

const $ = (id) => document.getElementById(id);
const brl = (v) => v.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const int = (v) => Math.round(v).toLocaleString("pt-BR");
const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

// ---------- configuração (fica só neste aparelho) ----------
const guardar = {
  ler(k, pad = "") { try { return localStorage.getItem(k) ?? pad; } catch { return pad; } },
  gravar(k, v) { try { localStorage.setItem(k, v); } catch { /* modo privado */ } },
};
const cfg = {
  get token() { return limparToken(guardar.ler("vmpay_token")); },
  get proxy() { return guardar.ler("vmpay_proxy").trim(); },
  get demo() { return guardar.ler("vmpay_demo") === "1"; },
};
function baseApi() {
  if (cfg.demo) return "https://demo.local/api/v1";
  if (!cfg.proxy) return API_VMPAY;
  return cfg.proxy.replace(/\/+$/, "").replace(/\/api\/v1$/, "") + "/api/v1";
}
const fetchImpl = (...a) => (cfg.demo ? fetchDemo(...a) : fetch(...a));

// ---------- estado ----------
let pasta = null, info = null, nomeArquivo = "", tipoArquivo = "", alterada = false, ocupado = false;

// ---------- perguntas (substituem MsgBox / InputBox) ----------
function perguntar(titulo, html, botoes) {
  return new Promise((ok) => {
    const dlg = $("dlgPergunta");
    $("pgTitulo").textContent = titulo;
    $("pgTexto").innerHTML = html;
    const caixa = $("pgBotoes"); caixa.innerHTML = "";
    for (const b of botoes) {
      const el = document.createElement("button");
      el.className = "btn" + (b.primario ? " primario" : ""); el.value = b.valor; el.textContent = b.texto;
      caixa.appendChild(el);
    }
    dlg.onclose = () => ok(dlg.returnValue || "cancelar");
    dlg.returnValue = "";
    dlg.showModal();
  });
}
const avisar = (titulo, html) => perguntar(titulo, html, [{ texto: "OK", valor: "ok", primario: true }]);

// ---------- interface ----------
function atualizarModo() {
  const f = $("avisoModo");
  if (cfg.demo) { f.hidden = false; f.innerHTML = "<b>Modo demonstração:</b> as vendas são inventadas. Desligue em ⚙︎ para usar a VMpay de verdade."; }
  else if (!cfg.token) { f.hidden = false; f.innerHTML = "Toque em ⚙︎ e cole o <b>token da VMpay</b> e o endereço do <b>proxy</b>. Ou ligue o modo demonstração para testar."; }
  else f.hidden = true;
}
function habilitar() {
  for (const el of document.querySelectorAll("[data-precisa-planilha] button, [data-precisa-planilha] input"))
    el.disabled = !info || ocupado;
  $("btnDiag").disabled = ocupado;
}
function marcarAlterada() {
  alterada = true;
  $("barraSalvar").hidden = false;
  $("txtSalvar").textContent = `${nomeArquivo} alterada`;
}
function erro(alvo, e) {
  console.error(e);
  const msg = String(e && e.message ? e.message : e);
  const dica = /recusou o token/.test(msg)
    ? `\n\nToken usado: ${resumoToken(cfg.token)}. Confira em ⚙︎ se é o mesmo que funcionou no navegador ` +
      `(cole só o token, sem "&page=..."). Proxy: ${baseApi()}`
    : "";
  $(alvo).innerHTML = `<div class="erro">Erro: ${esc(msg)}${esc(dica)}\n\nNenhum valor foi alterado.</div>`;
}
function progresso(alvo, txt, frac) {
  $(alvo).innerHTML = `<div class="progresso">${esc(txt)}<div class="trilho"><div class="barra-p" style="width:${Math.round(frac * 100)}%"></div></div></div>`;
}
function precisaApi() {
  if (cfg.demo || cfg.token) return true;
  avisar("Falta o token", "Abra as configurações (⚙︎) e cole o <b>access_token</b> da VMpay — ou ligue o modo demonstração.");
  return false;
}

function mostrarInfo() {
  const n = info.linhaDia.size;
  $("infoPlanilha").className = "info";
  $("infoPlanilha").innerHTML =
    `<b>${esc(nomeArquivo)}</b><br>Aba Diário: ${formatarData(info.primDia)} a ${formatarData(info.ultDia)} · ` +
    `${info.preenchidos} de ${n} dias com Realizado` +
    (info.primVazio ? `<br>Primeiro dia vazio: <b>${formatarData(info.primVazio)}</b>` : "");
  $("dtIni").value = serialParaISO(info.sugIni);
  $("dtFim").value = serialParaISO(info.sugFim);
  $("dtConf").value = serialParaISO(hojeSerial() - 1);
  verificar();
}

function verificar() {
  const caixa = $("avisoPeriodo"); caixa.innerHTML = "";
  $("optSobrescrever").hidden = true;
  if (!info) return null;
  const ini = isoParaSerial($("dtIni").value), fim = isoParaSerial($("dtFim").value);
  const av = (t) => caixa.insertAdjacentHTML("beforeend", `<div class="aviso">${t}</div>`);
  if (ini === null || fim === null) { av("Escolha as datas de início e fim."); return null; }
  if (fim < ini) { av("A data de fim tem que ser igual ou depois do início."); return null; }
  const v = M.verificarPeriodo(info, ini, fim);
  if (v.todosFora) { av(`Nenhum dia do período está na aba '${M.ABA_PROJ}' (coluna B).`); return null; }
  if (v.incluiHoje) av("O dia de hoje ainda não fechou e vai entrar parcial.");
  if (v.jaTem > 0) {
    $("optSobrescrever").hidden = false;
    av(`${v.jaTem} dia(s) do período já têm valor em Realizado.`);
  }
  return { ini, fim, ...v };
}

async function abrirArquivo(arq) {
  try {
    const buf = await arq.arrayBuffer();
    const p = await Pasta.abrir(buf);
    const i = await M.analisar(p);
    pasta = p; info = i; nomeArquivo = arq.name; tipoArquivo = arq.type; alterada = false;
    $("barraSalvar").hidden = true;
    $("resAtualizar").innerHTML = ""; $("resConferir").innerHTML = "";
    $("dicaExemplo").hidden = true;
    mostrarInfo();
  } catch (e) {
    pasta = info = null;
    $("infoPlanilha").className = "info";
    $("infoPlanilha").innerHTML = `<div class="erro" style="margin:0">${esc(e.message || e)}</div>`;
  }
  habilitar();
}

// ---------- 2. Atualizar vendas ----------
async function rodarAtualizar() {
  const per = verificar();
  if (!per || !precisaApi()) return;
  const sobrescrever = per.jaTem > 0 && $("sobrescrever").checked;
  if (per.incluiHoje) {
    const r = await perguntar("Dia ainda aberto", "O dia de hoje ainda não fechou e vai entrar parcial. Continuar?",
      [{ texto: "Não", valor: "nao" }, { texto: "Continuar", valor: "sim", primario: true }]);
    if (r !== "sim") return;
  }
  ocupado = true; habilitar();
  const t0 = performance.now();
  try {
    const res = await M.atualizarVendas(pasta, info, {
      token: cfg.token, base: baseApi(), ini: per.ini, fim: per.fim, sobrescrever, fetchImpl,
      progresso: ({ dia, i, n, pagina }) => progresso("resAtualizar",
        `Consultando ${formatarData(dia)} (${i}/${n})${pagina ? ` · página ${pagina}` : ""}…`, (i - 1) / n),
    });
    marcarAlterada();
    info = await M.analisar(pasta); // relê (dias preenchidos mudaram)
    const seg = Math.round((performance.now() - t0) / 1000);
    let html = `<div class="ok"><b>${res.gravados} dia(s) gravado(s)</b> em Realizado` +
      (res.mantidos ? `, ${res.mantidos} mantido(s) como estavam` : "") + `. Tempo: ${seg}s.</div>`;
    html += `<div class="tabela-wrap"><table><thead><tr><th>Dia</th><th class="n">API</th><th class="n">Antes</th><th class="n">Cupons</th><th></th></tr></thead><tbody>`;
    for (const d of res.dias) {
      html += `<tr><td>${formatarData(d.dia).slice(0, 5)}</td><td class="n">${brl(d.valor)}</td>` +
        `<td class="n">${typeof d.antes === "number" ? brl(d.antes) : "—"}</td><td class="n">${int(d.cupons)}</td>` +
        `<td>${d.gravado ? '<span class="tag">gravado</span>' : '<span class="tag manteve">manteve</span>'}</td></tr>`;
    }
    html += `</tbody></table></div>`;
    $("resAtualizar").innerHTML = html;
    mostrarInfoSemMexerDatas();

    if (res.ticket) {
      const atual = typeof res.ticket.atual === "number" ? int(res.ticket.atual) : String(res.ticket.atual ?? "vazio");
      const r = await perguntar("Ticket do mês",
        `Cupons estimados no mês (até ontem): <b>${int(res.ticket.estimado)}</b>\nCupons atuais no Painel: <b>${esc(atual)}</b>\n\nAtualizar os cupons do Painel com o valor estimado?`,
        [{ texto: "Não", valor: "nao" }, { texto: "Atualizar", valor: "sim", primario: true }]);
      if (r === "sim") {
        await M.aplicarTicket(pasta, res.ticket.estimado);
        $("resAtualizar").insertAdjacentHTML("beforeend", `<div class="ok">Painel!D7 atualizado para ${int(res.ticket.estimado)} cupons.</div>`);
      }
    }
  } catch (e) { erro("resAtualizar", e); }
  ocupado = false; habilitar();
}
function mostrarInfoSemMexerDatas() {
  const a = $("dtIni").value, b = $("dtFim").value;
  mostrarInfo(); $("dtIni").value = a; $("dtFim").value = b; verificar();
}

// ---------- 3. Conferir por loja ----------
async function rodarConferir() {
  const dia = isoParaSerial($("dtConf").value);
  if (dia === null) return avisar("Data inválida", "Escolha o dia para conferir.");
  if (!precisaApi()) return;
  ocupado = true; habilitar();
  try {
    const res = await M.conferirPorLoja(pasta, {
      token: cfg.token, base: baseApi(), dia, fetchImpl,
      progresso: ({ pagina }) => progresso("resConferir", `Conferindo ${formatarData(dia)}${pagina ? ` · página ${pagina}` : ""}…`, 0.5),
    });
    await M.gravarConferencia(pasta, res, formatarData);
    marcarAlterada();
    let html = `<div class="tabela-wrap"><table><thead><tr><th>Local</th><th>Máquina</th><th class="n">Valor</th><th class="n">Itens</th><th class="n">Cupons</th></tr></thead><tbody>`;
    for (const l of res.lojas) {
      html += `<tr class="${l.excluida ? "excl" : ""}"><td>${esc(l.local || l.cliente)}</td><td>${esc(l.maquina)}</td>` +
        `<td class="n">${brl(l.valor)}</td><td class="n">${int(l.itens)}</td><td class="n">${int(l.cupons)}</td></tr>`;
    }
    html += `</tbody></table></div><div class="totais">` +
      `<span>Total na planilha (sem as excluídas)</span><span>${brl(res.total)}</span>` +
      `<span>Lojas excluídas pelo filtro</span><span>${brl(res.totExc)}</span>` +
      `<span>Total geral da API</span><span>${brl(res.total + res.totExc)}</span>` +
      `<span>Vendas com valor zero (brinde/troca)</span><span>${int(res.zeradas)}</span></div>` +
      `<p class="dica">Também gravado na aba "${M.ABA_CONF}". Linhas em laranja = fora da conta pelo filtro.</p>`;
    $("resConferir").innerHTML = html;
  } catch (e) { erro("resConferir", e); }
  ocupado = false; habilitar();
}

// ---------- 4. Diagnóstico ----------
async function rodarDiag() {
  const dia = isoParaSerial($("dtDiag").value);
  if (dia === null) return avisar("Data inválida", "Escolha o dia.");
  if (!precisaApi()) return;
  const pre = $("resDiag"); pre.hidden = false; pre.textContent = `Consultando… (token: ${resumoToken(cfg.token)}; proxy: ${baseApi()})`;
  try {
    const dados = await M.diagnosticar({ token: cfg.token, base: baseApi(), dia, fetchImpl });
    pre.textContent = `Resposta de /vends (5 primeiras vendas de ${formatarData(dia)}):\n\n` + JSON.stringify(dados, null, 2);
  } catch (e) { pre.textContent = `Erro: ${e.message || e}\n\nToken usado: ${resumoToken(cfg.token)}\nProxy: ${baseApi()}`; }
}

// ---------- salvar ----------
function nomeSaida() { return nomeArquivo || "planilha.xlsx"; }
async function gerarArquivo() {
  const bytes = await pasta.gerar("uint8array");
  const macro = /\.xlsm$/i.test(nomeArquivo);
  const tipo = macro ? "application/vnd.ms-excel.sheet.macroEnabled.12" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  return new File([bytes], nomeSaida(), { type: tipo });
}
async function baixar() {
  const f = await gerarArquivo();
  const url = URL.createObjectURL(f);
  const a = document.createElement("a"); a.href = url; a.download = f.name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
  salva();
}
function salva() { alterada = false; $("txtSalvar").textContent = `${nomeArquivo} salva ✓`; }
async function compartilhar() {
  const f = await gerarArquivo();
  if (navigator.canShare && navigator.canShare({ files: [f] })) {
    try { await navigator.share({ files: [f], title: f.name }); salva(); return; }
    catch (e) { if (e.name === "AbortError") return; }
  }
  await baixar();
}

// ---------- configurações ----------
function abrirConfig() {
  $("cfgToken").value = cfg.token; $("cfgProxy").value = cfg.proxy; $("cfgDemo").checked = cfg.demo;
  $("cfgToken").classList.add("mascara"); $("cfgMostrar").checked = false;
  $("cfgResumo").textContent = "Salvo agora: " + resumoToken(cfg.token);
  $("dlgConfig").showModal();
}
$("dlgConfig").addEventListener("close", () => {
  if ($("dlgConfig").returnValue !== "salvar") return;
  guardar.gravar("vmpay_token", limparToken($("cfgToken").value));
  guardar.gravar("vmpay_proxy", $("cfgProxy").value.trim());
  guardar.gravar("vmpay_demo", $("cfgDemo").checked ? "1" : "0");
  atualizarModo();
});
$("cfgMostrar").addEventListener("change", (e) => { $("cfgToken").classList.toggle("mascara", !e.target.checked); });
$("cfgToken").addEventListener("input", () => { $("cfgResumo").textContent = "Vai salvar: " + resumoToken(limparToken($("cfgToken").value)); });

// ---------- eventos ----------
$("btnConfig").addEventListener("click", abrirConfig);
$("arquivo").addEventListener("change", (e) => { const f = e.target.files[0]; if (f) abrirArquivo(f); e.target.value = ""; });
$("dtIni").addEventListener("change", verificar);
$("dtFim").addEventListener("change", verificar);
$("btnAtualizar").addEventListener("click", rodarAtualizar);
$("btnConferir").addEventListener("click", rodarConferir);
$("btnDiag").addEventListener("click", rodarDiag);
$("btnBaixar").addEventListener("click", baixar);
$("btnCompartilhar").addEventListener("click", compartilhar);
window.addEventListener("beforeunload", (e) => { if (alterada) { e.preventDefault(); e.returnValue = ""; } });

$("dtDiag").value = serialParaISO(hojeSerial() - 1);
atualizarModo();
habilitar();
if ("serviceWorker" in navigator && location.protocol === "https:") navigator.serviceWorker.register("sw.js").catch(() => {});
