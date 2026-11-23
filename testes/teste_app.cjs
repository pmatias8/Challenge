// Teste de ponta a ponta do app (opção 3) num Chromium sem tela.
// Uso: node testes/teste_app.cjs <planilha-de-teste.xlsm> <pasta-de-saida>
// Roda o app em modo demonstração, executa as 3 macros e baixa a planilha.
const { chromium } = require("playwright");
const { spawn } = require("child_process");
const path = require("path");
const fs = require("fs");

const [, , entrada, saidaDir] = process.argv;
const raiz = path.resolve(__dirname, "../docs");
const PORTA = 8765;
const ok = (c, msg) => { if (!c) throw new Error("FALHOU: " + msg); console.log("  ✓ " + msg); };

(async () => {
  const srv = spawn("python3", ["-m", "http.server", String(PORTA), "--bind", "127.0.0.1"], { cwd: raiz, stdio: "ignore" });
  await new Promise((r) => setTimeout(r, 800));
  const nav = await chromium.launch();
  try {
    const ctx = await nav.newContext({ acceptDownloads: true, viewport: { width: 390, height: 844 }, deviceScaleFactor: 1, locale: "pt-BR" });
    const page = await ctx.newPage();
    const erros = [];
    page.on("pageerror", (e) => erros.push(e.message));
    page.on("console", (m) => { if (m.type() === "error") erros.push(m.text()); });
    await page.clock.setFixedTime(new Date("2026-11-23T15:00:00Z")); // 23/11/2026 12:00 em Brasília
    await page.addInitScript(() => localStorage.setItem("vmpay_demo", "1"));
    await page.goto(`http://127.0.0.1:${PORTA}/`);

    console.log("Abrir planilha");
    await page.setInputFiles("#arquivo", entrada);
    await page.waitForSelector("#infoPlanilha b");
    const infoTxt = await page.textContent("#infoPlanilha");
    ok(infoTxt.includes("15 de 30 dias"), "detectou 15 de 30 dias preenchidos: " + infoTxt);
    ok(await page.inputValue("#dtIni") === "2026-11-16" && await page.inputValue("#dtFim") === "2026-11-22",
      "sugeriu 16/11 a 22/11 (primeiro vazio até ontem)");
    await page.screenshot({ path: path.join(saidaDir, "1-aberta.png"), fullPage: true });

    console.log("Atualizar 16/11–22/11");
    await page.click("#btnAtualizar");
    await page.waitForSelector("#resAtualizar .ok", { timeout: 60000 });
    const r1 = await page.textContent("#resAtualizar .ok");
    ok(r1.includes("7 dia(s) gravado(s)"), r1);
    ok(!(await page.isVisible("#dlgPergunta")), "sem pergunta do ticket (log ainda não tem o mês todo)");

    console.log("Atualizar 01/11–22/11 mantendo os já preenchidos");
    await page.fill("#dtIni", "2026-11-01"); await page.dispatchEvent("#dtIni", "change");
    ok(await page.isVisible("#optSobrescrever"), "mostrou a opção de substituir");
    await page.click("#btnAtualizar");
    await page.waitForSelector("#dlgPergunta[open]", { timeout: 90000 });
    const pg = await page.textContent("#pgTexto");
    ok(pg.includes("Cupons atuais no Painel: 2.150"), "pergunta do ticket: " + pg.replace(/\n/g, " "));
    await page.screenshot({ path: path.join(saidaDir, "2-ticket.png") });
    await page.click("#pgBotoes button[value=sim]");
    await page.waitForSelector("text=Painel!D7 atualizado");
    const r2 = await page.textContent("#resAtualizar .ok");
    ok(r2.includes("0 dia(s) gravado(s)") && r2.includes("22 mantido(s)"), r2);
    await page.screenshot({ path: path.join(saidaDir, "3-resultado.png"), fullPage: true });

    console.log("Conferir 22/11 por loja");
    await page.fill("#dtConf", "2026-11-22");
    await page.click("#btnConferir");
    await page.waitForSelector("#resConferir .totais", { timeout: 60000 });
    ok(await page.locator("#resConferir tr.excl").count() === 1, "1 loja marcada como excluída pelo filtro");
    await page.screenshot({ path: path.join(saidaDir, "4-conferencia.png"), fullPage: true });

    console.log("Diagnóstico");
    await page.click("#secDiag summary");
    await page.click("#btnDiag");
    await page.waitForFunction(() => document.querySelector("#resDiag").textContent.includes("occurred_at"));
    ok(true, "diagnóstico mostrou JSON");

    console.log("Baixar");
    const [dl] = await Promise.all([page.waitForEvent("download"), page.click("#btnBaixar")]);
    const destino = path.join(saidaDir, "saida.xlsm");
    await dl.saveAs(destino);
    ok(dl.suggestedFilename() === path.basename(entrada), "mesmo nome de arquivo: " + dl.suggestedFilename());
    ok(erros.length === 0, "sem erros no console: " + erros.join(" | "));

    // valores esperados calculados direto dos dados falsos (independente do app)
    const demo = await import(path.join(raiz, "js/demo.js"));
    const esperado = {};
    for (let d = 16; d <= 22; d++) {
      const vs = demo.vendasDoDia(`${String(d).padStart(2, "0")}/11/2026 03:00:00`)
        .filter((v) => v.machine.name !== "Escritório Teste");
      esperado[d] = {
        valor: Math.round(vs.reduce((s, v) => s + v.value, 0) * 100) / 100,
        cupons: new Set(vs.map((v) => v.machine_id + "|" + v.occurred_at.slice(0, 19))).size,
      };
    }
    fs.writeFileSync(path.join(saidaDir, "esperado.json"), JSON.stringify(esperado));
    console.log("OK - app passou");
  } finally {
    await nav.close(); srv.kill();
  }
})().catch((e) => { console.error(e); process.exit(1); });
