// Teste do modo "planilha original no OneDrive" com o Microsoft Graph simulado.
// Uso: node testes/teste_onedrive.cjs <planilha-de-teste.xlsm> <pasta-de-saida>
const { chromium } = require("playwright");
const { spawn } = require("child_process");
const path = require("path");
const fs = require("fs");

const [, , entrada, saidaDir] = process.argv;
const raiz = path.resolve(__dirname, "../docs");
const PORTA = 8767;
const ok = (c, msg) => { if (!c) throw new Error("FALHOU: " + msg); console.log("  ✓ " + msg); };

(async () => {
  const srv = spawn("python3", ["-m", "http.server", String(PORTA), "--bind", "127.0.0.1"], { cwd: raiz, stdio: "ignore" });
  await new Promise((r) => setTimeout(r, 800));
  const nav = await chromium.launch();
  try {
    const page = await (await nav.newContext({ viewport: { width: 390, height: 844 }, locale: "pt-BR" })).newPage();
    const erros = [];
    page.on("pageerror", (e) => erros.push(e.message));

    // ---- OneDrive simulado ----
    let arquivo = fs.readFileSync(entrada), versao = 1, recusar412 = false;
    const puts = [];
    const etag = () => `"{AAAA-1111},${versao}"`;
    await page.route("https://graph.microsoft.com/v1.0/**", async (rota) => {
      const req = rota.request(), u = new URL(req.url());
      if (req.headers()["authorization"] !== "Bearer token-falso") return rota.fulfill({ status: 401, json: { error: { code: "InvalidAuthenticationToken" } } });
      if (u.pathname === "/v1.0/me/drive/recent") {
        return rota.fulfill({ json: { value: [
          { id: "x1", name: "Relatório.docx", remoteItem: { id: "R1", name: "Relatório.docx", parentReference: { driveId: "D9" } } },
          { id: "x2", name: "Faturamento.xlsm", remoteItem: { id: "I1", name: "Faturamento.xlsm", parentReference: { driveId: "D1", path: "/drives/D1/root:/Financeiro" } } },
        ] } });
      }
      if (u.pathname.startsWith("/v1.0/shares/u!")) {
        const link = Buffer.from(u.pathname.split("u!")[1].split("/")[0].replace(/-/g, "+").replace(/_/g, "/"), "base64").toString();
        ok(link === "https://empresa.sharepoint.com/:x:/g/Faturamento?e=abc", "link decodificado no formato do Graph");
        return rota.fulfill({ json: { id: "I1", name: "Faturamento.xlsm", parentReference: { driveId: "D1" } } });
      }
      if (u.pathname === "/v1.0/drives/D1/items/I1" && req.method() === "GET") {
        return rota.fulfill({ json: { id: "I1", name: "Faturamento.xlsm", eTag: etag(), webUrl: "https://empresa.sharepoint.com/x",
          "@microsoft.graph.downloadUrl": "https://download.test/arquivo" } });
      }
      if (u.pathname === "/v1.0/drives/D1/items/I1/content" && req.method() === "PUT") {
        puts.push({ ifMatch: req.headers()["if-match"], tamanho: req.postDataBuffer().length });
        if (recusar412 || req.headers()["if-match"] !== etag()) return rota.fulfill({ status: 412, json: { error: { code: "notAllowed", message: "ETag does not match" } } });
        arquivo = req.postDataBuffer(); versao++;
        return rota.fulfill({ json: { id: "I1", name: "Faturamento.xlsm", eTag: etag(), lastModifiedDateTime: new Date().toISOString() } });
      }
      return rota.fulfill({ status: 404, json: { error: { code: "itemNotFound" } } });
    });
    await page.route("https://download.test/**", (rota) => rota.fulfill({ body: arquivo, headers: { "Access-Control-Allow-Origin": "*" } }));

    await page.clock.setFixedTime(new Date("2026-11-23T15:00:00Z"));
    await page.addInitScript(() => {
      localStorage.setItem("vmpay_demo", "1");
      localStorage.setItem("ms_client", "11111111-2222-3333-4444-555555555555");
      globalThis.__TOKEN_TESTE__ = "token-falso";
    });
    await page.goto(`http://127.0.0.1:${PORTA}/`);

    console.log("Escolher planilha no OneDrive");
    await page.click("#btnOneDrive");
    await page.waitForSelector(".item-od");
    ok(await page.locator(".item-od").count() === 1, "lista só a planilha (ignora o .docx)");
    ok((await page.textContent(".item-od")).includes("/Financeiro"), "mostra a pasta");
    await page.click(".item-od");
    await page.waitForSelector("#infoPlanilha .tag");
    ok((await page.textContent("#infoPlanilha")).includes("salva na original"), "aberta no modo OneDrive");

    console.log("Atualizar vendas → salva sozinho na original");
    await page.click("#btnAtualizar");
    await page.waitForSelector("text=Salvo na planilha original", { timeout: 60000 });
    ok(puts.length === 1 && puts[0].ifMatch === '"{AAAA-1111},1"', "PUT com If-Match da versão aberta");
    fs.writeFileSync(path.join(saidaDir, "onedrive-v2.xlsm"), arquivo);
    await page.screenshot({ path: path.join(saidaDir, "5-onedrive.png"), fullPage: true });

    console.log("Conferir → salva de novo com a versão nova");
    await page.fill("#dtConf", "2026-11-22");
    await page.click("#btnConferir");
    await page.waitForFunction(() => document.querySelector("#resConferir")?.textContent.includes("Salvo na planilha original"), null, { timeout: 60000 });
    ok(puts.length === 2 && puts[1].ifMatch === '"{AAAA-1111},2"', "segundo PUT usa o eTag atualizado");

    console.log("Alguém alterou a planilha no meio do caminho");
    recusar412 = true;
    await page.fill("#dtIni", "2026-11-01"); await page.dispatchEvent("#dtIni", "change");
    await page.click("#btnAtualizar");
    await page.waitForSelector("#dlgPergunta[open]", { timeout: 90000 });
    await page.click("#pgBotoes button[value=nao]");
    await page.waitForSelector("#resAtualizar .erro");
    ok((await page.textContent("#resAtualizar .erro")).includes("foi alterada no OneDrive"), "412 vira aviso claro, sem sobrescrever");
    ok(await page.isVisible("#btnSalvarOriginal"), "botão para tentar salvar de novo");
    recusar412 = false;

    console.log("Abrir pelo link e reabrir a última");
    await page.reload();
    ok((await page.textContent("#odUltima")).includes("Abrir Faturamento.xlsm"), "atalho para a última planilha");
    await page.click("#btnOneDrive");
    await page.waitForSelector("#odLink");
    await page.fill("#odLink", "https://empresa.sharepoint.com/:x:/g/Faturamento?e=abc");
    await page.click("#btnOdLink");
    await page.waitForSelector("#infoPlanilha .tag");
    ok((await page.textContent("#infoPlanilha")).includes("7 dias") || (await page.textContent("#infoPlanilha")).includes("22 de 30"), "reabriu a versão salva: " + (await page.textContent("#infoPlanilha")));
    ok(erros.length === 0, "sem erros na página: " + erros.join(" | "));
    console.log("OK - modo OneDrive passou");
  } finally { await nav.close(); srv.kill(); }
})().catch((e) => { console.error(e); process.exit(1); });
