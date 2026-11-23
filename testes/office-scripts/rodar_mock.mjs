// Roda os Office Scripts fora do Excel, com uma pasta de trabalho simulada
// (ExcelScript em memória) e a API falsa do modo demonstração.
// Confere que dão o MESMO resultado que o app (opção 3).
// Uso: node testes/office-scripts/rodar_mock.mjs <caminho do typescript>
import { readFileSync } from "fs";
import vm from "vm";
import { createRequire } from "module";
import { fileURLToPath } from "url";
import path from "path";

const aqui = path.dirname(fileURLToPath(import.meta.url));
const raiz = path.resolve(aqui, "../..");
const ts = createRequire(import.meta.url)(process.argv[2] || "typescript");
const { fetchDemo, vendasDoDia } = await import(path.join(raiz, "docs/js/demo.js"));
const ok = (c, m) => { if (!c) throw new Error("FALHOU: " + m); console.log("  ✓ " + m); };

// ---------------- ExcelScript simulado ----------------
const colNum = (l) => [...l].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
class Aba {
  constructor(nome, pasta) { this.nome = nome; this.pasta = pasta; this.cel = new Map(); this.fmt = new Map(); this.congelada = 0; }
  getName() { return this.nome; }
  _get(r, c) { return this.cel.get(r + "," + c) ?? ""; }
  _set(r, c, v) { if (v === "" || v === null) this.cel.delete(r + "," + c); else this.cel.set(r + "," + c, v); }
  getRange(end) {
    let m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(end);
    if (m) {
      const r0 = +m[2] - 1, c0 = colNum(m[1]) - 1;
      const r1 = m[3] ? +m[4] - 1 : r0, c1 = m[3] ? colNum(m[3]) - 1 : c0;
      return new Intervalo(this, r0, c0, r1 - r0 + 1, c1 - c0 + 1);
    }
    m = /^([A-Z]+):([A-Z]+)$/.exec(end);
    if (m) return new Intervalo(this, 0, colNum(m[1]) - 1, 1048576, colNum(m[2]) - colNum(m[1]) + 1);
    throw new Error("endereço não suportado no simulador: " + end);
  }
  getRangeByIndexes(r, c, nr, nc) {
    if (nr < 1 || nc < 1) throw new Error("getRangeByIndexes com tamanho inválido (o Excel daria erro)");
    return new Intervalo(this, r, c, nr, nc);
  }
  getCell(r, c) { return new Intervalo(this, r, c, 1, 1); }
  getUsedRange() {
    let r1 = -1, c1 = -1;
    for (const k of this.cel.keys()) { const [r, c] = k.split(",").map(Number); r1 = Math.max(r1, r); c1 = Math.max(c1, c); }
    return r1 < 0 ? new Intervalo(this, 0, 0, 1, 1) : new Intervalo(this, 0, 0, r1 + 1, c1 + 1);
  }
  getFreezePanes() { return { freezeRows: (n) => { this.congelada = n; } }; }
  delete() { this.pasta.abas = this.pasta.abas.filter((a) => a !== this); }
  activate() { this.pasta.ativa = this.nome; }
}
class Intervalo {
  constructor(aba, r, c, nr, nc) { Object.assign(this, { aba, r, c, nr, nc }); }
  getValues() { return Array.from({ length: this.nr }, (_, i) => Array.from({ length: this.nc }, (_, j) => this.aba._get(this.r + i, this.c + j))); }
  setValues(v) {
    if (v.length !== this.nr || v.some((l) => l.length !== this.nc)) throw new Error(`setValues: tamanho ${v.length}x${v[0]?.length} ≠ intervalo ${this.nr}x${this.nc}`);
    v.forEach((l, i) => l.forEach((x, j) => this.aba._set(this.r + i, this.c + j, x)));
  }
  getValue() { return this.aba._get(this.r, this.c); }
  setValue(x) { this.aba._set(this.r, this.c, x); }
  setNumberFormat(f) { this.aba.fmt.set(`${this.r},${this.c}`, f); }
  getRowIndex() { return this.r; }
  getRowCount() { return this.nr; }
  getFormat() {
    const nada = () => {};
    return { getFont: () => ({ setBold: nada, setColor: nada, setSize: nada }), getFill: () => ({ setColor: nada }),
      setWrapText: nada, setRowHeight: nada, setColumnWidth: nada, autofitColumns: nada, setHorizontalAlignment: nada };
  }
}
class Pasta {
  constructor() { this.abas = []; }
  getWorksheets() { return [...this.abas]; }
  getWorksheet(n) { return this.abas.find((a) => a.nome === n); }
  addWorksheet(n) {
    if (this.abas.some((a) => a.nome.toLowerCase() === n.toLowerCase())) throw new Error("aba duplicada: " + n);
    const a = new Aba(n, this); this.abas.push(a); return a;
  }
}

// ---------------- planilha igual à de exemplo ----------------
const serial = (d, m, a) => Date.UTC(a, m - 1, d) / 86400000 + 25569;
function montarPasta() {
  const p = new Pasta();
  const painel = p.addWorksheet("Painel"); painel.getRange("D7").setValue(2150);
  const d = p.addWorksheet("Diário");
  for (let i = 0; i < 30; i++) {
    d.getCell(5 + i, 1).setValue(serial(1 + i, 11, 2026));
    d.getCell(5 + i, 3).setValue(4200);
    if (i < 15) d.getCell(5 + i, 4).setValue(Math.round((3900 + (i * 137) % 1900 + 0.37 * i) * 100) / 100);
  }
  const f = p.addWorksheet("Filtros vendas");
  f.getRange("A1").setValue("Máquinas / locais a NÃO somar no faturamento"); f.getRange("A2").setValue("ESCRITORIO TESTE");
  return p;
}

// ---------------- carrega um script ----------------
const AGORA = Date.parse("2026-11-23T15:00:00Z"); // 23/11/2026 12:00 em Brasília
function carregar(nome) {
  let fonte = readFileSync(path.join(raiz, "office-scripts", nome + ".ts"), "utf8");
  fonte = fonte.replace(/const API_BASE = ".*?";/, 'const API_BASE = "https://demo.local/api/v1";');
  const js = ts.transpileModule(fonte, { compilerOptions: { target: ts.ScriptTarget.ES2019, module: ts.ModuleKind.None } }).outputText;
  const logs = [];
  const ctx = vm.createContext({
    fetch: fetchDemo, console: { log: (...a) => logs.push(a.join(" ")) },
    ExcelScript: { HorizontalAlignment: { right: "Right", left: "Left", center: "Center", general: "General" } },
  });
  vm.runInContext(`Date.now = () => ${AGORA};`, ctx);
  vm.runInContext(js, ctx);
  return { main: vm.runInContext("main", ctx), logs };
}

// valores esperados direto dos dados falsos
const esperado = (dia) => {
  const vs = vendasDoDia(`${String(dia).padStart(2, "0")}/11/2026 03:00:00`).filter((v) => v.machine.name !== "Escritório Teste");
  return { valor: Math.round(vs.reduce((s, v) => s + v.value, 0) * 100) / 100,
           cupons: new Set(vs.map((v) => v.machine_id + "|" + v.occurred_at.slice(0, 19))).size };
};

console.log("AtualizarVendasAPI (sem parâmetros = sugestão da macro)");
const pasta = montarPasta();
const atualizar = carregar("AtualizarVendasAPI");
let r = await atualizar.main(pasta, undefined, undefined, "Não", "Não", "token-teste");
const diario = pasta.getWorksheet("Diário");
ok(r.includes("7 dia(s) gravado(s)") && r.includes("16/11/2026 a 22/11/2026"), r.split("\n")[0]);
for (let dia = 16; dia <= 22; dia++) {
  const v = diario.getCell(4 + dia, 4).getValue();
  if (v !== esperado(dia).valor) throw new Error(`dia ${dia}: ${v} ≠ ${esperado(dia).valor}`);
}
ok(true, "E21:E27 = mesmos valores que o app calculou (sem a loja filtrada)");
ok(diario.getCell(5, 4).getValue() === 3900 && diario.getCell(27, 4).getValue() === "", "dias 01 e 23 intactos");
ok(pasta.getWorksheet("Vendas API").getRange("I1").getValue().startsWith("7 dia(s)"), "resumo gravado em Vendas API!I1");

console.log("AtualizarVendasAPI 01/11–22/11, manter, atualizar cupons");
r = await atualizar.main(pasta, "01/11/2026", "", "Não", "Sim", "token-teste");
ok(r.includes("0 dia(s) gravado(s)") && r.includes("22 mantido(s)"), r.split("\n")[0]);
const log = pasta.getWorksheet("Vendas API");
ok(log.getRange("A2:A23").getValues().every((l) => typeof l[0] === "number") && log.getRange("A24").getValue() === "", "log com 22 dias, sem duplicar");
ok(pasta.getWorksheet("Painel").getRange("D7").getValue() === 21901, "Painel!D7 = 21901 (mesmo número do app)");
const l1 = log.getRange("A2:G23").getValues().find((l) => l[0] === serial(1, 11, 2026));
ok(l1[4] === 3900 && l1[5] === Math.round((esperado(1).valor - 3900) * 100) / 100, "log de 01/11 mostra valor antes (3900) e diferença " + l1[5]);

console.log("AtualizarVendasAPI substituindo");
r = await atualizar.main(pasta, "01/11/2026", "02/11/2026", "Sim", "Não", "token-teste");
ok(diario.getCell(5, 4).getValue() === esperado(1).valor, "01/11 substituído pelo valor da API");

console.log("Erros");
await atualizar.main(pasta, "31/02/2026x", "", "Não", "Não", "t").then(() => { throw new Error("devia falhar"); }, (e) => ok(/inválida/.test(e.message), e.message));
await atualizar.main(pasta, "", "", "Não", "Não", "").then(() => { throw new Error("devia falhar"); }, (e) => ok(/token/.test(e.message), e.message));

console.log("ConferirVendasPorLoja");
const conf = carregar("ConferirVendasPorLoja");
r = await conf.main(pasta, "22/11/2026", "token-teste");
const wc = pasta.getWorksheet("Conferência vendas");
const linhas = wc.getRange("A4:G8").getValues();
ok(linhas.filter((l) => l[6] === "Sim").length === 1 && linhas[0][3] >= linhas[4][3], "5 lojas, ordenadas por valor, 1 excluída");
ok(wc.getRange("D10").getValue() === esperado(22).valor, "total sem as excluídas = " + wc.getRange("D10").getValue());
await conf.main(pasta, "21/11/2026", "token-teste");
ok(pasta.getWorksheets().filter((a) => a.getName() === "Conferência vendas").length === 1, "rodar de novo recria a aba");

console.log("DiagnosticarVendasAPI");
const diag = carregar("DiagnosticarVendasAPI");
await diag.main(pasta, "", "token-teste");
ok(String(pasta.getWorksheet("DiagVendas").getRange("A3").getValue()).includes("occurred_at"), "resposta crua gravada em DiagVendas");
console.log("OK - Office Scripts passaram");
