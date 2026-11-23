// ============================================================
//  Editor "cirúrgico" de arquivos .xlsx / .xlsm
//
//  Em vez de reconstruir a planilha (o que perderia macros, gráficos,
//  formatação condicional etc.), abrimos o ZIP e alteramos só o XML
//  das células que a macro mexe. Todo o resto do arquivo - inclusive
//  o vbaProject.bin com as macros - volta intacto.
// ============================================================

const NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const NS_XML = "http://www.w3.org/XML/1998/namespace";
const TIPO_ABA = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const CT_ABA = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";

// ---------- endereços ----------

export function colParaNum(letras) {
  let n = 0;
  for (const c of letras.toUpperCase()) n = n * 26 + (c.charCodeAt(0) - 64);
  return n;
}
export function numParaCol(n) {
  let s = "";
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
export function separarRef(ref) {
  const m = /^([A-Z]+)(\d+)$/i.exec(ref.replace(/\$/g, ""));
  if (!m) throw new Error("Endereço inválido: " + ref);
  return { col: colParaNum(m[1]), lin: Number(m[2]) };
}
const ref = (lin, col) => numParaCol(col) + lin;

function filhos(el, nome) {
  const out = [];
  for (let n = el.firstChild; n; n = n.nextSibling) if (n.nodeType === 1 && n.localName === nome) out.push(n);
  return out;
}
const filho = (el, nome) => filhos(el, nome)[0] || null;

function caminhoRelativo(base, alvo) {
  if (alvo.startsWith("/")) return alvo.slice(1);
  const partes = base.split("/"); partes.pop();
  for (const p of alvo.split("/")) {
    if (p === "..") partes.pop(); else if (p !== ".") partes.push(p);
  }
  return partes.join("/");
}

// ============================================================
export class Pasta {
  constructor(zip, deps) {
    this.zip = zip;
    this.deps = deps;
    this.xml = new Map();      // caminho -> Document (parse sob demanda)
    this.sujos = new Set();    // caminhos alterados
    this.abas = new Map();     // nome em minúsculas -> Aba
  }

  /** Abre um ArrayBuffer/Uint8Array/Blob de um .xlsx ou .xlsm. */
  static async abrir(dados, deps = {}) {
    const d = {
      JSZip: deps.JSZip || globalThis.JSZip,
      DOMParser: deps.DOMParser || globalThis.DOMParser,
      XMLSerializer: deps.XMLSerializer || globalThis.XMLSerializer,
    };
    let zip;
    try { zip = await d.JSZip.loadAsync(dados); } catch {
      throw new Error("Esse arquivo não parece ser uma planilha .xlsx/.xlsm (o arquivo .xls antigo não é suportado).");
    }
    if (!zip.file("xl/workbook.xml")) throw new Error("Arquivo sem xl/workbook.xml: não é uma planilha do Excel.");
    const p = new Pasta(zip, d);
    await p._carregarIndice();
    return p;
  }

  async _doc(caminho) {
    if (this.xml.has(caminho)) return this.xml.get(caminho);
    const f = this.zip.file(caminho);
    if (!f) return null;
    const doc = new this.deps.DOMParser().parseFromString(await f.async("string"), "application/xml");
    if (doc.getElementsByTagName("parsererror").length) throw new Error("XML inválido em " + caminho);
    this.xml.set(caminho, doc);
    return doc;
  }
  _sujar(caminho) { this.sujos.add(caminho); }

  async _carregarIndice() {
    this.wb = await this._doc("xl/workbook.xml");
    this.wbRels = await this._doc("xl/_rels/workbook.xml.rels");
    this.tipos = await this._doc("[Content_Types].xml");
    const pr = this.wb.getElementsByTagNameNS(NS, "workbookPr")[0];
    this.data1904 = !!pr && /^(1|true)$/.test(pr.getAttribute("date1904") || "");

    const alvos = new Map();
    for (const r of filhos(this.wbRels.documentElement, "Relationship")) alvos.set(r.getAttribute("Id"), r);
    this.listaAbas = [];
    const sheets = this.wb.getElementsByTagNameNS(NS, "sheets")[0];
    for (const s of filhos(sheets, "sheet")) {
      const rid = s.getAttributeNS(NS_R, "id");
      const rel = alvos.get(rid);
      if (!rel || !/\/worksheet$/.test(rel.getAttribute("Type"))) continue; // ignora gráficos-aba etc.
      this.listaAbas.push({ nome: s.getAttribute("name"), caminho: caminhoRelativo("xl/workbook.xml", rel.getAttribute("Target")) });
    }
    // textos compartilhados
    this.textos = [];
    const relSS = filhos(this.wbRels.documentElement, "Relationship").find((r) => /\/sharedStrings$/.test(r.getAttribute("Type")));
    if (relSS) {
      const ss = await this._doc(caminhoRelativo("xl/workbook.xml", relSS.getAttribute("Target")));
      for (const si of ss.getElementsByTagNameNS(NS, "si")) {
        let t = "";
        const coletar = (el) => {
          for (let n = el.firstChild; n; n = n.nextSibling) {
            if (n.nodeType !== 1) continue;
            if (n.localName === "rPh") continue;      // texto fonético (japonês) não entra
            if (n.localName === "t") t += n.textContent; else coletar(n);
          }
        };
        coletar(si);
        this.textos.push(t);
      }
    }
  }

  nomesAbas() { return this.listaAbas.map((a) => a.nome); }

  /** Igual ao TemAba do VBA (não diferencia maiúsculas). */
  temAba(nome) { return this.listaAbas.some((a) => a.nome.toLowerCase() === nome.toLowerCase()); }

  async aba(nome) {
    const chave = nome.toLowerCase();
    if (this.abas.has(chave)) return this.abas.get(chave);
    const info = this.listaAbas.find((a) => a.nome.toLowerCase() === chave);
    if (!info) return null;
    const a = new Aba(this, info.nome, info.caminho, await this._doc(info.caminho));
    this.abas.set(chave, a);
    return a;
  }

  /** Cria uma aba no fim da pasta (igual a Worksheets.Add After:=última). */
  async criarAba(nome) {
    if (this.temAba(nome)) throw new Error("Já existe a aba " + nome);
    if (/[\\/?*[\]:]/.test(nome) || nome.length > 31) throw new Error("Nome de aba inválido: " + nome);
    const sheets = this.wb.getElementsByTagNameNS(NS, "sheets")[0];
    const ids = filhos(sheets, "sheet").map((s) => Number(s.getAttribute("sheetId")) || 0);
    const sheetId = Math.max(0, ...ids) + 1;
    const relsEl = this.wbRels.documentElement;
    const rids = new Set(filhos(relsEl, "Relationship").map((r) => r.getAttribute("Id")));
    let n = 1; while (rids.has("rId" + n)) n++;
    const rid = "rId" + n;
    let k = 1; while (this.zip.file(`xl/worksheets/sheet${k}.xml`)) k++;
    const caminho = `xl/worksheets/sheet${k}.xml`;

    this.zip.file(caminho,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
      `<worksheet xmlns="${NS}" xmlns:r="${NS_R}"><sheetViews><sheetView workbookViewId="0"/></sheetViews>` +
      `<sheetFormatPr defaultRowHeight="15"/><sheetData/>` +
      `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`);

    const rel = this.wbRels.createElementNS(NS_PKG_REL, "Relationship");
    rel.setAttribute("Id", rid); rel.setAttribute("Type", TIPO_ABA);
    rel.setAttribute("Target", `worksheets/sheet${k}.xml`);
    relsEl.appendChild(rel); this._sujar("xl/_rels/workbook.xml.rels");

    const s = this.wb.createElementNS(NS, sheets.prefix ? sheets.prefix + ":sheet" : "sheet");
    s.setAttribute("name", nome); s.setAttribute("sheetId", String(sheetId));
    s.setAttributeNS(NS_R, "r:id", rid);
    sheets.appendChild(s); this._sujar("xl/workbook.xml");

    const ov = this.tipos.createElementNS(NS_CT, "Override");
    ov.setAttribute("PartName", "/" + caminho); ov.setAttribute("ContentType", CT_ABA);
    this.tipos.documentElement.appendChild(ov); this._sujar("[Content_Types].xml");

    this.listaAbas.push({ nome, caminho });
    return this.aba(nome);
  }

  // ---------- estilos ----------

  /**
   * Devolve o índice de um estilo de célula (cellXfs) com as opções pedidas,
   * criando-o no styles.xml se preciso. Formatos numéricos usam os códigos
   * internos do Excel, que se adaptam ao idioma: 14 = data curta (dd/mm/aaaa
   * no Brasil), 22 = data e hora, 3 = #.##0, 4 = #.##0,00, 1 = 0.
   */
  async estilo({ numFmt = 0, negrito = false, corFonte = null, fundo = null, quebrar = false, alinhar = null } = {}) {
    const chave = JSON.stringify([numFmt, negrito, corFonte, fundo, quebrar, alinhar]);
    this._cacheEstilos = this._cacheEstilos || new Map();
    if (this._cacheEstilos.has(chave)) return this._cacheEstilos.get(chave);

    const relSt = filhos(this.wbRels.documentElement, "Relationship").find((r) => /\/styles$/.test(r.getAttribute("Type")));
    if (!relSt) return 0; // pasta sem styles.xml: fica no estilo padrão
    const caminho = caminhoRelativo("xl/workbook.xml", relSt.getAttribute("Target"));
    const st = await this._doc(caminho);
    const raiz = st.documentElement;
    const novo = (nome) => st.createElementNS(NS, raiz.prefix ? raiz.prefix + ":" + nome : nome);
    const contar = (el) => el.setAttribute("count", String(filhos(el, el.localName === "cellXfs" ? "xf" : el.localName.replace(/s$/, "")).length));

    let fontId = 0;
    if (negrito || corFonte) {
      const fonts = filho(raiz, "fonts");
      const f = filhos(fonts, "font")[0].cloneNode(true);
      for (const tag of ["b", "color"]) for (const x of filhos(f, tag)) f.removeChild(x);
      if (negrito) f.insertBefore(novo("b"), f.firstChild);
      if (corFonte) {
        const c = novo("color"); c.setAttribute("rgb", "FF" + corFonte);
        const depois = filho(f, "name") || filho(f, "family") || filho(f, "scheme");
        f.insertBefore(c, depois);
      }
      fonts.appendChild(f); contar(fonts);
      fontId = filhos(fonts, "font").length - 1;
    }
    let fillId = 0;
    if (fundo) {
      const fills = filho(raiz, "fills");
      const fl = novo("fill"), pf = novo("patternFill"), fg = novo("fgColor"), bg = novo("bgColor");
      pf.setAttribute("patternType", "solid"); fg.setAttribute("rgb", "FF" + fundo); bg.setAttribute("indexed", "64");
      pf.appendChild(fg); pf.appendChild(bg); fl.appendChild(pf); fills.appendChild(fl); contar(fills);
      fillId = filhos(fills, "fill").length - 1;
    }
    const xfs = filho(raiz, "cellXfs");
    const xf = novo("xf");
    xf.setAttribute("numFmtId", String(numFmt)); xf.setAttribute("fontId", String(fontId));
    xf.setAttribute("fillId", String(fillId)); xf.setAttribute("borderId", "0"); xf.setAttribute("xfId", "0");
    if (numFmt) xf.setAttribute("applyNumberFormat", "1");
    if (fontId) xf.setAttribute("applyFont", "1");
    if (fillId) xf.setAttribute("applyFill", "1");
    if (quebrar || alinhar) {
      xf.setAttribute("applyAlignment", "1");
      const al = novo("alignment");
      if (alinhar) al.setAttribute("horizontal", alinhar);
      if (quebrar) al.setAttribute("wrapText", "1");
      xf.appendChild(al);
    }
    xfs.appendChild(xf); contar(xfs);
    this._sujar(caminho);
    const idx = filhos(xfs, "xf").length - 1;
    this._cacheEstilos.set(chave, idx);
    return idx;
  }

  // ---------- gravação ----------

  /** Gera o arquivo final (mesmo formato do original: .xlsx ou .xlsm). */
  async gerar(tipo = "uint8array") {
    // calcChain.xml: o Excel reconstrói sozinho; se ficar desatualizado ele "repara" o arquivo.
    if (this.zip.file("xl/calcChain.xml")) {
      this.zip.remove("xl/calcChain.xml");
      for (const r of filhos(this.wbRels.documentElement, "Relationship"))
        if (/\/calcChain$/.test(r.getAttribute("Type"))) r.parentNode.removeChild(r);
      for (const o of filhos(this.tipos.documentElement, "Override"))
        if (o.getAttribute("PartName") === "/xl/calcChain.xml") o.parentNode.removeChild(o);
      this._sujar("xl/_rels/workbook.xml.rels"); this._sujar("[Content_Types].xml");
    }
    // Recalcular fórmulas ao abrir (os valores guardados ficaram velhos).
    let calc = this.wb.getElementsByTagNameNS(NS, "calcPr")[0];
    if (!calc) {
      const raiz = this.wb.documentElement;
      calc = this.wb.createElementNS(NS, raiz.prefix ? raiz.prefix + ":calcPr" : "calcPr");
      const depois = ["oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing",
        "fileRecoveryPr", "webPublishObjects", "extLst"].map((n) => filho(raiz, n)).find(Boolean);
      raiz.insertBefore(calc, depois || null);
    }
    calc.setAttribute("fullCalcOnLoad", "1");
    this._sujar("xl/workbook.xml");

    for (const a of this.abas.values()) a._finalizar();
    const ser = new this.deps.XMLSerializer();
    for (const caminho of this.sujos) {
      let txt = ser.serializeToString(this.xml.get(caminho));
      if (!txt.startsWith("<?xml")) txt = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + txt;
      this.zip.file(caminho, txt);
    }
    return this.zip.generateAsync({ type: tipo, compression: "DEFLATE", compressionOptions: { level: 6 } });
  }
}

// ============================================================
export class Aba {
  constructor(pasta, nome, caminho, doc) {
    this.pasta = pasta; this.nome = nome; this.caminho = caminho; this.doc = doc;
    this.raiz = doc.documentElement;
    this.pre = this.raiz.prefix ? this.raiz.prefix + ":" : "";
    this.dados = filho(this.raiz, "sheetData");
    this.linhas = new Map();
    for (const r of filhos(this.dados, "row")) this.linhas.set(Number(r.getAttribute("r")), r);
    this.colEstilo = new Map();
    const cols = filho(this.raiz, "cols");
    if (cols) for (const c of filhos(cols, "col")) {
      if (!c.getAttribute("style")) continue;
      for (let i = +c.getAttribute("min"); i <= +c.getAttribute("max"); i++) this.colEstilo.set(i, +c.getAttribute("style"));
    }
  }
  _novo(nome) { return this.doc.createElementNS(NS, this.pre + nome); }
  _sujar() { this.pasta._sujar(this.caminho); }

  _celula(lin, col, criar) {
    let row = this.linhas.get(lin);
    if (!row) {
      if (!criar) return null;
      row = this._novo("row"); row.setAttribute("r", String(lin));
      let depois = null;
      for (const r of filhos(this.dados, "row")) if (Number(r.getAttribute("r")) > lin) { depois = r; break; }
      this.dados.insertBefore(row, depois);
      this.linhas.set(lin, row);
    }
    const alvo = ref(lin, col);
    let depois = null;
    for (const c of filhos(row, "c")) {
      const r = c.getAttribute("r");
      const cc = r ? separarRef(r).col : null;
      if (r === alvo || cc === col) return c;
      if (cc !== null && cc > col) { depois = c; break; }
    }
    if (!criar) return null;
    const c = this._novo("c"); c.setAttribute("r", alvo);
    const est = this.colEstilo.get(col) ?? (row.getAttribute("customFormat") === "1" ? row.getAttribute("s") : null);
    if (est !== null && est !== undefined) c.setAttribute("s", String(est));
    row.insertBefore(c, depois);
    row.removeAttribute("spans"); // dica de otimização; tirar é sempre válido
    return c;
  }

  /** Valor da célula: number | string | boolean | null (vazia). Fórmulas devolvem o último valor calculado. */
  valor(endereco) {
    const { lin, col } = separarRef(endereco);
    return this.valorEm(lin, col);
  }
  valorEm(lin, col) {
    const c = this._celula(lin, col, false);
    if (!c) return null;
    const t = c.getAttribute("t") || "n";
    if (t === "inlineStr") { const is = filho(c, "is"); return is ? is.textContent : ""; }
    const v = filho(c, "v");
    if (!v) return null;
    const txt = v.textContent;
    if (t === "s") return this.pasta.textos[Number(txt)] ?? "";
    if (t === "str" || t === "e") return txt;
    if (t === "b") return txt === "1";
    if (t === "d") return txt; // raro (data ISO)
    const n = Number(txt);
    return Number.isFinite(n) ? n : null;
  }

  /** Grava número, texto ou null (limpa). Mantém o estilo existente da célula. */
  definir(endereco, valor, estilo) {
    const { lin, col } = separarRef(endereco);
    this.definirEm(lin, col, valor, estilo);
  }
  definirEm(lin, col, valor, estilo) {
    const c = this._celula(lin, col, true);
    for (const f of filhos(c, "f")) {
      if (f.getAttribute("t") === "shared" && f.getAttribute("ref"))
        throw new Error(`A célula ${this.nome}!${ref(lin, col)} é a origem de uma fórmula copiada; não vou sobrescrever.`);
      c.removeChild(f);
    }
    for (const tag of ["v", "is"]) for (const x of filhos(c, tag)) c.removeChild(x);
    c.removeAttribute("t");
    if (estilo !== undefined && estilo !== null && !c.getAttribute("s")) c.setAttribute("s", String(estilo));
    if (valor === null || valor === undefined || valor === "") { /* vazia */ }
    else if (typeof valor === "number") {
      if (!Number.isFinite(valor)) throw new Error("Número inválido para " + ref(lin, col));
      const v = this._novo("v"); v.textContent = String(valor); c.appendChild(v);
    } else if (typeof valor === "boolean") {
      c.setAttribute("t", "b"); const v = this._novo("v"); v.textContent = valor ? "1" : "0"; c.appendChild(v);
    } else {
      c.setAttribute("t", "inlineStr");
      const is = this._novo("is"), t = this._novo("t");
      t.setAttributeNS(NS_XML, "xml:space", "preserve");
      t.textContent = String(valor);
      is.appendChild(t); c.appendChild(is);
    }
    this._sujar();
  }

  /** Igual a Cells(Rows.Count, col).End(xlUp).Row: última linha com valor na coluna (1 se nenhuma). */
  ultimaLinha(col) {
    let ult = 1;
    for (const [lin] of this.linhas) {
      if (lin <= ult) continue;
      const v = this.valorEm(lin, col);
      if (v !== null && v !== "") ult = lin;
    }
    return ult;
  }

  /**
   * Largura (em "caracteres", como ColumnWidth do VBA) e, opcionalmente, o
   * estilo padrão das colunas de..ate (como Columns("B").NumberFormat).
   * Só para abas criadas aqui: se já houver definição nessas colunas, não mexe.
   */
  formatarColunas(de, ate, largura, estilo) {
    let cols = filho(this.raiz, "cols");
    if (!cols) { cols = this._novo("cols"); this.raiz.insertBefore(cols, this.dados); }
    for (const x of filhos(cols, "col")) if (+x.getAttribute("min") <= ate && +x.getAttribute("max") >= de) return;
    const c = this._novo("col");
    c.setAttribute("min", String(de)); c.setAttribute("max", String(ate));
    c.setAttribute("width", String(largura + 0.7109375)); c.setAttribute("customWidth", "1");
    if (estilo) { c.setAttribute("style", String(estilo)); for (let i = de; i <= ate; i++) this.colEstilo.set(i, estilo); }
    let depois = null;
    for (const x of filhos(cols, "col")) if (+x.getAttribute("min") > ate) { depois = x; break; }
    cols.insertBefore(c, depois);
    this._sujar();
  }

  /** Apaga todas as células da aba (para reescrever um relatório). */
  limparTudo() {
    for (const r of filhos(this.dados, "row")) this.dados.removeChild(r);
    this.linhas.clear();
    this._sujar();
  }

  alturaLinha(lin, pontos) {
    this._celula(lin, 1, true);
    const row = this.linhas.get(lin);
    row.setAttribute("ht", String(pontos)); row.setAttribute("customHeight", "1");
    this._sujar();
  }

  /** Congela a linha 1 (SplitRow = 1 + FreezePanes). */
  congelarLinha1() {
    let views = filho(this.raiz, "sheetViews");
    if (!views) { views = this._novo("sheetViews"); this.raiz.insertBefore(views, filho(this.raiz, "sheetFormatPr") || filho(this.raiz, "cols") || this.dados); }
    let v = filho(views, "sheetView");
    if (!v) { v = this._novo("sheetView"); v.setAttribute("workbookViewId", "0"); views.appendChild(v); }
    for (const x of filhos(v, "pane")) v.removeChild(x);
    for (const x of filhos(v, "selection")) v.removeChild(x);
    const p = this._novo("pane");
    p.setAttribute("ySplit", "1"); p.setAttribute("topLeftCell", "A2");
    p.setAttribute("activePane", "bottomLeft"); p.setAttribute("state", "frozen");
    v.insertBefore(p, v.firstChild);
    const s = this._novo("selection"); s.setAttribute("pane", "bottomLeft");
    p.parentNode.insertBefore(s, p.nextSibling);
    this._sujar();
  }

  _finalizar() {
    if (!this.pasta.sujos.has(this.caminho)) return;
    // atualiza <dimension> para cobrir todas as células
    let maxL = 1, maxC = 1;
    for (const [lin, row] of this.linhas) for (const c of filhos(row, "c")) {
      const r = c.getAttribute("r"); if (!r) continue;
      const p = separarRef(r); if (lin > maxL) maxL = lin; if (p.col > maxC) maxC = p.col;
    }
    const dim = filho(this.raiz, "dimension");
    if (dim) dim.setAttribute("ref", "A1:" + ref(maxL, maxC));
  }
}
