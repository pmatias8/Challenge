// Monta os 3 scripts finais (Office Scripts precisa de 1 arquivo por script).
// Uso: node office-scripts/montar.mjs
import { readFileSync, writeFileSync } from "fs";
const dir = new URL("./fonte/", import.meta.url);
const ler = (f) => readFileSync(new URL(f, dir), "utf8");
const config = ler("comum_config.ts"), comum = ler("comum.ts");
const aviso = "// ARQUIVO GERADO por office-scripts/montar.mjs a partir de office-scripts/fonte/ — copie e cole inteiro no Excel.\n";
for (const [fonte, saida] of [["atualizar.ts", "AtualizarVendasAPI.ts"], ["conferir.ts", "ConferirVendasPorLoja.ts"], ["diagnosticar.ts", "DiagnosticarVendasAPI.ts"]]) {
  const s = ler(fonte).replace("//@@COMUM_CONFIG@@\n", config);
  writeFileSync(new URL("./" + saida, import.meta.url), aviso + s.trimEnd() + "\n\n" + comum.trimEnd() + "\n");
  console.log("ok", saida);
}
