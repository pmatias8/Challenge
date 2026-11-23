# Macros de faturamento no celular e no iPad

Este guia mostra como rodar a macro `modVendasAPI` (vendas da VMpay → aba **Diário**) sem um computador com Windows. Há dois caminhos:

| | **Opção 2 – Office Scripts** | **Opção 3 – App web (protótipo)** |
|---|---|---|
| Onde roda | Excel na Web (navegador) | Qualquer navegador; dá para instalar na tela inicial |
| A planilha | Fica no OneDrive/SharePoint e é editada lá | Você abre o arquivo, o app altera e você salva de volta |
| Precisa de | Conta Microsoft 365 corporativa com Office Scripts liberado | Nada além do navegador |
| Arquivos | [`office-scripts/`](office-scripts/) | [`docs/`](docs/) |

> ⚠️ **Descoberta importante:** a API da VMpay **não envia os cabeçalhos CORS**. Isso foi testado em 23/11/2026: a resposta não traz `Access-Control-Allow-Origin`. Sem esse cabeçalho, o navegador bloqueia a resposta. Isso afeta o **Excel na Web** e também o **app**. O Excel no Windows não aplica essa regra, e é por isso que a macro VBA funciona no trabalho.
> A solução é um **proxy** pequeno e gratuito, explicado no passo 0. Ele vale para as duas opções.

---

## Passo 0 – Criar o proxy (uma vez, cerca de 5 minutos, grátis)

1. Crie uma conta grátis em <https://dash.cloudflare.com/sign-up>.
2. No menu, abra **Workers & Pages → Create → Create Worker**. Dê o nome `vmpay-proxy` e clique em **Deploy**.
3. Clique em **Edit code**, apague o que estiver lá e cole todo o conteúdo de [`proxy/cloudflare-worker.js`](proxy/cloudflare-worker.js). Depois clique em **Deploy**.
4. Anote o endereço, que tem o formato `https://vmpay-proxy.SEU-USUARIO.workers.dev`. Ao abrir esse endereço no navegador, deve aparecer "Proxy VMpay no ar".

O proxy só repassa pedidos `GET /api/v1/...` para a VMpay e não guarda nada. O token continua indo dentro de cada chamada, então o endereço sozinho não dá acesso a nada. O plano grátis permite 100 mil chamadas por dia, e a macro faz poucas dezenas por dia consultado.

---

## Opção 2 – Office Scripts (Excel na Web)

### Instalar
1. Abra a planilha de faturamento no **Excel na Web** (office.com → OneDrive → planilha).
2. Na aba **Automatizar**, clique em **Novo script**.
3. Apague o exemplo e cole o conteúdo de [`office-scripts/AtualizarVendasAPI.ts`](office-scripts/AtualizarVendasAPI.ts).
4. No topo do script, troque o valor de `API_BASE` pelo endereço do seu proxy seguido de `/api/v1`. Exemplo: `https://vmpay-proxy.SEU-USUARIO.workers.dev/api/v1`.
5. (Opcional) Cole o token em `TOKEN_VMPAY`. Leia o aviso no próprio script: se você criar um **botão** numa planilha compartilhada, o script vai junto com o token.
6. Salve com o nome **Atualizar vendas API**.
7. Repita os passos 2 a 6 para [`ConferirVendasPorLoja.ts`](office-scripts/ConferirVendasPorLoja.ts) e [`DiagnosticarVendasAPI.ts`](office-scripts/DiagnosticarVendasAPI.ts).
8. (Recomendado) No menu `…` de cada script, escolha **Adicionar no livro**. Isso cria um botão na planilha.

### Usar
Ao clicar em **Executar** ou no botão, o Excel mostra campos no lugar das caixas de pergunta da macro. Todos são opcionais:

| Campo | Equivale a | Vazio significa |
|---|---|---|
| `dataInicio` / `dataFim` (dd/mm/aaaa) | os 2 `InputBox` de datas | a mesma sugestão da macro: primeiro dia vazio até ontem |
| `substituirPreenchidos` | o "Sim/Não" sobre os dias já preenchidos | **Não**: mantém o que está e mostra a diferença na aba Vendas API |
| `atualizarCuponsPainel` | a pergunta do Ticket do mês | **Não**: só informa a estimativa |
| `token` | o `InputBox` do token | usa o valor de `TOKEN_VMPAY` |

O resumo da execução fica em **Vendas API!I1:I4**, porque o Office Scripts não tem `MsgBox`.

### No iPad e no Android
- Os apps nativos do Excel **não rodam** Office Scripts.
- **iPad:** abra a planilha no **Safari** (office.com), não no app do Excel. O Safari do iPad carrega a versão completa do Excel na Web, com a aba Automatizar e os botões de script.
- **Android:** no Chrome, ative **⋮ → Site para computador** antes de abrir a planilha no office.com.
- A Microsoft só suporta oficialmente o Excel na Web em computador, então **teste no seu aparelho** antes de depender disso.

### E o Power Automate (botão no celular)?
Para esta macro, ele **não serve**. O Power Automate bloqueia scripts que chamam APIs externas (`fetch`), e o motivo é a política de proteção de dados da Microsoft ([documentação](https://learn.microsoft.com/office/dev/scripts/develop/external-calls)). A alternativa seria um fluxo com o conector **HTTP**, que é premium e exige licença paga do Power Automate. Esse fluxo faria as chamadas e entregaria os dados para um script que só grava. Dá para montar, se valer a pena.

---

## Opção 3 – App web (protótipo)

### Publicar (uma vez)
No GitHub, abra o repositório e vá em **Settings → Pages**. Em *Build and deployment*, escolha **Deploy from a branch**, branch `claude/excel-macros-mobile-0xbsp0` e pasta **/docs**, e salve. Em 1 ou 2 minutos o app fica em `https://pmatias8.github.io/Challenge/`.

> O repositório é **público**: qualquer pessoa pode ver o código do app e dos scripts, que inclui nomes de abas e o endereço da API. **Nenhum token ou dado da empresa vai para o repositório**, porque o token fica só no aparelho. Se preferir manter o código privado, dá para hospedar o app de graça no próprio Cloudflare (Pages).

### Instalar no aparelho
Abra o endereço acima.
- **iPad/iPhone:** no Safari, toque em Compartilhar → **Adicionar à Tela de Início**.
- **Android:** no Chrome, toque em ⋮ → **Instalar app**.

### Usar
1. Toque em ⚙︎ e cole o **token** e o **endereço do proxy**. Esses dados ficam só no aparelho. Para testar sem nada disso, ligue o **modo demonstração** e use a [planilha de exemplo](docs/exemplo/Faturamento-exemplo.xlsx).
2. Toque em **Abrir planilha** e escolha o arquivo `.xlsm`. O seletor de arquivos do iPad e do Android acessa o OneDrive direto.
3. As datas já vêm sugeridas. Toque em **Atualizar vendas**. Se houver dias já preenchidos, aparece a opção **Substituir**. No fim, o app pergunta sobre os cupons do Painel, como a macro.
4. Toque em **Salvar / Abrir no Excel**.
   - No iPad: escolha **Salvar em Arquivos → OneDrive** e substitua o arquivo, ou abra direto no Excel.
   - No Android: escolha o OneDrive ou o Excel.

**O que o app preserva:** ele não recria a planilha. Abre o arquivo e altera **só as células que a macro altera**. As **macros VBA continuam dentro do .xlsm**, junto com gráficos, formatação e fórmulas, e o Excel recalcula tudo ao abrir. Por isso a mesma planilha continua funcionando no computador do trabalho.

**Cuidado:** não edite a mesma planilha em outro lugar enquanto usa o app. O arquivo salvo pelo app substitui o do OneDrive.

---

## Diferenças em relação ao VBA
- As regras de cálculo são as mesmas nas duas versões: dia de Brasília (UTC−3), páginas de 1000 vendas em lotes de 8, filtro sem diferenciar acentos e maiúsculas, cupom estimado por máquina + horário, e nada é gravado se a API der erro.
- Sobre a coluna das datas: a macro procura as datas na coluna **B** (`COL_DATA = 2`), mas as mensagens de erro do VBA falam em "coluna I". As versões novas dizem "coluna B".
- No app, a **Conferência por loja** aparece na tela e também é gravada na aba "Conferência vendas". O **Diagnóstico** aparece só na tela.

## Testes (para quem mexer no código)
Os testes usam uma API falsa e uma planilha gerada com o mesmo layout.
```bash
# planilha de teste (precisa de python3 + openpyxl + LibreOffice)
python3 testes/gerar_exemplo.py /tmp/f.xlsx --filtro "ESCRITORIO TESTE"
soffice --headless --calc --convert-to xlsx --outdir /tmp/lo /tmp/f.xlsx
python3 testes/injetar_vba.py /tmp/lo/f.xlsx /tmp/teste.xlsm
# app no navegador (Playwright) + conferência do arquivo baixado
node testes/teste_app.cjs /tmp/teste.xlsm /tmp/saida
python3 testes/verificar_saida.py /tmp/teste.xlsm /tmp/saida/saida.xlsm /tmp/saida/esperado.json /tmp/saida/lo
# Office Scripts numa pasta simulada (precisa do pacote typescript)
node office-scripts/montar.mjs
node testes/office-scripts/rodar_mock.mjs
```
Os scripts em `office-scripts/*.ts` são **gerados** a partir de `office-scripts/fonte/`. Edite os arquivos da pasta `fonte` e rode `montar.mjs`.
