# Fatura parcial por imagem (OCR) + Faturas em 2 blocos — Plano

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Subir um print do app do Itaú (imagem) como fatura **parcial**, lido por
OCR grátis no navegador, e separar a aba Faturas em dois blocos (parcial × fechada).

**Architecture:** App single-file (`index.html`, HTML+CSS+JS). OCR client-side com
**Tesseract.js** (CDN, carregado sob demanda). O parser separa a parte pura
(reconstrução das linhas a partir das palavras+posições — testável em node) da
cola de navegador (rodar o Tesseract e montar o resultado). A imagem sempre vira
`STATE.faturaAberta` (parcial). O **lugar** de upload (bloco parcial × fechada)
decide o tipo, sobrepondo a detecção por cabeçalho.

**Tech Stack:** JS vanilla; Tesseract.js 5 (CDN); reaproveita
`parseItauExcel`/`classificarTransacoes`/`mostrarRevisao`/`salvarFatura`/`conciliarPendentes`.

## Global Constraints (do spec)

- **Nada é apagado** — só reorganizado; registro manual ("Registrar gasto agora")
  fica idêntico.
- **A fatura fechada é a fonte da verdade;** a parcial (Excel/imagem) é retrato
  provisório que o fechamento descarta. Imagem = **sempre parcial**.
- Sem backend, sem chave de IA, sem custo. Tudo no navegador.
- Português (pt-BR). Valores no formato do banco (`1.234,56` → ponto=milhar).
- Cartão da imagem = `STATE.cartaoAtivo` (a imagem não mostra número).
- Verificação: (a) `node -e "new Function(script)"` (sintaxe), (b) harness node
  para a função pura, (c) navegador contra a imagem real, (d) revisão adversarial.

---

### Task 1: Aba Faturas em 2 blocos + flag `forcarAberta`

**Files:**
- Modify: `index.html` — HTML de `panel-upload` (bloco 1 ~linha 306, hoje um só);
  `processarArquivos` (~2212); handler dos file inputs.

**Interfaces:**
- Produz: `processarArquivos(files, forcarAberta)` onde `forcarAberta` é `true`
  (bloco parcial), `false` (bloco fechada) ou `undefined` (mantém a detecção por
  cabeçalho — retrocompat). `FATURA_TEMP.aberta` passa a ser
  `forcarAberta !== undefined ? forcarAberta : !!result.aberta`.

- [ ] **Step 1:** Reescrever o HTML de `panel-upload` em dois cards:
  - Card "Fatura parcial (acompanhamento)": texto + dropzone/botão que abre
    `#fileInputParcial` com `accept=".png,.jpg,.jpeg,.pdf,.csv,.xlsx"` e
    `onchange="processarArquivos(this.files, true)"`. Abaixo, mover pra cá o card
    existente `#cardLancamentosAberta`.
  - Card "Fatura fechada (definitiva)": texto + `#fileInputFechada` com
    `accept=".pdf,.csv,.xlsx"` e `onchange="processarArquivos(this.files, false)"`.
    Abaixo, o card "Faturas fechadas" (`#faturasList`).
  - Manter o `#fileInput` antigo? Não — remover e apontar o botão "↻ Atualizar
    fatura" do Meu Mês para `irParaAba('upload')` (já faz isso). Buscas por
    `getElementById('fileInput')` devem virar `fileInputParcial`.

- [ ] **Step 2:** Em `processarArquivos`, mudar assinatura para
  `async function processarArquivos(files, forcarAberta)` e trocar
  `aberta: !!result.aberta` por
  `aberta: (forcarAberta !== undefined ? forcarAberta : !!result.aberta)`.

- [ ] **Step 3:** Aviso não-bloqueante de incoerência: se `forcarAberta===false`
  e `result.aberta===true` (Excel aberto no bloco fechada), `console.warn` + um
  `alert` curto "Este arquivo parece ser uma fatura ABERTA, mas você subiu no
  bloco de fatura fechada. Vou tratar como fechada." (respeita o lugar).

- [ ] **Step 4: Verificar** — `node -e new Function` (sintaxe). No navegador
  (harness local): subir um Excel pelo bloco parcial → `FATURA_TEMP.aberta===true`;
  pelo bloco fechada → `false`. Botão "Atualizar fatura" leva à aba.

- [ ] **Step 5: Commit** — `git add index.html && git commit`.

---

### Task 2: `parseLinhasOCR` — reconstrução pura das linhas (testável em node)

**Files:**
- Modify: `index.html` — nova função pura `parseLinhasOCR`.
- Test: `scratchpad/test-ocr.js` (node, sem dependências).

**Interfaces:**
- Produz: `parseLinhasOCR(palavras, anoAtual, hojeDDMM) -> { transacoes, venc }`
  onde `palavras` é `[{ text, x0, y0, x1, y1, conf }]` (forma do Tesseract.js),
  `anoAtual` número, `hojeDDMM` `{d,m}` de hoje.
  `transacoes[i] = { data:'DD/MM/AAAA', estabelecimento, valor:Number,
  confianca:0-100, origem:'imagem' }`.

**Comportamento (formato observado no print real do app):**
- Agrupar palavras em **linhas** por sobreposição vertical (centro-Y próximo).
- **Cabeçalho de data**: linha cujo texto casa `/(\d{1,2})\s+de\s+(janeiro|
  fevereiro|março|abril|maio|junho|julho|agosto|setembro|outubro|novembro|
  dezembro)/i` → define `dataCorrente = DD/MM/anoAtual`; se `MM/DD` for depois de
  `hojeDDMM`, usar `anoAtual-1`.
- **Lançamento**: linha que contém um token de valor `/R\$\s*([\d.]+,\d{2})/`.
  Nome = concatenação das palavras à esquerda do token de valor (mesma linha).
  Valor = parse do token (ponto=milhar, vírgula=decimal). `confianca` = menor
  `conf` entre o token de valor e as palavras do nome.
- **Ignorar**: linhas de subtipo (`apple pay`, `cartão adicional físico`), linha
  `US$ ...` (internacional — o R$ já veio na mesma linha ou na de cima), número
  de página (`/^\d+\s+de\s+\d+$/`), e linhas sem valor e sem data.
- `venc`: não vem na imagem → retornar `null` (o chamador decide o default).
- Sem lançamentos → `{ transacoes: [], venc: null }`.

- [ ] **Step 1: Harness** — `scratchpad/test-ocr.js` com um array de `palavras`
  simulando o print real (dados vistos: cabeçalho "31 de julho"; linhas
  "mp *gaby / R$ 103,00", "vercel inc. / R$ 108,60 / US$ 20,00",
  "mercadolivre*mercadol / R$ 96,00", "shein *shein.com / R$ 220,59"; "1 de 2").
  Copiar a lógica de `parseLinhasOCR` e imprimir as transações.

- [ ] **Step 2: Implementar `parseLinhasOCR`** em `index.html` conforme acima.

- [ ] **Step 3: Rodar o harness** — `node scratchpad/test-ocr.js`. Conferir:
  `mp *gaby` 103,00; `vercel inc.` 108,60 (US$ ignorado); `mercadolivre*mercadol`
  96,00; `shein *shein.com` 220,59; datas = 31/07 no bloco "hoje"; "1 de 2" fora.

- [ ] **Step 4: Verificar sintaxe** — `node -e new Function`.

- [ ] **Step 5: Commit.**

---

### Task 3: Tesseract.js sob demanda + `parseItauImagem` (cola de navegador)

**Files:**
- Modify: `index.html` — loader do Tesseract; `parseItauImagem`; roteamento em
  `processarArquivos` para imagens.

**Interfaces:**
- Consome: `parseLinhasOCR` (Task 2), `classificarTransacoes` (existente),
  `STATE.cartaoAtivo`, `STATE.faturaAberta?.venc`.
- Produz: `parseItauImagem(file) -> Promise<{ cartoes, venc, postagem:null,
  subtotaisDeclarados:{}, totalDeclarado, aberta:true }>` — mesma forma de
  `parseItauExcel`. Um cartão só: `cartoes[cartaoAtivo] = { nome, final, transacoes }`,
  cada transação com `cartao: cartaoAtivo`, `categoria` (via dicionário), e os
  campos `confianca`/`origem:'imagem'`.
- Produz: `carregarTesseract() -> Promise<Tesseract>` (idempotente; injeta o
  script do CDN `https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js`
  uma vez).

- [ ] **Step 1: Loader** — `carregarTesseract()`: se `window.Tesseract` existe,
  resolve; senão injeta o `<script>` do CDN e resolve no `onload` (rejeita no
  `onerror` com mensagem "não consegui carregar o leitor de imagem; verifique a
  conexão").

- [ ] **Step 2: `parseItauImagem(file)`** — `await carregarTesseract()`; rodar
  `Tesseract.recognize(file, 'por', { logger })` com o `logger` atualizando o
  `#uploadStatus` ("Lendo a imagem… NN%"). Da resposta, montar `palavras` a
  partir de `data.words` (`text`, `bbox.x0/y0/x1/y1`, `confidence`); chamar
  `parseLinhasOCR(palavras, anoAtual, hojeDDMM)`. Se `transacoes.length===0`,
  lançar erro "Não consegui ler a imagem; tente um print mais nítido ou use o
  Excel." Marcar cada transação com `cartao = STATE.cartaoAtivo`. `venc` =
  `STATE.faturaAberta?.venc` || padrão (dia 17 do mês seguinte, formato
  DD/MM/AAAA). Retornar a estrutura no formato de `parseItauExcel` com
  `aberta:true` e `cartoes[cartaoAtivo]`.

- [ ] **Step 3: Rotear no upload** — em `processarArquivos`, detectar
  `isImg = /\.(png|jpe?g)$/i.test(file.name)`; incluir no gate de formato; quando
  `isImg`, `result = await parseItauImagem(file)` e seguir para
  `classificarTransacoes()` direto (cartão único, pular seleção — como Nubank),
  com `FATURA_TEMP.cartoesSelecionados = [STATE.cartaoAtivo]` e
  `FATURA_TEMP.origem = 'imagem'`.

- [ ] **Step 4: Verificar** — sintaxe; no navegador, subir o print real de
  exemplo pelo bloco parcial → os lançamentos aparecem na conferência com valores
  próximos do print; falha graciosa numa imagem sem texto.

- [ ] **Step 5: Commit.**

---

### Task 4: Múltiplos prints somados numa parcial só

**Files:**
- Modify: `index.html` — `processarArquivos` (acúmulo de imagens).

**Interfaces:**
- Consome: `parseItauImagem`. Produz: um único `FATURA_TEMP` com as transações de
  todas as imagens selecionadas juntas (as 2 páginas).

- [ ] **Step 1:** Quando o lote de `files` contém imagens, processá-las **todas**
  antes de abrir a conferência: concatenar as `transacoes` de cada
  `parseItauImagem` num só conjunto; então montar um único `FATURA_TEMP`
  (`aberta:true`, `origem:'imagem'`, `cartoesSelecionados:[cartaoAtivo]`) e chamar
  `classificarTransacoes()` uma vez. (Não usar o loop 1-arquivo-por-vez das
  faturas fechadas para imagens.)

- [ ] **Step 2: Verificar** — navegador: subir 2 prints juntos → conferência com
  a soma dos dois (sem descartar o primeiro).

- [ ] **Step 3: Commit.**

---

### Task 5: Conferência — valor/descrição editáveis + destaque de baixa confiança

**Files:**
- Modify: `index.html` — `renderTransacoesEditavel` (~2548); novos handlers
  `alterarEstabelecimentoTemp`, `alterarValorTemp`.

**Interfaces:**
- Consome: `FATURA_TEMP.origem`, `t.confianca`, `parseValorBR` (existente).
- Produz: `alterarEstabelecimentoTemp(idx, val)` e `alterarValorTemp(idx, val)`
  que gravam em `FATURA_TEMP.transacoes[idx]` e re-renderizam a conferência.

- [ ] **Step 1:** Em `renderTransacoesEditavel`, quando
  `FATURA_TEMP.origem === 'imagem'`: a célula de **estabelecimento** vira
  `<input type="text" value=... onchange="alterarEstabelecimentoTemp(${idx}, this.value)">`
  e a de **valor** vira `<input type="text" inputmode="decimal" value="${String(t.valor).replace('.',',')}"
  onchange="alterarValorTemp(${idx}, this.value)">`. Para outras origens, manter
  texto (como hoje).

- [ ] **Step 2:** Destaque de baixa confiança: se `origem==='imagem'` e
  `(t.confianca||100) < 75`, aplicar `bg = 'background:var(--amber-soft)'` e um
  selo `⚠ confira` na coluna Origem. (Não sobrescrever o verde de conciliado.)

- [ ] **Step 3:** Handlers: `alterarValorTemp(idx,val)` usa `parseValorBR(val)`
  (guardar 0 se inválido) e marca `confianca=100` (revisado à mão);
  `alterarEstabelecimentoTemp(idx,val)` grava `.trim()` e `confianca=100`. Ambos
  chamam `mostrarRevisao()` de novo (atualiza total/realce).

- [ ] **Step 4:** Ordenação por "a conferir": incluir a coluna `origem` no
  `setRevSort` já existente ordenando baixa-confiança primeiro quando
  `origem==='imagem'` (chave = `t.confianca ?? 100`).

- [ ] **Step 5: Verificar** — navegador: com uma parcial de imagem, editar um
  valor (com vírgula) → total atualiza; linha de baixa confiança fica amarela;
  editar zera o alerta.

- [ ] **Step 6: Commit.**

---

### Task 6: Conciliação visível ao salvar a fatura fechada

**Files:**
- Modify: `index.html` — `conciliarPendentes` (~2386, retornar detalhes) e o
  salvamento da fatura fechada (`salvarFatura`, ramo fechada) para exibir o resumo.

**Interfaces:**
- Consome: resultado de `conciliarPendentes`. Produz: um resumo mostrado no
  `#uploadStatus` após salvar a fechada: nº de anotados que casaram e nº que
  sobraram para o próximo ciclo.

- [ ] **Step 1:** Fazer `conciliarPendentes` devolver `{ casados, naoCasados }`
  (contagens) além de marcar os pendentes. Hoje retorna um **número** atribuído a
  `FATURA_TEMP.numConciliados` em `classificarTransacoes` — então **atualizar o
  chamador**: `const conc = conciliarPendentes(); FATURA_TEMP.numConciliados =
  FATURA_TEMP.aberta ? 0 : conc.casados; FATURA_TEMP.numNaoConciliados =
  FATURA_TEMP.aberta ? 0 : conc.naoCasados;`. `naoCasados` = pendentes não
  conciliados do(s) cartão(ões) desta fatura. Não mudar o critério de casamento.
  Conferir todos os usos de `numConciliados` (ex.: banner em `mostrarRevisao`)
  continuam válidos (ainda é número).

- [ ] **Step 2:** No ramo da fatura **fechada** de `salvarFatura`, após
  `renderFaturasList()`, montar a mensagem de sucesso incluindo:
  `✓ N gasto(s) anotado(s) casaram com esta fatura (arquivados). ⚠ M não casaram —
  continuam para o próximo ciclo.` (só a parte ⚠ se M>0).

- [ ] **Step 3: Verificar** — navegador: com 2 pendentes (1 que casa, 1 que não),
  subir a fechada correspondente → mensagem "✓ 1 … ⚠ 1 …"; o que casou vira
  conciliado, o outro permanece pendente.

- [ ] **Step 4: Commit.**

---

### Task 7: Revisão adversarial + deploy

- [ ] **Step 1:** `node -e new Function` (sintaxe) do `index.html` inteiro.
- [ ] **Step 2:** Revisão adversarial (Workflow, lentes: parse OCR/datas/valor;
  roteamento parcial×fechada e flag; múltiplas imagens; edição/realce na
  conferência; resumo da conciliação; sync/estado). Corrigir achados confirmados.
- [ ] **Step 3:** `git commit` das correções e `vercel --prod`.

---

## Self-Review (cobertura do spec)

- Imagem por OCR grátis, cai na conferência: Tasks 2–3. ✓
- Aba Faturas em 2 blocos (lugar decide): Task 1. ✓
- Conferência edita valor/descrição + destaque baixa confiança: Task 5. ✓
- Fechada: descarta parcial (já em `salvarFatura`) + conciliação visível: Task 6. ✓
- Registro manual inalterado: nenhuma task o toca. ✓
- Múltiplos prints somados: Task 4. ✓
- Cartão = cartaoAtivo; venc reaproveitado/padrão: Task 3. ✓
- Nada de backend/chave/custo: Tesseract client-side (Tasks 2–3). ✓

Assinaturas consistentes entre tasks: `parseLinhasOCR`, `parseItauImagem`,
`carregarTesseract`, `processarArquivos(files, forcarAberta)`,
`alterarValorTemp`/`alterarEstabelecimentoTemp`, `conciliarPendentes -> {casados,
naoCasados}`.
