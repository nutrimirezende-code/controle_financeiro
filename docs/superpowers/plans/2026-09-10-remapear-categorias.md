# Remapear categorias (De→Para) — Plano de implementação

> **For agentic workers:** implementar tarefa a tarefa; passos com checkbox.

**Goal:** Ferramenta em Configurações que aplica um De→Para em todo o histórico
(com backup antes), trocando a classificação antiga pela nova.

**Architecture:** App single-file (`index.html`). Núcleo é a função **pura**
`aplicarDePara(state, mapa, novasCategorias)` (testável em node). A UI só coleta
o mapa, mostra resumo, faz backup e chama a função. Constantes de código
atualizadas para a nova classificação.

**Tech Stack:** JS vanilla. Sem backend novo.

## Global Constraints

- Nada é apagado sem backup: a ferramenta baixa o backup JSON **antes** de aplicar.
- Função `aplicarDePara` é **pura** (não muta o `state` recebido; retorna cópia).
- Categoria desconhecida (fora do mapa) é **mantida** (nunca vira branco por engano).
- Destino "branco" = `'??'` em transações; remove entradas em orçamento/dicionário.
- pt-BR. Verificação: node harness da função pura + `node -e new Function` +
  navegador + revisão adversarial.

---

### Task 1: `aplicarDePara` (função pura) + harness node

**Files:**
- Modify: `index.html` — nova função `aplicarDePara`.
- Test: `scratchpad/test-remap.js` (node).

**Interfaces:**
- Produz: `aplicarDePara(state, mapa, novasCategorias) -> { state: novoState,
  resumo: { nLancamentos, nBranco, nOrcamentos, nDictRemovidos } }`.
  `mapa` = `{ 'Categoria Antiga': 'CATEGORIA NOVA' | '' }` (`''`/`'??'`/`null` =
  branco). `novasCategorias` = array das 16.
- Helper interno `mapCat(c)`: se `c` está no `mapa`, retorna o destino
  (branco→`'??'`); senão retorna `c` inalterado.

- [ ] **Step 1: Harness** — `scratchpad/test-remap.js` monta um mini-STATE:
  categorias antigas; 1 fatura com transações (`Supermercado`, `Feira`, `Casa`,
  `Business Michelle`, `XPTO desconhecida`); `pendentes` (`Meninos`);
  `assinaturas` (`APPs Streaming`); `orcamentos` `{Supermercado:1000, Feira:500,
  Casa:300, Carro:200}`; `dicionario` `{"SUP":{cat:"Supermercado",n:5,
  alt:[["Feira",2]]}, "SUELY":{cat:"Casa",n:3}, "FACE":{cat:"Business Michelle",
  n:4,alt:[["Michelle",1]]}}`. Copiar a lógica de `aplicarDePara` e imprimir o
  resultado + `resumo`.

- [ ] **Step 2: Implementar `aplicarDePara`** em `index.html`:
  - `const novo = JSON.parse(JSON.stringify(state))`.
  - `mapCat(c)`: `mapa.hasOwnProperty(c) ? (mapa[c]||'??') : c` (branco→`'??'`).
  - Transações: `for f in novo.faturas: for t in f.transacoes: t.categoria =
    mapCat(t.categoria)`. Idem `novo.faturaAberta?.transacoes`, `novo.pendentes`,
    `novo.assinaturas` (campo `.categoria`).
  - orcamentos: novo objeto; para cada `[k,v]`: `nk = mapCat(k)`; se `nk==='??'`
    pula; senão `out[nk] = (out[nk]||0) + Number(v||0)` (soma quando junta).
  - dicionario: novo objeto; para cada `[chave, r]`: `nc = mapCat(r.cat)`; se
    `nc==='??'` pula (remove); senão `r.cat = nc`; remapeia `r.alt`: mapeia cada
    `[c,n]`, descarta `'??'`, soma `n` por cat, remove alt igual a `nc`.
  - `novo.categorias = [...novasCategorias]`.
  - `resumo`: contar transações remapeadas (categoria mudou), quantas foram pra
    `'??'`, orçamentos resultantes, entradas de dicionário removidas.
  - `return { state: novo, resumo }`.

- [ ] **Step 3: Rodar harness** — `node scratchpad/test-remap.js`. Conferir:
  Supermercado/Feira→COMIDA (orçamento COMIDA = 1500), Casa→`'??'` (orçamento Casa
  sumiu), Carro→DIVERSOS (200), Business Michelle→MICHELLE, `XPTO`→`XPTO`
  (inalterada), dicionário `SUELY` removido, `SUP.cat`=COMIDA com `alt` Feira→COMIDA
  fundido/removido, `FACE.cat`=MICHELLE com alt Michelle→MICHELLE removido (=cat).

- [ ] **Step 4:** `node -e new Function` (sintaxe).
- [ ] **Step 5: Commit.**

---

### Task 2: Atualizar constantes de código para a nova classificação

**Files:**
- Modify: `index.html` — `CATEGORIAS_DEFAULT`, `ITAU_PARA_MICHELLE`, `DICT_INICIAL`.

- [ ] **Step 1:** `CATEGORIAS_DEFAULT` = as 16 novas (mesma ordem do spec).

- [ ] **Step 2:** `ITAU_PARA_MICHELLE` novo:
  `{"ALIMENTAÇÃO":"RESTAURANTES","VESTUÁRIO":"MICHELLE","SAÚDE":"SAÚDE",
  "DIVERSOS":"DIVERSOS","VEÍCULOS":"DIVERSOS","TURISMO E ENTRETENIM":"LAZER",
  "MORADIA":"MORADIA","EDUCAÇÃO":"EDUCAÇÃO","HOBBY":"LAZER"}`.

- [ ] **Step 3:** `DICT_INICIAL`: remapear cada `.cat`/`.alt` pelo mesmo De→Para,
  removendo entradas que caem em branco (Casa/Serviços/Meninos). Fazer via script
  node no scratchpad que lê o `DICT_INICIAL` atual, aplica o mapa e imprime o novo
  literal; colar o resultado. (Consistência de fresh install; a Michelle não é
  afetada por isto — o dado dela é migrado pela ferramenta.)

- [ ] **Step 4:** `node -e new Function` (sintaxe).
- [ ] **Step 5: Commit.**

---

### Task 3: Tela "Trocar classificação" em Configurações + aplicar

**Files:**
- Modify: `index.html` — HTML de um card novo em `panel-config`; funções
  `renderRemapConfig`, `setRemapDestino`, `aplicarRemap`; hook em `renderConfig`.

**Interfaces:**
- Consome: `aplicarDePara` (Task 1), `exportarBackup` (existente), `salvarState`,
  `NOVAS_CATEGORIAS` (const das 16) e `DE_PARA_PADRAO` (const com o mapa acordado).

- [ ] **Step 1:** Consts no script: `NOVAS_CATEGORIAS = [...16...]` e
  `DE_PARA_PADRAO = { 'Supermercado':'COMIDA','Feira':'COMIDA',
  'APPs Streaming':'ASSINATURAS','Internet':'MORADIA','Restaurante':'RESTAURANTES',
  'Saúde':'SAÚDE','Lazer':'LAZER','Margô':'MARGÔ','Michelle':'MICHELLE',
  'Business Michelle':'MICHELLE','Rogério':'ROGÉRIO','Presente':'DIVERSOS',
  'Carro':'DIVERSOS','Outros':'DIVERSOS','Casa':'','Serviços':'','Meninos':'' }`.

- [ ] **Step 2:** Card em `panel-config` (perto de Categorias): título "Trocar
  classificação", aviso de que mexe em todo o histórico, `<div id="remapConfig">`,
  e botões "Fazer backup e aplicar" (`aplicarRemap()`) e um resumo `<div id="remapResumo">`.

- [ ] **Step 3:** `renderRemapConfig()`: coletar as categorias antigas presentes
  nos dados (`STATE.categorias` ∪ todas as categorias em uso). Para cada, uma
  linha com o nome e um `<select onchange="setRemapDestino('cat', this.value)">`
  com opções: `(deixar em branco)` + as 16 novas — valor inicial de
  `remapMapa[cat]` (default `DE_PARA_PADRAO[cat]` se houver, senão `''`). Guardar
  em `let remapMapa = {}`. Mostrar prévia chamando `aplicarDePara` (sem salvar) e
  exibindo o `resumo` em `#remapResumo`.

- [ ] **Step 4:** `setRemapDestino(cat, val)`: `remapMapa[cat]=val`; re-render +
  atualizar resumo.

- [ ] **Step 5:** `aplicarRemap()`: `confirm` duplo ("Isto vai mudar a categoria
  de TODO o histórico. Já vou baixar um backup antes. Continuar?"); chamar
  `exportarBackup()`; `const { state } = aplicarDePara(STATE, remapMapa,
  NOVAS_CATEGORIAS); STATE = normalizarEstado(state); salvarState();
  renderConfig(); renderAbaAtiva();` mensagem de sucesso + lembrete de reclassificar
  os `'??'` no Histórico.

- [ ] **Step 6:** `renderConfig` chama `renderRemapConfig()`.

- [ ] **Step 7: Verificar** — navegador: injetar mini-STATE antigo, abrir Config,
  ver a tabela pré-preenchida, aplicar → categorias novas, orçamentos somados,
  Casa/Serviços/Meninos em branco, backup baixado; Histórico filtra "sem categoria".

- [ ] **Step 8: Commit.**

---

### Task 4: Revisão adversarial + deploy

- [ ] **Step 1:** `node -e new Function` (sintaxe geral).
- [ ] **Step 2:** Revisão adversarial (Workflow): pureza/imutabilidade de
  `aplicarDePara`, soma de orçamentos, dicionário (.alt dedupe/remoção), categoria
  desconhecida preservada, backup antes de aplicar, sync após aplicar, constantes
  de código coerentes.
- [ ] **Step 3:** Corrigir achados; `git commit`; `vercel --prod`.

---

## Self-Review (cobertura do spec)

- Transforma faturas/parcial/pendentes/assinaturas/orçamentos/dicionário/lista:
  Task 1. ✓
- Branco = `'??'`, desconhecida preservada, orçamento somado, dicionário
  removido/dedupe: Task 1. ✓
- Constantes de código (fresh install / importações futuras): Task 2. ✓
- UI editável pré-preenchida + resumo + backup + confirmação dupla: Task 3. ✓
- Backup antes de aplicar: Task 3 Step 5. ✓
- extras/avulsas/contasFixas não tocados: nenhuma task os inclui. ✓

Assinaturas consistentes: `aplicarDePara(state, mapa, novasCategorias) ->
{state, resumo}`, `NOVAS_CATEGORIAS`, `DE_PARA_PADRAO`, `renderRemapConfig`,
`setRemapDestino`, `aplicarRemap`.
