# Fase 1 — Excel + Meu Mês (base) — Plano de Implementação

> Baseado no spec `docs/superpowers/specs/2026-07-27-redesenho-meu-mes-design.md`.
> Executado inline nesta sessão (autor único, com contexto completo do código).

**Goal:** Ler a fatura Itaú em Excel (aberta/fechada) e mostrar o "Meu Mês" —
gasto por categoria do ciclo atual do cartão em uso, com teto total + meta por
categoria.

**Arquitetura:** App single-file (`index.html`, HTML+CSS+JS). Estado em `STATE`
(localStorage + Google Sheets via Apps Script, com merge por `updatedAt` +
lixeira). Sem framework de testes: verificação por (a) checagem de sintaxe
`node -e new Function`, (b) teste do parser em node contra o arquivo real da
Michelle, (c) revisão adversarial via Workflow, (d) conferência no navegador.

**Tech:** JS vanilla; SheetJS (`XLSX`, já carregado por CDN) para ler .xlsx;
Vercel para deploy (`vercel --prod`).

## Global Constraints (do spec)

- **Nada é apagado** — só reorganizado; todas as coleções/telas existentes
  permanecem.
- **A fatura é a fonte da verdade;** lançamentos diários são palpite que se
  ajusta na conciliação.
- Ciclo do cartão, não calendário. Cartão em uso hoje: **5074**.
- Português (pt-BR) em toda a interface, como o resto do app.
- Merge de sincronização: qualquer coleção/campo novo precisa de migração em
  `normalizarEstado` e tratamento no `mesclarEstados`.

---

## Escopo da Fase 1

Fontes do Meu Mês nesta fase: **lançamentos diários (pendentes) + fatura aberta
importada**, conciliados entre si. **Assinaturas provisionadas ficam para a
Fase 2.**

---

### Task 1: Parser de Excel da fatura Itaú (aberta e fechada)

**Files:**
- Modify: `index.html` (novo `parseItauExcel`; roteamento em `processarArquivos`;
  `accept` do input de arquivo).
- Test: `scratchpad/test-parse-xlsx.js` (node + pacote `xlsx`).

**Interfaces:**
- Produz: `parseItauExcel(file) -> Promise<{ cartoes, venc, postagem,
  subtotaisDeclarados, totalDeclarado, aberta }>` — mesma forma que
  `parseItau`, com `aberta: boolean`. `cartoes[final] = { nome, final,
  transacoes: [{ data:'DD/MM', estabelecimento, valor, parcela, catItau:null,
  cidade:null, secao }] }`.

**Comportamento do parser (formato observado no arquivo real):**
- Cabeçalho tem "Fatura Aberta - Agosto/2026" → `aberta = true` se o texto contém
  "aberta" (case-insensitive); senão `false`.
- A planilha tem uma tabela de Lançamentos por cartão. Colunas relevantes:
  Data, Lançamento, Parcelamento, Valor, Titularidade, Nº do cartão. Cada bloco
  de cartão começa por uma linha "… - final NNNN" (ex.: "…Mastercard - final
  6233") e os lançamentos daquele cartão vêm abaixo até o próximo bloco/subtotal.
- Valor pode vir com "R$", vírgula decimal e milhar; normalizar como no
  `parseItau`. Estornos (negativos) mantêm sinal.
- Internacionais: linha com "IOF Compra Internacional" e linhas "USD x,xx
  (cotação R$ y,yy)" — o valor em reais fica na coluna Valor; usar esse. Não
  criar transação para a linha da cotação isolada.
- Parcelamento "Parcela N de M" → `parcela = 'NN/MM'` (mesmo formato do PDF).
- `venc`: derivar do cabeçalho "Fatura … - <Mês>/<Ano>" → vencimento
  'DD/MM/AAAA' (dia do vencimento informado no cabeçalho, ou dia 1 se ausente).

- [ ] **Step 1: Harness node** — instalar `xlsx` no scratchpad e escrever
  `test-parse-xlsx.js` que lê o arquivo real
  (`.../transfers/2026-30/Cartão parcial vencimento 17-08.xlsx`), roda a lógica
  de parse (portada) e imprime: nº de cartões, nº de transações por cartão,
  soma por cartão, e 5 transações de exemplo. Rodar e conferir contra o
  subtotal declarado na planilha.

- [ ] **Step 2: Implementar `parseItauExcel`** em `index.html` (usa
  `XLSX.read(await file.arrayBuffer())` → `XLSX.utils.sheet_to_json(sheet,
  {header:1})` para varrer linhas). Portar a lógica validada no harness.

- [ ] **Step 3: Rotear no upload** — em `processarArquivos`, aceitar `.xlsx`
  (`isXlsx = /\.xlsx$/i.test`), chamar `parseItauExcel`, e propagar `aberta`
  para `FATURA_TEMP.aberta`. Atualizar `accept` do `#fileInput` para incluir
  `.xlsx` e o texto do dropzone.

- [ ] **Step 4: Verificar** — `node -e "new Function(script)"` (sintaxe OK);
  deploy; no navegador, subir o arquivo real e conferir que os lançamentos
  aparecem na revisão com valores certos.

- [ ] **Step 5: Commit** — `git add index.html && git commit`.

---

### Task 2: Cartão em uso + teto mensal (estado + config)

**Files:**
- Modify: `index.html` (`stateDefault`, `normalizarEstado`, `mesclarEstados`,
  seção Configurações).

**Interfaces:**
- Produz: `STATE.cartaoAtivo` (string, final do cartão, default `'5074'` ou o
  primeiro de `cartoesMichelle`) e `STATE.tetoMensal` (número, default `0` = sem
  teto). Getters simples via `STATE.cartaoAtivo` / `STATE.tetoMensal`.

- [ ] **Step 1:** Adicionar `cartaoAtivo` e `tetoMensal` em `stateDefault()` e
  defaults em `normalizarEstado()`. No `mesclarEstados`, tratá-los como campos
  escalares do lado base (vencendo pelo `STATE.updatedAt` global, como os demais
  escalares — nenhuma lógica extra necessária).

- [ ] **Step 2:** Em Configurações, adicionar: seletor "Cartão em uso"
  (dropdown de `cartoesMichelle` + opção Nubank) e campo "Teto do mês (R$)".
  `onchange` grava e `salvarState()`.

- [ ] **Step 3: Verificar** — sintaxe; no navegador, mudar cartão/teto, recarregar
  e confirmar persistência.

- [ ] **Step 4: Commit.**

---

### Task 3: Armazenar fatura aberta + conciliar com os lançamentos

**Files:**
- Modify: `index.html` (`salvarFatura` ou novo `salvarFaturaAberta`;
  `stateDefault`/`normalizarEstado` para `faturaAberta`).

**Interfaces:**
- Produz: `STATE.faturaAberta = { cartao, transacoes:[...], importadaEm } | null`
  — snapshot da fatura aberta do cartão ativo; **substituído** a cada
  reimportação daquele cartão. Usa a mesma forma de transação das faturas.

- [ ] **Step 1:** `faturaAberta` em `stateDefault`/`normalizarEstado`
  (default `null`). No merge, vence o `updatedAt` global (escalar).

- [ ] **Step 2:** No salvamento da revisão, se `FATURA_TEMP.aberta === true`:
  gravar em `STATE.faturaAberta` (não em `faturas[]`), rodar `conciliarPendentes`
  contra essas transações (marca `pendentes` que casam como conciliados,
  `faturaId = 'aberta'`), e **NÃO** fechar ciclo. Reimportar substitui o snapshot
  e re-concilia (reabrir pendentes cujo `faturaId==='aberta'` antes de reconciliar,
  para não travar itens de um snapshot antigo).

- [ ] **Step 3:** Fatura **fechada** continua no fluxo atual (`faturas.push`),
  sem mudança. Quando uma fatura fechada do cartão ativo é salva, limpar
  `STATE.faturaAberta` desse cartão (o ciclo fechou).

- [ ] **Step 4: Verificar** — sintaxe; navegador: subir aberta → pendentes
  casados somem do "anotados"; reimportar → sem duplicar; subir fechada → aberta
  some.

- [ ] **Step 5: Commit.**

---

### Task 4: Tela "Meu Mês" (cálculo + render) na Início

**Files:**
- Modify: `index.html` (HTML da aba Início; `computeMeuMes`, `renderMeuMes`;
  chamada em `renderInicio`; reaproveitar CSS de `.budget-*`).

**Interfaces:**
- Consome: `STATE.faturaAberta`, `STATE.pendentes`, `STATE.orcamentos`,
  `STATE.tetoMensal`, `STATE.cartaoAtivo`.
- Produz: `computeMeuMes() -> { total, teto, porCategoria: [{cat, gasto,
  limite}], itens }`. Card "Meu Mês" no topo da Início.

**Cálculo (ciclo atual do cartão ativo):**
- `gastos = (faturaAberta?.cartao === cartaoAtivo ? faturaAberta.transacoes :
  []) + pendentes.filter(p => !p.conciliado && (p.cartao===cartaoAtivo ||
  !p.cartao) )`. Cada item com `{categoria||'??', valor}`.
- Agrupar por categoria (somando valor). `total = soma`.
- `porCategoria`: para cada categoria com gasto ou com orçamento definido,
  `{cat, gasto, limite: STATE.orcamentos[cat] || 0}`, ordenado por gasto desc.

- [ ] **Step 1:** `computeMeuMes()` conforme acima (função pura sobre STATE).

- [ ] **Step 2:** `renderMeuMes()` — card no topo da Início: barra do **teto
  total** (gasto/teto, cor por faixa) + uma barra por categoria
  (gasto/limite; sem limite mostra só o valor). Botão "Atualizar" →
  `document.getElementById('fileInput').click()`.

- [ ] **Step 3:** Adicionar o container `#meuMes` no HTML da Início (acima do
  registro rápido) e chamar `renderMeuMes()` dentro de `renderInicio()`.

- [ ] **Step 4: Verificar** — sintaxe; navegador: com pendentes + fatura aberta,
  as barras batem com a soma; teto e metas aparecem certos; estourar categoria
  fica vermelho.

- [ ] **Step 5: Revisão adversarial** (Workflow, 1 lente) do cálculo do Meu Mês
  e da conciliação da fatura aberta (dupla contagem, cartão errado, reimport).
  Corrigir achados.

- [ ] **Step 6: Commit + deploy final da Fase 1.**

---

## Self-Review (cobertura do spec, Fase 1)

- Parser de Excel (aberta/fechada): Task 1. ✓
- Meu Mês (teto total + meta por categoria, ciclo do cartão): Tasks 2–4. ✓
- Conciliação lançamentos × fatura (aberta) + carryover: Task 3 (aberta) e o
  fluxo de fechada já existente cobre o carryover (pendentes não casados
  permanecem). ✓
- Cartão em uso: Task 2. ✓
- Assinaturas provisionadas: **Fase 2** (fora desta). ✓
- Reorganização de telas / "Mais" / link do Rogério: **Fases 3–4**. ✓

Sem placeholders; assinaturas de função consistentes entre tasks
(`parseItauExcel`, `computeMeuMes`, `renderMeuMes`, `STATE.faturaAberta`,
`STATE.cartaoAtivo`, `STATE.tetoMensal`).
