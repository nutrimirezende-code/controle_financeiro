# Redesenho: sistema mais simples, centrado no "Meu Mês"

**Data:** 2026-07-27
**Autora do produto:** Michelle Rezende

## Problema

O sistema cresceu com dezenas de funções (6 abas, análises, bot, conciliação,
parcelas, assinaturas, destaques). A Michelle relatou que ficou **confuso e
difícil de usar** — marcou as quatro dores: telas demais, Histórico com coisa
demais, não conseguir ver o gasto do mês, e o acerto complicado.

Ao mesmo tempo, surgiu uma necessidade central e concreta: **acompanhar, no
decorrer do mês, quanto já se gastou em cada classificação, para não estourar o
limite** — usando a fatura ABERTA (parcial) que o Itaú permite exportar em Excel
a qualquer momento, além dos lançamentos diários.

## Objetivo

Reduzir o uso diário a poucas telas claras, com o **"Meu Mês"** (gasto por
categoria, em tempo real) como coração do sistema. **Nada é apagado** — as
funções avançadas saem do caminho para um menu "Mais".

## Princípios

- **Nada se perde:** todas as funções e dados existentes continuam; a
  simplificação é de organização/visibilidade, não de exclusão.
- **O dia a dia cabe em 2–3 telas.** O resto fica acessível, mas escondido.
- **A fatura é a fonte da verdade;** os lançamentos diários são o palpite em
  tempo real que se ajusta quando a fatura chega (conciliação).

## Nova estrutura de telas

De 6 abas para 4 visíveis:

1. **Início** (reorganizada) — mantém o que já existe + o novo "Meu Mês":
   - **Registrar gasto agora** (valor + descrição + escolha: cartão ou acerto) — mantido
   - **Meu Mês** (novo, destaque no topo) — gasto por categoria com limite
   - **Próximo acerto** (card) — mantido
   - **Pendências** (card) — mantido
   - **Gastos de cartão anotados** (lista) — mantido
   - Botão **Atualizar/Importar fatura** (abre o upload de fatura aberta/fechada)

2. **Fixos / Assinaturas** (nova) — cobranças automáticas do cartão (Netflix,
   Apple, Spotify, Vivo…) que a Michelle não lança à mão, provisionadas no mês.

3. **Acerto** — **inalterada** (a Michelle confirmou que está boa).

4. **Mais** (menu escondido) — Upload de faturas antigas, Histórico/análises
   (parcelas, assinaturas detectadas, comparativo, pivot, gráficos), Dicionário,
   Backup, Configurações. O bot do Telegram fica desativado (código dorminte).

## Meu Mês (detalhado)

O coração do sistema. Mostra o gasto do **ciclo atual do cartão em uso** por
classificação, com dois níveis de limite.

### Cartão em uso

- A Michelle usa um cartão principal por período (hoje o 5074). O **cartão em
  uso** é escolhido nas configurações (ou assume o último importado); o Meu Mês
  acompanha o ciclo desse cartão. Trocar o cartão em uso é uma ação simples.

### Período (ciclo do cartão)

- O "mês" segue o **ciclo do cartão**, não o calendário (o cartão 5074 fecha
  ~dia 10, sem data cravada).
- O ciclo atual = tudo que **ainda não entrou numa fatura fechada** (lançamentos
  anotados não conciliados + assinaturas provisionadas do ciclo).
- Ao subir a **fatura fechada**, o ciclo "fecha": os lançamentos que casam com a
  fatura são arquivados nela; os que **não casam** (caíram para a próxima fatura)
  ficam pendentes e **entram sozinhos no próximo ciclo**. Nada se perde nem conta
  em dobro.

### Limites (os dois)

- **Teto total do mês:** um valor único (ex.: R$6.000) com uma barra no topo.
- **Meta por categoria:** um limite por classificação (ex.: Supermercado
  R$1.500) — reaproveita o campo de orçamento que já existe. Cada categoria tem
  sua barra (verde = ok, amarelo = perto, vermelho = estourou).

### Fontes de dados somadas no Meu Mês

1. **Lançamentos diários** (pendentes não conciliados) — a descrição é o texto
   que a Michelle digita, já que não sabe o nome que aparecerá na fatura.
2. **Assinaturas provisionadas** — uma instância por assinatura ativa por ciclo,
   contando como se já tivesse sido lançada (ver "Fixos/Assinaturas").
3. **Conciliação da fatura** — quando a fatura (aberta em Excel ou fechada em
   PDF) é importada, os valores reais substituem os palpites nos itens que casam.

## Início reorganizada

Mesmos blocos de hoje, com hierarquia mais clara:

1. **Registrar gasto agora** (topo, ação principal): valor + descrição + dois
   botões (💳 cartão / 💰 acerto Rogério). Igual ao atual.
2. **Meu Mês** (destaque, logo abaixo): barra do teto total + barras por
   categoria; botão "Atualizar" que abre a importação de fatura.
3. **Status compacto:** próximo acerto (com dias restantes) e pendências.
4. **Gastos anotados** (lista) — os lançamentos do ciclo, editáveis/excluíveis.

## Fixos / Assinaturas

Tela nova para cadastrar cobranças automáticas do cartão que a Michelle não
lança manualmente.

- Cada assinatura: **descrição, valor, categoria, cartão, dia aproximado,
  ativo/inativo**.
- **Pré-preenchida** com as recorrências já detectadas no histórico (a função
  `detectarAssinaturas` já existe): Netflix, Apple, Spotify, Vivo, Canva etc. A
  Michelle só confere/edita/desativa.
- **Provisionamento:** cada assinatura ativa aparece no "Meu Mês" do ciclo atual
  como um lançamento provisionado (com selo "provisionado"), somando na sua
  categoria. Quando a cobrança real aparece na fatura importada, ela **concilia**
  com a provisão (não duplica) e a provisão vira "confirmada". No ciclo seguinte,
  provisiona de novo.
- Coleção nova no estado: `assinaturas[]` (distinta de `contasFixas[]`, que são
  contas do acerto com o Rogério).

## Importação de fatura (PDF, CSV e — novo — Excel)

- **Novo parser de Excel** (`.xlsx`) da fatura Itaú, incluindo a **fatura aberta
  (parcial)**. Formato observado: tabela com Data, Lançamento, Parcelamento,
  Valor, por seção de cartão, com lançamentos internacionais em USD +
  cotação. Produz a mesma estrutura de transações do parser de PDF.
- **Aberta vs fechada:**
  - **Aberta** (Excel parcial): é um "refresh" — concilia e atualiza os valores
    reais no Meu Mês, mas **não fecha o ciclo**. Reimportar substitui o snapshot
    aberto anterior daquele cartão (não acumula).
  - **Fechada** (PDF ou Excel fechado): arquiva em `faturas[]`, concilia, os não
    casados ficam pendentes para o próximo ciclo, e o ciclo **reseta**.

## Conciliação e carryover

Reaproveita o motor de conciliação já existente (casamento por cartão + valor
quase exato + data ±3 dias), agora rodando também na fatura aberta e nas
provisões de assinatura.

- Lançamento/assinatura que **casa** com a fatura → confirmado com o valor real.
- Lançamento que **não casa** ao fechar a fatura → continua pendente e vai para o
  próximo ciclo (o "esperar a próxima fatura" que a Michelle perguntou).
- Sem duplicação: um item nunca é contado na fatura e como pendente ao mesmo
  tempo.

## Acerto (inalterado)

A aba Acerto e sua lógica (contas fixas, extras com acerto de destino, avulsas,
mensagem, histórico) **não mudam**. Continua como está hoje.

## Link do Rogério (Resumo + Detalhar)

A página de compartilhamento (`compartilhar.html`) ganha **dois botões**:

- **Resumo:** tabela de **classificação × meses** (categoria nas linhas, meses
  nas colunas, lado a lado) — visão de para onde o dinheiro foi ao longo do
  tempo. Calculada no próprio navegador a partir dos dados já enviados.
- **Detalhar:** a lista de lançamentos com filtros/ordenação que já existe hoje.

Sem mudança de backend — os dados enviados no link já contêm o necessário
(vencimento, categoria, valor por lançamento).

## Modelo de dados (mudanças)

Acréscimos ao `STATE` (com migração e merge de sincronização, como as demais
coleções):

- `assinaturas[]` — recorrências do cartão para provisionar no Meu Mês.
- `tetoMensal` — valor do teto total do mês (número).
- `orcamentos{}` — já existe (metas por categoria); passa a alimentar o Meu Mês.
- `pendentes[]` — já existe; ganha, se necessário, marcação de "provisionado"
  para as instâncias de assinatura, e continua sendo a base do Meu Mês.
- Fatura importada ganha marca `aberta`/`fechada` para o refresh vs. fechamento.

Coleções e telas existentes (faturas, extras, avulsas, contasFixas, acertos,
dicionário, categorias, orçamentos, lixeira) permanecem.

## Fora de escopo (por ora)

- Reativar o bot do Telegram (fica dorminte; pode voltar como serviço no Vercel
  no futuro).
- Apagar funções — nada é removido, só reorganizado.
- Mudar a lógica do Acerto.

## Fases sugeridas de construção

1. **Excel + Meu Mês base:** parser de Excel da fatura aberta; tela Meu Mês
   (teto total + metas por categoria) somando lançamentos e conciliação.
2. **Fixos/Assinaturas + provisionamento:** tela nova, pré-preenchida, e a
   provisão no Meu Mês.
3. **Reorganização das telas:** Início organizada + menu "Mais" recolhendo o
   avançado.
4. **Link do Rogério:** botões Resumo (pivot) + Detalhar.

Cada fase é entregue e testada antes da próxima, com o app sempre funcionando.
