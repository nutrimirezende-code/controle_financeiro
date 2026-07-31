# Fatura parcial por imagem (OCR) + Faturas em dois blocos

**Data:** 2026-07-31
**Autora do produto:** Michelle Rezende

## Problema

Hoje a fatura parcial (aberta) só entra por Excel — exportado no computador. No
dia a dia, a Michelle quer subir um **print do app do Itaú** direto do celular
para acompanhar os gastos por categoria durante o mês. Além disso, o app hoje
"adivinha" se a fatura é parcial ou fechada pelo cabeçalho do Excel — o que não
funciona para imagem (não tem cabeçalho) e é ambíguo.

## Objetivo

1. Permitir subir **imagem** (PNG/JPG) da fatura parcial, lida por **OCR grátis
   no navegador** (sem chave, sem servidor, sem custo), caindo na mesma tela de
   conferência antes de salvar.
2. Separar a aba Faturas em **dois blocos** — parcial e fechada — para que o
   **lugar** (não o conteúdo do arquivo) decida o tipo. Remove a ambiguidade.

## Princípios (herdados do redesenho)

- **Nada é apagado** — só reorganizado. O registro manual ("Registrar gasto
  agora") continua **idêntico**.
- **A fatura fechada é a fonte da verdade.** A parcial é só um retrato provisório
  para acompanhar o mês; não é conciliada linha a linha.
- Português (pt-BR). Sem backend novo, sem chave de IA, sem custo.

## Modelo mental: o que a parcial é (e o que ela NÃO é)

- A **parcial** (Excel ou imagem) vira `STATE.faturaAberta` — um snapshot que
  alimenta o Meu Mês e é **substituído** a cada nova parcial.
- Ela **não** consome os gastos anotados; a dedução no Meu Mês é só de exibição.
- Quando a **fechada** entra, `STATE.faturaAberta` é **descartado inteiro**
  (`= null`) e os valores reais do banco assumem. Erro de OCR na parcial é
  inofensivo: some quando a fechada chega.
- O que é conciliado de verdade com a fechada são os **gastos anotados à mão
  (pendentes)** + **assinaturas**, por cartão + valor (±R$0,10) + data (±3 dias)
  — motor `conciliarPendentes` já existente. Não muda.

## Aba Faturas em dois blocos

### Bloco 1 — Fatura parcial (acompanhamento)

- Aceita **imagem (PNG/JPG)** ou **Excel** parcial. Sempre tratada como
  **aberta** (`forcarAberta = true`), independentemente do conteúdo.
- Salva em `STATE.faturaAberta` (substitui a anterior). Alimenta o Meu Mês.
- Abaixo, o card "Fatura aberta (parcial)" com os lançamentos editáveis (já
  existe: `cardLancamentosAberta`).

### Bloco 2 — Fatura fechada (definitiva)

- Aceita **Excel ou PDF** (o arquivo exato do banco — **sem imagem** aqui, para
  ser preciso; é o que alimenta acerto e histórico). Sempre tratada como
  **fechada** (`forcarAberta = false`).
- Ao salvar: descarta a parcial (`faturaAberta = null`), roda a conciliação dos
  anotados/assinaturas, arquiva em `faturas[]`, reinicia o ciclo.
- Abaixo, a lista de faturas fechadas (`faturasList`).

> O **lugar** define aberta/fechada, sobrepondo a detecção por cabeçalho. Se
> houver conflito óbvio (ex.: Excel com "Fatura Aberta" subido no bloco fechada),
> mostrar um aviso não-bloqueante, mas respeitar o lugar escolhido.

## Parser de imagem (`parseItauImagem`)

- Carrega **Tesseract.js** por CDN, **sob demanda** (só quando uma imagem é
  subida) — não pesa o resto do app.
- Reconhece as palavras com posição (bounding boxes) e reconstrói as linhas:
  - **Cabeçalho de data** — regex `\d{1,2} de <mês>` → define a data
    (DD/MM/AAAA) dos lançamentos abaixo. Ano = atual; se a data cair no futuro,
    usa o ano anterior.
  - **Lançamento** — linha com nome à esquerda e **`R$ x,xx`** à direita. A
    linha cinza de subtipo ("apple pay" / "cartão adicional físico") é ignorada.
  - **Internacional** — pega o valor em **R$**; ignora a linha `US$`.
  - Ruído (números de página "1 de 2", ícones do topo) é descartado.
- Valor: formato do banco (`1.234,56`) → ponto=milhar, vírgula=decimal.
- Cartão: **não aparece na imagem** → atribui ao **cartão em uso**
  (`STATE.cartaoAtivo`); a Michelle ajusta na conferência.
- Produz a **mesma estrutura** de `parseItauExcel`
  (`{ cartoes, venc, subtotaisDeclarados, totalDeclarado, aberta:true }`).
  Aplica o **dicionário** (via `classificarTransacoes`) para pré-classificar.
- **Vencimento:** a imagem não mostra; reaproveita o `venc` da parcial atual
  (`STATE.faturaAberta?.venc`) se houver, senão um padrão (dia 17 do próximo mês).
  É cosmético para a parcial.
- **Múltiplos prints** (2 páginas): as imagens selecionadas juntas são
  **somadas** numa única parcial (um `FATURA_TEMP`), antes da conferência.
- **Falha de leitura:** se o OCR não achar lançamentos, mensagem clara — "Não
  consegui ler a imagem; tente um print mais nítido ou use o Excel."

## Conferência com correção (importante para OCR)

- A tela de conferência da **parcial** passa a permitir **corrigir valor e
  descrição** por linha (além de categoria/cartão/excluir), porque o OCR pode
  errar. É a rede de segurança.
- Ordenação já pedida anteriormente permanece.

## Conciliação visível na fatura fechada

- Na conferência da **fechada**, antes de confirmar, mostrar um **resumo do que
  casou** com os gastos anotados/assinaturas, ex.:
  > "✓ 5 dos seus gastos anotados casaram com esta fatura (serão arquivados).
  > ⚠ 2 não casaram — continuam para o próximo ciclo."
- Só então salva. (Hoje a conciliação acontece no salvar, porém silenciosa.)

## Registro manual (inalterado)

- "Registrar gasto agora" (cartão → `pendentes`; acerto → `extras`) continua
  **igual**. Os `pendentes` são exatamente o que concilia com a fechada.
- No Meu Mês, as três fontes convivem sem dobra (dedup já existente):
  parcial importada + anotados + assinaturas.

## Fora de escopo

- Ler a **fatura fechada** por imagem (fica em Excel/PDF por precisão).
- Backend / chave de IA / custo — tudo roda no navegador.
- Mudar o motor de conciliação (`conciliarPendentes`) — reaproveitado como está.

## Modelo de dados

- Sem coleção nova. `STATE.faturaAberta` já existe.
- `parseItauImagem` produz a mesma forma de transação dos outros parsers.
- Transação de imagem carrega `origem: 'imagem'` (opcional) só para rótulo/aviso.
