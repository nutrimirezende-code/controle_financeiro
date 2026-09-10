# Módulo Supermercado — cupom fiscal por IA + gastos por categoria

**Data:** 2026-09-10
**Autora do produto:** Michelle Rezende

## Problema

Além do controle de cartões/acerto, a Michelle quer um **controle separado de
compras de supermercado**: fotografar o cupom fiscal, ter os itens extraídos
(nome, quantidade, valor unitário, valor total, estabelecimento) e acompanhar
**gasto por categoria de mercado**, filtrando por mês/período.

## Objetivo

Uma aba nova **Supermercado** onde ela sobe a **foto do cupom**, a **IA** lê e
devolve os itens estruturados, ela confere/classifica (o sistema aprende), salva,
e depois **filtra e vê totais por categoria e por mês**.

## Decisões acordadas (brainstorm)

- **Leitura por IA** (não OCR grátis): cupom fiscal tem dezenas de itens em 2
  linhas cada; OCR grátis erraria muito. A IA lê o cupom inteiro com precisão.
- **Categorias de mercado próprias** (15, editáveis), distintas das 16
  financeiras: Hortifruti · Carnes e aves · Peixes e frutos do mar · Laticínios e
  ovos · Padaria · Mercearia · Bebidas · Bebidas alcoólicas · Congelados · Doces e
  snacks · Limpeza · Higiene e beleza · Pet · Bazar e utilidades · Outros.
- **Módulo independente** do controle de cartões: a fatura continua mostrando a
  compra como 1 linha "COMIDA"; o módulo Supermercado é o detalhamento por dentro.
  Não se misturam (evita dupla contagem).
- Sincroniza igual ao resto (localStorage + nuvem, merge por `updatedAt` +
  lixeira), com backup.

## Backend de IA (Vercel Function + AI Gateway)

- Função serverless nova em `api/ler-cupom.js` (Node, **sem dependências npm** —
  usa `fetch` global) que recebe a imagem (base64) e faz POST para o **Vercel AI
  Gateway** (`https://ai-gateway.vercel.sh/v1/chat/completions`, compatível
  OpenAI) com um prompt que pede **JSON estruturado** dos itens.
- **Autenticação (preferir OIDC, sem chave):** funções deployadas na Vercel
  autenticam no AI Gateway via **OIDC automático** — a Michelle **não** precisa
  criar nem colar chave nenhuma. Requer apenas o **AI Gateway ativo com créditos**
  na conta Vercel (tem crédito grátis inicial; depois precisa de billing). Só se o
  OIDC não servir, cai no fallback de `AI_GATEWAY_API_KEY` (a Michelle cria a chave
  e a define via `vercel env` — ela mesma digita o segredo; eu não manipulo chave).
  Usar OIDC pode exigir o pacote `ai` (+ build na Vercel); avaliar no build.
- **Modelo:** um modelo de visão barato (confirmar o ID atual no build via
  `curl https://ai-gateway.vercel.sh/v1/models`; candidato: Google Gemini Flash).
  Custo ~centavos por cupom.
- **Contrato da função:** entrada `{ imagemBase64 }` (aceita múltiplas imagens do
  mesmo cupom via array). Saída JSON:
  `{ estabelecimento, data:'DD/MM/AAAA', total, itens:[{descricao, qtd, unidade,
  valorUnit, valorTotal}] }`. A função valida o JSON; em erro devolve
  `{ erro }` com mensagem amigável.
- O `index.html` (estático) chama `POST /api/ler-cupom` com a imagem; o resto
  (revisão, categorização, salvar) é client-side.
- **Nada de chave no cliente** — a chave vive só no backend (env var da Vercel).

## Modelo de dados (client)

Acréscimos ao `STATE` (com migração em `normalizarEstado`, merge por id + lixeira):

- `comprasMercado[]` — cada compra:
  `{ id, estabelecimento, data:'DD/MM/AAAA', total, formaPagamento?, importadaEm,
  updatedAt, itens:[{ descricao, qtd, unidade, valorUnit, valorTotal, categoria }] }`.
- `categoriasMercado[]` — as 15 categorias (editável), com default em código.
- `dicMercado{}` — dicionário que aprende `descrição do item → categoria`
  (mesma ideia do `dicionario` das faturas: `{ CHAVE: { cat, n, alt } }`), para
  auto-classificar itens nos próximos cupons.

## Fluxo de uso (aba Supermercado)

1. **Nova compra:** botão/dropzone → escolhe a(s) foto(s) do cupom → "Lendo o
   cupom…" (chama `api/ler-cupom`).
2. **Conferência:** tela com estabelecimento/data/total (editáveis) e a lista de
   itens; cada item com `categoria` (pré-classificada pelo `dicMercado`; os sem
   regra ficam destacados). Ela corrige/edita valor/nome/categoria, exclui item se
   preciso. Aprende no `dicMercado` ao salvar (como as faturas).
3. **Salvar:** grava em `comprasMercado[]`.
4. **Lista de compras:** estabelecimento · data · total · nº itens; abrir mostra
   os itens.
5. **Relatório por categoria:** filtro por **mês/período** e estabelecimento;
   tabela "categoria × total" (e opcional: categoria × meses, como o pivot do
   Rogério). Total geral do período.

## UI / navegação

- Nova aba **Supermercado** na navegação principal (ou no menu "Mais" se lotar —
  decidir no plano; provavelmente principal, é uso frequente).
- Reaproveita componentes existentes (dropzone/upload, tela de conferência
  editável com destaque de baixa confiança, filtros/sort, barras por categoria).

## Fora de escopo

- Ler cupom por OCR grátis ou pela chave/QR da SEFAZ (SP bloqueia; inviável).
- Fundir gasto de mercado com o Meu Mês do cartão.
- Comparar preços entre mercados / histórico de preço por item (futuro possível,
  os dados por item permitem, mas não agora).

## Pré-requisitos e riscos

- **Setup da IA na Vercel** (chave AI Gateway + créditos) é obrigatório e depende
  de ação da Michelle no painel — guiar passo a passo; validar cedo (Task 1 do
  plano) antes de construir o resto.
- Custo por uso (centavos/cupom) — aceito pela Michelle.
- Revisão adversarial antes do deploy final.
