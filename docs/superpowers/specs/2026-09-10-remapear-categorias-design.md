# Remapear categorias (De→Para) — troca radical de classificação

**Data:** 2026-09-10
**Autora do produto:** Michelle Rezende

## Problema

A Michelle reestruturou a classificação das contas: saiu de ~17 categorias
antigas para **16 categorias novas** com nomes e agrupamentos diferentes.
Precisa que **todo o histórico já lançado** passe a usar a nova classificação —
não só daqui pra frente. Os dados vivem no app (localStorage + nuvem), não no
código; então a mudança é uma **transformação dos dados dela**, feita por uma
ferramenta dentro do app.

## Objetivo

Uma ferramenta **"Trocar classificação"** (em Configurações) que aplica um
mapa **De→Para** em tudo de uma vez, com **backup automático antes** e um
**resumo antes de aplicar**. Editável e pré-preenchida com o combinado — reusável
se as categorias mudarem de novo.

## Categorias novas (16, nível único)

MORADIA · COMIDA · MANUTENÇÃO CASA · EDUCAÇÃO · SAÚDE · RESTAURANTES ·
ASSINATURAS · LAZER · VIDA ATIVA · ROGÉRIO · MICHELLE · MENINOS · MARGÔ ·
VIAGEM · DOAÇÃO · DIVERSOS

Nível único (uma etiqueta por lançamento); os subitens que a Michelle listou
(água, energia, açougue…) são só guia — o detalhe fino vai na **observação**.

## De→Para acordado

| Antiga | → Nova |
|---|---|
| Supermercado, Feira | COMIDA |
| APPs Streaming | ASSINATURAS |
| Internet | MORADIA |
| Restaurante | RESTAURANTES |
| Saúde | SAÚDE |
| Lazer | LAZER |
| Margô | MARGÔ |
| Michelle, Business Michelle | MICHELLE |
| Rogério | ROGÉRIO |
| Presente, Carro, Outros | DIVERSOS |
| Casa, Serviços, Meninos | **(em branco: `??`)** |

MANUTENÇÃO CASA, EDUCAÇÃO, VIDA ATIVA, VIAGEM, DOAÇÃO começam vazias; recebem o
que a Michelle reclassificar depois (dos que ficaram em branco).

## O que é transformado (dados da Michelle, no STATE)

A categoria aparece em: `faturas[].transacoes[].categoria`,
`faturaAberta.transacoes[].categoria`, `pendentes[].categoria`,
`assinaturas[].categoria`, `orcamentos{}` (chaves) e `dicionario{}`
(`.cat` + `.alt[][0]`). Além de `STATE.categorias` (a lista).
**NÃO** têm categoria e não são tocados: `extras`, `avulsas`, `contasFixas`,
`acertos`, `tetoMensal`.

### Regras da transformação (`aplicarDePara`)

- **Transações / pendentes / assinaturas:** `categoria = map(categoria)`; quando
  o destino é branco, vira `'??'` (não classificado — a Michelle reetiqueta).
  Categoria fora do mapa (desconhecida) → mantém como está (não perde dado).
- **orcamentos:** remapeia as chaves; quando várias antigas caem na mesma nova
  (ex.: Supermercado+Feira→COMIDA), **soma** os limites; chaves que vão pra
  branco são **descartadas**.
- **dicionario:** para cada entrada, `cat = map(cat)`; se o destino for branco,
  **remove a entrada** (pra reaprender certo quando ela reclassificar). Em `alt`,
  remapeia cada `[c,n]`, descarta os que vão pra branco, **soma** contagens de
  cats repetidas e remove alt igual à `cat` principal.
- **STATE.categorias:** substituída pela lista nova (16).
- Idempotente e puro: `aplicarDePara(state, mapa, novasCategorias) -> novoState`
  (não muta o original; retorna cópia transformada + um `resumo`).

## Constantes no código (fresh install / importações futuras)

Atualizar no `index.html` para as categorias novas (senão importações futuras
classificam no antigo):
- `CATEGORIAS_DEFAULT` → as 16 novas.
- `ITAU_PARA_MICHELLE` → destinos novos: ALIMENTAÇÃO→RESTAURANTES,
  VESTUÁRIO→MICHELLE, SAÚDE→SAÚDE, DIVERSOS→DIVERSOS, VEÍCULOS→DIVERSOS,
  TURISMO E ENTRETENIM→LAZER, MORADIA→MORADIA, EDUCAÇÃO→EDUCAÇÃO, HOBBY→LAZER.
- `DICT_INICIAL` → remapear os `.cat`/`.alt` pelo mesmo De→Para (consistência de
  fresh install). Entradas que caem em branco são removidas do seed.

## Fluxo da ferramenta (Configurações → "Trocar classificação")

1. Tela mostra a **lista nova** (editável) e uma **tabela De→Para**: cada
   categoria antiga presente nos dados → dropdown (novas + "deixar em branco"),
   pré-preenchida com o acordado.
2. **Resumo antes de aplicar:** "N lançamentos remapeados · M vão ficar sem
   categoria · L orçamentos ajustados".
3. Botão **"Fazer backup e aplicar"**: baixa o backup JSON automaticamente,
   então aplica `aplicarDePara`, salva, re-renderiza. Mensagem de sucesso +
   lembrete: "reclassifique os sem categoria no Histórico (filtro 'sem
   categoria')".
4. Confirmação dupla (a ação mexe em todo o histórico).

## Fora de escopo

- Sub-categorias (2 níveis).
- Reclassificar automaticamente os que foram pra branco (a Michelle faz no
  Histórico, que já tem edição de categoria em lote + filtro).
- Mudar o motor de classificação/conciliação.

## Segurança

- Backup automático (download) imediatamente antes de aplicar.
- Função pura testada em node contra um mini-STATE realista.
- Revisão adversarial antes do deploy.
