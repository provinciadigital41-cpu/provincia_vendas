# Napkin Runbook

## Curation Rules
- Re-prioritize on every read.
- Keep recurring, high-value notes only.
- Max 10 items per category.
- Each item includes date + "Do instead".

## Domain Behavior Guardrails

1. **[2026-04-24] Template Marca usa única função para sem-risco e com-risco**
   A função `montarVarsParaTemplateMarca` serve ambos os templates de marca (`TEMPLATE_UUID_CONTRATO` e `TEMPLATE_UUID_CONTRATO_MARCA_RISCO`). `montarVarsParaTemplateOutros` só serve o terceiro template.
   Do instead: editar `montarVarsParaTemplateMarca` para afetar ambos os templates de marca ao mesmo tempo.

2. **[2026-04-24] Campos de Descrição de Marca isolados por slot**
   Os novos templates de marca usam descrição isolada por slot: `Descrição do serviço - MARCA` (hífen) para Marca 1, e `Descrição do serviço – MARCA 2..5` (travessão –) para Marcas 2 a 5. Campo `Quantidade depósitos/processos de MARCA` foi removido dos templates de marca.
   Do instead: usar `d.desc_servico_marca_1..5` (formato "1 {SERVIÇO} JUNTO AO INPI") em vez do antigo `d.desc_servico_marca`.

3. **[2026-04-24] Caráter diferente nos nomes dos campos MARCA 2-5**
   `Descrição do serviço - MARCA` usa hífen (-), mas `Descrição do serviço – MARCA 2` até `5` usam travessão (–). Copiar/colar do template é obrigatório para não errar o caractere.
   Do instead: verificar o caractere no nome do campo antes de editar, não assumir que é igual ao campo da Marca 1.

## Execution & Validation

1. **[2026-04-24] server.js é enorme (86k tokens) — não ler inteiro**
   Do instead: usar Grep para localizar seções específicas pelo nome do campo ou variável antes de ler com offset+limit.
