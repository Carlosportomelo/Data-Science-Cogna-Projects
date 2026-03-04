**Análise Executiva — Metodologia e Plano de Ação**

Resumo rápido
- Objetivo: detalhar como priorizamos leads, calcular quantos leads adicionais por canal são necessários para atingir uma meta de +10% em matrículas e responder perguntas operacionais (ex.: atacar offline, incentivos a secretárias).
- Fonte: tabela de resumo fornecida (snapshot de canais).

Snapshot (dados fornecidos)
- Cadastro da unidade: Leads gerados = 8.035 ; Matrículas = 5.466 ; Conversão = 68%
- Digital - Orgânico: Leads gerados = 2.153 ; Matrículas = 653 ; Conversão = 30%
- Digital - Pago: Leads gerados = 8.595 ; Matrículas = 426 ; Conversão = 5%
- Total: Leads = 18.781 ; Matrículas = 6.545 ; Conversão média = 35%

Meta e interpretação
- Meta solicitada: aumentar matrículas em 10% (baseline).
- Matrículas atuais: 6.545. Meta (10%): 6.545 * 1.10 = 7.199 (arredondado).
- Matrículas adicionais necessárias: ~655 (7.199 - 6.545).

Metodologia de cálculo (simples, transparente)
- Premissas: usamos a taxa de conversão observada por canal (Matrículas / Leads) como estimativa de eficiência futura. Isto é um modelo determinístico e não incorpora elasticidade de mídia, saturação ou custo.
- Fórmula básica:
  - Leads necessários para gerar N matrículas = N / (Taxa de conversão do canal)

Exemplos (leads necessários para gerar os +655 matrículas):
- Se investirmos 100% no canal "Cadastro da unidade" (Conv = 68%): Leads necessários ≈ 963 leads
- Se investirmos 100% em "Digital - Orgânico" (Conv = 30%): Leads necessários ≈ 2.184 leads
- Se investirmos 100% em "Digital - Pago" (Conv = 5%): Leads necessários = 13.100 leads

Mix recomendado (realista e eficiente)
- Prioridade prática: combinar canais de alta eficiência com capacidade de escala.
- Exemplo de alocação conservadora (70% offline / 30% orgânico):
  - Matrículas alvo via Cadastro da Unidade: 0.70 * 655 ≈ 459 → Leads ≈ 459 / 0.68 ≈ 675 leads
  - Matrículas alvo via Digital Orgânico: 0.30 * 655 ≈ 196 → Leads ≈ 196 / 0.30 ≈ 654 leads
  - Total leads adicionais (estimado) ≈ 1.329 leads

Por que esse mix?
- O canal 'Cadastro da unidade' tem conversão muito alta (68%) — gerar leads aqui rende mais matrículas por unidade de esforço.
- 'Digital - Orgânico' tem conversão razoável (30%) e escala sem custo direto por lead se você ativar SEO/Conteúdo; é o segundo canal mais eficiente.
- 'Digital - Pago' hoje tem conversão muito baixa (5%). Só recomenda-se escalar pago se houver teste criativo/segmentação com expectativa real de elevar conversão, ou se houver orçamento para escalar o volume.

Scoring de prioridade de lead (Plano Bala - resumo metodológico)
- Objetivo: dar nota a cada lead (0–100) que represente a propensão à matrícula, para priorizar ações operacionais.
- Variáveis usadas (heurística implementada no código):
  1. Taxa histórica do canal (Conv_Rate) — canais com melhor conv. recebem maior peso.
  2. Recência (Ano do lead) — leads mais recentes recebem peso maior (escala 0.5–1.0 entre ano mínimo e máximo observado).
  3. Stage multiplier (Etapa do negócio) — leads em "em qualificação" ou "novo negócio" ganham multiplicador (>1.0).
- Fórmula (resumida):
  - score_raw = Conv_Norm * Recency_W * Stage_Mult
  - score_normalizado = (score_raw - min) / (max - min) * 100  (escala 0–100)
- Interpretação rápida:
  - 80–100: altíssima prioridade — contatar imediatamente, priorizar follow-up humano
  - 60–80: alta prioridade — incluir em campanha de retorno + contato humano rápido
  - 40–60: média — nutrir + automação
  - <40: baixa — manter em fluxo automatizado / reengajamento programado

Requisitos de dados (sempre presentes)
- Para cada lead entregue na recomendação ou CSV de operação devemos trazer:
  - `lead_id` (ID do registro)
  - `canal` (ex.: Cadastro da unidade, Digital - Orgânico, Digital - Pago)
  - `origem` / `Fonte original do tráfego` (string detalhada: ex.: 'Social pago', 'Pesquisa paga', 'CRM_UI')
  - `Etapa do negócio` (para aplicar multiplicador)
  - `Data de criação` (para recência)
  - `Plano_Bala_Score` (nota 0–100)

Resposta às perguntas operacionais
- 1) "Se o cadastro da unidade diminuiu, temos que atacar mais em ações offline nas unidades?"
  - Nem sempre. Porque o canal offline tem a maior conversão (68%), deslocar esforços para gerar leads offline é eficiente — mas o mais importante é garantir que os leads gerados convertam: qualidade do atendimento, processo de agendamento e follow-up fazem grande diferença.
  - Se a diminuição no cadastro da unidade decorre de falha operacional (ex.: secretárias não cadastrando), então sim: reativar/treinar e incentivar o cadastro pode recuperar matrículas com baixo volume de leads adicionais.
  - Se a diminuição veio por falta de demanda, pode ser melhor focar em melhorar conversão nos leads existentes (treinamento, scripts, SLA) antes de elevar captação offline.

- 2) "Soltamos um incentivo de cadastro de leads para as secretárias?"
  - Recomendação: incentive por resultados (matrículas) e não apenas por cadastro.
  - Por quê? Cadastro puro (volume) pode gerar leads de baixa qualidade; pagar por matrícula (ou por meta híbrida: base + bônus por conversão) alinha comportamento com o resultado desejado.
  - Operacional: implementar KPI mensal por unidade (ex.: conversão de leads para matrícula) e bônus por taxa de conversão incremental.

Implementação prática e próximos passos
1) Produzir lista operacional (CSV) com colunas obrigatórias (`lead_id`, `canal`, `origem`, `Etapa do negócio`, `Data de criação`, `Plano_Bala_Score`) — já gerada pelo script `Scritps/executive_report.py` como `Outputs/Executive_Looker_<date>.csv`.
2) Gerar a lista top-N para ação imediata (ex.: top 250 pelo `Plano_Bala_Score`) — gerada na aba `5_Plano_Bala_Top` do Excel.
3) Atribuir execução:
   - Equipe comercial/unidades: receber top leads com nota alta (80+), SLA < 2 horas
   - Equipe de operação/secretárias: KPI de conversão, bônus por incremento
   - Marketing: rodar ações de SEO e conteúdo para orgânico; testar criativos e segmentação para pago antes de escalonar

Observações finais e limitações
- Cálculos acima assumem taxa de conversão estável por canal. Na prática, conversão pode variar com a qualidade do tráfego, mudanças de processo e sazonalidade.
- Recomendo rodar o cálculo de necessidade de leads semanalmente usando o dataset atualizado (o script consolidado `Scritps/executive_report.py` já lê a base e produz os arquivos necessários).

Appendix — números (resumo de cálculo)
- Cenários:
  - 100% em Cadastro da unidade (68%): 655 matrículas → ~963 leads
  - 100% em Digital - Orgânico (30%): 655 matrículas → ~2.184 leads
  - 100% em Digital - Pago (5%): 655 matrículas → 13.100 leads
  - Mix sugerido (70% offline / 30% orgânico): 655 matrículas → ~1.329 leads

Se quiser, eu:
- gero uma versão PDF deste MD para apresentação; ou
- ajusto os percentuais do mix (ex.: 60/40) e mostro o impacto em leads necessários; ou
- integro estes cálculos direto ao `Scritps/analysis_tool.py` como um comando `plan_leads` para gerar automaticamente cenários.

— Fim —

*** Arquivo gerado automaticamente por análise; contate-me se quiser números por unidade ou simulações com custos.
