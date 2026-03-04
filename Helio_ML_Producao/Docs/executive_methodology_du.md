# Planejamento de Performance por Dias Úteis (DU) — Metodologia e Guia para Defesa

Resumo rápido (prático):
- Meta temporal: atingir a meta numérica acordada até `31/03/2026`.
- Abordagem: distribuímos o gap de conversões restantes uniformemente por Dias Úteis entre hoje (ou data informada) e a data final; identificamos semanas que exigem esforço acima do normal ("semanas sensíveis").

Entradas e suposições:
- Usamos a base de leads atual (`Data/hubspot_leads.csv`) e o arquivo de scores ML (`Outputs/executive_ml_scored_<data>.csv`) quando disponível.
- Se o usuário não fornecer uma meta absoluta, usamos um incremento padrão de +250 matrículas (configurável).
- Dias Úteis: segunda a sexta; feriados não são considerados por padrão (posso incluir feriados regionais se desejar).

Como o script funciona (versão curta e didática):
1. Conta quantas matrículas já existem hoje (wins conhecidos).
2. Calcula quantas conversões faltam para a meta (target - wins atual).
3. Conta quantos Dias Úteis restam até `31/03/2026`.
4. Divide uniformemente as conversões restantes pelos Dias Úteis → obtém a meta diária.
5. Agrupa por semanas (início na segunda) para obter metas semanais.
6. Compara a meta semanal por DU com a taxa histórica de wins por DU (média + 1 desvio-padrão). Se a meta semanal for maior que esse limite, a semana é marcada como sensível.

Por que isso é defensável em reunião:
- Transparente: é um plano deterministicamente calculado em função de quantos dias úteis restam e de quantas conversões faltam.
- Conservador: o critério de sensibilidade usa média histórica + 1 desvio-padrão, o que destaca semanas realmente fora do padrão.
- Auditable: o script salva os arquivos e todas as suposições (datas, alvo, contagem de dias úteis), permitindo reproduzir os números na reunião.

Outputs gerados automaticamente pelo script `Scritps/generate_performance_plan.py`:
- `Outputs/performance_plan_<date>_daily.csv` — plano por dia útil (datas e meta diária).
- `Outputs/performance_plan_<date>_weekly.csv` — agregação por semana com flag `sensitive`.
- `Outputs/performance_plan_<date>_qualification_scored.csv` — todos os leads em qualificação com score (pronto para ação operacional).

Recomendações operacionais imediatas:
- Para semanas sensíveis: aumentar follow-up humano, priorizar os canais com maior conversão e considerar campanhas de push (SMS/WhatsApp) nessas janelas.
- Para leads em qualificação com score alto (top percentil): definir SLA de 2 horas e checklist mínimo de contato (telefone + WhatsApp + tentativa de reenvio de proposta).

Se quiser, na próxima entrega eu:
- incluo feriados regionais no calendário de Dias Úteis, para maior precisão; ou
- calibro as probabilidades do modelo (Platt/Isotonic) para que os scores sejam interpretáveis como probabilidades; ou
- gero um PDF pronto para apresentação com as tabelas principais já formatadas.

---

Arquivo gerado automaticamente por análise; me avise se quer que eu una este conteúdo ao `Docs/executive_methodology.md` ou gere um PDF pronto para apresentação.
