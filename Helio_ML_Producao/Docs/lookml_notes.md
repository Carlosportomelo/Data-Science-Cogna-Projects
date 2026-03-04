**Looker Dataset Notes**

Purpose
- Flat CSV prepared for Looker to build dashboards and explore lead propensity.

File
- `Outputs/looker_dataset_<date>.csv` — main dataset. Contains one row per lead with ML score and supporting fields.

Key columns (name -> description, type)
- `lead_id` -> Unique lead identifier (string/integer). Primary key. (string)
- `data_criacao` -> Lead creation datetime. (datetime)
- `channel` -> High-level channel (`Fonte original do tráfego`). (string)
- `channel_detail` -> Detailed channel / campaign info. (string)
- `unit` -> `Unidade Desejada` (target unit). (string)
- `stage` -> Current `Etapa do negócio`. (string)
- `motivo_perda` -> If lead matched a lost-deal, textual reason from `Motivo do negócio perdido`. (string)
- `num_activities` -> Número de atividades de vendas (numeric). (integer)
- `value` -> Valor na moeda da empresa (numeric). (float)
- `owner` -> Proprietário do negócio (sales owner). (string)
- `plano_bala_score` -> Heuristic score (0-100) computed by rule-based method. (float)
- `ml_score` -> ML model score (probability-like) indicating propensity to convert. (float)

Notes and recommendations
- Use `ml_score` to sort/prioritize leads in dashboards and to create filters for Top N lists.
- Use `motivo_perda` to analyze common loss reasons and to feed back into marketing/operations.
- Suggested Looker measures:
  - Count of leads, Sum of `num_activities`, Avg(`ml_score`), Conversion rate by grouping by `channel`/`unit`.
- Suggested dimensions:
  - `lead_id`, `data_criacao` (date + month), `channel`, `channel_detail`, `unit`, `stage`, `owner`, `motivo_perda`.

Refresh cadence
- Re-generate CSV weekly (or nightly) depending on operational needs. Retrain ML weekly if labels update frequently.

Contact
- For questions about columns or to change the export, contact the analytics owner.
