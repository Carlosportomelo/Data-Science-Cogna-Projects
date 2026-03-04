#!/usr/bin/env python3
"""
Generate Top-N lists (Top50, Top250, Top1000) and Top10 per unit using ML scores.

Behavior:
- Loads latest `Outputs/executive_ml_scored_*.csv` and `Data/hubspot_leads.csv`.
- Ensures merge on Record ID (casts to str) and fills missing scores with median.
- Identifies 'in qualification' leads and produces:
  - `Outputs/top50_<date>.csv`, `top250`, `top1000`
  - `Outputs/top10_per_unit_<date>.csv` (unit, top10 leads)
  - `Outputs/top_lists_<date>.xlsx` with sheets for each list

Usage:
  python Scritps/generate_top_n_lists.py
"""
import os
import glob
from datetime import datetime
import pandas as pd
import numpy as np
import sys


def find_latest_scored_file():
    # search recursively in Outputs (including archive)
    candidates = glob.glob(os.path.join('Outputs', '**', 'executive_ml_scored_*.csv'), recursive=True)
    if not candidates:
        return None
    candidates.sort(key=os.path.getmtime, reverse=True)
    return candidates[0]


def find_stage_col(df):
    for col in df.columns:
        low = col.lower()
        if 'etapa' in low or 'stage' in low or 'status' in low:
            return col
    return None


def find_unit_col(df):
    for col in df.columns:
        low = col.lower()
        if 'unidade' in low or 'unit' in low or 'desired' in low:
            return col
    return None


def guess_id_col(df):
    for col in df.columns:
        if col.lower() in ('record id', 'record_id', 'recordid', 'id', 'lead id', 'lead_id'):
            return col
    # fallback: return first column named like 'id'
    for col in df.columns:
        if col.lower().endswith('id'):
            return col
    return None


def main():
    os.makedirs('Outputs', exist_ok=True)
    leads_path = os.path.join('Data', 'hubspot_leads.csv')
    if not os.path.exists(leads_path):
        print('ERROR: leads file not found', file=sys.stderr)
        sys.exit(1)
    leads = pd.read_csv(leads_path, dtype=str, low_memory=False)

    scored_file = find_latest_scored_file()
    if not scored_file:
        print('ERROR: scored file not found in Outputs/', file=sys.stderr)
        sys.exit(1)
    scored = pd.read_csv(scored_file, low_memory=False)
    print('Using scored file:', scored_file)

    # Guess id columns
    left_id = guess_id_col(leads) or next(iter(leads.columns))
    right_id = guess_id_col(scored) or next(iter(scored.columns))
    # cast to str
    leads[left_id] = leads[left_id].astype(str)
    scored[right_id] = scored[right_id].astype(str)

    merged = leads.merge(scored, how='left', left_on=left_id, right_on=right_id, suffixes=('', '_scored'))

    # Prefer calibrated ML score column names if present
    preferred = ['ML_Score_0_100', 'ML_Prob_Calibrated', 'ML_Prob_Raw', 'ML_Score_Pct']
    score_col = None
    for p in preferred:
        if p in merged.columns:
            score_col = p
            break
    if score_col is None:
        # fallback to common names
        for c in merged.columns:
            if c.lower() in ('score', 'probability', 'prob', 'p_score', 'p_pred'):
                score_col = c
                break
    if score_col is None:
        # fallback to first numeric column
        numeric = [c for c in merged.columns if pd.api.types.is_numeric_dtype(merged[c])]
        if numeric:
            score_col = numeric[0]
    if score_col is None:
        merged['__score__'] = 0.0
        score_col = '__score__'

    # ensure numeric and fill missing with median (or 0 when all NA)
    merged[score_col] = pd.to_numeric(merged[score_col], errors='coerce')
    if merged[score_col].isna().all():
        merged[score_col] = 0.0
    else:
        merged[score_col].fillna(merged[score_col].median(), inplace=True)

    # identify qualification
    stage_col = find_stage_col(merged)
    if stage_col:
        merged['in_qualification'] = merged[stage_col].fillna('').str.lower().str.contains('qual')
    else:
        merged['in_qualification'] = True

    qual = merged[merged['in_qualification']].copy()

    # select columns to export
    keep = [c for c in ['Record ID', left_id] if c in qual.columns]
    # prefer some columns if present
    for c in ['lead_id', 'lead id', 'Record ID', 'Email', 'Telefone', 'canal', 'Canal', 'canal_principal', 'Etapa do negócio', 'Data de criação']:
        if c in qual.columns and c not in keep:
            keep.append(c)
    # always include score
    if score_col not in keep:
        keep.append(score_col)

    qual_out = qual[keep].copy()
    qual_out = qual_out.sort_values(score_col, ascending=False)
    qual_out['rank'] = range(1, len(qual_out) + 1)

    stamp = datetime.now().strftime('%Y-%m-%d')
    top50 = qual_out.head(50)
    top250 = qual_out.head(250)
    top1000 = qual_out.head(1000)

    top50_csv = f'Outputs/top50_{stamp}.csv'
    top250_csv = f'Outputs/top250_{stamp}.csv'
    top1000_csv = f'Outputs/top1000_{stamp}.csv'
    qual_csv = f'Outputs/qualificacao_todos_pontuados_{stamp}.csv'

    top50.to_csv(top50_csv, index=False)
    top250.to_csv(top250_csv, index=False)
    top1000.to_csv(top1000_csv, index=False)
    qual_out.to_csv(qual_csv, index=False)

    # Top10 per unit
    unit_col = find_unit_col(merged)
    top10_unit_csv = f'Outputs/top10_por_unidade_{stamp}.csv'
    if unit_col:
        grouped = qual.groupby(unit_col)
        rows = []
        for unit, df in grouped:
            df2 = df.sort_values(score_col, ascending=False).head(10)
            df2 = df2.assign(unidade=unit)
            rows.append(df2)
        if rows:
            out_unit = pd.concat(rows, axis=0)
            # include useful columns
            cols = [c for c in [unit_col, left_id, 'Email', 'Telefone', 'Etapa do negócio', score_col] if c in out_unit.columns]
            out_unit[cols].to_csv(top10_unit_csv, index=False)
        else:
            # empty
            pd.DataFrame().to_csv(top10_unit_csv, index=False)
    else:
        # export empty with note
        pd.DataFrame({'note': ['unit column not found']}).to_csv(top10_unit_csv, index=False)

    # Write excel workbook
    xlsx = f'Outputs/listas_top_{stamp}.xlsx'
    with pd.ExcelWriter(xlsx) as w:
        top50.to_excel(w, sheet_name='Top50', index=False)
        top250.to_excel(w, sheet_name='Top250', index=False)
        top1000.to_excel(w, sheet_name='Top1000', index=False)
        if unit_col and rows:
            out_unit[cols].to_excel(w, sheet_name='Top10_por_unidade', index=False)
        # rename rank to Portuguese
        qual_out = qual_out.rename(columns={'rank': 'Posição'}) if 'rank' in qual_out.columns else qual_out
        qual_out.to_excel(w, sheet_name='Qualificacao_Todos', index=False)

    print('Top lists saved:')
    print(' -', top50_csv)
    print(' -', top250_csv)
    print(' -', top1000_csv)
    print(' -', top10_unit_csv)
    print(' -', xlsx)


if __name__ == '__main__':
    main()
