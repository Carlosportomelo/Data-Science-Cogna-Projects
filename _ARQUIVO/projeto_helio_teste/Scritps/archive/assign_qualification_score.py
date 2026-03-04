#!/usr/bin/env python3
"""
Assign a 1-5 qualification score to all leads based on ML scores.

Behavior:
- Loads latest `Outputs/executive_ml_scored_*.csv` (searches Outputs and archive).
- Merges scores into `Data/hubspot_leads.csv` by Record ID (casts to str).
- Maps the model score/probability into an integer 1..5 using a configurable method:
  - `quantile` (default): equal-sized quintiles (top 20% -> 5)
  - `thresholds`: provide explicit probability thresholds (ex: 0.9,0.75,0.5,0.25)
- Saves `Outputs/qualification_score_1to5_<date>.csv` and returns also an Excel sheet.

Usage examples:
  python Scritps/assign_qualification_score.py
  python Scritps/assign_qualification_score.py --method thresholds --thresh 0.95 0.80 0.60 0.30
"""
import os
import glob
import argparse
from datetime import datetime
import pandas as pd
import numpy as np
import sys


def find_latest_scored_file():
    candidates = glob.glob(os.path.join('Outputs', '**', 'executive_ml_scored_*.csv'), recursive=True)
    if not candidates:
        return None
    candidates.sort(key=os.path.getmtime, reverse=True)
    return candidates[0]


def guess_id_col(df):
    for col in df.columns:
        if col.lower() in ('record id', 'record_id', 'recordid', 'id', 'lead id', 'lead_id'):
            return col
    for col in df.columns:
        if col.lower().endswith('id'):
            return col
    return None


def find_score_col(df):
    for c in df.columns:
        low = c.lower()
        if low in ('score', 'probability', 'prob', 'p', 'p_score', 'pred') or 'prob' in low or 'score' in low:
            return c
    # fallback: numeric column with name starting with p or pred
    numeric = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    if numeric:
        return numeric[0]
    return None


def map_quantile(scores):
    # create quintiles; highest quintile -> 5
    try:
        labels = [1,2,3,4,5]
        q = pd.qcut(scores.rank(method='first'), q=5, labels=labels)
        # qcut returns lowest label for lowest values; we want 1=lowest,5=highest -> ok
        return q.astype(int)
    except Exception:
        # fallback: use percentiles
        pct = np.nanpercentile(scores.dropna(), [20,40,60,80])
        bins = [-np.inf] + pct.tolist() + [np.inf]
        return pd.cut(scores, bins=bins, labels=[1,2,3,4,5]).astype(float).astype('Int64')


def map_thresholds(scores, thresholds):
    # thresholds is list high->low e.g. [0.95,0.8,0.6,0.3] -> bins 0..0.3->1, 0.3-0.6->2, etc
    thr = sorted(thresholds)
    bins = [-np.inf] + thr + [np.inf]
    labels = list(range(1, len(bins)))
    return pd.cut(scores, bins=bins, labels=labels).astype(float).astype('Int64')


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--method', choices=['quantile','thresholds'], default='quantile')
    parser.add_argument('--thresh', nargs='+', type=float, help='Thresholds for --method thresholds (space separated, low->high)')
    args = parser.parse_args()

    leads_path = os.path.join('Data','hubspot_leads.csv')
    if not os.path.exists(leads_path):
        print('ERROR: Data/hubspot_leads.csv not found', file=sys.stderr)
        sys.exit(1)
    leads = pd.read_csv(leads_path, dtype=str, low_memory=False)

    scored_file = find_latest_scored_file()
    if not scored_file:
        print('ERROR: No scored ML file found in Outputs/', file=sys.stderr)
        sys.exit(1)
    scored = pd.read_csv(scored_file, low_memory=False)
    print('Using scored file:', scored_file)

    left_id = guess_id_col(leads) or next(iter(leads.columns))
    right_id = guess_id_col(scored) or next(iter(scored.columns))
    leads[left_id] = leads[left_id].astype(str)
    scored[right_id] = scored[right_id].astype(str)

    merged = leads.merge(scored, how='left', left_on=left_id, right_on=right_id, suffixes=('','_scored'))

    score_col = find_score_col(merged)
    if score_col is None:
        print('ERROR: No score/probability column found in scored file', file=sys.stderr)
        sys.exit(1)

    merged[score_col] = pd.to_numeric(merged[score_col], errors='coerce')
    # fill missing with median
    if merged[score_col].isna().all():
        merged[score_col].fillna(0.0, inplace=True)
    else:
        merged[score_col].fillna(merged[score_col].median(), inplace=True)

    if args.method == 'quantile':
        merged['qualification_score_1_5'] = map_quantile(merged[score_col])
    else:
        if not args.thresh:
            print('ERROR: thresholds required for method=thresholds', file=sys.stderr)
            sys.exit(1)
        merged['qualification_score_1_5'] = map_thresholds(merged[score_col], args.thresh)

    # Ensure integer dtype
    merged['qualification_score_1_5'] = merged['qualification_score_1_5'].astype('Int64')

    # Save outputs
    stamp = datetime.now().strftime('%Y-%m-%d')
    out_csv = f'Outputs/qualification_score_1to5_{stamp}.csv'
    out_xlsx = f'Outputs/qualification_score_1to5_{stamp}.xlsx'

    merged.to_csv(out_csv, index=False)
    # write a light xlsx with only needed cols
    cols = [c for c in [left_id, 'qualification_score_1_5', score_col, 'Etapa do negócio', 'Email', 'Telefone', 'canal'] if c in merged.columns]
    pd.DataFrame(merged)[cols].to_excel(out_xlsx, index=False)

    print('Saved:', out_csv)
    print('Saved:', out_xlsx)


if __name__ == '__main__':
    main()
