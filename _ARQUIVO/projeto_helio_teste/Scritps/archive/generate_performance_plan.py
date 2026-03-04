#!/usr/bin/env python3
"""
Generate a business-day (Dias Úteis) performance plan to hit a target by a deadline.

Behavior:
- Loads the latest scored dataset from `Outputs/` if available and raw leads from `Data/`.
- Identifies leads in qualification (heuristic: stage contains 'qual' OR not won/lost).
- Ensures every lead in qualification has a score (merges scored file; fills missing with fallback)
- Computes required conversions per business day and per week to reach the target by the deadline.
- Flags "sensitive" weeks where required conversions per BD are unusually high vs historical.
- Saves results to `Outputs/performance_plan_<date>.xlsx` and CSVs.

Usage examples:
  python Scritps/generate_performance_plan.py --target_conversions 1200
  python Scritps/generate_performance_plan.py             # uses default +250 increment
"""
import argparse
import os
import sys
from datetime import datetime
import glob

import pandas as pd
import numpy as np


def find_latest_scored_file():
    candidates = glob.glob(os.path.join('Outputs', 'executive_ml_scored_*.csv'))
    if not candidates:
        return None
    candidates.sort(key=os.path.getmtime, reverse=True)
    return candidates[0]


def guess_id_column(df_left, df_right):
    common = set([c.lower() for c in df_left.columns]).intersection(set([c.lower() for c in df_right.columns]))
    # prefer Record ID style
    for pref in ['record id', 'record_id', 'recordid', 'id', 'lead id', 'lead_id']:
        if pref in common:
            # map back to original column name
            left = next((col for col in df_left.columns if col.lower() == pref), None)
            right = next((col for col in df_right.columns if col.lower() == pref), None)
            return left, right
    # fallback: try names matching exactly
    for col_left in df_left.columns:
        for col_right in df_right.columns:
            if col_left.lower() == col_right.lower():
                return col_left, col_right
    return None, None


def find_stage_column(df):
    for col in df.columns:
        low = col.lower()
        if 'etapa' in low or 'stage' in low or 'status' in low or 'pipeline' in low:
            return col
    return None


def find_created_col(df):
    for col in df.columns:
        low = col.lower()
        if 'created' in low or 'create' in low or 'data' in low and ('created' in low or 'data de' in low or 'data' in low):
            return col
    # fallback: try parse
    for col in df.columns:
        try:
            pd.to_datetime(df[col].dropna().iloc[:5], dayfirst=True)
            return col
        except Exception:
            continue
    return None


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--target_conversions', type=int, help='Absolute target conversions by deadline')
    parser.add_argument('--target_increment', type=int, default=250, help='If no absolute target, add this many conversions to current wins')
    parser.add_argument('--end_date', type=str, default='2026-03-31', help='Deadline date (YYYY-MM-DD)')
    parser.add_argument('--start_date', type=str, default=None, help='Plan start date (YYYY-MM-DD), default today')
    parser.add_argument('--scored_file', type=str, default=None, help='Path to scored CSV (optional)')
    args = parser.parse_args()

    os.makedirs('Outputs', exist_ok=True)

    # Load leads
    leads_path = os.path.join('Data', 'hubspot_leads.csv')
    if not os.path.exists(leads_path):
        print('ERROR: Data/hubspot_leads.csv not found', file=sys.stderr)
        sys.exit(1)
    leads = pd.read_csv(leads_path, dtype=str, low_memory=False)

    # Load scored data
    scored_file = args.scored_file or find_latest_scored_file()
    scored = None
    if scored_file and os.path.exists(scored_file):
        print(f'Using scored file: {scored_file}')
        scored = pd.read_csv(scored_file, low_memory=False)
    else:
        print('No scored file found in Outputs/. Proceeding with leads only; scores will be missing/fallback')

    # Normalize column names to help merges
    # Try to find id column to merge
    if scored is not None:
        left_id, right_id = guess_id_column(leads, scored)
    else:
        left_id, right_id = None, None

    if left_id is None or right_id is None:
        # try common names
        for candidate in ['Record ID', 'record_id', 'record id', 'id']:
            if candidate in leads.columns:
                left_id = candidate
                break
        if scored is not None:
            for candidate in ['Record ID', 'record_id', 'record id', 'id']:
                if candidate in scored.columns:
                    right_id = candidate
                    break

    if scored is not None and left_id and right_id:
        # ensure both id columns are string to avoid dtype merge errors
        try:
            leads[left_id] = leads[left_id].astype(str)
            scored[right_id] = scored[right_id].astype(str)
        except Exception:
            leads[left_id] = leads[left_id].astype(str)
        merged = leads.merge(scored, how='left', left_on=left_id, right_on=right_id, suffixes=('', '_scored'))
    elif scored is not None:
        # fallback: merge on index if sizes match
        print('Warning: ID columns not found for reliable merge; attempting concat by index', file=sys.stderr)
        scored = scored.reset_index(drop=True)
        merged = pd.concat([leads.reset_index(drop=True), scored.reset_index(drop=True)], axis=1)
    else:
        merged = leads.copy()

    # Find stage column and created column
    stage_col = find_stage_column(merged)
    created_col = find_created_col(merged)

    # Identify qualification leads: heuristic
    if stage_col:
        qual_mask = merged[stage_col].fillna('').str.lower().str.contains('qual')
    else:
        # fallback: assume not won and not lost -> in qualification
        qual_mask = ~(
            merged.filter(regex='(?i)win|won|matr|matrícula|matricula|conclu|perdido|lost|closed').astype(str)
            .apply(lambda col: col.str.lower().str.contains('true|1|yes|matr'), axis=1)
            .any(axis=1)
        )

    # Find score column
    score_col = None
    for c in merged.columns:
        if c.lower() in ('score', 'probability', 'prob', 'score_model', 'pred'):
            score_col = c
            break
    if score_col is None:
        # fallback: numeric columns likely to be probabilities
        numeric_cols = [c for c in merged.columns if pd.api.types.is_numeric_dtype(merged[c])]
        if numeric_cols:
            # choose column with name containing 'prob' or 'score' if possible
            for c in numeric_cols:
                if 'prob' in c.lower() or 'score' in c.lower() or c.lower().startswith('p'):
                    score_col = c
                    break
            if score_col is None:
                score_col = numeric_cols[0]

    # Fill missing scores with column median or 0
    if score_col is None:
        merged['__score__'] = np.nan
        score_col = '__score__'
    merged[score_col] = pd.to_numeric(merged[score_col], errors='coerce')
    if merged[score_col].isna().all():
        merged[score_col].fillna(0.0, inplace=True)
    else:
        merged[score_col].fillna(merged[score_col].median(), inplace=True)

    # Mark qualification leads and ensure all are classified (i.e., have a score)
    merged['in_qualification'] = qual_mask.fillna(False)

    # Determine current wins
    # Try 'is_won' or label columns
    is_won_col = None
    for c in merged.columns:
        if c.lower() in ('is_won', 'iswon', 'won', 'is_win'):
            is_won_col = c
            break
    if is_won_col is not None:
        current_wins = int(merged[is_won_col].astype(bool).sum())
    else:
        if 'label' in merged.columns:
            current_wins = int((merged['label'] == 1).sum())
        else:
            # fallback: try stage contains matrícula
            if stage_col:
                current_wins = int(merged[stage_col].fillna('').str.lower().str.contains('matr').sum())
            else:
                current_wins = 0

    # Determine target
    if args.target_conversions:
        target = int(args.target_conversions)
    else:
        target = current_wins + int(args.target_increment)

    remaining_needed = max(target - current_wins, 0)

    # dates
    if args.start_date:
        start = pd.to_datetime(args.start_date).normalize()
    else:
        start = pd.Timestamp.now().normalize()
    end = pd.to_datetime(args.end_date).normalize()

    bdays = pd.bdate_range(start, end)
    if len(bdays) == 0:
        print('No business days in the range. Check dates.', file=sys.stderr)
        sys.exit(1)

    per_bd = remaining_needed / len(bdays)

    daily_plan = pd.DataFrame({'date': bdays})
    daily_plan['weekday'] = daily_plan['date'].dt.day_name()
    daily_plan['required_conversions'] = per_bd
    daily_plan['cum_required'] = daily_plan['required_conversions'].cumsum()

    # Weekly aggregation (weeks starting Monday)
    daily_plan['week_start'] = daily_plan['date'] - pd.to_timedelta(daily_plan['date'].dt.weekday, unit='D')
    weekly = daily_plan.groupby('week_start').agg(business_days=('date', 'count'),
                                                 weekly_required=('required_conversions', 'sum')).reset_index()

    # Historical weekly conversion rate
    hist_weekly = None
    if created_col and (is_won_col or 'label' in merged.columns):
        tmp = merged[[created_col]].copy()
        tmp[created_col] = pd.to_datetime(merged[created_col], dayfirst=True, errors='coerce')
        if is_won_col:
            tmp['is_won'] = merged[is_won_col].astype(bool)
        else:
            tmp['is_won'] = (merged.get('label') == 1)
        tmp = tmp.dropna(subset=[created_col])
        if not tmp.empty:
            tmp['week_start'] = tmp[created_col] - pd.to_timedelta(tmp[created_col].dt.weekday, unit='D')
            hist_weekly = tmp.groupby('week_start').agg(wins=('is_won', 'sum'))
            # business days per week (Mon-Fri)
            hist_weekly = hist_weekly.reset_index()
            hist_weekly['bdays'] = hist_weekly['week_start'].apply(lambda d: len(pd.bdate_range(d, d + pd.Timedelta(days=6))))
            hist_weekly['wins_per_bd'] = hist_weekly['wins'] / hist_weekly['bdays'].replace(0, np.nan)

    # Sensitivity: weeks where required per BD > historical mean + 1*std
    weekly['required_per_bd'] = weekly['weekly_required'] / weekly['business_days']
    sens_thresh = None
    if hist_weekly is not None and not hist_weekly['wins_per_bd'].dropna().empty:
        mean = hist_weekly['wins_per_bd'].mean()
        std = hist_weekly['wins_per_bd'].std()
        sens_thresh = mean + std
        weekly['sensitive'] = weekly['required_per_bd'] > sens_thresh
    else:
        weekly['sensitive'] = False

    # Prepare qualification scored list
    qual_df = merged[merged['in_qualification']].copy()
    # ensure score exists
    qual_df['score'] = merged[score_col]

    # Save outputs
    stamp = datetime.now().strftime('%Y-%m-%d')
    out_prefix = f'Outputs/performance_plan_{stamp}'
    daily_csv = out_prefix + '_daily.csv'
    weekly_csv = out_prefix + '_weekly.csv'
    qual_csv = out_prefix + '_qualification_scored.csv'
    xlsx = out_prefix + '.xlsx'

    daily_plan.to_csv(daily_csv, index=False)
    weekly.to_csv(weekly_csv, index=False)
    qual_df.to_csv(qual_csv, index=False)

    with pd.ExcelWriter(xlsx) as w:
        daily_plan.to_excel(w, sheet_name='Daily_Plan', index=False)
        weekly.to_excel(w, sheet_name='Weekly_Plan', index=False)
        qual_df.to_excel(w, sheet_name='Qualification_Scored', index=False)
        if hist_weekly is not None:
            hist_weekly.to_excel(w, sheet_name='Historical_Weekly', index=False)

    # Print short summary
    print('Performance plan saved:')
    print(' -', xlsx)
    print(' -', daily_csv)
    print(' -', weekly_csv)
    print(' -', qual_csv)
    print(f'Current wins: {current_wins}; Target: {target}; Remaining needed: {remaining_needed}')
    if sens_thresh is not None:
        print(f'Historical wins per BD mean+std threshold: {sens_thresh:.4f}')


if __name__ == '__main__':
    main()
