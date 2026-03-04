#!/usr/bin/env python3
"""
Calibrate model probabilities using isotonic regression on training base and re-score leads.

Outputs:
- Outputs/qualification_score_calibrated_<date>.csv
- Outputs/qualification_score_calibrated_<date>.xlsx

Produces columns: Record ID, ML_Prob_Raw, ML_Prob_Calibrated, ML_Score_Pct (0-100 percentile), ML_Score_0_100 (0-100 calibrated)
"""
import os
import glob
from datetime import datetime
import pandas as pd
import numpy as np
import joblib
import sys

from sklearn.isotonic import IsotonicRegression


def find_model():
    p = os.path.join('models','qualification_model.pkl')
    return p if os.path.exists(p) else None


def find_latest_training_base():
    files = glob.glob(os.path.join('Outputs','archive','ml_training_base_*.csv'))
    if not files:
        return None
    files.sort(key=os.path.getmtime, reverse=True)
    return files[0]


def guess_id_col(df):
    for col in df.columns:
        if col.lower() in ('record id','record_id','recordid','id','lead id','lead_id'):
            return col
    for col in df.columns:
        if col.lower().endswith('id'):
            return col
    return df.columns[0]


def infer_anomes(df):
    if 'Ano' not in df.columns or 'Mes' not in df.columns:
        created_col = None
        for c in df.columns:
            low = c.lower()
            if 'create' in low or 'data' in low or 'created' in low:
                created_col = c
                break
        if created_col:
            tmp = pd.to_datetime(df[created_col], dayfirst=True, errors='coerce')
            df['Ano'] = tmp.dt.year
            df['Mes'] = tmp.dt.month
    return df


def main():
    model_path = find_model()
    if not model_path:
        print('Model not found in models/; run build_and_apply_continuous_score.py first', file=sys.stderr)
        sys.exit(1)
    model = joblib.load(model_path)

    train_csv = find_latest_training_base()
    if not train_csv:
        print('Training base not found', file=sys.stderr)
        sys.exit(1)
    train = pd.read_csv(train_csv, low_memory=False)
    train = infer_anomes(train)

    # determine feature columns from the pipeline if possible
    feat_cols = []
    try:
        pre = model.named_steps.get('pre')
        for name, trans, cols in pre.transformers:
            if cols == 'remainder' or cols is None:
                continue
            feat_cols.extend(list(cols))
    except Exception:
        # fallback: common list
        for c in ['Fonte original do tráfego','Detalhamento da fonte original do tráfego 1','Unidade Desejada','Etapa do negócio','Número de atividades de vendas','Valor na moeda da empresa','Ano','Mes']:
            if c in train.columns:
                feat_cols.append(c)

    feat_cols = [c for c in feat_cols if c in train.columns]
    if not feat_cols:
        print('No feature columns found in training base', file=sys.stderr)
        sys.exit(1)

    X_train = train[feat_cols].copy()
    y_train = train['label'].astype(float)

    # get raw predicted probabilities on training set
    probs_train = model.predict_proba(X_train)[:,1]

    # drop rows without labels
    mask = y_train.notna()
    probs_train = probs_train[mask]
    y_train = y_train[mask]
    if len(y_train) == 0:
        print('No labeled rows in training base to calibrate.', file=sys.stderr)
        sys.exit(1)

    # fit isotonic regression for calibration
    ir = IsotonicRegression(out_of_bounds='clip')
    ir.fit(probs_train, y_train)

    # now score current leads
    leads_path = os.path.join('Data','hubspot_leads.csv')
    if not os.path.exists(leads_path):
        print('Leads file not found', file=sys.stderr)
        sys.exit(1)
    leads = pd.read_csv(leads_path, dtype=str, low_memory=False)
    leads = infer_anomes(leads)

    # ensure features present
    feat_cols_leads = [c for c in feat_cols if c in leads.columns]
    if not feat_cols_leads:
        print('No matching feature columns in leads to score', file=sys.stderr)
        sys.exit(1)
    X_leads = leads[feat_cols_leads].copy()

    probs_raw = model.predict_proba(X_leads)[:,1]
    probs_cal = ir.transform(probs_raw)

    # percentiles of calibrated probs
    pct = pd.Series(probs_cal).rank(pct=True).values * 100

    leads['ML_Prob_Raw'] = np.round(probs_raw, 6)
    leads['ML_Prob_Calibrated'] = np.round(probs_cal, 6)
    leads['ML_Score_Pct'] = np.round(pct, 3)
    leads['ML_Score_0_100'] = np.round(probs_cal * 100, 3)

    id_col = guess_id_col(leads)
    out_cols = [c for c in [id_col, 'ML_Prob_Raw','ML_Prob_Calibrated','ML_Score_Pct','ML_Score_0_100','Etapa do negócio','Data de criação','Unidade Desejada','Email','Telefone'] if c in leads.columns or c.startswith('ML_')]

    stamp = datetime.now().strftime('%Y-%m-%d')
    out_csv = f'Outputs/qualification_score_calibrated_{stamp}.csv'
    out_xlsx = f'Outputs/qualification_score_calibrated_{stamp}.xlsx'

    leads.to_csv(out_csv, index=False)
    leads[out_cols].to_excel(out_xlsx, index=False)

    print('Saved calibrated scores to:', out_csv)
    print('Saved xlsx to:', out_xlsx)


if __name__ == '__main__':
    main()
