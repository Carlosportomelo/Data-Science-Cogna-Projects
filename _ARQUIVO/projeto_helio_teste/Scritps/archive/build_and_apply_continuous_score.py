#!/usr/bin/env python3
"""
Train (if needed) and apply an ML model to produce a continuous qualification score (0-100)
for all leads in `Data/hubspot_leads.csv` and save results.

Behavior:
- If a serialized model exists in `models/qualification_model.pkl`, loads it and scores leads.
- Otherwise, trains a simple model using `Outputs/archive/ml_training_base_*.csv` and saves it.
- Produces `Outputs/qualification_score_continuous_<date>.csv` with column `ML_Score` (0-100, 2 decimals)

Notes:
- The training pipeline is intentionally simple (basic categorical encoding and numeric imputation).
- This script focuses on producing reproducible scores for future leads; for production, replace with the full pipeline.
"""
import os
import glob
from datetime import datetime
import pandas as pd
import numpy as np
import joblib
import sys

from sklearn.compose import ColumnTransformer
from sklearn.impute import SimpleImputer
from sklearn.preprocessing import OneHotEncoder, StandardScaler
from sklearn.ensemble import HistGradientBoostingClassifier
from sklearn.pipeline import Pipeline


MODEL_DIR = 'models'
os.makedirs(MODEL_DIR, exist_ok=True)


def find_latest_training_base():
    files = glob.glob(os.path.join('Outputs', 'archive', 'ml_training_base_*.csv'))
    if not files:
        return None
    files.sort(key=os.path.getmtime, reverse=True)
    return files[0]


def find_latest_scored_file():
    files = glob.glob(os.path.join('Outputs', '**', 'executive_ml_scored_*.csv'), recursive=True)
    if not files:
        return None
    files.sort(key=os.path.getmtime, reverse=True)
    return files[0]


def guess_id_col(df):
    for col in df.columns:
        if col.lower() in ('record id', 'record_id', 'recordid', 'id', 'lead id', 'lead_id'):
            return col
    for col in df.columns:
        if col.lower().endswith('id'):
            return col
    return df.columns[0]


def train_model(train_csv, model_path):
    print('Training model from', train_csv)
    df = pd.read_csv(train_csv, low_memory=False)
    # drop rows without label
    if 'label' not in df.columns:
        raise RuntimeError('Training base has no label column')
    df = df[df['label'].notna()].copy()
    df['label'] = df['label'].astype(float)

    # simple feature list
    cat_feats = []
    num_feats = []
    for c in ['Fonte original do tráfego', 'Detalhamento da fonte original do tráfego 1', 'Unidade Desejada', 'Etapa do negócio']:
        if c in df.columns:
            cat_feats.append(c)
    for c in ['Número de atividades de vendas', 'Valor na moeda da empresa', 'Ano', 'Mes']:
        if c in df.columns:
            num_feats.append(c)

    if not cat_feats and not num_feats:
        raise RuntimeError('No features found to train on')

    X = df[cat_feats + num_feats].copy()
    y = df['label']

    # preprocessing
    transformers = []
    if num_feats:
        transformers.append(('num', Pipeline([('impute', SimpleImputer(strategy='median')), ('scale', StandardScaler())]), num_feats))
    if cat_feats:
        transformers.append(('cat', Pipeline([('impute', SimpleImputer(strategy='constant', fill_value='missing')), ('ohe', OneHotEncoder(handle_unknown='ignore', sparse_output=False))]), cat_feats))

    preprocessor = ColumnTransformer(transformers=transformers, remainder='drop', sparse_threshold=0.0)

    clf = HistGradientBoostingClassifier(random_state=42)
    pipe = Pipeline([('pre', preprocessor), ('clf', clf)])

    pipe.fit(X, y)
    joblib.dump(pipe, model_path)
    print('Saved trained model to', model_path)
    return pipe


def load_or_train_model():
    model_path = os.path.join(MODEL_DIR, 'qualification_model.pkl')
    if os.path.exists(model_path):
        print('Loading existing model:', model_path)
        return joblib.load(model_path)
    train_csv = find_latest_training_base()
    if not train_csv:
        print('No training base found; aborting.', file=sys.stderr)
        sys.exit(1)
    return train_model(train_csv, model_path)


def main():
    # load leads
    leads_csv = os.path.join('Data', 'hubspot_leads.csv')
    if not os.path.exists(leads_csv):
        print('Leads file not found:', leads_csv, file=sys.stderr)
        sys.exit(1)
    leads = pd.read_csv(leads_csv, dtype=str, low_memory=False)

    # Ensure Ano/Mes exist by parsing creation date if needed
    if 'Ano' not in leads.columns or 'Mes' not in leads.columns:
        created_col = None
        for c in leads.columns:
            low = c.lower()
            if 'create' in low or 'data' in low or 'created' in low:
                created_col = c
                break
        if created_col:
            tmp_dates = pd.to_datetime(leads[created_col], dayfirst=True, errors='coerce')
            leads['Ano'] = tmp_dates.dt.year
            leads['Mes'] = tmp_dates.dt.month

    # load model (or train)
    model = load_or_train_model()

    # find features expected by the model
    # get feature names from preprocessor: tricky; instead reuse same feature list as training
    # We'll attempt to use columns used in training pipeline by checking transformers
    pre = model.named_steps.get('pre')
    feat_cols = []
    # ColumnTransformer stores transformers_ after fitting
    try:
        for name, trans, cols in pre.transformers:
            if cols == 'remainder' or cols is None:
                continue
            feat_cols.extend(list(cols))
    except Exception:
        # fallback: try common columns
        for c in ['Fonte original do tráfego', 'Detalhamento da fonte original do tráfego 1', 'Unidade Desejada', 'Etapa do negócio', 'Número de atividades de vendas', 'Valor na moeda da empresa', 'Ano', 'Mes']:
            if c in leads.columns:
                feat_cols.append(c)

    feat_cols = [c for c in feat_cols if c in leads.columns]
    if not feat_cols:
        print('No feature columns found in leads to score. Found none of expected columns.', file=sys.stderr)
        sys.exit(1)

    X_leads = leads[feat_cols].copy()
    # Ensure numeric columns are numeric
    for c in X_leads.columns:
        if X_leads[c].dtype == object:
            # leave as-is for categorical; numeric attempt
            try:
                X_leads[c] = pd.to_numeric(X_leads[c], errors='ignore')
            except Exception:
                pass

    # predict probabilities
    try:
        probs = model.predict_proba(X_leads)
        # assume positive class is 1 and second column
        if probs.shape[1] == 1:
            scores = probs.ravel()
        else:
            scores = probs[:, 1]
    except Exception as e:
        # fallback: use decision_function and map to 0-1 via logistic
        try:
            dec = model.decision_function(X_leads)
            from scipy.special import expit
            scores = expit(dec)
        except Exception as e2:
            print('Model prediction failed:', e, e2, file=sys.stderr)
            sys.exit(1)

    # create continuous score 0-100 with 2 decimals
    leads['ML_Score'] = (scores * 100).round(2)

    # Mark qualification leads (heuristic)
    stage_col = None
    for col in leads.columns:
        if 'etapa' in col.lower() or 'stage' in col.lower():
            stage_col = col
            break
    if stage_col:
        leads['in_qualification'] = leads[stage_col].fillna('').str.lower().str.contains('qual')
    else:
        leads['in_qualification'] = True

    out_cols = [c for c in [guess_id_col(leads), 'ML_Score', 'in_qualification', stage_col, 'Data de criação', 'Unidade Desejada', 'Fonte original do tráfego', 'Email', 'Telefone'] if c in leads.columns or c in ['ML_Score','in_qualification']]
    stamp = datetime.now().strftime('%Y-%m-%d')
    out_csv = f'Outputs/qualification_score_continuous_{stamp}.csv'
    out_xlsx = f'Outputs/qualification_score_continuous_{stamp}.xlsx'

    leads.to_csv(out_csv, index=False)
    leads[out_cols].to_excel(out_xlsx, index=False)

    print('Saved continuous scores to:', out_csv)
    print('Saved summary xlsx to:', out_xlsx)


if __name__ == '__main__':
    main()
