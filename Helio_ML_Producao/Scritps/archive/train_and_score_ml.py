import os
import glob
import pandas as pd
import numpy as np
from datetime import datetime

from sklearn.model_selection import train_test_split
from sklearn.linear_model import LogisticRegression
from sklearn.ensemble import HistGradientBoostingClassifier
from sklearn.metrics import roc_auc_score, precision_score

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(BASE, 'Outputs')
os.makedirs(OUT, exist_ok=True)

DATE = datetime.now().strftime('%Y-%m-%d')

def find_ml_base():
    files = glob.glob(os.path.join(OUT, 'ml_training_base_*.csv'))
    if not files:
        raise FileNotFoundError('No ml_training_base_*.csv found in Outputs/')
    return max(files, key=os.path.getmtime)

def preprocess(df):
    # Copy important cols
    df = df.copy()
    # Fill missing text fields
    text_cols = ['Fonte original do tráfego', 'Detalhamento da fonte original do tráfego 1', 'Detalhamento da fonte original do tráfego 2', 'Etapa do negócio', 'Unidade Desejada', 'Proprietário do negócio', 'motivo_perda']
    for c in text_cols:
        if c in df.columns:
            df[c] = df[c].fillna('(missing)').astype(str)
        else:
            df[c] = '(missing)'

    # Numeric features
    if 'Número de atividades de vendas' in df.columns:
        df['num_atividades'] = pd.to_numeric(df['Número de atividades de vendas'], errors='coerce').fillna(0)
    else:
        df['num_atividades'] = 0

    if 'Valor na moeda da empresa' in df.columns:
        df['valor_clean'] = pd.to_numeric(df['Valor na moeda da empresa'].astype(str).str.replace('[^0-9\.-]', '', regex=True), errors='coerce').fillna(0)
    else:
        df['valor_clean'] = 0

    # Recency: use Ano and Mes
    if 'Data de criação' in df.columns:
        df['data_cri'] = pd.to_datetime(df['Data de criação'], errors='coerce')
        df['days_since'] = (pd.Timestamp.now() - df['data_cri']).dt.days.fillna(9999)
    else:
        df['days_since'] = 9999

    # Reduce cardinality for categorical: keep top categories
    def top_n_onehot(series, prefix, n=30):
        top = series.value_counts().nlargest(n).index.tolist()
        s = series.where(series.isin(top), other='__other__')
        return pd.get_dummies(s, prefix=prefix)

    X_text = pd.DataFrame(index=df.index)
    X_text = pd.concat([X_text,
                        top_n_onehot(df['Fonte original do tráfego'], 'chan', n=40),
                        top_n_onehot(df['Detalhamento da fonte original do tráfego 1'], 'chan_det', n=50),
                        top_n_onehot(df['Unidade Desejada'], 'unit', n=50)
                        ], axis=1)

    X_num = df[['num_atividades', 'valor_clean', 'days_since']].copy()

    X = pd.concat([X_num.reset_index(drop=True), X_text.reset_index(drop=True)], axis=1).fillna(0)
    return X, df

def precision_at_k(y_true, y_score, k=250):
    idx = np.argsort(y_score)[::-1][:k]
    return precision_score(y_true.iloc[idx], np.ones(len(idx)))

def main():
    base_file = find_ml_base()
    print('[ML TRAIN] Using base:', base_file)
    df = pd.read_csv(base_file, encoding='utf-8')

    # Preprocess and feature matrix
    X, df = preprocess(df)

    # Labels
    if 'label' not in df.columns:
        raise SystemExit('label column not found in training base')

    y = df['label']

    # Split: time-based if Ano exists
    if 'Ano' in df.columns and df['Ano'].notna().any():
        train_mask = df['Ano'] <= 2024
        val_mask = df['Ano'] == 2025
        if train_mask.sum() < 100 or val_mask.sum() < 100:
            # fallback to random
            X_train, X_val, y_train, y_val = train_test_split(X[y.notna()], y[y.notna()], test_size=0.2, random_state=42, stratify=y[y.notna()])
        else:
            X_train = X[train_mask & y.notna()]
            y_train = y[train_mask & y.notna()]
            X_val = X[val_mask & y.notna()]
            y_val = y[val_mask & y.notna()]
    else:
        X_train, X_val, y_train, y_val = train_test_split(X[y.notna()], y[y.notna()], test_size=0.2, random_state=42, stratify=y[y.notna()])

    print(f'[ML TRAIN] Train rows: {len(X_train)}, Val rows: {len(X_val)}')

    # Logistic Regression baseline
    print('[ML TRAIN] Training Logistic Regression...')
    lr = LogisticRegression(max_iter=1000, class_weight='balanced')
    lr.fit(X_train, y_train)
    lr_scores = lr.predict_proba(X_val)[:, 1]
    lr_auc = roc_auc_score(y_val, lr_scores)
    lr_p250 = precision_at_k(y_val.reset_index(drop=True), lr_scores, k=min(250, len(y_val)))

    print(f'  LR AUC: {lr_auc:.4f}, Precision@250: {lr_p250:.4f}')

    # Gradient boosting
    print('[ML TRAIN] Training HistGradientBoostingClassifier...')
    gb = HistGradientBoostingClassifier(max_iter=200)
    gb.fit(X_train, y_train)
    gb_scores = gb.predict_proba(X_val)[:, 1]
    gb_auc = roc_auc_score(y_val, gb_scores)
    gb_p250 = precision_at_k(y_val.reset_index(drop=True), gb_scores, k=min(250, len(y_val)))
    print(f'  HGB AUC: {gb_auc:.4f}, Precision@250: {gb_p250:.4f}')

    # Choose best model (by AUC)
    best_model = gb if gb_auc >= lr_auc else lr
    best_name = 'HistGradientBoosting' if gb_auc >= lr_auc else 'LogisticRegression'
    print(f'[ML TRAIN] Best model: {best_name}')

    # Score full dataset (where label is NaN — candidates)
    full_scores = best_model.predict_proba(X.fillna(0))[:, 1]
    df['ML_Score'] = full_scores

    scored_csv = os.path.join(OUT, f'executive_ml_scored_{DATE}.csv')
    cols_to_export = [c for c in ['Record ID', 'Nome do negócio', 'Data de criação', 'Etapa do negócio', 'Fonte original do tráfego', 'Detalhamento da fonte original do tráfego 1', 'Unidade Desejada', 'motivo_perda', 'label'] if c in df.columns]
    cols_to_export += ['ML_Score']
    df[cols_to_export].to_csv(scored_csv, index=False)
    print('[ML TRAIN] Scored CSV saved:', scored_csv)

    # Top-N recommendations among unlabeled (or across all) - choose unlabeled
    candidates = df[df['label'].isna()].copy()
    topn = candidates.sort_values('ML_Score', ascending=False).head(250)
    top_csv = os.path.join(OUT, f'top_recommendations_{DATE}.csv')
    topn[cols_to_export].to_csv(top_csv, index=False)
    print('[ML TRAIN] Top recommendations saved:', top_csv)

    # Save simple metrics report
    report = os.path.join(OUT, f'ml_report_{DATE}.txt')
    with open(report, 'w') as f:
        f.write(f'Logistic AUC: {lr_auc:.4f} Precision@250: {lr_p250:.4f}\n')
        f.write(f'HGB AUC: {gb_auc:.4f} Precision@250: {gb_p250:.4f}\n')
        f.write(f'Chosen model: {best_name}\n')
    print('[ML TRAIN] Report saved:', report)

if __name__ == "__main__":
    main()
