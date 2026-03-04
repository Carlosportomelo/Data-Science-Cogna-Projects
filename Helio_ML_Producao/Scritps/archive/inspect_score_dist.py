#!/usr/bin/env python3
import pandas as pd
import glob
import os

path = glob.glob(os.path.join('Outputs','qualification_score_continuous_*.csv'))
if not path:
    path = glob.glob(os.path.join('Outputs','**','executive_ml_scored_*.csv'), recursive=True)
    if not path:
        print('No score files found')
        raise SystemExit(1)
    path = [path[0]]
print('Inspecting', path[0])
df = pd.read_csv(path[0], low_memory=False)
if 'ML_Score' not in df.columns:
    print('ML_Score column not found; available columns:', df.columns.tolist())
    raise SystemExit(1)
scores = pd.to_numeric(df['ML_Score'], errors='coerce')
print('count', scores.count())
print('min', scores.min(), 'max', scores.max())
print('unique values sample:', sorted(scores.dropna().unique())[:10])
print('percentiles:')
print(scores.quantile([0,0.01,0.05,0.1,0.25,0.5,0.75,0.9,0.95,0.99,1.0]))
print('value counts top 20:')
print(scores.value_counts().head(20))
