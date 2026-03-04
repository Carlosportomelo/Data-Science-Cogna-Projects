import os
import glob
import pandas as pd
from datetime import datetime

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(BASE, 'Data')
OUT = os.path.join(BASE, 'Outputs')
os.makedirs(OUT, exist_ok=True)

DATE = datetime.now().strftime('%Y-%m-%d')
OUT_CSV = os.path.join(OUT, f'looker_dataset_{DATE}.csv')

def latest(pattern):
    files = glob.glob(os.path.join(OUT, pattern))
    return max(files, key=os.path.getmtime) if files else None

def read_csv(path):
    try:
        return pd.read_csv(path, encoding='utf-8')
    except Exception:
        return pd.read_csv(path, encoding='latin1')

def main():
    # source files
    hubspot = os.path.join(DATA, 'hubspot_leads.csv')
    scored = latest('executive_ml_scored_*.csv')
    plano = latest('Executive_Looker_*.csv')

    if not os.path.exists(hubspot):
        raise SystemExit('hubspot_leads.csv not found in Data/')

    print('Reading hubspot leads...')
    df_h = read_csv(hubspot)
    df_h.columns = df_h.columns.str.strip()

    if scored:
        print('Reading scored file:', scored)
        df_s = read_csv(scored)
        df_s.columns = df_s.columns.str.strip()
    else:
        df_s = pd.DataFrame()

    # normalize key
    key = None
    for c in df_h.columns:
        if c.strip().lower() in ['record id', 'record_id', 'id', 'hs_object_id']:
            key = c
            break
    if not key:
        raise SystemExit('No id column found in hubspot CSV')

    df_h['_id_str'] = df_h[key].astype(str)
    if not df_s.empty:
        # try to find id in scored
        sid = None
        for c in df_s.columns:
            if c.strip().lower() in ['record id', 'record_id', 'id', 'lead_id']:
                sid = c
                break
        if sid:
            df_s['_id_str'] = df_s[sid].astype(str)
        else:
            # if scored doesn't have id, assume same order
            df_s = df_s.reset_index().rename(columns={'index':'_id_index'})

    # merge
    if not df_s.empty and '_id_str' in df_s.columns:
        df = df_h.merge(df_s.drop_duplicates('_id_str'), left_on='_id_str', right_on='_id_str', how='left', suffixes=('', '_ml'))
    else:
        df = df_h.copy()

    # select columns for Looker dataset
    mapping = {
        key: 'lead_id',
        'Data de criação': 'data_criacao',
        'Fonte original do tráfego': 'channel',
        'Detalhamento da fonte original do tráfego 1': 'channel_detail',
        'Unidade Desejada': 'unit',
        'Etapa do negócio': 'stage',
        'Motivo do negócio perdido': 'motivo_perda',
        'Número de atividades de vendas': 'num_activities',
        'Valor na moeda da empresa': 'value',
        'Proprietário do negócio': 'owner'
    }

    out_cols = []
    for src, dst in mapping.items():
        if src in df.columns:
            out_cols.append((src, dst))

    # also include scores
    score_cols = []
    if 'Plano_Bala_Score' in df.columns:
        score_cols.append(('Plano_Bala_Score', 'plano_bala_score'))
    if 'ML_Score' in df.columns:
        score_cols.append(('ML_Score', 'ml_score'))

    # build output dataframe
    out_df = pd.DataFrame()
    for src, dst in out_cols:
        out_df[dst] = df[src]
    for src, dst in score_cols:
        out_df[dst] = df[src]

    # fallback: include Record ID if mapping didn't include it
    if 'lead_id' not in out_df.columns:
        out_df['lead_id'] = df_h[key]

    out_df.to_csv(OUT_CSV, index=False)
    print('Looker dataset exported to:', OUT_CSV)

if __name__ == '__main__':
    main()
