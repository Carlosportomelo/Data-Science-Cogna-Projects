import os
import pandas as pd
import numpy as np
from datetime import datetime

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LEADS_CSV = os.path.join(BASE, 'Data', 'hubspot_leads.csv')
LOST_CSV = os.path.join(BASE, 'Data', 'hubspot_negocios_perdidos.csv')
OUT = os.path.join(BASE, 'Outputs')
os.makedirs(OUT, exist_ok=True)

DATE = datetime.now().strftime('%Y-%m-%d')
OUT_CSV = os.path.join(OUT, f'ml_training_base_{DATE}.csv')

def read_csv(path):
    try:
        return pd.read_csv(path, encoding='utf-8')
    except Exception:
        return pd.read_csv(path, encoding='latin1')

def prepare():
    print('[ML PREP] Reading leads...')
    leads = read_csv(LEADS_CSV)
    print(f"  - leads rows: {len(leads):,}")

    print('[ML PREP] Reading lost-deals...')
    lost = read_csv(LOST_CSV)
    print(f"  - lost rows: {len(lost):,}")

    # Normalize column names
    leads.columns = leads.columns.str.strip()
    lost.columns = lost.columns.str.strip()

    # Key: try to use 'Record ID' if present
    key = None
    for cand in ['Record ID', 'RecordID', 'record id', 'record_id']:
        if cand in leads.columns:
            key = cand
            break
    if key is None:
        raise SystemExit('Record ID column not found in leads file')

    # Ensure key exists in lost file
    if key not in lost.columns:
        # try common alternatives
        if 'Record ID' in lost.columns:
            lost_key = 'Record ID'
        else:
            # fallback: no join possible -> mark none as lost
            print('[ML PREP] WARNING: lost file has no Record ID column; no matches will be created')
            lost_key = None
    else:
        lost_key = key

    # Minimal date parse
    if 'Data de criação' in leads.columns:
        leads['Data de criação'] = pd.to_datetime(leads['Data de criação'], errors='coerce')
        leads['Ano'] = leads['Data de criação'].dt.year
        leads['Mes'] = leads['Data de criação'].dt.month
    else:
        leads['Ano'] = np.nan
        leads['Mes'] = np.nan

    # is_won: Etapa do negócio contains 'MATRÍCULA'
    if 'Etapa do negócio' in leads.columns:
        leads['is_won'] = leads['Etapa do negócio'].astype(str).str.contains('MATRÍCULA', case=False, na=False)
    else:
        leads['is_won'] = False

    # is_lost: match Record ID with lost file and capture motivo
    leads['_rec_id_str'] = leads[key].astype(str)
    leads['motivo_perda'] = None
    if lost_key is not None:
        lost_ids = lost[lost_key].dropna().astype(str).astype(object).unique()
        # Compare as strings
        leads['is_lost'] = leads['_rec_id_str'].isin([str(x) for x in lost_ids])
        # Build mapping from lost file for motivo
        if 'Motivo do negócio perdido' in lost.columns:
            # normalize lost keys to strings
            lost_map = lost.set_index(lost_key)['Motivo do negócio perdido'].dropna().to_dict()
            # ensure string keys
            lost_map = {str(k): v for k, v in lost_map.items()}
            leads['motivo_perda'] = leads['_rec_id_str'].map(lost_map)
    else:
        leads['is_lost'] = False

    # label: 1 = won, 0 = lost, NaN else
    def label_row(r):
        if r['is_won']:
            return 1
        if r['is_lost']:
            return 0
        return np.nan

    leads['label'] = leads.apply(label_row, axis=1)

    # keep useful fields for ML training
    keep_cols = [key, 'Data de criação', 'Ano', 'Mes', 'Etapa do negócio', 'Fonte original do tráfego',
                 'Detalhamento da fonte original do tráfego 1', 'Detalhamento da fonte original do tráfego 2',
                 'Unidade Desejada', 'Número de atividades de vendas', 'Pipeline', 'Valor na moeda da empresa',
                 'Proprietário do negócio', 'is_won', 'is_lost', 'label']

    cols_exist = [c for c in keep_cols if c in leads.columns]
    out = leads[cols_exist].copy()

    # Save
    out.to_csv(OUT_CSV, index=False)
    print(f"[ML PREP] Training base saved: {OUT_CSV}")

    # Summary counts
    total = len(out)
    n_won = int(out['is_won'].sum()) if 'is_won' in out.columns else 0
    n_lost = int(out['is_lost'].sum()) if 'is_lost' in out.columns else 0
    n_label = int(out['label'].notna().sum())
    print('--- SUMMARY ---')
    print(f'  total rows: {total:,}')
    print(f'  winners (is_won): {n_won:,}')
    print(f'  lost (is_lost): {n_lost:,}')
    print(f'  labeled rows (won or lost): {n_label:,}')

    return OUT_CSV

if __name__ == '__main__':
    prepare()
