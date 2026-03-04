import os
import pandas as pd
from datetime import datetime

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
IN_PATH = os.path.join(BASE, 'Data', 'hubspot_leads.csv')
OUT = os.path.join(BASE, 'Outputs')
os.makedirs(OUT, exist_ok=True)

DATA_HOJE = datetime.now().strftime('%Y-%m-%d')
OUT_CSV = os.path.join(OUT, f'looker_dataset_{DATA_HOJE}.csv')

def read_data(path):
    try:
        df = pd.read_csv(path, encoding='utf-8')
    except Exception:
        df = pd.read_csv(path, encoding='latin1')
    df.columns = df.columns.str.strip()
    # robust date parsing
    raw_dates = df['Data de criação'].astype(str)
    parsed_default = pd.to_datetime(raw_dates, errors='coerce')
    parsed_dayfirst = pd.to_datetime(raw_dates, dayfirst=True, errors='coerce')
    df['Data de criação'] = parsed_default.fillna(parsed_dayfirst)
    df['Ano'] = df['Data de criação'].dt.year
    df['Mes'] = df['Data de criação'].dt.month
    df['Matricula'] = df['Etapa do negócio'].str.contains('MATRÍCULA', case=False, na=False).astype(int)
    df['Fonte original do tráfego'] = df['Fonte original do tráfego'].fillna('(Sem fonte - Nulo)')
    df['Detalhamento da fonte original do tráfego 1'] = df['Detalhamento da fonte original do tráfego 1'].fillna('(Sem detalhamento)')
    return df

def plano_bala_score_simple(df):
    # Use same heuristic as report script: channel conv norm, stage multiplier, recency
    ch_rate = df.groupby('Fonte original do tráfego').agg(Leads=('Fonte original do tráfego', 'count'), Matriculas=('Matricula', 'sum')).reset_index()
    ch_rate['Conv_Rate'] = ch_rate['Matriculas'] / ch_rate['Leads']
    minr = ch_rate['Conv_Rate'].min()
    maxr = ch_rate['Conv_Rate'].max()
    ch_rate['Conv_Norm'] = (ch_rate['Conv_Rate'] - minr) / (maxr - minr) if maxr > minr else 0.5
    df = df.merge(ch_rate[['Fonte original do tráfego', 'Conv_Norm']], on='Fonte original do tráfego', how='left')
    years = sorted(df['Ano'].dropna().unique())
    ymin, ymax = (min(years), max(years)) if years else (None, None)
    if ymin is None:
        df['Recency_W'] = 1.0
    else:
        df['Recency_W'] = df['Ano'].fillna(ymin).apply(lambda y: 0.5 + 0.5 * ((y - ymin) / (ymax - ymin)) if ymax and ymax > ymin else 1.0)
    def stage_mult(s):
        if pd.isna(s):
            return 1.0
        ss = s.upper()
        if 'QUALIF' in ss:
            return 1.25
        if 'NOVO' in ss or 'NEGOCIO' in ss or 'NEGÓCIO' in ss:
            return 1.1
        return 1.0
    df['Stage_Mult'] = df['Etapa do negócio'].apply(stage_mult)
    df['Plano_Bala_Score_raw'] = df['Conv_Norm'] * df['Recency_W'] * df['Stage_Mult']
    rmin = df['Plano_Bala_Score_raw'].min()
    rmax = df['Plano_Bala_Score_raw'].max()
    if rmax > rmin:
        df['Plano_Bala_Score'] = ((df['Plano_Bala_Score_raw'] - rmin) / (rmax - rmin)) * 100
    else:
        df['Plano_Bala_Score'] = df['Plano_Bala_Score_raw'] * 100
    return df

def create_looker_dataset(df, out_csv):
    # Select and normalize columns useful for Looker
    cols = []
    # try to find ID column
    id_col = None
    for c in df.columns:
        if c.strip() in ['#', 'Id', 'ID'] or 'id' in c.lower() or 'hs_object' in c.lower():
            id_col = c
            break
    if id_col is None:
        df['lead_id'] = range(1, len(df)+1)
        id_col = 'lead_id'

    # compute scores on full df (function expects original column names)
    scored = plano_bala_score_simple(df)

    out = df[[id_col, 'Data de criação', 'Ano', 'Mes', 'Fonte original do tráfego', 'Detalhamento da fonte original do tráfego 1', 'Unidade Desejada', 'Etapa do negócio', 'Matricula']].copy()
    out = out.rename(columns={id_col: 'lead_id', 'Fonte original do tráfego': 'channel', 'Detalhamento da fonte original do tráfego 1': 'channel_detail', 'Unidade Desejada': 'unit', 'Etapa do negócio': 'stage'})

    # merge the score from scored (match on position/index or id if exists)
    # scored keeps original index; prefer to match on a detected id column
    if id_col in scored.columns:
        scored_id_col = id_col
    else:
        scored = scored.reset_index().rename(columns={'index':'lead_idx'})
        scored_id_col = None

    if scored_id_col is not None and scored_id_col in out.columns:
        merged = out.merge(scored[[scored_id_col, 'Plano_Bala_Score']], left_on='lead_id', right_on=scored_id_col, how='left')
        merged = merged.drop(columns=[scored_id_col], errors='ignore')
    else:
        # fallback: align by row order
        merged = out.copy()
        merged['Plano_Bala_Score'] = scored['Plano_Bala_Score'].values if 'Plano_Bala_Score' in scored.columns else None

    merged.to_csv(out_csv, index=False)
    print('Dataset para Looker salvo em:', out_csv)

def main():
    df = read_data(IN_PATH)
    create_looker_dataset(df, OUT_CSV)

if __name__ == '__main__':
    main()
