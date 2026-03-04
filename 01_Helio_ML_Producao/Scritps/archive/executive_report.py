import os
import pandas as pd
import numpy as np
from datetime import datetime

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
IN_PATH = os.path.join(BASE, 'Data', 'hubspot_leads.csv')
OUT = os.path.join(BASE, 'Outputs')
os.makedirs(OUT, exist_ok=True)

DATE = datetime.now().strftime('%Y-%m-%d')
OUT_XLSX = os.path.join(OUT, f'Executive_Marketing_Report_{DATE}.xlsx')
OUT_CSV = os.path.join(OUT, f'Executive_Looker_{DATE}.csv')


def read_data(path):
    try:
        df = pd.read_csv(path, encoding='utf-8')
    except Exception:
        df = pd.read_csv(path, encoding='latin1')
    df.columns = df.columns.str.strip()
    # robust date parsing
    if 'Data de criação' in df.columns:
        raw_dates = df['Data de criação'].astype(str)
        parsed_default = pd.to_datetime(raw_dates, errors='coerce')
        parsed_dayfirst = pd.to_datetime(raw_dates, dayfirst=True, errors='coerce')
        df['Data de criação'] = parsed_default.fillna(parsed_dayfirst)
        df['Ano'] = df['Data de criação'].dt.year
        df['Mes'] = df['Data de criação'].dt.month
    else:
        df['Ano'] = np.nan
        df['Mes'] = np.nan

    # simple flags
    if 'Etapa do negócio' in df.columns:
        df['Matricula'] = df['Etapa do negócio'].str.contains('MATRÍCULA', case=False, na=False).astype(int)
    else:
        df['Matricula'] = 0

    # safe fill
    if 'Fonte original do tráfego' in df.columns:
        df['Fonte original do tráfego'] = df['Fonte original do tráfego'].fillna('(Sem fonte)')
    else:
        df['Fonte original do tráfego'] = '(Sem fonte)'

    if 'Detalhamento da fonte original do tráfego 1' not in df.columns:
        df['Detalhamento da fonte original do tráfego 1'] = '(Sem detalhamento)'

    return df


def resumo_anos(df):
    resumo = df.groupby('Ano', dropna=True).agg(
        Total_Leads=('Ano', 'count'),
        Matriculas=('Matricula', 'sum')
    ).reset_index().sort_values('Ano')
    resumo['Taxa_Conversao'] = resumo['Matriculas'] / resumo['Total_Leads']
    return resumo


def channel_performance(df):
    ch = df.groupby(['Fonte original do tráfego', 'Ano'], dropna=True).agg(
        Leads=('Fonte original do tráfego', 'count'),
        Matriculas=('Matricula', 'sum')
    ).reset_index()
    ch['Taxa_Conv'] = ch['Matriculas'] / ch['Leads']
    pivot_leads = ch.pivot(index='Fonte original do tráfego', columns='Ano', values='Leads').fillna(0).astype(int)
    pivot_conv = ch.pivot(index='Fonte original do tráfego', columns='Ano', values='Taxa_Conv').fillna(0)
    return ch, pivot_leads, pivot_conv


def units_2025(df):
    if 'Unidade Desejada' not in df.columns:
        return pd.DataFrame()
    df25 = df[df['Ano'] == 2025]
    unid = df25.groupby('Unidade Desejada').agg(
        Leads=('Unidade Desejada', 'count'),
        Matriculas=('Matricula', 'sum')
    ).reset_index()
    unid['Taxa_Conversao'] = unid['Matriculas'] / unid['Leads'].replace(0, np.nan)
    unid = unid.sort_values('Leads', ascending=False)
    return unid


def plano_bala(df):
    # channel conv
    ch = df.groupby('Fonte original do tráfego').agg(Leads=('Fonte original do tráfego', 'count'), Matriculas=('Matricula', 'sum')).reset_index()
    ch['Conv_Rate'] = ch['Matriculas'] / ch['Leads']
    minr, maxr = ch['Conv_Rate'].min(), ch['Conv_Rate'].max()
    if maxr > minr:
        ch['Conv_Norm'] = (ch['Conv_Rate'] - minr) / (maxr - minr)
    else:
        ch['Conv_Norm'] = 0.5

    df = df.merge(ch[['Fonte original do tráfego', 'Conv_Norm']], on='Fonte original do tráfego', how='left')

    years = sorted(df['Ano'].dropna().unique())
    if years:
        ymin, ymax = years[0], years[-1]
        if ymax > ymin:
            df['Recency_W'] = df['Ano'].fillna(ymin).apply(lambda y: 0.5 + 0.5 * ((y - ymin) / (ymax - ymin)))
        else:
            df['Recency_W'] = 1.0
    else:
        df['Recency_W'] = 1.0

    def stage_mult(s):
        if pd.isna(s):
            return 1.0
        ss = str(s).upper()
        if 'QUALIF' in ss:
            return 1.25
        if 'NOVO' in ss or 'NEGOCIO' in ss or 'NEGÓCIO' in ss:
            return 1.1
        return 1.0

    if 'Etapa do negócio' in df.columns:
        df['Stage_Mult'] = df['Etapa do negócio'].apply(stage_mult)
    else:
        df['Stage_Mult'] = 1.0

    df['Plano_Bala_raw'] = df['Conv_Norm'].fillna(0) * df['Recency_W'] * df['Stage_Mult']
    rmin, rmax = df['Plano_Bala_raw'].min(), df['Plano_Bala_raw'].max()
    if rmax > rmin:
        df['Plano_Bala_Score'] = ((df['Plano_Bala_raw'] - rmin) / (rmax - rmin)) * 100
    else:
        df['Plano_Bala_Score'] = df['Plano_Bala_raw'] * 100

    # top recommendations among target stages
    terms = ['NOVO', 'QUALIF', 'NEGOCIO', 'NEGÓCIO']
    def is_target(s):
        if pd.isna(s):
            return False
        ss = str(s).upper()
        return any(t in ss for t in terms)

    df['Is_Target'] = df['Etapa do negócio'].apply(is_target) if 'Etapa do negócio' in df.columns else False

    recos = df[df['Is_Target']].copy()
    # ensure id
    id_col = None
    for c in recos.columns:
        if c.strip() in ['#', 'Id', 'ID'] or 'id' in c.lower() or 'hs_object' in c.lower():
            id_col = c
            break
    if id_col is None:
        recos = recos.reset_index(drop=True)
        recos['lead_id'] = range(1, len(recos) + 1)
        id_col = 'lead_id'

    cols_keep = [id_col, 'Data de criação', 'Ano', 'Fonte original do tráfego', 'Detalhamento da fonte original do tráfego 1', 'Unidade Desejada', 'Etapa do negócio', 'Matricula', 'Plano_Bala_Score']
    cols_exist = [c for c in cols_keep if c in recos.columns]
    recos = recos[cols_exist].rename(columns={id_col: 'ID'})
    top = recos.sort_values('Plano_Bala_Score', ascending=False).head(250)
    return df, ch, top


def write_outputs(df, resumo, pivot_leads, pivot_conv, units25, top_reco, out_xlsx, out_csv):
    # Excel
    with pd.ExcelWriter(out_xlsx, engine='openpyxl') as w:
        resumo.to_excel(w, sheet_name='1_Resumo_Geral', index=False)
        pivot_leads.to_excel(w, sheet_name='2_Channel_Leads')
        pivot_conv.to_excel(w, sheet_name='3_Channel_Conv')
        if not units25.empty:
            units25.to_excel(w, sheet_name='4_Units_2025', index=False)
        top_reco.to_excel(w, sheet_name='5_Plano_Bala_Top', index=False)

    # Looker CSV (lead-level)
    id_col = None
    for c in df.columns:
        if c.strip() in ['#', 'Id', 'ID'] or 'id' in c.lower() or 'hs_object' in c.lower():
            id_col = c
            break
    out = df.copy()
    if id_col is None:
        out = out.reset_index(drop=True)
        out['lead_id'] = range(1, len(out) + 1)
        id_col = 'lead_id'

    fields = [id_col, 'Data de criação', 'Ano', 'Mes', 'Fonte original do tráfego', 'Detalhamento da fonte original do tráfego 1', 'Unidade Desejada', 'Etapa do negócio', 'Matricula', 'Plano_Bala_Score']
    fields = [f for f in fields if f in out.columns]
    out[fields].to_csv(out_csv, index=False)


def main():
    print('[EXECUTIVE] Reading data...')
    df = read_data(IN_PATH)
    print(f'  - Leads loaded: {len(df):,}')

    resumo = resumo_anos(df)
    ch_df, pivot_leads, pivot_conv = channel_performance(df)
    units25 = units_2025(df)
    scored_df, ch_rate, top_reco = plano_bala(df)

    # merge score into df if exists
    if 'Plano_Bala_Score' not in df.columns and 'Plano_Bala_Score' in scored_df.columns:
        df['Plano_Bala_Score'] = scored_df['Plano_Bala_Score'].values

    print('[EXECUTIVE] Writing outputs...')
    write_outputs(df, resumo, pivot_leads, pivot_conv, units25, top_reco, OUT_XLSX, OUT_CSV)
    print(f"[DONE] Excel: {OUT_XLSX}")
    print(f"[DONE] CSV:   {OUT_CSV}")


if __name__ == '__main__':
    main()
