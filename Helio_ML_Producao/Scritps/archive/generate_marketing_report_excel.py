import os
import pandas as pd
import numpy as np
from datetime import datetime

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
IN_PATH = os.path.join(BASE, 'Data', 'hubspot_leads.csv')
OUT = os.path.join(BASE, 'Outputs')
os.makedirs(OUT, exist_ok=True)

DATA_HOJE = datetime.now().strftime('%Y-%m-%d')
OUT_XLSX = os.path.join(OUT, f'Marketing_Report_{DATA_HOJE}.xlsx')

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

def channel_performance(df):
    # Per channel per year: count leads and matriculas
    ch = df.groupby(['Fonte original do tráfego', 'Ano']).agg(
        Total_Leads=('Fonte original do tráfego', 'count'),
        Matriculas=('Matricula', 'sum')
    ).reset_index()
    ch['Taxa_Conv'] = ch['Matriculas'] / ch['Total_Leads']
    pivot = ch.pivot(index='Fonte original do tráfego', columns='Ano', values='Total_Leads').fillna(0).astype(int)
    conv_pivot = ch.pivot(index='Fonte original do tráfego', columns='Ano', values='Taxa_Conv').fillna(0)
    return ch, pivot, conv_pivot

def units_overview(df):
    units = df.groupby(['Unidade Desejada', 'Ano']).agg(
        Leads=('Unidade Desejada', 'count'),
        Matriculas=('Matricula', 'sum')
    ).reset_index()
    units_pivot = units.pivot(index='Unidade Desejada', columns='Ano', values='Leads').fillna(0).astype(int)
    # detect YoY decrease for the most recent year available
    years = sorted([y for y in df['Ano'].dropna().unique()])
    flag = []
    if len(years) >= 2:
        y_latest = years[-1]
        y_prev = years[-2]
        summary = units.groupby('Unidade Desejada').sum().reset_index()
        for u in units_pivot.index:
            latest = units_pivot.loc[u].get(y_latest, 0)
            prev = units_pivot.loc[u].get(y_prev, 0)
            pct_change = (latest - prev) / prev if prev != 0 else np.nan
            flag.append((u, prev, latest, pct_change))
        flags_df = pd.DataFrame(flag, columns=['Unidade Desejada', f'Leads_{y_prev}', f'Leads_{y_latest}', 'Pct_Change'])
        flags_df['Flag_Decrease'] = flags_df['Pct_Change'] < 0
    else:
        flags_df = pd.DataFrame(columns=['Unidade Desejada'])
    return units, units_pivot, flags_df

def plano_bala_score(df):
    # Compute historical channel conversion rates
    ch_rate = df.groupby('Fonte original do tráfego').agg(
        Leads=('Fonte original do tráfego', 'count'),
        Matriculas=('Matricula', 'sum')
    ).reset_index()
    ch_rate['Conv_Rate'] = ch_rate['Matriculas'] / ch_rate['Leads']
    # normalize conv rate to 0-1
    minr = ch_rate['Conv_Rate'].min()
    maxr = ch_rate['Conv_Rate'].max()
    ch_rate['Conv_Norm'] = (ch_rate['Conv_Rate'] - minr) / (maxr - minr) if maxr > minr else 0.5

    # target stages: novos negocios & em qualificacao heuristics
    stage_terms = ['NOVO', 'NOVOS', 'NEGÓCIO', 'NEGOCIO', 'QUALIF', 'QUALIFICA']
    def is_target_stage(s):
        if pd.isna(s):
            return False
        ss = s.upper()
        return any(t in ss for t in stage_terms)

    df['Is_Target_Stage'] = df['Etapa do negócio'].apply(is_target_stage)

    # Merge conv norm into df
    df = df.merge(ch_rate[['Fonte original do tráfego', 'Conv_Norm', 'Conv_Rate']], on='Fonte original do tráfego', how='left')

    # recency weight: newer year gets higher weight
    years = sorted(df['Ano'].dropna().unique())
    if len(years) == 0:
        df['Recency_W'] = 1.0
    else:
        ymin, ymax = min(years), max(years)
        df['Recency_W'] = df['Ano'].fillna(ymin).apply(lambda y: 0.5 + 0.5 * ((y - ymin) / (ymax - ymin)) if ymax > ymin else 1.0)

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

    # Base score
    df['Plano_Bala_Score_raw'] = df['Conv_Norm'] * df['Recency_W'] * df['Stage_Mult']
    # normalize to 0-100
    rmin = df['Plano_Bala_Score_raw'].min()
    rmax = df['Plano_Bala_Score_raw'].max()
    if rmax > rmin:
        df['Plano_Bala_Score'] = ((df['Plano_Bala_Score_raw'] - rmin) / (rmax - rmin)) * 100
    else:
        df['Plano_Bala_Score'] = df['Plano_Bala_Score_raw'] * 100

    # Keep relevant columns for recommendations; ensure we have an ID column
    recommendations = df[df['Is_Target_Stage']].copy()
    # detect id column
    id_col = None
    for c in df.columns:
        if c.strip() in ['#', 'Id', 'ID'] or 'id' in c.lower() or 'hs_object' in c.lower():
            id_col = c
            break
    if id_col is None:
        recommendations = recommendations.reset_index(drop=True)
        recommendations['lead_id'] = range(1, len(recommendations) + 1)
        id_col = 'lead_id'

    cols_req = [id_col, 'Data de criação', 'Ano', 'Fonte original do tráfego', 'Detalhamento da fonte original do tráfego 1', 'Unidade Desejada', 'Etapa do negócio', 'Matricula', 'Plano_Bala_Score']
    # keep only columns that exist to avoid KeyError
    cols_exist = [c for c in cols_req if c in recommendations.columns]
    recommendations = recommendations[cols_exist].copy()
    # rename ID column to standardized name
    recommendations = recommendations.rename(columns={id_col: 'ID'})

    # Top recommendations
    if 'Plano_Bala_Score' in recommendations.columns:
        top_reco = recommendations.sort_values('Plano_Bala_Score', ascending=False).head(500)
    else:
        top_reco = recommendations.head(500)
    return df, ch_rate, top_reco

def write_report(df, ch_pivot, ch_conv, units_pivot, flags_df, top_reco, out_path):
    with pd.ExcelWriter(out_path, engine='openpyxl') as x:
        # Resumo geral por ano
        resumo = df.groupby('Ano').agg(Total_Leads=('Ano', 'count'), Matriculas=('Matricula', 'sum')).reset_index()
        resumo['Taxa_Conversao'] = resumo['Matriculas'] / resumo['Total_Leads']
        resumo.to_excel(x, sheet_name='1_Resumo_Geral', index=False)

        ch_pivot.to_excel(x, sheet_name='2_Channel_Leads')
        ch_conv.to_excel(x, sheet_name='3_Channel_Conv')

        units_pivot.to_excel(x, sheet_name='4_Units_Leads')
        flags_df.to_excel(x, sheet_name='5_Units_Flags', index=False)

        top_reco.to_excel(x, sheet_name='6_Plano_Bala_Recom', index=False)

        # Add a sheet with instructions / goals
        goals = pd.DataFrame([
            ['Objetivo baseline', 'Aumentar performance em 10% dos canais (baseline)'],
            ['Objetivo otimizado', 'Meta otimizada por canal (ex.: +25%)'],
            ['Moonshot', 'Meta ambiciosa (ex.: +100% em canais prioritários)']
        ], columns=['Nome', 'Descrição'])
        goals.to_excel(x, sheet_name='7_Goals', index=False)

    print('Relatório salvo em:', out_path)

def main():
    df = read_data(IN_PATH)
    ch_df, ch_pivot, ch_conv = channel_performance(df)
    units_df, units_pivot, flags_df = units_overview(df)
    scored_df, ch_rate, top_reco = plano_bala_score(df)
    write_report(df, ch_pivot, ch_conv, units_pivot, flags_df, top_reco, OUT_XLSX)

if __name__ == '__main__':
    main()
