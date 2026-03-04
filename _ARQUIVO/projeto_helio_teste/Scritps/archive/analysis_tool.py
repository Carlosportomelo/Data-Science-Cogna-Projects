"""
Consolidated analysis tool

Usage:
  python analysis_tool.py report    # generate Excel marketing report
  python analysis_tool.py looker    # generate CSV for Looker
  python analysis_tool.py diagnose  # run quick diagnostics
  python analysis_tool.py all       # run report + looker

This file consolidates the key workflows into one script. Historical scripts
were copied into `Scritps/archive/` to keep history.
"""
import argparse
import os
import pandas as pd
import numpy as np
from datetime import datetime

BASE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(BASE)
IN_PATH = os.path.join(ROOT, 'Data', 'hubspot_leads.csv')
OUT = os.path.join(ROOT, 'Outputs')
os.makedirs(OUT, exist_ok=True)

DATA_HOJE = datetime.now().strftime('%Y-%m-%d')

def read_data(path):
    try:
        df = pd.read_csv(path, encoding='utf-8')
    except Exception:
        df = pd.read_csv(path, encoding='latin1')
    df.columns = df.columns.str.strip()
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

def generate_marketing_report(df, out_path=None):
    if out_path is None:
        out_path = os.path.join(OUT, f'Marketing_Report_{DATA_HOJE}.xlsx')

    # Channel pivots
    ch = df.groupby(['Fonte original do tráfego', 'Ano']).agg(Total_Leads=('Fonte original do tráfego','count'), Matriculas=('Matricula','sum')).reset_index()
    ch_pivot = ch.pivot(index='Fonte original do tráfego', columns='Ano', values='Total_Leads').fillna(0).astype(int)
    ch_conv = ch.pivot(index='Fonte original do tráfego', columns='Ano', values='Matriculas').fillna(0)

    # Units
    units = df.groupby(['Unidade Desejada','Ano']).agg(Leads=('Unidade Desejada','count'), Matriculas=('Matricula','sum')).reset_index()
    units_pivot = units.pivot(index='Unidade Desejada', columns='Ano', values='Leads').fillna(0).astype(int)
    years = sorted([y for y in df['Ano'].dropna().unique()])
    flags_df = pd.DataFrame()
    if len(years) >= 2:
        y_latest = years[-1]; y_prev = years[-2]
        flags = []
        for u in units_pivot.index:
            latest = units_pivot.loc[u].get(y_latest,0)
            prev = units_pivot.loc[u].get(y_prev,0)
            pct = (latest - prev)/prev if prev!=0 else np.nan
            flags.append((u, prev, latest, pct))
        flags_df = pd.DataFrame(flags, columns=['Unidade Desejada', f'Leads_{y_prev}', f'Leads_{y_latest}', 'Pct_Change'])
        flags_df['Flag_Decrease'] = flags_df['Pct_Change'] < 0

    # Score (Plano Bala de Prata - simplified)
    ch_rate = df.groupby('Fonte original do tráfego').agg(Leads=('Fonte original do tráfego','count'), Matriculas=('Matricula','sum')).reset_index()
    ch_rate['Conv_Rate'] = ch_rate['Matriculas'] / ch_rate['Leads']
    minr = ch_rate['Conv_Rate'].min(); maxr = ch_rate['Conv_Rate'].max()
    ch_rate['Conv_Norm'] = (ch_rate['Conv_Rate'] - minr) / (maxr - minr) if maxr>minr else 0.5
    df2 = df.merge(ch_rate[['Fonte original do tráfego','Conv_Norm']], on='Fonte original do tráfego', how='left')
    ymin = min(years) if years else None
    ymax = max(years) if years else None
    if ymin is None:
        df2['Recency_W']=1.0
    else:
        df2['Recency_W'] = df2['Ano'].fillna(ymin).apply(lambda y: 0.5+0.5*((y-ymin)/(ymax-ymin)) if ymax>ymin else 1.0)
    def stage_mult(s):
        if pd.isna(s): return 1.0
        s = s.upper()
        if 'QUALIF' in s: return 1.25
        if 'NOVO' in s or 'NEGOCIO' in s or 'NEGÓCIO' in s: return 1.1
        return 1.0
    df2['Stage_Mult'] = df2['Etapa do negócio'].apply(stage_mult)
    df2['Plano_Bala_Score_raw'] = df2['Conv_Norm'] * df2['Recency_W'] * df2['Stage_Mult']
    rmin = df2['Plano_Bala_Score_raw'].min(); rmax = df2['Plano_Bala_Score_raw'].max()
    df2['Plano_Bala_Score'] = ((df2['Plano_Bala_Score_raw'] - rmin)/(rmax-rmin))*100 if rmax>rmin else df2['Plano_Bala_Score_raw']*100

    recos = df2[df2['Etapa do negócio'].astype(str).str.upper().str.contains('NOVO|QUALIF|NEGOCIO|NEGÓCIO')].copy()
    # Try to standardize ID column
    id_col = next((c for c in recos.columns if c.strip() in ['#','Id','ID'] or 'id' in c.lower() or 'hs_object' in c.lower()), None)
    if id_col is None:
        recos = recos.reset_index(drop=True); recos['ID']=range(1,len(recos)+1)
    else:
        recos = recos.rename(columns={id_col:'ID'})

    top_reco = recos.sort_values('Plano_Bala_Score', ascending=False).head(500)

    # write Excel
    with pd.ExcelWriter(out_path, engine='openpyxl') as w:
        resumo = df.groupby('Ano').agg(Total_Leads=('Ano','count'), Matriculas=('Matricula','sum')).reset_index()
        resumo['Taxa_Conversao'] = resumo['Matriculas']/resumo['Total_Leads']
        resumo.to_excel(w, sheet_name='1_Resumo_Geral', index=False)
        ch_pivot.to_excel(w, sheet_name='2_Channel_Leads')
        ch_conv.to_excel(w, sheet_name='3_Channel_Conv')
        units_pivot.to_excel(w, sheet_name='4_Units_Leads')
        flags_df.to_excel(w, sheet_name='5_Units_Flags', index=False)
        top_reco.to_excel(w, sheet_name='6_Plano_Bala_Recom', index=False)
        pd.DataFrame([['Objetivo baseline', 'Aumentar performance em 10% dos canais (baseline)'], ['Objetivo otimizado', 'Meta otimizada por canal (ex.: +25%)'], ['Moonshot','Meta ambiciosa (+100%)']], columns=['Nome','Descrição']).to_excel(w, sheet_name='7_Goals', index=False)

    print('Relatório salvo em:', out_path)
    return out_path

def generate_looker_csv(df, out_csv=None):
    if out_csv is None:
        out_csv = os.path.join(OUT, f'looker_dataset_{DATA_HOJE}.csv')
    # compute score as in report
    df2 = df.copy()
    ch_rate = df2.groupby('Fonte original do tráfego').agg(Leads=('Fonte original do tráfego','count'), Matriculas=('Matricula','sum')).reset_index()
    ch_rate['Conv_Rate'] = ch_rate['Matriculas']/ch_rate['Leads']
    minr = ch_rate['Conv_Rate'].min(); maxr = ch_rate['Conv_Rate'].max()
    ch_rate['Conv_Norm'] = (ch_rate['Conv_Rate']-minr)/(maxr-minr) if maxr>minr else 0.5
    df2 = df2.merge(ch_rate[['Fonte original do tráfego','Conv_Norm']], on='Fonte original do tráfego', how='left')
    years = sorted([y for y in df2['Ano'].dropna().unique()])
    ymin = min(years) if years else None; ymax = max(years) if years else None
    if ymin is None:
        df2['Recency_W'] = 1.0
    else:
        df2['Recency_W'] = df2['Ano'].fillna(ymin).apply(lambda y: 0.5+0.5*((y-ymin)/(ymax-ymin)) if ymax>ymin else 1.0)
    def stage_mult(s):
        if pd.isna(s): return 1.0
        ss = s.upper()
        if 'QUALIF' in ss: return 1.25
        if 'NOVO' in ss or 'NEGOCIO' in ss or 'NEGÓCIO' in ss: return 1.1
        return 1.0
    df2['Stage_Mult'] = df2['Etapa do negócio'].apply(stage_mult)
    df2['Plano_Bala_Score_raw'] = df2['Conv_Norm'] * df2['Recency_W'] * df2['Stage_Mult']
    rmin = df2['Plano_Bala_Score_raw'].min(); rmax = df2['Plano_Bala_Score_raw'].max()
    df2['Plano_Bala_Score'] = ((df2['Plano_Bala_Score_raw']-rmin)/(rmax-rmin))*100 if rmax>rmin else df2['Plano_Bala_Score_raw']*100

    # detect id
    id_col = next((c for c in df2.columns if c.strip() in ['#','Id','ID'] or 'id' in c.lower() or 'hs_object' in c.lower()), None)
    out = df2[[id_col if id_col else None, 'Data de criação', 'Ano', 'Mes', 'Fonte original do tráfego', 'Detalhamento da fonte original do tráfego 1', 'Unidade Desejada', 'Etapa do negócio', 'Matricula', 'Plano_Bala_Score']].copy()
    if id_col is None:
        out = out.reset_index().rename(columns={'index':'lead_id'})
    else:
        out = out.rename(columns={id_col:'lead_id'})
    out.to_csv(out_csv, index=False)
    print('Dataset para Looker salvo em:', out_csv)
    return out_csv

def diagnose(path=IN_PATH):
    # quick counts
    total_lines = 0
    with open(path, 'r', encoding='utf-8', errors='replace') as f:
        for _ in f: total_lines += 1
    df = read_data(path)
    print('Linhas no arquivo (incl header):', total_lines)
    print('Linhas lidas pelo pandas:', len(df))
    print('Anos detectados:', sorted(df['Ano'].dropna().unique()))
    return True

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('task', choices=['report','looker','diagnose','all'], help='Task to run')
    args = parser.parse_args()
    df = read_data(IN_PATH)
    if args.task == 'report':
        p = generate_marketing_report(df)
        print('Report saved:', p)
    elif args.task == 'looker':
        p = generate_looker_csv(df)
        print('Looker CSV saved:', p)
    elif args.task == 'diagnose':
        diagnose()
    elif args.task == 'all':
        p = generate_marketing_report(df)
        q = generate_looker_csv(df)
        print('Done:', p, q)

if __name__ == '__main__':
    main()
