"""
==============================================================================
RELATÓRIO EXECUTIVO - PERFORMANCE DE CANAIS (VERSÃO CORRIGIDA)
==============================================================================
Script: 8_Relatorio_Executivo_Performance_Canais_v2.py
Correção: Classificação de canais baseada na FONTE MACRO (não no detalhe)
Data: 2025-12-11
==============================================================================
"""

import pandas as pd
import os
from datetime import datetime

# --- CONFIGURAÇÃO ---
CAMINHO_BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CAMINHO_ENTRADA = os.path.join(CAMINHO_BASE, 'Data', 'hubspot_leads.csv')
PASTA_OUTPUT = os.path.join(CAMINHO_BASE, 'Outputs')
os.makedirs(PASTA_OUTPUT, exist_ok=True)

DATA_HOJE = datetime.now().strftime('%Y-%m-%d')
CAMINHO_SAIDA = os.path.join(PASTA_OUTPUT, f'Board_Report_Performance_Canais_{DATA_HOJE}_v2.xlsx')

print(f"[INIT] Gerando Relatório Executivo (VERSÃO CORRIGIDA)")
print(f"[INFO] Entrada: {CAMINHO_ENTRADA}")

try:
    # 1. CARREGAMENTO
    try:
        df = pd.read_csv(CAMINHO_ENTRADA, sep=',', encoding='utf-8', on_bad_lines='skip')
    except:
        df = pd.read_csv(CAMINHO_ENTRADA, sep=',', encoding='latin1', on_bad_lines='skip')

    df.columns = df.columns.str.strip()
    
    # Renomear colunas para facilitar uso
    mapa = {
        'Record ID': 'ID_Lead',
        'Data de criação': 'Data_Entrada',
        'Fonte original do tráfego': 'Fonte_Macro', 
        'Detalhamento da fonte original do tráfego 1': 'Detalhe_Tecnico', 
        'Etapa do negócio': 'Status_Atual'
    }
    df = df.rename(columns=mapa)
    df['Data_Entrada'] = pd.to_datetime(df['Data_Entrada'], errors='coerce')

    # 2. CLASSIFICAÇÃO CORRIGIDA
    # IMPORTANTE: A fonte macro já está bem categorizada pelo HubSpot!
    # Não precisamos olhar o detalhe para a maioria dos casos
    
    def classificar_canal_corrigido(row):
        """
        Classificação baseada na FONTE MACRO do HubSpot
        que já vem bem estruturada
        """
        fonte = str(row['Fonte_Macro']).lower().strip()
        
        # Mapeamento direto da fonte macro
        if 'off-line' in fonte or 'offline' in fonte:
            return 'Offline (Unidade/Balcão)'
        elif 'social pago' in fonte or 'paid social' in fonte:
            return 'Meta Ads (Social Pago)'
        elif 'pesquisa paga' in fonte or 'paid search' in fonte:
            return 'Google Ads (Pesquisa Paga)'
        elif 'pesquisa orgânica' in fonte or 'organic search' in fonte:
            return 'Google Orgânico (SEO)'
        elif 'social orgânico' in fonte or 'organic social' in fonte:
            return 'Social Orgânico'
        elif 'tráfego direto' in fonte or 'direct traffic' in fonte:
            return 'Tráfego Direto'
        elif 'referências' in fonte or 'referrals' in fonte:
            return 'Referências'
        elif 'e-mail' in fonte or 'email' in fonte:
            return 'E-mail Marketing'
        else:
            return 'Outros'

    df['Canal_Estrategico'] = df.apply(classificar_canal_corrigido, axis=1)
    df['Matricula_Realizada'] = df['Status_Atual'].astype(str).str.contains('MATRÍCULA', case=False).astype(int)

    print(f"[OK] Dados carregados: {len(df):,} leads")
    
    # Verificação de sanidade
    print("\n[CHECK] Distribuição por Canal:")
    print(df['Canal_Estrategico'].value_counts())

    # 3. GERAÇÃO DOS KPIs

    # KPI 1: Overview de Performance 2025
    df_2025 = df[df['Data_Entrada'].dt.year == 2025].copy()
    
    kpi_performance = df_2025.groupby('Canal_Estrategico').agg(
        Total_Leads=('ID_Lead', 'count'),
        Total_Matriculas=('Matricula_Realizada', 'sum')
    ).reset_index()
    
    # Taxas em formato decimal (Excel formata depois)
    kpi_performance['Taxa_Conversao'] = kpi_performance['Total_Matriculas'] / kpi_performance['Total_Leads']
    kpi_performance['Share_Leads'] = kpi_performance['Total_Leads'] / kpi_performance['Total_Leads'].sum()
    kpi_performance['Share_Matriculas'] = kpi_performance['Total_Matriculas'] / kpi_performance['Total_Matriculas'].sum()
    kpi_performance = kpi_performance.sort_values('Total_Matriculas', ascending=False)

    # KPI 2: YOY Analysis (Jan-Nov para comparação justa)
    mes_corte = df_2025['Data_Entrada'].max().month
    
    df_yoy = df[
        (df['Data_Entrada'].dt.month <= mes_corte) & 
        (df['Data_Entrada'].dt.year.isin([2024, 2025]))
    ].copy()
    
    df_yoy['Ano'] = df_yoy['Data_Entrada'].dt.year
    
    # Leads por canal/ano
    leads_yoy = df_yoy.pivot_table(
        index='Canal_Estrategico', 
        columns='Ano', 
        values='ID_Lead', 
        aggfunc='count',
        fill_value=0
    ).reset_index()
    
    # Matrículas por canal/ano
    mat_yoy = df_yoy.pivot_table(
        index='Canal_Estrategico', 
        columns='Ano', 
        values='Matricula_Realizada', 
        aggfunc='sum',
        fill_value=0
    ).reset_index()
    
    kpi_yoy = leads_yoy.copy()
    kpi_yoy = kpi_yoy.rename(columns={2024: 'Leads_2024', 2025: 'Leads_2025'})
    kpi_yoy['Mat_2024'] = mat_yoy[2024]
    kpi_yoy['Mat_2025'] = mat_yoy[2025]
    
    # Calcular variações
    kpi_yoy['Var_Leads_%'] = (kpi_yoy['Leads_2025'] / kpi_yoy['Leads_2024']) - 1
    kpi_yoy['Var_Mat_%'] = (kpi_yoy['Mat_2025'] / kpi_yoy['Mat_2024']) - 1
    kpi_yoy['Conv_2024'] = kpi_yoy['Mat_2024'] / kpi_yoy['Leads_2024']
    kpi_yoy['Conv_2025'] = kpi_yoy['Mat_2025'] / kpi_yoy['Leads_2025']
    kpi_yoy['Var_Conv_pp'] = kpi_yoy['Conv_2025'] - kpi_yoy['Conv_2024']
    
    # Tratar infinitos e NaN
    kpi_yoy = kpi_yoy.replace([float('inf'), float('-inf')], 0)
    kpi_yoy = kpi_yoy.fillna(0)

    # KPI 3: Trend Mensal Offline
    df_offline = df[df['Canal_Estrategico'] == 'Offline (Unidade/Balcão)'].copy()
    df_offline['Mes_Ref'] = df_offline['Data_Entrada'].dt.to_period('M')
    df_offline['Ano'] = df_offline['Data_Entrada'].dt.year
    
    kpi_trend = df_offline.groupby(['Mes_Ref', 'Ano']).agg(
        Volume_Leads=('ID_Lead', 'count'),
        Matriculas=('Matricula_Realizada', 'sum')
    ).reset_index()
    kpi_trend['Mes_Ref'] = kpi_trend['Mes_Ref'].astype(str)

    # 4. EXPORTAÇÃO
    print(f"\n[EXPORTANDO] Salvando arquivo...")
    
    with pd.ExcelWriter(CAMINHO_SAIDA, engine='openpyxl') as writer:
        kpi_performance.to_excel(writer, sheet_name='1_Performance_2025', index=False)
        kpi_yoy.to_excel(writer, sheet_name='2_Analise_YOY', index=False)
        kpi_trend.to_excel(writer, sheet_name='3_Trend_Offline', index=False)
        
        # Drill down com amostra
        df_2025[['Data_Entrada', 'Canal_Estrategico', 'Fonte_Macro', 'Status_Atual']].head(2000).to_excel(
            writer, sheet_name='4_Drill_Down_Data', index=False
        )

    print(f"\n{'='*60}")
    print(f"[SUCESSO] Relatório gerado: {CAMINHO_SAIDA}")
    print(f"{'='*60}")
    print("\nNOTA: No Excel, selecione colunas de % e use 'Ctrl+Shift+%'")

except Exception as e:
    import traceback
    print(f"[ERRO CRÍTICO] {e}")
    traceback.print_exc()