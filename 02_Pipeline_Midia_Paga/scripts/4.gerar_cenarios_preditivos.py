#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
4.gerar_cenarios_preditivos.py

OBJETIVO:
    Ler a base consolidada (Blend) e gerar uma aba de cenários para o Looker Studio.
    Compara a performance Realizada (Nova Gestão) vs Simulação (Estratégia 2024).

INPUT:
    - Último arquivo 'meta_googleads_blend_*.xlsx' na pasta outputs/

OUTPUT:
    - outputs/cenarios_preditivos_dashboard.xlsx
"""

import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime
import glob
import os
import sys

# --- CONFIGURAÇÕES ---
try:
    BASE_DIR = Path(__file__).resolve().parent.parent
except NameError:
    BASE_DIR = Path.cwd()

OUTPUT_DIR = BASE_DIR / "outputs"
DATA_CORTE_GESTAO = pd.Timestamp('2025-08-01') # Data da mudança de gestão
MESES_PROJECAO = 12 # Quantos meses projetar para frente

def find_latest_blend_file(directory):
    """Encontra o arquivo de blend mais recente na pasta de outputs."""
    pattern = str(directory / "meta_googleads_blend_*.xlsx")
    files = glob.glob(pattern)
    if not files:
        return None
    # Ordenar por data de modificação (mais recente primeiro)
    latest_file = max(files, key=os.path.getmtime)
    return Path(latest_file)

def calcular_kpis_periodo(df, start_date, end_date=None):
    """Calcula médias de KPIs para um período específico."""
    if end_date:
        mask = (df['Mes_Ref'] >= start_date) & (df['Mes_Ref'] < end_date)
    else:
        mask = (df['Mes_Ref'] >= start_date)
    
    df_periodo = df[mask]
    
    if df_periodo.empty:
        return None

    invest_medio = df_periodo['Midia_Paga'].mean()
    leads_medio = df_periodo['Total_Negocios'].mean()
    matriculas_medio = df_periodo['Matriculas'].mean()
    rvo_medio = df_periodo['RVO'].mean()
    
    # Evitar divisão por zero
    cpl = invest_medio / leads_medio if leads_medio > 0 else 0
    taxa_conv = matriculas_medio / leads_medio if leads_medio > 0 else 0
    ticket_medio = rvo_medio / matriculas_medio if matriculas_medio > 0 else 0
    
    return {
        'Investimento': invest_medio,
        'Leads': leads_medio,
        'Matriculas': matriculas_medio,
        'RVO': rvo_medio,
        'CPL': cpl,
        'Taxa_Conversao': taxa_conv,
        'Ticket_Medio': ticket_medio
    }

def main():
    print("="*70)
    print("GERADOR DE CENÁRIOS PREDITIVOS")
    print("="*70)

    # 1. Encontrar e Carregar Arquivo
    blend_file = find_latest_blend_file(OUTPUT_DIR)
    if not blend_file:
        print("ERRO: Nenhum arquivo 'meta_googleads_blend_*.xlsx' encontrado em 'outputs/'.")
        print("   Execute o script 3 primeiro.")
        sys.exit(1)
    
    print(f"Lendo arquivo base: {blend_file.name}")
    try:
        # Ler a aba granular para ter os dados brutos
        df_raw = pd.read_excel(blend_file, sheet_name='Visao_Granular_Final')
    except Exception as e:
        print(f"Erro ao ler Excel: {e}")
        sys.exit(1)

    # 2. Preparar Dados Mensais (Realizado)
    print("Processando dados históricos...")
    df_raw['Data'] = pd.to_datetime(df_raw['Data'])
    df_raw['Mes_Ref'] = df_raw['Data'].dt.to_period('M').dt.to_timestamp()
    
    # Agrupar por mês
    df_mensal = df_raw.groupby('Mes_Ref').agg({
        'Midia_Paga': 'sum',
        'Total_Negocios': 'sum',
        'Matriculas': 'sum',
        'RVO': 'sum'
    }).reset_index()
    
    # 3. Calcular KPIs das Estratégias
    print(f"Calculando KPIs (Corte: {DATA_CORTE_GESTAO.strftime('%m/%Y')})...")
    
    # Estratégia Antiga (Jan 2024 até Jul 2025)
    kpis_antiga = calcular_kpis_periodo(df_mensal, '2024-01-01', DATA_CORTE_GESTAO)
    
    # Estratégia Nova (Ago 2025 até Hoje)
    kpis_nova = calcular_kpis_periodo(df_mensal, DATA_CORTE_GESTAO)
    
    if not kpis_antiga or not kpis_nova:
        print("Aviso: Dados insuficientes para calcular KPIs de um dos períodos.")
        # Fallback simples se faltar dados
        if not kpis_nova: kpis_nova = kpis_antiga
    
    print(f"   Estratégia 2024 (Antiga): Invest ~R${kpis_antiga['Investimento']:,.0f}, Conv {kpis_antiga['Taxa_Conversao']:.1%}")
    print(f"   Estratégia Atual (Nova):   Invest ~R${kpis_nova['Investimento']:,.0f}, Conv {kpis_nova['Taxa_Conversao']:.1%}")

    # 4. Gerar Tabela de Cenários
    print("Gerando projeções...")
    
    cenarios = []
    ultimo_mes_real = df_mensal['Mes_Ref'].max()
    data_fim_projecao = ultimo_mes_real + pd.DateOffset(months=MESES_PROJECAO)
    
    # Loop mês a mês desde o início de 2024 até o futuro
    for data in pd.date_range(start='2024-01-01', end=data_fim_projecao, freq='MS'):
        
        # A. DADOS REAIS (REALIZADO)
        if data <= ultimo_mes_real:
            row_real = df_mensal[df_mensal['Mes_Ref'] == data]
            if not row_real.empty:
                r = row_real.iloc[0]
                cenarios.append({
                    'Data': data, 'Cenario': '1. Realizado', 
                    'Investimento': r['Midia_Paga'], 'Leads': r['Total_Negocios'], 
                    'Matriculas': r['Matriculas'], 'RVO': r['RVO']
                })
        
        # B. SIMULAÇÕES (A partir da mudança de gestão)
        if data >= DATA_CORTE_GESTAO:
            
            # Função auxiliar para calcular linha projetada
            def add_projecao(nome_cenario, kpis_base, fator_otimista=1.0):
                inv = kpis_base['Investimento']
                cpl = kpis_base['CPL']
                conv = kpis_base['Taxa_Conversao'] * fator_otimista
                ticket = kpis_base['Ticket_Medio']
                
                leads_proj = inv / cpl if cpl > 0 else 0
                mat_proj = leads_proj * conv
                rvo_proj = mat_proj * ticket
                
                cenarios.append({
                    'Data': data, 'Cenario': nome_cenario,
                    'Investimento': inv, 'Leads': leads_proj,
                    'Matriculas': mat_proj, 'RVO': rvo_proj
                })

            # Cenário 2: Se mantivesse a estratégia antiga (Investimento Antigo + Eficiência Antiga)
            add_projecao('2. Estratégia 2024 (Simulação)', kpis_antiga)
            
            # Cenário 3: Tendência Atual (Investimento Novo + Eficiência Nova) - Projetado para o futuro
            if data > ultimo_mes_real:
                add_projecao('3. Tendência Atual (Projeção)', kpis_nova)
                
                # Cenário 4: Otimista (Investimento Novo + Eficiência Melhorada em 10%)
                add_projecao('4. Cenário Otimista', kpis_nova, fator_otimista=1.1)

    df_cenarios = pd.DataFrame(cenarios)
    
    # 5. Salvar
    arquivo_saida = OUTPUT_DIR / "cenarios_preditivos_dashboard.xlsx"
    print(f"\nSalvando arquivo em: {arquivo_saida}")
    df_cenarios.to_excel(arquivo_saida, index=False, sheet_name="Cenarios_Looker")
    print("Concluído com sucesso!")

if __name__ == "__main__":
    main()