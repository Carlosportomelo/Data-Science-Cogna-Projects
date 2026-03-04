"""
================================================================================
ANÁLISE COMPLETA DO FUNIL RED BALLOON - CICLO 26.1
Gerador automático de relatório Excel
================================================================================
Autor: Carlos Porto de Melo | Marketing Analytics & Growth
Data: Março/2026

COMO USAR:
1. Atualize as variáveis na seção CONFIGURAÇÕES
2. Rode: python analise_funil_automatico.py
3. Abra o Excel gerado com todas as tabelas prontas

SAÍDA:
- Analise_Funil_PPT_{DATA}.xlsx com todas as abas
================================================================================
"""

import pandas as pd
import numpy as np
import re
from datetime import datetime
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
import warnings
warnings.filterwarnings('ignore')

# ============================================================================
# ⚠️ CONFIGURAÇÕES - ATUALIZAR ANTES DE RODAR
# ============================================================================

# Caminho do arquivo HubSpot (adicione o seu caminho aqui)
CANDIDATE_PATHS = [
    Path(r"C:\Users\a483650\Projetos\_DADOS_CENTRALIZADOS\hubspot\hubspot_leads_ATUAL.csv"),
    Path("/mnt/user-data/uploads/hubspot_leads_atual.csv"),
]

# Pasta de saída
OUTPUT_DIR = Path(r"C:\Users\a483650\Projetos\06_Analise_Funil_RedBalloon\outputs")

# Ciclo atual
INICIO_CICLO = '2025-10-01'
FIM_CICLO = '2026-03-31'
NOME_CICLO = '26.1'

# ============================================================================
# ⚠️ ATUALIZAR ESTES VALORES ANTES DE CADA EXECUÇÃO
# ============================================================================
META_ALUNOS = 800
REALIZADO_ALUNOS = 595  # Atualizar com dado real
# ============================================================================

# Configurações financeiras (Política Comercial 2026)
PARCELAS = 12
KIT_KIDS = 996.77
KIT_JUNIORS_TEENS = 1217.77
KIT_MEDIO = (KIT_KIDS + KIT_JUNIORS_TEENS) / 2

# Unidades próprias
UNIDADES_PROPRIAS = [
    'VILA LEOPOLDINA', 'MORUMBI', 'ITAIM BIBI', 'PACAEMBU',
    'PINHEIROS', 'JARDINS', 'PERDIZES', 'SANTANA'
]

PRECOS_UNIDADE = {
    'VILA LEOPOLDINA': 1319.77, 'MORUMBI': 1319.77, 'ITAIM BIBI': 1319.77,
    'PACAEMBU': 1319.77, 'PINHEIROS': 1319.77, 'JARDINS': 1319.77,
    'PERDIZES': 1319.77, 'SANTANA': 1307.77,
}

PRECO_PADRAO = 1319.77

# ============================================================================
# FUNÇÕES AUXILIARES
# ============================================================================

def encontrar_arquivo():
    for path in CANDIDATE_PATHS:
        if path.exists():
            return path
    raise FileNotFoundError("Arquivo não encontrado. Adicione o caminho correto em CANDIDATE_PATHS.")

def extrair_idade_aluno(nome):
    if pd.isna(nome): return None
    match = re.search(r'\s*-\s*(\d{1,2})\s*-\s*', str(nome))
    if match:
        idade = int(match.group(1))
        if 3 <= idade <= 17: return idade
    match = re.search(r'\s*-\s*(\d{1,2})\s*$', str(nome))
    if match:
        idade = int(match.group(1))
        if 3 <= idade <= 17: return idade
    return None

def classificar_produto(idade):
    if pd.isna(idade): return 'Não informada'
    elif idade <= 6: return 'Kids'
    elif idade <= 10: return 'Juniors'
    elif idade <= 14: return 'Early Teens'
    return 'Late Teens'

def classificar_faixa(dias, detalhada=False):
    if pd.isna(dias): return None
    elif dias <= 1: return 'D-1'
    elif dias <= 7: return '2-7 dias'
    elif dias <= 30: return '8-30 dias'
    elif detalhada:
        if dias <= 60: return '31-60 dias'
        elif dias <= 90: return '61-90 dias'
        else: return '> 90 dias'
    else:
        if dias <= 90: return '31-90 dias'
        else: return '> 90 dias'

def obter_kit(produto):
    if produto == 'Kids': return KIT_KIDS
    elif produto in ['Juniors', 'Early Teens', 'Late Teens']: return KIT_JUNIORS_TEENS
    return KIT_MEDIO

# ============================================================================
# CARREGAR DADOS
# ============================================================================

def carregar_dados():
    print("="*70)
    print("📊 GERADOR DE RELATÓRIO - FUNIL RED BALLOON")
    print("="*70)
    
    arquivo = encontrar_arquivo()
    print(f"\n📂 Carregando: {arquivo}")
    
    df = pd.read_csv(arquivo)
    data_extracao = datetime.now().strftime('%d/%m/%Y')
    print(f"   Total: {len(df):,} registros")
    print(f"   Data: {data_extracao}")
    
    df['Data de criação'] = pd.to_datetime(df['Data de criação'], errors='coerce')
    df['Data de fechamento'] = pd.to_datetime(df['Data de fechamento'], errors='coerce')
    df['Data da última atividade'] = pd.to_datetime(df['Data da última atividade'], errors='coerce')
    
    return df, data_extracao

# ============================================================================
# 1. DIAGNÓSTICO DE QUALIDADE
# ============================================================================

def gerar_diagnostico(df):
    print("\n🔍 Gerando diagnóstico de qualidade...")
    
    total = len(df)
    ativ_nula = df['Data da última atividade'].isna().sum()
    
    mat = df[df['Etapa do negócio'].str.contains('MATRÍCULA CONCLUÍDA', na=False)].copy()
    total_mat = len(mat)
    mat_sem_ativ = mat['Data da última atividade'].isna().sum()
    
    mat['dias'] = (mat['Data de fechamento'] - mat['Data de criação']).dt.days
    mat_d1 = mat[mat['dias'] <= 1]
    mat_d1_offline = (mat_d1['Fonte original do tráfego'] == 'Fontes off-line').sum()
    media_ativ_d1 = mat_d1['Número de atividades de vendas'].mean()
    
    leads_offline = (df['Fonte original do tráfego'] == 'Fontes off-line').sum()
    mat_offline = (mat['Fonte original do tráfego'] == 'Fontes off-line').sum()
    leads_social = (df['Fonte original do tráfego'] == 'Social pago').sum()
    mat_social = (mat['Fonte original do tráfego'] == 'Social pago').sum()
    
    dados = [
        ['Leads com última atividade nula', f'{ativ_nula:,}', f'{ativ_nula/total*100:.1f}%'],
        ['Matrículas sem data de atividade', f'{mat_sem_ativ:,}', f'{mat_sem_ativ/total_mat*100:.1f}%'],
        ['Matrículas D-1 (conversão até 1 dia)', f'{len(mat_d1):,}', f'{len(mat_d1)/total_mat*100:.1f}%'],
        ['Matrículas D-1 off-line', f'{mat_d1_offline:,}', f'{mat_d1_offline/len(mat_d1)*100:.1f}%'],
        ['Média atividades D-1', f'{media_ativ_d1:.2f}', '—'],
        ['', '', ''],
        ['Off-line - Leads', f'{leads_offline:,}', f'{leads_offline/total*100:.1f}%'],
        ['Off-line - Matrículas', f'{mat_offline:,}', f'{mat_offline/total_mat*100:.1f}%'],
        ['Off-line - Taxa conversão', f'{mat_offline/leads_offline*100:.1f}%', '—'],
        ['', '', ''],
        ['Social Pago - Leads', f'{leads_social:,}', f'{leads_social/total*100:.1f}%'],
        ['Social Pago - Matrículas', f'{mat_social:,}', f'{mat_social/total_mat*100:.1f}%'],
        ['Social Pago - Taxa conversão', f'{mat_social/leads_social*100:.1f}%', '—'],
    ]
    
    return pd.DataFrame(dados, columns=['Métrica', 'Valor', '%'])

# ============================================================================
# 2. ERROS PRÓPRIAS VS FRANQUEADAS
# ============================================================================

def gerar_erros_proprias_franqueadas(df):
    print("🏢 Gerando comparativo próprias vs franqueadas...")
    
    df_c = df[(df['Data de criação'] >= INICIO_CICLO) & (df['Data de criação'] <= FIM_CICLO)].copy()
    df_c['unidade_upper'] = df_c['Unidade Desejada'].str.upper().str.strip()
    df_c['propria'] = df_c['unidade_upper'].isin(UNIDADES_PROPRIAS)
    
    df_c['e1'] = df_c['Data de fechamento'] < df_c['Data de criação']
    df_c['e2'] = df_c['Data da última atividade'].isna()
    df_c['e3'] = df_c['Data da última atividade'] < df_c['Data de criação']
    
    p = df_c[df_c['propria']]
    f = df_c[~df_c['propria']]
    
    dados = [
        ['Total leads', len(p), len(f)],
        ['Data fechamento < criação', f'{p["e1"].sum()/len(p)*100:.1f}%', f'{f["e1"].sum()/len(f)*100:.1f}%'],
        ['Última atividade nula', f'{p["e2"].sum()/len(p)*100:.1f}%', f'{f["e2"].sum()/len(f)*100:.1f}%'],
        ['Última atividade < criação', f'{p["e3"].sum()/len(p)*100:.1f}%', f'{f["e3"].sum()/len(f)*100:.1f}%'],
    ]
    
    return pd.DataFrame(dados, columns=['Tipo de Erro', 'Próprias', 'Franqueadas'])

# ============================================================================
# 3. ANÁLISE POR FAIXA
# ============================================================================

def gerar_analise_faixa(df, tipo='proprias'):
    print(f"📊 Gerando análise por faixa ({tipo})...")
    
    df_c = df[(df['Data de criação'] >= INICIO_CICLO) & (df['Data de criação'] <= FIM_CICLO)].copy()
    df_c = df_c[~(df_c['Data de fechamento'] < df_c['Data de criação'])]
    df_c = df_c[df_c['Data da última atividade'].notna()]
    df_c = df_c[~(df_c['Data da última atividade'] < df_c['Data de criação'])]
    
    df_c['unidade_upper'] = df_c['Unidade Desejada'].str.upper().str.strip()
    if tipo == 'proprias':
        df_c = df_c[df_c['unidade_upper'].isin(UNIDADES_PROPRIAS)]
    else:
        df_c = df_c[~df_c['unidade_upper'].isin(UNIDADES_PROPRIAS)]
    
    df_c['recencia'] = (df_c['Data da última atividade'] - df_c['Data de criação']).dt.days
    df_c['faixa'] = df_c['recencia'].apply(lambda d: classificar_faixa(d, detalhada=True))
    df_c['mat'] = df_c['Etapa do negócio'].str.contains('MATRÍCULA CONCLUÍDA', na=False)
    
    faixas = ['D-1', '2-7 dias', '8-30 dias', '31-60 dias', '61-90 dias', '> 90 dias']
    
    dados = []
    for f in faixas:
        s = df_c[df_c['faixa'] == f]
        l = len(s)
        m = s['mat'].sum()
        t = m/l*100 if l > 0 else 0
        al = s['Número de atividades de vendas'].mean() if l > 0 else 0
        am = s[s['mat']]['Número de atividades de vendas'].mean() if m > 0 else 0
        st = '🔴' if f == 'D-1' else ('🟡' if f == '2-7 dias' else '✅')
        dados.append([f, l, int(m), f'{t:.1f}%', round(al, 1), round(am, 1), st])
    
    return pd.DataFrame(dados, columns=['Faixa', 'Leads', 'Mat', 'Taxa', 'Méd Ativ (leads)', 'Méd Ativ (mat)', 'Status'])

# ============================================================================
# 4. PROJEÇÃO POR FAIXA
# ============================================================================

def gerar_projecao_faixa(df):
    print("🎯 Gerando projeção por faixa...")
    
    df_c = df[(df['Data de criação'] >= INICIO_CICLO) & (df['Data de criação'] <= FIM_CICLO)].copy()
    df_c = df_c[~(df_c['Data de fechamento'] < df_c['Data de criação'])]
    df_c = df_c[df_c['Data da última atividade'].notna()]
    df_c = df_c[~(df_c['Data da última atividade'] < df_c['Data de criação'])]
    df_c['unidade_upper'] = df_c['Unidade Desejada'].str.upper().str.strip()
    df_c = df_c[df_c['unidade_upper'].isin(UNIDADES_PROPRIAS)]
    
    df_c['recencia'] = (df_c['Data da última atividade'] - df_c['Data de criação']).dt.days
    df_c['faixa'] = df_c['recencia'].apply(classificar_faixa)
    df_c['mat'] = df_c['Etapa do negócio'].str.contains('MATRÍCULA CONCLUÍDA', na=False)
    df_c['qual'] = df_c['Etapa do negócio'].str.contains('QUALIFICAÇÃO', na=False)
    df_c['mensalidade'] = df_c['unidade_upper'].map(PRECOS_UNIDADE).fillna(PRECO_PADRAO)
    df_c['receita'] = (df_c['mensalidade'] * PARCELAS) + KIT_MEDIO
    
    faixas = ['D-1', '2-7 dias', '8-30 dias', '31-90 dias', '> 90 dias']
    
    dados = []
    mat_por_faixa = {}
    
    for faixa in faixas:
        df_f = df_c[df_c['faixa'] == faixa]
        eq, me, rp = 0, 0.0, 0.0
        
        for u in UNIDADES_PROPRIAS:
            du = df_f[df_f['unidade_upper'] == u]
            dq = du[du['qual']]
            n = len(dq)
            if n == 0: continue
            taxa = du['mat'].sum() / len(du) if len(du) > 0 else 0
            mat = n * taxa
            rec = mat * dq['receita'].mean()
            eq += n; me += mat; rp += rec
        
        mat_por_faixa[faixa] = me
        dados.append([faixa, eq, round(me, 0), round(rp, 2)])
    
    # Totais
    teq = sum(d[1] for d in dados)
    tme = sum(d[2] for d in dados)
    trp = sum(d[3] for d in dados)
    dados.append(['TOTAL', teq, tme, trp])
    
    df_result = pd.DataFrame(dados, columns=['Faixa', 'Em Qualificação', 'Matrículas Esperadas', 'Receita Potencial'])
    
    return df_result, mat_por_faixa

# ============================================================================
# 5. PROJEÇÃO POR PRODUTO
# ============================================================================

def gerar_projecao_produto(df):
    print("🎒 Gerando projeção por produto...")
    
    df_c = df[(df['Data de criação'] >= INICIO_CICLO) & (df['Data de criação'] <= FIM_CICLO)].copy()
    df_c = df_c[~(df_c['Data de fechamento'] < df_c['Data de criação'])]
    df_c = df_c[df_c['Data da última atividade'].notna()]
    df_c = df_c[~(df_c['Data da última atividade'] < df_c['Data de criação'])]
    df_c['unidade_upper'] = df_c['Unidade Desejada'].str.upper().str.strip()
    df_c = df_c[df_c['unidade_upper'].isin(UNIDADES_PROPRIAS)]
    df_c['recencia'] = (df_c['Data da última atividade'] - df_c['Data de criação']).dt.days
    df_c = df_c[(df_c['recencia'] >= 8) & (df_c['recencia'] <= 30)]
    
    df_c['idade_aluno'] = df_c['Nome do negócio'].apply(extrair_idade_aluno)
    df_c['produto'] = df_c['idade_aluno'].apply(classificar_produto)
    df_c['mensalidade'] = df_c['unidade_upper'].map(PRECOS_UNIDADE).fillna(PRECO_PADRAO)
    df_c['kit'] = df_c['produto'].apply(obter_kit)
    df_c['receita'] = (df_c['mensalidade'] * PARCELAS) + df_c['kit']
    df_c['mat'] = df_c['Etapa do negócio'].str.contains('MATRÍCULA CONCLUÍDA', na=False)
    df_c['qual'] = df_c['Etapa do negócio'].str.contains('QUALIFICAÇÃO', na=False)
    
    dados = []
    for p in ['Kids', 'Juniors', 'Early Teens', 'Late Teens', 'Não informada']:
        dp = df_c[df_c['produto'] == p]
        dq = dp[dp['qual']]
        nl, nq, nm = len(dp), len(dq), dp['mat'].sum()
        if nl == 0: continue
        taxa = nm / nl
        me = nq * taxa
        msg = dq['mensalidade'].mean() if nq > 0 else PRECO_PADRAO
        kit = dq['kit'].mean() if nq > 0 else KIT_MEDIO
        rl = (msg * PARCELAS) + kit
        rp = me * rl
        dados.append([p, nq, f'{taxa*100:.1f}%', round(me, 0), round(msg, 2), round(kit, 2), round(rl, 2), round(rp, 2)])
    
    # Totais
    tl = sum(d[1] for d in dados)
    tm = sum(d[3] for d in dados)
    tr = sum(d[7] for d in dados)
    dados.append(['TOTAL', tl, '—', tm, '—', '—', '—', tr])
    
    return pd.DataFrame(dados, columns=['Produto', 'Leads', 'Taxa', 'Mat Esp', 'Mensalidade', 'Kit', 'Receita/Lead', 'Receita Pot'])

# ============================================================================
# 6. PROJEÇÃO POR CANAL
# ============================================================================

def gerar_projecao_canal(df):
    print("📢 Gerando projeção por canal...")
    
    df_c = df[(df['Data de criação'] >= INICIO_CICLO) & (df['Data de criação'] <= FIM_CICLO)].copy()
    df_c = df_c[~(df_c['Data de fechamento'] < df_c['Data de criação'])]
    df_c = df_c[df_c['Data da última atividade'].notna()]
    df_c = df_c[~(df_c['Data da última atividade'] < df_c['Data de criação'])]
    df_c['unidade_upper'] = df_c['Unidade Desejada'].str.upper().str.strip()
    df_c = df_c[df_c['unidade_upper'].isin(UNIDADES_PROPRIAS)]
    df_c['recencia'] = (df_c['Data da última atividade'] - df_c['Data de criação']).dt.days
    df_c = df_c[(df_c['recencia'] >= 8) & (df_c['recencia'] <= 30)]
    
    df_c['mensalidade'] = df_c['unidade_upper'].map(PRECOS_UNIDADE).fillna(PRECO_PADRAO)
    df_c['receita'] = (df_c['mensalidade'] * PARCELAS) + KIT_MEDIO
    df_c['mat'] = df_c['Etapa do negócio'].str.contains('MATRÍCULA CONCLUÍDA', na=False)
    df_c['qual'] = df_c['Etapa do negócio'].str.contains('QUALIFICAÇÃO', na=False)
    
    dados = []
    for canal in df_c['Fonte original do tráfego'].dropna().unique():
        dc = df_c[df_c['Fonte original do tráfego'] == canal]
        dq = dc[dc['qual']]
        nl, nq, nm = len(dc), len(dq), dc['mat'].sum()
        if nl == 0 or nq == 0: continue
        taxa = nm / nl
        me = nq * taxa
        rm = dq['receita'].mean()
        rp = me * rm
        dados.append([canal, nq, f'{taxa*100:.1f}%', round(me, 0), round(rm, 2), round(rp, 2)])
    
    # Ordenar por receita
    dados = sorted(dados, key=lambda x: x[5], reverse=True)
    
    # Totais
    teq = sum(d[1] for d in dados)
    tme = sum(d[3] for d in dados)
    trp = sum(d[5] for d in dados)
    dados.append(['TOTAL', teq, '—', tme, '—', trp])
    
    return pd.DataFrame(dados, columns=['Canal', 'Em Qual', 'Taxa', 'Mat Esp', 'Receita/Lead', 'Receita Pot'])

# ============================================================================
# 7. PROJEÇÃO POR UNIDADE
# ============================================================================

def gerar_projecao_unidade(df):
    print("🏫 Gerando projeção por unidade...")
    
    df_c = df[(df['Data de criação'] >= INICIO_CICLO) & (df['Data de criação'] <= FIM_CICLO)].copy()
    df_c = df_c[~(df_c['Data de fechamento'] < df_c['Data de criação'])]
    df_c = df_c[df_c['Data da última atividade'].notna()]
    df_c = df_c[~(df_c['Data da última atividade'] < df_c['Data de criação'])]
    df_c['unidade_upper'] = df_c['Unidade Desejada'].str.upper().str.strip()
    df_c = df_c[df_c['unidade_upper'].isin(UNIDADES_PROPRIAS)]
    df_c['recencia'] = (df_c['Data da última atividade'] - df_c['Data de criação']).dt.days
    df_c = df_c[(df_c['recencia'] >= 8) & (df_c['recencia'] <= 30)]
    
    df_c['mensalidade'] = df_c['unidade_upper'].map(PRECOS_UNIDADE).fillna(PRECO_PADRAO)
    df_c['receita'] = (df_c['mensalidade'] * PARCELAS) + KIT_MEDIO
    df_c['mat'] = df_c['Etapa do negócio'].str.contains('MATRÍCULA CONCLUÍDA', na=False)
    df_c['qual'] = df_c['Etapa do negócio'].str.contains('QUALIFICAÇÃO', na=False)
    
    dados = []
    for u in UNIDADES_PROPRIAS:
        du = df_c[df_c['unidade_upper'] == u]
        dq = du[du['qual']]
        nl, nq, nm = len(du), len(dq), du['mat'].sum()
        if nl == 0: continue
        taxa = nm / nl
        me = nq * taxa
        rm = dq['receita'].mean() if nq > 0 else 0
        rp = me * rm
        dados.append([u, nq, f'{taxa*100:.1f}%', round(me, 0), round(rp, 2)])
    
    # Ordenar por receita
    dados = sorted(dados, key=lambda x: x[4], reverse=True)
    
    # Totais
    teq = sum(d[1] for d in dados)
    tme = sum(d[3] for d in dados)
    trp = sum(d[4] for d in dados)
    dados.append(['TOTAL', teq, '—', tme, trp])
    
    return pd.DataFrame(dados, columns=['Unidade', 'Em Qual', 'Taxa', 'Mat Esp', 'Receita Pot'])

# ============================================================================
# 8. META VS PROJEÇÃO
# ============================================================================

def gerar_meta_vs_projecao(mat_por_faixa):
    print("🎯 Gerando meta vs projeção...")
    
    d1 = mat_por_faixa.get('D-1', 0)
    d27 = mat_por_faixa.get('2-7 dias', 0)
    d830 = mat_por_faixa.get('8-30 dias', 0)
    d3190 = mat_por_faixa.get('31-90 dias', 0)
    d90 = mat_por_faixa.get('> 90 dias', 0)
    
    cons = d830
    conf = d830 + d3190 + d90
    otim = d1 + d27 + d830 + d3190 + d90
    
    dados = [
        ['META Ciclo ' + NOME_CICLO, META_ALUNOS, '—', '—'],
        ['Realizado', REALIZADO_ALUNOS, '—', '—'],
        ['GAP', META_ALUNOS - REALIZADO_ALUNOS, '—', '—'],
        ['', '', '', ''],
        ['Conservador (só 8-30d)', int(cons), int(REALIZADO_ALUNOS + cons), int(REALIZADO_ALUNOS + cons - META_ALUNOS)],
        ['Confiável (8-30d + 31-90d + >90d)', int(conf), int(REALIZADO_ALUNOS + conf), int(REALIZADO_ALUNOS + conf - META_ALUNOS)],
        ['Otimista (todas)', int(otim), int(REALIZADO_ALUNOS + otim), int(REALIZADO_ALUNOS + otim - META_ALUNOS)],
        ['PARA BATER', META_ALUNOS - REALIZADO_ALUNOS, META_ALUNOS, 0],
    ]
    
    return pd.DataFrame(dados, columns=['Cenário', 'Mat Esp', 'Total', 'vs Meta'])

# ============================================================================
# 9. PROJEÇÃO ACUMULADA
# ============================================================================

def gerar_projecao_acumulada(mat_por_faixa, eq_por_faixa):
    print("📈 Gerando projeção acumulada...")
    
    faixas = ['D-1', '2-7 dias', '8-30 dias', '31-90 dias', '> 90 dias']
    
    dados = []
    acum = 0.0
    
    for f in faixas:
        eq = eq_por_faixa.get(f, 0)
        me = mat_por_faixa.get(f, 0)
        acum += me
        rp = REALIZADO_ALUNOS + acum
        vs = rp - META_ALUNOS
        st = "✅" if vs >= 0 else "❌"
        dados.append([f, eq, int(me), int(acum), int(rp), int(vs), st])
    
    # Total
    otim = sum(mat_por_faixa.values())
    dados.append(['TOTAL', '—', int(otim), '—', int(REALIZADO_ALUNOS + otim), int(REALIZADO_ALUNOS + otim - META_ALUNOS), '❌' if REALIZADO_ALUNOS + otim < META_ALUNOS else '✅'])
    
    return pd.DataFrame(dados, columns=['Faixa', 'Em Qual', 'Mat Esp', 'Acumulado', 'Real + Proj', 'vs Meta', 'Status'])

# ============================================================================
# 10. MÉTRICAS RESUMO (FAIXA 8-30)
# ============================================================================

def gerar_metricas_resumo(df):
    print("📋 Gerando métricas resumo...")
    
    df_c = df[(df['Data de criação'] >= INICIO_CICLO) & (df['Data de criação'] <= FIM_CICLO)].copy()
    df_c = df_c[~(df_c['Data de fechamento'] < df_c['Data de criação'])]
    df_c = df_c[df_c['Data da última atividade'].notna()]
    df_c = df_c[~(df_c['Data da última atividade'] < df_c['Data de criação'])]
    df_c['unidade_upper'] = df_c['Unidade Desejada'].str.upper().str.strip()
    df_c = df_c[df_c['unidade_upper'].isin(UNIDADES_PROPRIAS)]
    df_c['recencia'] = (df_c['Data da última atividade'] - df_c['Data de criação']).dt.days
    df_c = df_c[(df_c['recencia'] >= 8) & (df_c['recencia'] <= 30)]
    
    df_c['mensalidade'] = df_c['unidade_upper'].map(PRECOS_UNIDADE).fillna(PRECO_PADRAO)
    df_c['receita'] = (df_c['mensalidade'] * PARCELAS) + KIT_MEDIO
    df_c['mat'] = df_c['Etapa do negócio'].str.contains('MATRÍCULA CONCLUÍDA', na=False)
    df_c['qual'] = df_c['Etapa do negócio'].str.contains('QUALIFICAÇÃO', na=False)
    
    teq, tme, trp = 0, 0.0, 0.0
    for u in UNIDADES_PROPRIAS:
        du = df_c[df_c['unidade_upper'] == u]
        dq = du[du['qual']]
        n = len(dq)
        if n == 0: continue
        taxa = du['mat'].sum() / len(du) if len(du) > 0 else 0
        mat = n * taxa
        rec = mat * dq['receita'].mean()
        teq += n; tme += mat; trp += rec
    
    taxa_geral = tme / teq if teq > 0 else 0
    receita_media = trp / tme if tme > 0 else 0
    
    dados = [
        ['Leads em Qualificação', teq],
        ['Taxa de Conversão', f'{taxa_geral*100:.1f}%'],
        ['Matrículas Esperadas', int(tme)],
        ['Receita Média/Matrícula', f'R$ {receita_media/1000:.0f}k'],
        ['Receita Potencial', f'R$ {trp/1000:.0f} mil'],
    ]
    
    return pd.DataFrame(dados, columns=['Métrica', 'Valor'])

# ============================================================================
# EXPORTAR EXCEL
# ============================================================================

def exportar_excel(dfs, data_extracao):
    # Criar pasta se não existir
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    
    # Nome do arquivo com data
    data_arquivo = datetime.now().strftime('%Y%m%d')
    arquivo = OUTPUT_DIR / f'Analise_Funil_PPT_{data_arquivo}.xlsx'
    
    print(f"\n💾 Exportando para: {arquivo}")
    
    with pd.ExcelWriter(arquivo, engine='openpyxl') as writer:
        # Aba Resumo
        resumo = pd.DataFrame([
            ['ANÁLISE DO FUNIL RED BALLOON', ''],
            [f'Ciclo: {NOME_CICLO}', ''],
            [f'Data de extração: {data_extracao}', ''],
            [f'Meta: {META_ALUNOS} alunos', ''],
            [f'Realizado: {REALIZADO_ALUNOS} alunos', ''],
            [f'Gap: {META_ALUNOS - REALIZADO_ALUNOS} alunos', ''],
        ], columns=['Info', 'Valor'])
        resumo.to_excel(writer, sheet_name='Resumo', index=False)
        
        # Demais abas
        for nome, df in dfs.items():
            df.to_excel(writer, sheet_name=nome[:31], index=False)
    
    return arquivo

# ============================================================================
# EXECUÇÃO PRINCIPAL
# ============================================================================

def main():
    # Carregar dados
    df, data_extracao = carregar_dados()
    
    # Gerar todas as análises
    dfs = {}
    
    dfs['1. Diagnóstico'] = gerar_diagnostico(df)
    dfs['2. Erros Prop vs Franq'] = gerar_erros_proprias_franqueadas(df)
    dfs['3. Faixa Próprias'] = gerar_analise_faixa(df, 'proprias')
    dfs['4. Faixa Franqueadas'] = gerar_analise_faixa(df, 'franqueadas')
    
    df_projecao, mat_por_faixa = gerar_projecao_faixa(df)
    dfs['5. Projeção Faixa'] = df_projecao
    
    # Extrair em_qualificacao por faixa
    eq_por_faixa = dict(zip(df_projecao['Faixa'][:-1], df_projecao['Em Qualificação'][:-1]))
    
    dfs['6. Projeção Produto'] = gerar_projecao_produto(df)
    dfs['7. Projeção Canal'] = gerar_projecao_canal(df)
    dfs['8. Projeção Unidade'] = gerar_projecao_unidade(df)
    dfs['9. Meta vs Projeção'] = gerar_meta_vs_projecao(mat_por_faixa)
    dfs['10. Projeção Acumulada'] = gerar_projecao_acumulada(mat_por_faixa, eq_por_faixa)
    dfs['11. Métricas Resumo'] = gerar_metricas_resumo(df)
    
    # Exportar
    arquivo = exportar_excel(dfs, data_extracao)
    
    # Resumo final
    print("\n" + "="*70)
    print("✅ RELATÓRIO GERADO COM SUCESSO!")
    print("="*70)
    print(f"\n📁 Arquivo: {arquivo}")
    print(f"\n📊 Abas geradas:")
    for i, nome in enumerate(dfs.keys(), 1):
        print(f"   {i}. {nome}")
    
    print(f"\n🎯 Resumo:")
    print(f"   Meta: {META_ALUNOS} | Realizado: {REALIZADO_ALUNOS} | Gap: {META_ALUNOS - REALIZADO_ALUNOS}")
    
    otim = sum(mat_por_faixa.values())
    print(f"   Projeção otimista: {int(REALIZADO_ALUNOS + otim)} ({int(REALIZADO_ALUNOS + otim - META_ALUNOS):+d} vs meta)")
    
    print("\n💡 Para atualizar, edite META_ALUNOS e REALIZADO_ALUNOS no início do script.")
    print("="*70)

if __name__ == "__main__":
    main()