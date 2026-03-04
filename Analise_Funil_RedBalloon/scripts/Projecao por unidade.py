# ============================================================================
# PROJEÇÃO DE FUNIL RED BALLOON - UNIDADES PRÓPRIAS
# Cálculo consistente: soma das partes por unidade
# ============================================================================
# Autor: Carlos Porto de Melo | Marketing Analytics & Growth
# Data: Fev/2026
# ============================================================================

import pandas as pd
import numpy as np
import re
from datetime import datetime
import os
import warnings
warnings.filterwarnings('ignore')

# ============================================================================
# CONFIGURAÇÕES
# ============================================================================

INPUT_PATH = r"C:\Users\a483650\Projetos\_DADOS_CENTRALIZADOS\hubspot\hubspot_leads_atual.csv"
OUTPUT_DIR = r"C:\Users\a483650\Projetos\06_Analise_Funil_RedBalloon\outputs"

INICIO_CICLO = '2025-10-01'
FIM_CICLO = '2026-03-31'
PARCELAS = 12

# Kit médio para leads sem idade informada
KIT_MEDIO = (996.77 + 1217.77) / 2  # R$ 1.107,27

# ============================================================================
# UNIDADES PRÓPRIAS
# ============================================================================

UNIDADES_PROPRIAS = [
    'VILA LEOPOLDINA',
    'MORUMBI',
    'ITAIM BIBI',
    'PACAEMBU',
    'PINHEIROS',
    'JARDINS',
    'PERDIZES',
    'SANTANA'
]

# ============================================================================
# TABELA DE PREÇOS - UNIDADES PRÓPRIAS (SP Capital)
# ============================================================================

PRECOS_UNIDADE = {
    'VILA LEOPOLDINA': 1319.77,
    'MORUMBI': 1319.77,
    'ITAIM BIBI': 1319.77,
    'PACAEMBU': 1319.77,
    'PINHEIROS': 1319.77,
    'JARDINS': 1319.77,
    'PERDIZES': 1319.77,
    'SANTANA': 1307.77,
}

PRECO_PADRAO = 1319.77

# ============================================================================
# FUNÇÕES AUXILIARES
# ============================================================================

def extrair_idade_aluno(nome):
    if pd.isna(nome):
        return None
    match = re.search(r'\s*-\s*(\d{1,2})\s*-\s*', str(nome))
    if match:
        idade = int(match.group(1))
        if 3 <= idade <= 17:
            return idade
    match = re.search(r'\s*-\s*(\d{1,2})\s*$', str(nome))
    if match:
        idade = int(match.group(1))
        if 3 <= idade <= 17:
            return idade
    return None

def classificar_produto(idade):
    if pd.isna(idade):
        return 'Não informada'
    elif idade <= 6:
        return 'Kids'
    elif idade <= 10:
        return 'Juniors'
    elif idade <= 14:
        return 'Early Teens'
    else:
        return 'Late Teens'

def obter_preco_unidade(unidade):
    if pd.isna(unidade):
        return PRECO_PADRAO
    return PRECOS_UNIDADE.get(str(unidade).upper().strip(), PRECO_PADRAO)

def obter_preco_kit(produto):
    if produto == 'Kids':
        return 996.77
    elif produto in ['Juniors', 'Early Teens', 'Late Teens']:
        return 1217.77
    else:
        return KIT_MEDIO  # Kit médio para não informados

def classificar_faixa_recencia(dias):
    if pd.isna(dias):
        return None
    elif dias <= 1:
        return 'D-1'
    elif dias <= 7:
        return '2-7 dias'
    elif dias <= 30:
        return '8-30 dias'
    elif dias <= 60:
        return '31-60 dias'
    elif dias <= 90:
        return '61-90 dias'
    else:
        return '> 90 dias'

def limpar_nome_aba(nome):
    nome_limpo = str(nome).replace('/', '-').replace('\\', '-').replace('*', '').replace('?', '')
    nome_limpo = nome_limpo.replace('[', '').replace(']', '').replace(':', '')
    return nome_limpo[:31]

# ============================================================================
# FUNÇÕES DE PROCESSAMENTO
# ============================================================================

def carregar_e_processar(input_path, inicio_ciclo, fim_ciclo):
    print(f"📂 Carregando dados de: {input_path}")
    df = pd.read_csv(input_path)
    print(f"   Total de registros: {len(df):,}")
    
    # Converter datas
    print("📅 Convertendo datas...")
    df['Data de criação'] = pd.to_datetime(df['Data de criação'], errors='coerce')
    df['Data de fechamento'] = pd.to_datetime(df['Data de fechamento'], errors='coerce')
    df['Data da última atividade'] = pd.to_datetime(df['Data da última atividade'], errors='coerce')
    
    # Filtrar ciclo
    inicio_dt = pd.to_datetime(inicio_ciclo)
    fim_dt = pd.to_datetime(fim_ciclo)
    df = df[(df['Data de criação'] >= inicio_dt) & (df['Data de criação'] <= fim_dt)].copy()
    print(f"📊 Leads do ciclo ({inicio_ciclo} a {fim_ciclo}): {len(df):,}")
    
    # Limpar base
    print("🧹 Limpando base...")
    total_inicial = len(df)
    
    erro_fechamento = df['Data de fechamento'] < df['Data de criação']
    erro_atividade_nula = df['Data da última atividade'].isna()
    erro_atividade_anterior = df['Data da última atividade'] < df['Data de criação']
    
    print(f"   ❌ Data fechamento < criação: {erro_fechamento.sum():,}")
    print(f"   ❌ Última atividade nula: {erro_atividade_nula.sum():,}")
    print(f"   ❌ Última atividade < criação: {erro_atividade_anterior.sum():,}")
    
    df = df[~erro_fechamento & ~erro_atividade_nula & ~erro_atividade_anterior].copy()
    print(f"   ✅ Base limpa: {len(df):,} ({total_inicial - len(df):,} removidos)")
    
    # Filtrar apenas unidades próprias
    print(f"🏢 Filtrando unidades próprias...")
    df['unidade_upper'] = df['Unidade Desejada'].str.upper().str.strip()
    df = df[df['unidade_upper'].isin(UNIDADES_PROPRIAS)].copy()
    print(f"   ✅ Leads em unidades próprias: {len(df):,}")
    
    # Enriquecer dados
    print("🔧 Enriquecendo dados...")
    
    df['idade_recencia'] = (df['Data da última atividade'] - df['Data de criação']).dt.days
    df['faixa_recencia'] = df['idade_recencia'].apply(classificar_faixa_recencia)
    df['idade_aluno'] = df['Nome do negócio'].apply(extrair_idade_aluno)
    df['produto'] = df['idade_aluno'].apply(classificar_produto)
    df['mensalidade'] = df['Unidade Desejada'].apply(obter_preco_unidade)
    df['kit_didatico'] = df['produto'].apply(obter_preco_kit)
    df['receita_anual'] = (df['mensalidade'] * PARCELAS) + df['kit_didatico']
    df['matricula'] = df['Etapa do negócio'].str.contains('MATRÍCULA CONCLUÍDA', na=False)
    df['em_qualificacao'] = df['Etapa do negócio'].str.contains('QUALIFICAÇÃO', na=False)
    
    return df

def calcular_projecao_por_faixa(df):
    """Calcula projeção por faixa - SOMA DAS UNIDADES"""
    
    faixas_ordem = ['D-1', '2-7 dias', '8-30 dias', '31-60 dias', '61-90 dias', '> 90 dias']
    resultados = []
    
    for faixa in faixas_ordem:
        df_faixa = df[df['faixa_recencia'] == faixa]
        
        # Calcular por unidade e somar (método correto)
        total_em_qual = 0
        total_mat_esp = 0
        total_receita = 0
        
        for unidade in UNIDADES_PROPRIAS:
            df_unid = df_faixa[df_faixa['unidade_upper'] == unidade]
            df_unid_qual = df_unid[df_unid['em_qualificacao'] == True]
            
            n_em_qual = len(df_unid_qual)
            if n_em_qual == 0:
                continue
            
            # Taxa específica da unidade nesta faixa
            taxa_unid = df_unid['matricula'].sum() / len(df_unid) if len(df_unid) > 0 else 0
            mat_esp = n_em_qual * taxa_unid
            receita_media = df_unid_qual['receita_anual'].mean()
            receita_pot = mat_esp * receita_media
            
            total_em_qual += n_em_qual
            total_mat_esp += mat_esp
            total_receita += receita_pot
        
        # Receita média é a divisão do total
        receita_media_faixa = total_receita / total_mat_esp if total_mat_esp > 0 else 0
        
        resultados.append({
            'Faixa Recência': faixa,
            'Em Qualificação': total_em_qual,
            'Matrículas Esperadas': total_mat_esp,
            'Receita Média': receita_media_faixa,
            'Receita Potencial': total_receita
        })
    
    return pd.DataFrame(resultados)

def calcular_projecao_por_unidade(df, faixa='8-30 dias'):
    """Calcula projeção por unidade para uma faixa específica"""
    
    df_faixa = df[df['faixa_recencia'] == faixa]
    resultados = []
    
    for unidade in UNIDADES_PROPRIAS:
        df_unid = df_faixa[df_faixa['unidade_upper'] == unidade]
        df_unid_qual = df_unid[df_unid['em_qualificacao'] == True]
        
        n_leads = len(df_unid)
        n_mat = df_unid['matricula'].sum()
        n_em_qual = len(df_unid_qual)
        
        if n_leads == 0:
            continue
        
        taxa_unid = n_mat / n_leads
        mat_esp = n_em_qual * taxa_unid
        receita_media = df_unid_qual['receita_anual'].mean() if n_em_qual > 0 else 0
        receita_pot = mat_esp * receita_media
        
        resultados.append({
            'Unidade': unidade,
            'Total Leads': n_leads,
            'Matrículas Históricas': n_mat,
            'Taxa Conversão': taxa_unid,
            'Em Qualificação': n_em_qual,
            'Matrículas Esperadas': mat_esp,
            'Receita Média': receita_media,
            'Receita Potencial': receita_pot
        })
    
    df_result = pd.DataFrame(resultados)
    df_result = df_result.sort_values('Receita Potencial', ascending=False)
    
    return df_result

def calcular_projecao_por_canal(df, faixa='8-30 dias'):
    """Calcula projeção por canal para uma faixa específica"""
    
    df_faixa = df[df['faixa_recencia'] == faixa]
    df_faixa_qual = df_faixa[df_faixa['em_qualificacao'] == True]
    
    resultados = []
    
    for canal in df_faixa['Fonte original do tráfego'].dropna().unique():
        df_canal = df_faixa[df_faixa['Fonte original do tráfego'] == canal]
        df_canal_qual = df_faixa_qual[df_faixa_qual['Fonte original do tráfego'] == canal]
        
        n_leads = len(df_canal)
        n_mat = df_canal['matricula'].sum()
        n_em_qual = len(df_canal_qual)
        
        if n_leads == 0 or n_em_qual == 0:
            continue
        
        taxa_canal = n_mat / n_leads
        mat_esp = n_em_qual * taxa_canal
        receita_media = df_canal_qual['receita_anual'].mean()
        receita_pot = mat_esp * receita_media
        
        resultados.append({
            'Canal': canal,
            'Total Leads': n_leads,
            'Matrículas Históricas': n_mat,
            'Taxa Conversão': taxa_canal,
            'Em Qualificação': n_em_qual,
            'Matrículas Esperadas': mat_esp,
            'Receita Média': receita_media,
            'Receita Potencial': receita_pot
        })
    
    df_result = pd.DataFrame(resultados)
    df_result = df_result.sort_values('Receita Potencial', ascending=False)
    
    return df_result

def calcular_projecao_por_produto(df, faixa='8-30 dias'):
    """Calcula projeção por produto para uma faixa específica"""
    
    df_faixa = df[df['faixa_recencia'] == faixa]
    df_faixa_qual = df_faixa[df_faixa['em_qualificacao'] == True]
    
    resultados = []
    
    for produto in ['Kids', 'Juniors', 'Early Teens', 'Late Teens', 'Não informada']:
        df_prod = df_faixa[df_faixa['produto'] == produto]
        df_prod_qual = df_faixa_qual[df_faixa_qual['produto'] == produto]
        
        n_leads = len(df_prod)
        n_mat = df_prod['matricula'].sum()
        n_em_qual = len(df_prod_qual)
        
        if n_leads == 0:
            continue
        
        taxa_prod = n_mat / n_leads
        mat_esp = n_em_qual * taxa_prod
        mensalidade_media = df_prod_qual['mensalidade'].mean() if n_em_qual > 0 else PRECO_PADRAO
        kit_medio = df_prod_qual['kit_didatico'].mean() if n_em_qual > 0 else KIT_MEDIO
        receita_media = df_prod_qual['receita_anual'].mean() if n_em_qual > 0 else 0
        receita_pot = mat_esp * receita_media
        
        resultados.append({
            'Produto': produto,
            'Total Leads': n_leads,
            'Matrículas Históricas': n_mat,
            'Taxa Conversão': taxa_prod,
            'Em Qualificação': n_em_qual,
            'Matrículas Esperadas': mat_esp,
            'Mensalidade Média': mensalidade_media,
            'Kit Médio': kit_medio,
            'Receita Média': receita_media,
            'Receita Potencial': receita_pot
        })
    
    return pd.DataFrame(resultados)

def preparar_leads_unidade(df, unidade, faixa='8-30 dias'):
    """Prepara dataframe de leads de uma unidade para exportação"""
    
    df_unid = df[(df['unidade_upper'] == unidade) & 
                  (df['faixa_recencia'] == faixa) & 
                  (df['em_qualificacao'] == True)].copy()
    
    if len(df_unid) == 0:
        return None
    
    colunas = {
        'Record ID': 'ID',
        'Nome do negócio': 'Lead',
        'Data de criação': 'Data Criação',
        'Etapa do negócio': 'Status',
        'Fonte original do tráfego': 'Canal',
        'idade_aluno': 'Idade',
        'produto': 'Produto',
        'Número de atividades de vendas': 'Atividades',
        'mensalidade': 'Mensalidade',
        'kit_didatico': 'Kit',
        'receita_anual': 'Receita Anual'
    }
    
    df_export = df_unid[list(colunas.keys())].copy()
    df_export.columns = list(colunas.values())
    df_export['Data Criação'] = pd.to_datetime(df_export['Data Criação']).dt.strftime('%d/%m/%Y')
    
    return df_export

def exportar_excel(output_dir, df, df_projecao_faixa, df_projecao_unidade, 
                   df_projecao_canal, df_projecao_produto):
    """Exporta resultados para Excel"""
    
    os.makedirs(output_dir, exist_ok=True)
    
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"projecao_proprias_{timestamp}.xlsx"
    filepath = os.path.join(output_dir, filename)
    
    print(f"\n💾 Exportando para: {filepath}")
    
    with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
        
        # ABA 1: RESUMO POR FAIXA
        print("   📝 Criando aba: Resumo por Faixa")
        df_faixa_fmt = df_projecao_faixa.copy()
        df_faixa_fmt['Matrículas Esperadas'] = df_faixa_fmt['Matrículas Esperadas'].apply(lambda x: f"{x:.0f}")
        df_faixa_fmt['Receita Média'] = df_faixa_fmt['Receita Média'].apply(lambda x: f"R$ {x:,.2f}")
        df_faixa_fmt['Receita Potencial'] = df_faixa_fmt['Receita Potencial'].apply(lambda x: f"R$ {x:,.2f}")
        df_faixa_fmt.to_excel(writer, sheet_name='Resumo por Faixa', index=False)
        
        # ABA 2: PROJEÇÃO POR UNIDADE
        print("   📝 Criando aba: Projeção por Unidade")
        df_unid_fmt = df_projecao_unidade.copy()
        df_unid_fmt['Taxa Conversão'] = df_unid_fmt['Taxa Conversão'].apply(lambda x: f"{x*100:.1f}%")
        df_unid_fmt['Matrículas Esperadas'] = df_unid_fmt['Matrículas Esperadas'].apply(lambda x: f"{x:.0f}")
        df_unid_fmt['Receita Média'] = df_unid_fmt['Receita Média'].apply(lambda x: f"R$ {x:,.2f}")
        df_unid_fmt['Receita Potencial'] = df_unid_fmt['Receita Potencial'].apply(lambda x: f"R$ {x:,.2f}")
        df_unid_fmt.to_excel(writer, sheet_name='Projeção por Unidade', index=False)
        
        # ABA 3: PROJEÇÃO POR CANAL
        print("   📝 Criando aba: Projeção por Canal")
        df_canal_fmt = df_projecao_canal.copy()
        df_canal_fmt['Taxa Conversão'] = df_canal_fmt['Taxa Conversão'].apply(lambda x: f"{x*100:.1f}%")
        df_canal_fmt['Matrículas Esperadas'] = df_canal_fmt['Matrículas Esperadas'].apply(lambda x: f"{x:.0f}")
        df_canal_fmt['Receita Média'] = df_canal_fmt['Receita Média'].apply(lambda x: f"R$ {x:,.2f}")
        df_canal_fmt['Receita Potencial'] = df_canal_fmt['Receita Potencial'].apply(lambda x: f"R$ {x:,.2f}")
        df_canal_fmt.to_excel(writer, sheet_name='Projeção por Canal', index=False)
        
        # ABA 4: PROJEÇÃO POR PRODUTO
        print("   📝 Criando aba: Projeção por Produto")
        df_prod_fmt = df_projecao_produto.copy()
        df_prod_fmt['Taxa Conversão'] = df_prod_fmt['Taxa Conversão'].apply(lambda x: f"{x*100:.1f}%")
        df_prod_fmt['Matrículas Esperadas'] = df_prod_fmt['Matrículas Esperadas'].apply(lambda x: f"{x:.0f}")
        df_prod_fmt['Mensalidade Média'] = df_prod_fmt['Mensalidade Média'].apply(lambda x: f"R$ {x:,.2f}")
        df_prod_fmt['Kit Médio'] = df_prod_fmt['Kit Médio'].apply(lambda x: f"R$ {x:,.2f}")
        df_prod_fmt['Receita Média'] = df_prod_fmt['Receita Média'].apply(lambda x: f"R$ {x:,.2f}")
        df_prod_fmt['Receita Potencial'] = df_prod_fmt['Receita Potencial'].apply(lambda x: f"R$ {x:,.2f}")
        df_prod_fmt.to_excel(writer, sheet_name='Projeção por Produto', index=False)
        
        # ABAS POR UNIDADE (leads em qualificação 8-30 dias)
        for unidade in UNIDADES_PROPRIAS:
            df_leads = preparar_leads_unidade(df, unidade)
            if df_leads is not None and len(df_leads) > 0:
                nome_aba = limpar_nome_aba(unidade)
                print(f"   📝 Criando aba: {nome_aba} ({len(df_leads)} leads)")
                
                df_leads['Mensalidade'] = df_leads['Mensalidade'].apply(lambda x: f"R$ {x:,.2f}")
                df_leads['Kit'] = df_leads['Kit'].apply(lambda x: f"R$ {x:,.2f}")
                df_leads['Receita Anual'] = df_leads['Receita Anual'].apply(lambda x: f"R$ {x:,.2f}")
                
                df_leads.to_excel(writer, sheet_name=nome_aba, index=False)
    
    print(f"   ✅ Arquivo gerado com sucesso!")
    return filepath

# ============================================================================
# EXECUÇÃO PRINCIPAL
# ============================================================================

def main():
    print("\n" + "="*70)
    print("📊 PROJEÇÃO RED BALLOON - UNIDADES PRÓPRIAS")
    print("   Método: Soma das partes por unidade")
    print("="*70 + "\n")
    
    # 1. Carregar e processar
    df = carregar_e_processar(INPUT_PATH, INICIO_CICLO, FIM_CICLO)
    
    # 2. Calcular projeções
    print("\n📈 Calculando projeções...")
    df_projecao_faixa = calcular_projecao_por_faixa(df)
    df_projecao_unidade = calcular_projecao_por_unidade(df)
    df_projecao_canal = calcular_projecao_por_canal(df)
    df_projecao_produto = calcular_projecao_por_produto(df)
    
    # 3. Exibir resultados
    print("\n" + "="*70)
    print("📋 PROJEÇÃO POR FAIXA DE RECÊNCIA")
    print("="*70)
    print(f"\n{'Faixa':<12} {'Em Qual':>10} {'Mat Esp':>10} {'Receita Média':>15} {'Receita Potencial':>20}")
    print("-"*70)
    for _, row in df_projecao_faixa.iterrows():
        print(f"{row['Faixa Recência']:<12} {row['Em Qualificação']:>10,} {row['Matrículas Esperadas']:>10.0f} R$ {row['Receita Média']:>12,.2f} R$ {row['Receita Potencial']:>17,.2f}")
    
    print("\n" + "="*70)
    print("🏫 PROJEÇÃO POR UNIDADE (Faixa 8-30 dias)")
    print("="*70)
    print(f"\n{'Unidade':<20} {'Em Qual':>10} {'Taxa':>8} {'Mat Esp':>10} {'Receita Potencial':>20}")
    print("-"*70)
    
    total_em_qual = df_projecao_unidade['Em Qualificação'].sum()
    total_mat = df_projecao_unidade['Matrículas Esperadas'].sum()
    total_receita = df_projecao_unidade['Receita Potencial'].sum()
    
    for _, row in df_projecao_unidade.iterrows():
        print(f"{row['Unidade']:<20} {row['Em Qualificação']:>10,} {row['Taxa Conversão']*100:>7.1f}% {row['Matrículas Esperadas']:>10.0f} R$ {row['Receita Potencial']:>17,.2f}")
    
    print("-"*70)
    print(f"{'TOTAL':<20} {total_em_qual:>10,} {'':<8} {total_mat:>10.0f} R$ {total_receita:>17,.2f}")
    
    print("\n" + "="*70)
    print("📢 PROJEÇÃO POR CANAL (Faixa 8-30 dias)")
    print("="*70)
    print(f"\n{'Canal':<25} {'Em Qual':>10} {'Taxa':>8} {'Mat Esp':>10} {'Receita Potencial':>20}")
    print("-"*75)
    for _, row in df_projecao_canal.iterrows():
        print(f"{row['Canal']:<25} {row['Em Qualificação']:>10,} {row['Taxa Conversão']*100:>7.1f}% {row['Matrículas Esperadas']:>10.0f} R$ {row['Receita Potencial']:>17,.2f}")
    
    # 4. Exportar Excel
    filepath = exportar_excel(OUTPUT_DIR, df, df_projecao_faixa, df_projecao_unidade, 
                               df_projecao_canal, df_projecao_produto)
    
    print("\n" + "="*70)
    print("✅ RESUMO FINAL (Faixa 8-30 dias)")
    print("="*70)
    print(f"\n  Leads em Qualificação:   {total_em_qual:,}")
    print(f"  Matrículas Esperadas:    {total_mat:.0f}")
    print(f"  Receita Potencial:       R$ {total_receita:,.2f}")
    print("\n" + "="*70 + "\n")
    
    return filepath

if __name__ == "__main__":
    main()