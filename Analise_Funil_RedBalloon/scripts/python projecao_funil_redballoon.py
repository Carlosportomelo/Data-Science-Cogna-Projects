# ============================================================================
# PROJEÇÃO DE FUNIL RED BALLOON - VERSÃO COMPLETA (CORRIGIDA)
# Cruzamento: Idade do Lead → Produto | Unidade → Preço da Praça
# ============================================================================
# Autor: Carlos Porto de Melo | Marketing Analytics & Growth
# Data: Fev/2026
# Correção: Kit médio para leads sem idade informada
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

# Ciclo 26.1 (Out/25 - Mar/26)
INICIO_CICLO = '2025-10-01'
FIM_CICLO = '2026-03-31'

PARCELAS = 12  # Contrato anual

# ============================================================================
# TABELA DE KITS DIDÁTICOS (Política Comercial 2026)
# ============================================================================

KIT_KIDS = 996.77
KIT_JUNIORS_TEENS = 1217.77
KIT_MEDIO = (KIT_KIDS + KIT_JUNIORS_TEENS) / 2  # R$ 1.107,27 para não informados

# ============================================================================
# TABELA DE PREÇOS POR UNIDADE (Política Comercial 2026)
# Mensalidade 04h (carga horária padrão)
# ============================================================================

PRECOS_UNIDADE = {
    # SÃO PAULO CAPITAL - Padrão (04h = R$ 1.319,77)
    'MOEMA': 1319.77,
    'ITAIM BIBI': 1319.77,
    'JARDINS': 1319.77,
    'CAMPO BELO': 1319.77,
    'MORUMBI': 1319.77,
    'PINHEIROS': 1319.77,
    'PERDIZES': 1319.77,
    'PACAEMBU': 1319.77,
    'POMPEIA': 1319.77,
    'PARAISO': 1319.77,
    'VILA MARIANA': 1319.77,
    'IPIRANGA': 1319.77,
    'SAUDE': 1319.77,
    'MOOCA': 1319.77,
    'TATUAPE': 1319.77,
    'AGUA FRIA': 1319.77,
    'TUCURUVI': 1319.77,
    'CHACARA KLABIN': 1319.77,
    'VILA LEOPOLDINA': 1319.77,
    'CANTAREIRA': 1307.77,
    
    # SÃO PAULO CAPITAL - Exceções
    'BUTANTA': 1307.77,
    'SANTANA': 1307.77,
    'VILA SAO FRANCISCO': 1307.77,
    
    # GRANDE SÃO PAULO
    'ALPHAVILLE': 1293.77,
    'GRANJA VIANA': 1119.77,
    'GUARULHOS': 1279.00,
    'MOGI DAS CRUZES': 853.77,
    'SANTO ANDRE': 1091.77,
    'SAO BERNARDO DO CAMPO': 1091.77,
    'SAO CAETANO DO SUL': 1091.77,
    
    # SÃO PAULO INTERIOR
    'SAO JOSE DOS CAMPOS': 876.77,
    'TAQUARAL': 1163.77,
    'NOVA CAMPINAS': 1163.77,
    'SOROCABA': 1094.77,
    'RIBEIRAO PRETO': 983.44,
    'JUNDIAI': 814.77,
    'AMERICANA': 900.00,
    'INDAIATUBA': 900.00,
    'PIRACICABA': 950.00,
    'LIMEIRA': 900.00,
    'BAURU': 900.00,
    'SAO JOSE DO RIO PRETO': 950.00,
    'ARACATUBA': 850.00,
    'ATIBAIA': 900.00,
    'ITU': 900.00,
    'PAULINIA': 950.00,
    'VALINHOS': 950.00,
    
    # RIO DE JANEIRO
    'IPANEMA': 1188.76,
    'BARRA BLUE SQUARE': 1150.00,
    'BARRA JARDIM OCEANICO': 1150.00,
    'BOSQUE DA BARRA': 1150.00,
    'JACAREPAGUA': 1100.00,
    'TIJUCA': 1100.00,
    'ICARAI': 1100.00,
    'REGIAO OCEANICA': 1100.00,
    
    # MINAS GERAIS
    'LOURDES': 1089.83,
    'SION': 1089.83,
    'PAMPULHA': 1050.00,
    'NOVA LIMA': 1089.83,
    'UBERLANDIA': 950.00,
    
    # BRASÍLIA
    'ASA SUL': 1133.77,
    
    # SUL
    'PORTO ALEGRE': 1045.00,
    'CURITIBA': 1000.00,
    'LONDRINA': 900.00,
    
    # NORDESTE
    'SALVADOR': 950.00,
    'BOA VIAGEM': 1000.00,
    'GRACAS': 1000.00,
    'CASA FORTE': 1000.00,
    'ALDEOTA': 950.00,
    
    # CENTRO-OESTE
    'GOIANIA': 950.00,
    'ANAPOLIS': 850.00,
    'CUIABA': 900.00,
    'CAMPO GRANDE': 900.00,
    
    # NORTE
    'MANAUS': 818.61,
}

PRECO_PADRAO = 1000.00

# ============================================================================
# FUNÇÕES
# ============================================================================

def carregar_dados(path):
    """Carrega a base do HubSpot"""
    print(f"📂 Carregando dados de: {path}")
    df = pd.read_csv(path)
    print(f"   Total de registros: {len(df):,}")
    return df

def preparar_datas(df):
    """Converte colunas de data"""
    print("📅 Convertendo datas...")
    df['Data de criação'] = pd.to_datetime(df['Data de criação'], errors='coerce')
    df['Data de fechamento'] = pd.to_datetime(df['Data de fechamento'], errors='coerce')
    df['Data da última atividade'] = pd.to_datetime(df['Data da última atividade'], errors='coerce')
    return df

def filtrar_ciclo(df, inicio, fim):
    """Filtra leads do ciclo especificado"""
    inicio_dt = pd.to_datetime(inicio)
    fim_dt = pd.to_datetime(fim)
    
    df_ciclo = df[(df['Data de criação'] >= inicio_dt) & (df['Data de criação'] <= fim_dt)].copy()
    print(f"📊 Leads do ciclo ({inicio} a {fim}): {len(df_ciclo):,}")
    return df_ciclo

def limpar_base(df):
    """Remove registros com erros de preenchimento"""
    print("🧹 Limpando base...")
    
    total_inicial = len(df)
    
    erro_fechamento = df['Data de fechamento'] < df['Data de criação']
    erro_atividade_nula = df['Data da última atividade'].isna()
    erro_atividade_anterior = df['Data da última atividade'] < df['Data de criação']
    
    print(f"   ❌ Data fechamento < criação: {erro_fechamento.sum():,}")
    print(f"   ❌ Última atividade nula: {erro_atividade_nula.sum():,}")
    print(f"   ❌ Última atividade < criação: {erro_atividade_anterior.sum():,}")
    
    df_limpo = df[~erro_fechamento & ~erro_atividade_nula & ~erro_atividade_anterior].copy()
    
    removidos = total_inicial - len(df_limpo)
    print(f"   ✅ Base limpa: {len(df_limpo):,} ({removidos:,} removidos)")
    
    return df_limpo

def extrair_idade_aluno(nome):
    """Extrai a idade do aluno do campo 'Nome do negócio'"""
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
    """Classifica o produto baseado na idade do aluno"""
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
    """Retorna o preço da mensalidade 04h para a unidade"""
    if pd.isna(unidade):
        return PRECO_PADRAO
    
    unidade_upper = str(unidade).upper().strip()
    return PRECOS_UNIDADE.get(unidade_upper, PRECO_PADRAO)

def obter_preco_kit(produto):
    """Retorna o preço do kit didático baseado no produto"""
    if produto == 'Kids':
        return KIT_KIDS
    elif produto in ['Juniors', 'Early Teens', 'Late Teens']:
        return KIT_JUNIORS_TEENS
    else:
        return KIT_MEDIO  # ✅ CORRIGIDO: Kit médio para não informados

def calcular_recencia(df):
    """Calcula a idade do lead baseada na recência"""
    print("⏱️ Calculando recência...")
    df['idade_recencia'] = (df['Data da última atividade'] - df['Data de criação']).dt.days
    return df

def classificar_faixa_recencia(dias):
    """Classifica o lead por faixa de idade (recência)"""
    if pd.isna(dias):
        return None
    elif dias <= 1:
        return 'D-1'
    elif dias <= 7:
        return '2-7 dias'
    elif dias <= 30:
        return '8-30 dias'
    elif dias <= 90:
        return '31-90 dias'  # ✅ CORRIGIDO: Unificado 31-90
    else:
        return '> 90 dias'

def enriquecer_dados(df):
    """Extrai idade, produto, preços e cria flags"""
    print("🔧 Enriquecendo dados...")
    
    df['idade_aluno'] = df['Nome do negócio'].apply(extrair_idade_aluno)
    df['produto'] = df['idade_aluno'].apply(classificar_produto)
    df['mensalidade'] = df['Unidade Desejada'].apply(obter_preco_unidade)
    df['kit_didatico'] = df['produto'].apply(obter_preco_kit)
    df['receita_anual'] = (df['mensalidade'] * PARCELAS) + df['kit_didatico']
    df['faixa_recencia'] = df['idade_recencia'].apply(classificar_faixa_recencia)
    df['matricula'] = df['Etapa do negócio'].str.contains('MATRÍCULA CONCLUÍDA', na=False)
    df['em_qualificacao'] = df['Etapa do negócio'].str.contains('QUALIFICAÇÃO', na=False)
    
    com_idade = (df['idade_aluno'].notna()).sum()
    print(f"   ✅ Leads com idade extraída: {com_idade:,} ({com_idade/len(df)*100:.1f}%)")
    
    return df

def calcular_taxa_conversao(df):
    """Calcula taxa de conversão histórica por faixa de recência"""
    print("📈 Calculando taxas de conversão...")
    
    faixas_ordem = ['D-1', '2-7 dias', '8-30 dias', '31-90 dias', '> 90 dias']
    
    resultados = []
    
    for faixa in faixas_ordem:
        subset = df[df['faixa_recencia'] == faixa]
        leads = len(subset)
        matriculas = subset['matricula'].sum()
        taxa = (matriculas / leads) if leads > 0 else 0
        media_ativ = subset['Número de atividades de vendas'].mean()
        
        resultados.append({
            'Faixa Recência': faixa,
            'Total Leads': leads,
            'Matrículas': matriculas,
            'Taxa Conversão': taxa,
            'Média Atividades': media_ativ
        })
    
    return pd.DataFrame(resultados)

def calcular_projecao_por_faixa(df, df_taxas):
    """Calcula projeção de matrículas e receita por faixa de recência"""
    print("🎯 Calculando projeção por faixa de recência...")
    
    faixas_ordem = ['D-1', '2-7 dias', '8-30 dias', '31-90 dias', '> 90 dias']
    
    resultados = []
    
    for faixa in faixas_ordem:
        subset = df[(df['faixa_recencia'] == faixa) & (df['em_qualificacao'] == True)]
        em_qual = len(subset)
        
        taxa = df_taxas[df_taxas['Faixa Recência'] == faixa]['Taxa Conversão'].values[0]
        mat_esperadas = em_qual * taxa
        receita_media = subset['receita_anual'].mean() if len(subset) > 0 else 0
        receita_potencial = mat_esperadas * receita_media
        
        resultados.append({
            'Faixa Recência': faixa,
            'Em Qualificação': em_qual,
            'Taxa Conversão': taxa,
            'Matrículas Esperadas': mat_esperadas,
            'Receita Média/Lead': receita_media,
            'Receita Potencial': receita_potencial
        })
    
    return pd.DataFrame(resultados)

def calcular_projecao_por_produto(df, df_taxas):
    """Calcula projeção por produto (Kids/Juniors/Teens)"""
    print("📊 Calculando projeção por produto...")
    
    taxa_8_30 = df_taxas[df_taxas['Faixa Recência'] == '8-30 dias']['Taxa Conversão'].values[0]
    df_8_30_qual = df[(df['faixa_recencia'] == '8-30 dias') & (df['em_qualificacao'] == True)]
    
    resultados = []
    
    for produto in ['Kids', 'Juniors', 'Early Teens', 'Late Teens', 'Não informada']:
        subset = df_8_30_qual[df_8_30_qual['produto'] == produto]
        em_qual = len(subset)
        
        df_prod = df[(df['faixa_recencia'] == '8-30 dias') & (df['produto'] == produto)]
        mat_prod = df_prod['matricula'].sum()
        taxa_prod = (mat_prod / len(df_prod)) if len(df_prod) > 0 else taxa_8_30
        
        mat_esperadas = em_qual * taxa_prod
        
        mensalidade_media = subset['mensalidade'].mean() if len(subset) > 0 else PRECO_PADRAO
        kit_medio = subset['kit_didatico'].mean() if len(subset) > 0 else KIT_MEDIO
        receita_anual_media = (mensalidade_media * PARCELAS) + kit_medio
        receita_potencial = mat_esperadas * receita_anual_media
        
        resultados.append({
            'Produto': produto,
            'Em Qualificação': em_qual,
            'Taxa Conversão': taxa_prod,
            'Matrículas Esperadas': mat_esperadas,
            'Mensalidade Média': mensalidade_media,
            'Kit Médio': kit_medio,
            'Receita Anual/Lead': receita_anual_media,
            'Receita Potencial': receita_potencial
        })
    
    return pd.DataFrame(resultados)

def calcular_projecao_por_unidade(df, df_taxas):
    """Calcula projeção por unidade (top 20)"""
    print("🏫 Calculando projeção por unidade...")
    
    taxa_8_30 = df_taxas[df_taxas['Faixa Recência'] == '8-30 dias']['Taxa Conversão'].values[0]
    df_8_30_qual = df[(df['faixa_recencia'] == '8-30 dias') & (df['em_qualificacao'] == True)]
    
    resultados = []
    
    for unidade in df_8_30_qual['Unidade Desejada'].unique():
        subset = df_8_30_qual[df_8_30_qual['Unidade Desejada'] == unidade]
        em_qual = len(subset)
        
        if em_qual == 0:
            continue
        
        df_unid = df[(df['faixa_recencia'] == '8-30 dias') & (df['Unidade Desejada'] == unidade)]
        mat_unid = df_unid['matricula'].sum()
        taxa_unid = (mat_unid / len(df_unid)) if len(df_unid) > 0 else taxa_8_30
        
        mat_esperadas = em_qual * taxa_unid
        
        mensalidade = obter_preco_unidade(unidade)
        receita_anual_media = subset['receita_anual'].mean()
        receita_potencial = mat_esperadas * receita_anual_media
        
        resultados.append({
            'Unidade': unidade,
            'Em Qualificação': em_qual,
            'Taxa Conversão': taxa_unid,
            'Matrículas Esperadas': mat_esperadas,
            'Mensalidade 04h': mensalidade,
            'Receita Potencial': receita_potencial
        })
    
    df_result = pd.DataFrame(resultados)
    df_result = df_result.sort_values('Receita Potencial', ascending=False)
    
    return df_result

def calcular_projecao_por_canal(df, df_taxas):
    """Calcula projeção por canal de origem"""
    print("📢 Calculando projeção por canal...")
    
    df_8_30_qual = df[(df['faixa_recencia'] == '8-30 dias') & (df['em_qualificacao'] == True)]
    df_8_30_hist = df[df['faixa_recencia'] == '8-30 dias']
    
    resultados = []
    
    for canal in df_8_30_qual['Fonte original do tráfego'].unique():
        subset = df_8_30_qual[df_8_30_qual['Fonte original do tráfego'] == canal]
        em_qual = len(subset)
        
        if em_qual == 0:
            continue
        
        df_canal = df_8_30_hist[df_8_30_hist['Fonte original do tráfego'] == canal]
        mat_canal = df_canal['matricula'].sum()
        taxa_canal = (mat_canal / len(df_canal)) if len(df_canal) > 0 else 0
        
        mat_esperadas = em_qual * taxa_canal
        
        receita_anual_media = subset['receita_anual'].mean()
        receita_potencial = mat_esperadas * receita_anual_media
        
        resultados.append({
            'Canal': canal,
            'Em Qualificação': em_qual,
            'Taxa Conversão': taxa_canal,
            'Matrículas Esperadas': mat_esperadas,
            'Receita Média/Lead': receita_anual_media,
            'Receita Potencial': receita_potencial
        })
    
    df_result = pd.DataFrame(resultados)
    df_result = df_result.sort_values('Receita Potencial', ascending=False)
    
    return df_result

def formatar_mat_esperadas(valor):
    """Formata matrículas esperadas com asterisco se arredondado"""
    if valor == int(valor):
        return f"{int(valor)}"
    else:
        return f"{valor:.1f}*"

def gerar_resumo(df, df_projecao_faixa):
    """Gera resumo executivo"""
    
    total_leads = len(df)
    total_matriculas = df['matricula'].sum()
    
    f8_30 = df_projecao_faixa[df_projecao_faixa['Faixa Recência'] == '8-30 dias'].iloc[0]
    receita_media = df[df['em_qualificacao'] == True]['receita_anual'].mean()
    
    resumo = {
        'Métrica': [
            'Total de Leads (base limpa)',
            'Total de Matrículas Históricas',
            'Taxa Conversão Geral',
            'Receita Anual Média/Lead',
            '--- FAIXA PRIORITÁRIA: 8-30 DIAS ---',
            'Leads em Qualificação',
            'Taxa de Conversão',
            'Matrículas Esperadas',
            'Receita Potencial'
        ],
        'Valor': [
            f"{total_leads:,}",
            f"{total_matriculas:,}",
            f"{total_matriculas/total_leads*100:.1f}%",
            f"R$ {receita_media:,.2f}",
            '---',
            f"{f8_30['Em Qualificação']:,}",
            f"{f8_30['Taxa Conversão']*100:.1f}%",
            formatar_mat_esperadas(f8_30['Matrículas Esperadas']),
            f"R$ {f8_30['Receita Potencial']:,.2f}"
        ]
    }
    
    return pd.DataFrame(resumo)

def exportar_excel(output_dir, df_taxas, df_projecao_faixa, df_projecao_produto, 
                   df_projecao_unidade, df_projecao_canal, df_resumo):
    """Exporta resultados para Excel"""
    
    os.makedirs(output_dir, exist_ok=True)
    
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"projecao_funil_completa_{timestamp}.xlsx"
    filepath = os.path.join(output_dir, filename)
    
    print(f"\n💾 Exportando para: {filepath}")
    
    with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
        
        df_resumo.to_excel(writer, sheet_name='Resumo', index=False)
        
        df_taxas_fmt = df_taxas.copy()
        df_taxas_fmt['Taxa Conversão'] = df_taxas_fmt['Taxa Conversão'].apply(lambda x: f"{x*100:.1f}%")
        df_taxas_fmt['Média Atividades'] = df_taxas_fmt['Média Atividades'].apply(lambda x: f"{x:.1f}")
        df_taxas_fmt.to_excel(writer, sheet_name='Taxas Conversão', index=False)
        
        df_faixa_fmt = df_projecao_faixa.copy()
        df_faixa_fmt['Taxa Conversão'] = df_faixa_fmt['Taxa Conversão'].apply(lambda x: f"{x*100:.1f}%")
        df_faixa_fmt['Matrículas Esperadas'] = df_faixa_fmt['Matrículas Esperadas'].apply(formatar_mat_esperadas)
        df_faixa_fmt['Receita Média/Lead'] = df_faixa_fmt['Receita Média/Lead'].apply(lambda x: f"R$ {x:,.2f}")
        df_faixa_fmt['Receita Potencial'] = df_faixa_fmt['Receita Potencial'].apply(lambda x: f"R$ {x:,.2f}")
        df_faixa_fmt.to_excel(writer, sheet_name='Projeção por Faixa', index=False)
        
        df_prod_fmt = df_projecao_produto.copy()
        df_prod_fmt['Taxa Conversão'] = df_prod_fmt['Taxa Conversão'].apply(lambda x: f"{x*100:.1f}%")
        df_prod_fmt['Matrículas Esperadas'] = df_prod_fmt['Matrículas Esperadas'].apply(formatar_mat_esperadas)
        df_prod_fmt['Mensalidade Média'] = df_prod_fmt['Mensalidade Média'].apply(lambda x: f"R$ {x:,.2f}")
        df_prod_fmt['Kit Médio'] = df_prod_fmt['Kit Médio'].apply(lambda x: f"R$ {x:,.2f}")
        df_prod_fmt['Receita Anual/Lead'] = df_prod_fmt['Receita Anual/Lead'].apply(lambda x: f"R$ {x:,.2f}")
        df_prod_fmt['Receita Potencial'] = df_prod_fmt['Receita Potencial'].apply(lambda x: f"R$ {x:,.2f}")
        df_prod_fmt.to_excel(writer, sheet_name='Projeção por Produto', index=False)
        
        df_unid_fmt = df_projecao_unidade.head(20).copy()
        df_unid_fmt['Taxa Conversão'] = df_unid_fmt['Taxa Conversão'].apply(lambda x: f"{x*100:.1f}%")
        df_unid_fmt['Matrículas Esperadas'] = df_unid_fmt['Matrículas Esperadas'].apply(formatar_mat_esperadas)
        df_unid_fmt['Mensalidade 04h'] = df_unid_fmt['Mensalidade 04h'].apply(lambda x: f"R$ {x:,.2f}")
        df_unid_fmt['Receita Potencial'] = df_unid_fmt['Receita Potencial'].apply(lambda x: f"R$ {x:,.2f}")
        df_unid_fmt.to_excel(writer, sheet_name='Projeção por Unidade', index=False)
        
        df_canal_fmt = df_projecao_canal.copy()
        df_canal_fmt['Taxa Conversão'] = df_canal_fmt['Taxa Conversão'].apply(lambda x: f"{x*100:.1f}%")
        df_canal_fmt['Matrículas Esperadas'] = df_canal_fmt['Matrículas Esperadas'].apply(formatar_mat_esperadas)
        df_canal_fmt['Receita Média/Lead'] = df_canal_fmt['Receita Média/Lead'].apply(lambda x: f"R$ {x:,.2f}")
        df_canal_fmt['Receita Potencial'] = df_canal_fmt['Receita Potencial'].apply(lambda x: f"R$ {x:,.2f}")
        df_canal_fmt.to_excel(writer, sheet_name='Projeção por Canal', index=False)
    
    print(f"   ✅ Arquivo gerado com sucesso!")
    return filepath

# ============================================================================
# EXECUÇÃO PRINCIPAL
# ============================================================================

def main():
    print("\n" + "="*70)
    print("📊 PROJEÇÃO DE FUNIL RED BALLOON - VERSÃO COMPLETA")
    print("   Cruzamento: Idade → Produto | Unidade → Preço")
    print("   Kit médio para não informados: R$ {:.2f}".format(KIT_MEDIO))
    print("="*70 + "\n")
    
    df = carregar_dados(INPUT_PATH)
    df = preparar_datas(df)
    df_ciclo = filtrar_ciclo(df, INICIO_CICLO, FIM_CICLO)
    df_limpo = limpar_base(df_ciclo)
    df_limpo = calcular_recencia(df_limpo)
    df_limpo = enriquecer_dados(df_limpo)
    
    df_taxas = calcular_taxa_conversao(df_limpo)
    df_projecao_faixa = calcular_projecao_por_faixa(df_limpo, df_taxas)
    df_projecao_produto = calcular_projecao_por_produto(df_limpo, df_taxas)
    df_projecao_unidade = calcular_projecao_por_unidade(df_limpo, df_taxas)
    df_projecao_canal = calcular_projecao_por_canal(df_limpo, df_taxas)
    df_resumo = gerar_resumo(df_limpo, df_projecao_faixa)
    
    print("\n" + "="*70)
    print("📋 TAXAS DE CONVERSÃO POR FAIXA:")
    print("="*70)
    print(df_taxas.to_string(index=False))
    
    print("\n" + "="*70)
    print("🎯 PROJEÇÃO POR FAIXA DE RECÊNCIA:")
    print("="*70)
    print(df_projecao_faixa.to_string(index=False))
    
    print("\n" + "="*70)
    print("🎒 PROJEÇÃO POR PRODUTO (Faixa 8-30 dias):")
    print("="*70)
    print(df_projecao_produto.to_string(index=False))
    
    print("\n" + "="*70)
    print("🏫 PROJEÇÃO POR UNIDADE - TOP 10 (Faixa 8-30 dias):")
    print("="*70)
    print(df_projecao_unidade.head(10).to_string(index=False))
    
    print("\n" + "="*70)
    print("📢 PROJEÇÃO POR CANAL (Faixa 8-30 dias):")
    print("="*70)
    print(df_projecao_canal.to_string(index=False))
    
    filepath = exportar_excel(
        OUTPUT_DIR, 
        df_taxas, 
        df_projecao_faixa, 
        df_projecao_produto,
        df_projecao_unidade,
        df_projecao_canal, 
        df_resumo
    )
    
    print("\n" + "="*70)
    print("✅ PROCESSAMENTO CONCLUÍDO!")
    print("   * Valores com asterisco indicam arredondamento")
    print("="*70 + "\n")
    
    return filepath

if __name__ == "__main__":
    main()