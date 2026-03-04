"""
Dashboard Executivo - Red Balloon
Autor: Analista de Dados Sênior
Data: 2026-02-11

Objetivo: Criar relatório executivo visual para tomada de decisão estratégica
Público: Diretoria (Marcelo)
"""

import pandas as pd
import xlsxwriter
from pathlib import Path
from datetime import datetime

# ============================================================================
# CONFIGURAÇÕES
# ============================================================================

# Define o diretório de trabalho
PASTA_OUTPUTS = Path(r"c:\Users\a483650\Projetos\06_Analise_Funil_RedBalloon\outputs")

# Arquivo Excel fonte (com todas as abas)
ARQUIVO_EXCEL_FONTE = "funil_redballoon_20260211_174832.xlsx"

# Nomes das abas necessárias
ABA_PROJECAO = "10f_Projecao_Realista"
ABA_PERFORMANCE = "4_Performance_Fonte"
ABA_PRIORIZACAO = "10g_Priorizacao_Leads"
ABA_ATENDIMENTO = "10c_Taxa_Atendimento"

# Arquivo de saída
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
ARQUIVO_SAIDA = f"Dashboard_Executivo_RedBalloon_{timestamp}.xlsx"


# ============================================================================
# FUNÇÕES DE LEITURA E PROCESSAMENTO
# ============================================================================

def ler_projecao():
    """Lê dados de projeção e extrai informações do veredito"""
    df = pd.read_excel(PASTA_OUTPUTS / ARQUIVO_EXCEL_FONTE, sheet_name=ABA_PROJECAO)
    
    # Filtra linha de Segmentação
    linha_seg = df[df['Metodo'].str.contains('Segmentacao', case=False, na=False)]
    
    if not linha_seg.empty:
        row = linha_seg.iloc[0]
        return {
            'leads_atuais': int(row['Leads_Ciclo_26_1']),
            'meta_faltante': int(row['Meta_Faltante']),
            'gap_projetado': int(row['Gap'])
        }
    
    # Fallback: usa primeira linha
    row = df.iloc[0]
    return {
        'leads_atuais': int(row['Leads_Ciclo_26_1']),
        'meta_faltante': int(row['Meta_Faltante']),
        'gap_projetado': int(row['Gap'])
    }


def ler_performance_canais():
    """Lê performance por fonte/canal"""
    df = pd.read_excel(PASTA_OUTPUTS / ARQUIVO_EXCEL_FONTE, sheet_name=ABA_PERFORMANCE)
    
    # Filtra apenas os canais principais
    canais_principais = ['Meta Ads', 'Google Ads', 'Offline', 'Direto']
    df_filtrado = df[df['Fonte_Categoria'].isin(canais_principais)].copy()
    
    # Cria DataFrame resultado com as colunas corretas
    df_resultado = pd.DataFrame({
        'Canal': df_filtrado['Fonte_Categoria'],
        'Volume': df_filtrado['Total_Leads'].astype(int),
        'Conversão %': df_filtrado['Taxa_Conversao_%'].astype(float)
    })
    
    # Ordena por volume
    df_resultado = df_resultado.sort_values('Volume', ascending=False).reset_index(drop=True)
    
    return df_resultado


def ler_priorizacao():
    """Lê dados de priorização de leads"""
    df = pd.read_excel(PASTA_OUTPUTS / ARQUIVO_EXCEL_FONTE, sheet_name=ABA_PRIORIZACAO)
    
    # Lê os dados já agregados da aba
    resumo = pd.DataFrame({
        'Prioridade': df['Prioridade'].str.title(),  # Alta, Media, Baixa
        'Quantidade': df['Quantidade_Leads'].astype(int),
        '% do Total': df['Percentual_Total_%'].astype(float)
    })
    
    # Ordena: Alta, Média, Baixa
    ordem = {'Alta': 1, 'Media': 2, 'Baixa': 3}
    resumo['ordem'] = resumo['Prioridade'].map(ordem)
    resumo = resumo.sort_values('ordem').drop('ordem', axis=1).reset_index(drop=True)
    
    # Substitui 'Media' por 'Média'
    resumo['Prioridade'] = resumo['Prioridade'].replace('Media', 'Média')
    
    return resumo


def ler_temperatura():
    """Lê dados de temperatura/envelhecimento dos leads"""
    df = pd.read_excel(PASTA_OUTPUTS / ARQUIVO_EXCEL_FONTE, sheet_name=ABA_ATENDIMENTO)
    
    # Filtra apenas as categorias mais relevantes
    categorias_importantes = ['Ativos (0-7 dias)', 'Muito Frios (>90 dias)']
    df_filtrado = df[df['Faixa_Tempo'].isin(categorias_importantes)].copy()
    
    # Cria DataFrame resultado
    resumo = pd.DataFrame({
        'Categoria': df_filtrado['Faixa_Tempo'],
        'Quantidade': df_filtrado['Quantidade'].astype(int),
        '% do Total': df_filtrado['Percentual'].astype(float)
    }).reset_index(drop=True)
    
    return resumo


# ============================================================================
# FUNÇÃO PRINCIPAL - CRIAÇÃO DO DASHBOARD
# ============================================================================

def criar_dashboard():
    """Cria o Dashboard Executivo em Excel"""
    
    print("=" * 80)
    print(" CRIANDO DASHBOARD EXECUTIVO - RED BALLOON")
    print("=" * 80)
    
    # Lê os dados
    print("\n[1/5] Lendo dados de projeção...")
    dados_veredito = ler_projecao()
    
    print("[2/5] Lendo performance de canais...")
    df_performance = ler_performance_canais()
    
    print("[3/5] Lendo priorização de leads...")
    df_priorizacao = ler_priorizacao()
    
    print("[4/5] Lendo temperatura da base...")
    df_temperatura = ler_temperatura()
    
    # Cria o workbook
    print("[5/5] Gerando arquivo Excel...")
    caminho_saida = PASTA_OUTPUTS / ARQUIVO_SAIDA
    workbook = xlsxwriter.Workbook(str(caminho_saida))
    worksheet = workbook.add_worksheet('VISÃO EXECUTIVA')
    
    # Esconde as linhas de grade
    worksheet.hide_gridlines(2)
    
    # ========================================================================
    # FORMATOS
    # ========================================================================
    
    # Título principal
    fmt_titulo = workbook.add_format({
        'bold': True,
        'font_size': 18,
        'font_color': '#FFFFFF',
        'bg_color': '#1F4E78',
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })
    
    # Cabeçalho de seção
    fmt_secao = workbook.add_format({
        'bold': True,
        'font_size': 14,
        'font_color': '#1F4E78',
        'bg_color': '#D9E2F3',
        'align': 'left',
        'valign': 'vcenter',
        'border': 1
    })
    
    # Cabeçalho de tabela
    fmt_cab_tabela = workbook.add_format({
        'bold': True,
        'font_color': '#FFFFFF',
        'bg_color': '#1F4E78',
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })
    
    # Célula normal
    fmt_celula = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })
    
    # Número
    fmt_numero = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'num_format': '#,##0'
    })
    
    # Percentual
    fmt_percentual = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'num_format': '0.0"%"'
    })
    
    # Alerta vermelho
    fmt_alerta = workbook.add_format({
        'bold': True,
        'font_size': 12,
        'font_color': '#C00000',
        'bg_color': '#FFC7CE',
        'align': 'center',
        'valign': 'vcenter',
        'border': 2
    })
    
    # Sucesso verde
    fmt_sucesso = workbook.add_format({
        'bold': True,
        'font_size': 12,
        'font_color': '#006100',
        'bg_color': '#C6EFCE',
        'align': 'center',
        'valign': 'vcenter',
        'border': 2
    })
    
    # Alta prioridade (verde claro)
    fmt_alta_prior = workbook.add_format({
        'bg_color': '#C6EFCE',
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })
    
    # Baixa prioridade (cinza)
    fmt_baixa_prior = workbook.add_format({
        'bg_color': '#D9D9D9',
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })
    
    # Percentual com cores
    fmt_perc_alta = workbook.add_format({
        'bg_color': '#C6EFCE',
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'num_format': '0.0"%"'
    })
    
    fmt_perc_baixa = workbook.add_format({
        'bg_color': '#D9D9D9',
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'num_format': '0.0"%"'
    })
    
    fmt_perc_medio = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'num_format': '0.0"%"'
    })
    
    # ========================================================================
    # LAYOUT DO DASHBOARD
    # ========================================================================
    
    linha_atual = 0
    
    # TÍTULO PRINCIPAL
    worksheet.merge_range(linha_atual, 0, linha_atual, 7, 
                         'DIAGNÓSTICO ESTRATÉGICO - CICLO 26.1', fmt_titulo)
    worksheet.set_row(linha_atual, 30)
    linha_atual += 2
    
    # ========================================================================
    # BLOCO 1: O VEREDITO
    # ========================================================================
    
    worksheet.merge_range(linha_atual, 0, linha_atual, 7,
                         '📊 BLOCO 1: O VEREDITO - Precisamos de novos leads ou o estoque resolve?',
                         fmt_secao)
    worksheet.set_row(linha_atual, 20)
    linha_atual += 1
    
    # Cabeçalhos
    worksheet.write(linha_atual, 1, 'Leads Atuais (Estoque)', fmt_cab_tabela)
    worksheet.write(linha_atual, 2, 'Meta Faltante', fmt_cab_tabela)
    worksheet.write(linha_atual, 3, 'Gap Projetado', fmt_cab_tabela)
    linha_atual += 1
    
    # Valores
    worksheet.write_number(linha_atual, 1, dados_veredito['leads_atuais'], fmt_numero)
    worksheet.write_number(linha_atual, 2, dados_veredito['meta_faltante'], fmt_numero)
    worksheet.write_number(linha_atual, 3, dados_veredito['gap_projetado'], fmt_numero)
    linha_atual += 1
    
    # Conclusão (Gap positivo = faltam matrículas, negativo = sobram)
    gap = dados_veredito['gap_projetado']
    if gap > 0:
        conclusao = f"⚠️ RISCO: Estoque Insuficiente. Faltam {int(gap)} matrículas."
        fmt_conclusao = fmt_alerta
    else:
        conclusao = f"✓ OK: Estoque Suficiente. Margem de {abs(int(gap))} matrículas."
        fmt_conclusao = fmt_sucesso
    
    worksheet.merge_range(linha_atual, 0, linha_atual, 7, conclusao, fmt_conclusao)
    worksheet.set_row(linha_atual, 25)
    linha_atual += 3
    
    # ========================================================================
    # BLOCOS 2 e 3 - LADO A LADO
    # ========================================================================
    
    linha_bloco2 = linha_atual
    
    # BLOCO 2: CAUSA RAIZ (Colunas 0-3)
    worksheet.merge_range(linha_atual, 0, linha_atual, 3,
                         '🔍 BLOCO 2: CAUSA RAIZ - Por que a conversão caiu?',
                         fmt_secao)
    worksheet.set_row(linha_atual, 20)
    linha_atual += 1
    
    # Cabeçalhos
    worksheet.write(linha_atual, 0, 'Canal', fmt_cab_tabela)
    worksheet.write(linha_atual, 1, 'Volume', fmt_cab_tabela)
    worksheet.write(linha_atual, 2, 'Conversão %', fmt_cab_tabela)
    inicio_tabela_perf = linha_atual
    linha_atual += 1
    
    # Dados de performance
    for idx, row in df_performance.iterrows():
        worksheet.write(linha_atual, 0, row['Canal'], fmt_celula)
        worksheet.write_number(linha_atual, 1, row['Volume'], fmt_numero)
        
        # Conversão com formatação condicional
        conversao = row['Conversão %']
        if conversao < 5:
            fmt_conv = fmt_alerta
        elif conversao > 7:
            fmt_conv = fmt_sucesso
        else:
            fmt_conv = fmt_percentual
        
        worksheet.write_number(linha_atual, 2, conversao / 100, fmt_conv)
        linha_atual += 1
    
    fim_bloco2 = linha_atual
    
    # BLOCO 3: SOLUÇÃO/AÇÃO (Colunas 4-7)
    linha_atual = linha_bloco2
    worksheet.merge_range(linha_atual, 4, linha_atual, 7,
                         '🎯 BLOCO 3: AÇÃO IMEDIATA - Onde o time deve focar?',
                         fmt_secao)
    worksheet.set_row(linha_atual, 20)
    linha_atual += 1
    
    # Cabeçalhos
    worksheet.write(linha_atual, 4, 'Prioridade', fmt_cab_tabela)
    worksheet.write(linha_atual, 5, 'Quantidade', fmt_cab_tabela)
    worksheet.write(linha_atual, 6, '% do Total', fmt_cab_tabela)
    linha_atual += 1
    
    # Dados de priorização
    for idx, row in df_priorizacao.iterrows():
        prioridade = row['Prioridade']
        
        # Define formato baseado na prioridade
        if prioridade == 'Alta':
            fmt_linha = fmt_alta_prior
            fmt_perc = fmt_perc_alta
        elif prioridade == 'Baixa':
            fmt_linha = fmt_baixa_prior
            fmt_perc = fmt_perc_baixa
        else:
            fmt_linha = fmt_celula
            fmt_perc = fmt_perc_medio
        
        worksheet.write(linha_atual, 4, prioridade, fmt_linha)
        worksheet.write_number(linha_atual, 5, row['Quantidade'], fmt_linha)
        worksheet.write_number(linha_atual, 6, row['% do Total'] / 100, fmt_perc)
        linha_atual += 1
    
    # Sincroniza as linhas
    linha_atual = max(fim_bloco2, linha_atual) + 2
    
    # ========================================================================
    # BLOCO 4: TEMPERATURA DA BASE
    # ========================================================================
    
    worksheet.merge_range(linha_atual, 0, linha_atual, 7,
                         '🌡️ BLOCO 4: TEMPERATURA DA BASE - Como está o envelhecimento?',
                         fmt_secao)
    worksheet.set_row(linha_atual, 20)
    linha_atual += 1
    
    # Cabeçalhos
    worksheet.write(linha_atual, 1, 'Categoria', fmt_cab_tabela)
    worksheet.write(linha_atual, 2, 'Quantidade', fmt_cab_tabela)
    worksheet.write(linha_atual, 3, '% do Total', fmt_cab_tabela)
    linha_atual += 1
    
    # Dados de temperatura
    for idx, row in df_temperatura.iterrows():
        worksheet.write(linha_atual, 1, row['Categoria'], fmt_celula)
        worksheet.write_number(linha_atual, 2, row['Quantidade'], fmt_numero)
        worksheet.write_number(linha_atual, 3, row['% do Total'] / 100, fmt_percentual)
        linha_atual += 1
    
    # ========================================================================
    # AJUSTES FINAIS
    # ========================================================================
    
    # Ajusta largura das colunas
    worksheet.set_column(0, 0, 20)  # Canal/Categoria
    worksheet.set_column(1, 3, 18)  # Dados numéricos
    worksheet.set_column(4, 4, 15)  # Prioridade
    worksheet.set_column(5, 6, 15)  # Quantidade e %
    worksheet.set_column(7, 7, 15)  # Coluna extra
    
    # Fecha o workbook
    workbook.close()
    
    print(f"\n✓ Dashboard criado com sucesso!")
    print(f"📁 Arquivo: {caminho_saida}")
    print(f"📊 Aba: VISÃO EXECUTIVA")
    print("\n" + "=" * 80)
    
    return caminho_saida


# ============================================================================
# EXECUÇÃO PRINCIPAL
# ============================================================================

if __name__ == "__main__":
    try:
        arquivo_criado = criar_dashboard()
        print("\n✓ PROCESSO CONCLUÍDO COM SUCESSO!")
        print(f"\nAbra o arquivo: {arquivo_criado.name}")
        print("Para apresentar à diretoria.")
        
    except FileNotFoundError as e:
        print(f"\n❌ ERRO: Arquivo não encontrado.")
        print(f"Detalhes: {e}")
        print("\nVerifique se os CSVs estão na pasta 'outputs'.")
        
    except Exception as e:
        print(f"\n❌ ERRO INESPERADO: {e}")
        import traceback
        traceback.print_exc()
