"""
Respostas Diretas para o Marcelo
Baseado na Análise do Funil RedBalloon
"""

import pandas as pd
import xlsxwriter
from pathlib import Path
from datetime import datetime

# ============================================================================
# CONFIGURAÇÕES
# ============================================================================

PASTA_OUTPUTS = Path(r"c:\Users\a483650\Projetos\06_Analise_Funil_RedBalloon\outputs")
ARQUIVO_FONTE = "funil_redballoon_20260211_174832.xlsx"

# ============================================================================
# LEITURA DOS DADOS
# ============================================================================

def ler_dados():
    """Lê todas as abas necessárias"""
    arquivo = PASTA_OUTPUTS / ARQUIVO_FONTE
    
    # Lê dados do comparativo (ciclos 26.1 vs 25.1) - DADOS ESPECÍFICOS POR CICLO
    df_comparativo = pd.read_excel(arquivo, sheet_name="7b_Comparativo_26vs25")
    
    # Filtra apenas ciclo 26.1
    performance_261 = df_comparativo[df_comparativo['Ciclo'] == '26.1'].copy()
    performance_261 = performance_261[~performance_261['Canal'].isna()].copy()
    
    dados = {
        'projecao': pd.read_excel(arquivo, sheet_name="10f_Projecao_Realista"),
        'performance': performance_261,  # AGORA USA DADOS DO CICLO 26.1 APENAS
        'priorizacao': pd.read_excel(arquivo, sheet_name="10g_Priorizacao_Leads"),
        'temperatura': pd.read_excel(arquivo, sheet_name="10c_Taxa_Atendimento"),
        'status': pd.read_excel(arquivo, sheet_name="2_Status_Atual"),
        'atividades': pd.read_excel(arquivo, sheet_name="10d_Distribuicao_Atividades"),
    }
    
    return dados


# ============================================================================
# CRIAÇÃO DO RELATÓRIO
# ============================================================================

def criar_relatorio():
    """Cria relatório direto com 2 abas"""
    
    print("=" * 80)
    print(" GERANDO RESPOSTAS DIRETAS PARA MARCELO")
    print("=" * 80)
    
    # Lê os dados
    print("\n[1/3] Lendo dados...")
    dados = ler_dados()
    
    # Cria o arquivo
    print("[2/3] Criando relatório...")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    arquivo_saida = PASTA_OUTPUTS / f"Respostas_Diretas_Marcelo_{timestamp}.xlsx"
    
    workbook = xlsxwriter.Workbook(str(arquivo_saida))
    
    # ========================================================================
    # FORMATOS
    # ========================================================================
    
    fmt_titulo = workbook.add_format({
        'bold': True,
        'font_size': 16,
        'font_color': '#FFFFFF',
        'bg_color': '#1F4E78',
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })
    
    fmt_pergunta = workbook.add_format({
        'bold': True,
        'font_size': 12,
        'font_color': '#1F4E78',
        'bg_color': '#E7E6E6',
        'align': 'left',
        'valign': 'vcenter',
        'border': 1,
        'text_wrap': True
    })
    
    fmt_cabecalho = workbook.add_format({
        'bold': True,
        'font_color': '#FFFFFF',
        'bg_color': '#4472C4',
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })
    
    fmt_celula = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })
    
    fmt_numero = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'num_format': '#,##0'
    })
    
    fmt_percentual = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'num_format': '0.0"%"'
    })
    
    fmt_alerta_vermelho = workbook.add_format({
        'bold': True,
        'font_color': '#C00000',
        'bg_color': '#FFC7CE',
        'align': 'center',
        'valign': 'vcenter',
        'border': 2
    })
    
    fmt_sucesso_verde = workbook.add_format({
        'bold': True,
        'font_color': '#006100',
        'bg_color': '#C6EFCE',
        'align': 'center',
        'valign': 'vcenter',
        'border': 2
    })
    
    fmt_destaque = workbook.add_format({
        'bold': True,
        'font_size': 14,
        'font_color': '#1F4E78',
        'bg_color': '#FFF2CC',
        'align': 'left',
        'valign': 'vcenter',
        'border': 2,
        'text_wrap': True
    })
    
    # Formatos para priorização com percentual
    fmt_perc_alta = workbook.add_format({
        'bg_color': '#C6EFCE',
        'align': 'center',
        'border': 1,
        'num_format': '0.0"%"'
    })
    
    fmt_perc_baixa = workbook.add_format({
        'bg_color': '#FFC7CE',
        'align': 'center',
        'border': 1,
        'num_format': '0.0"%"'
    })
    
    fmt_perc_media = workbook.add_format({
        'align': 'center',
        'border': 1,
        'num_format': '0.0"%"'
    })
    
    # ========================================================================
    # ABA 1: DIAGNÓSTICO DOS LEADS
    # ========================================================================
    
    ws1 = workbook.add_worksheet('1_DIAGNÓSTICO DOS LEADS')
    ws1.hide_gridlines(2)
    
    linha = 0
    
    # Título
    ws1.merge_range(linha, 0, linha, 5, '📊 DIAGNÓSTICO DOS LEADS ATUAIS - CICLO 26.1', fmt_titulo)
    ws1.set_row(linha, 25)
    linha += 2
    
    # ========================================================================
    # PERGUNTA 1: QUALIDADE DOS LEADS
    # ========================================================================
    
    ws1.merge_range(linha, 0, linha, 5, '1️⃣ Qual a qualidade dos leads por canal?', fmt_pergunta)
    ws1.set_row(linha, 25)
    linha += 1
    
    # Cabeçalhos
    ws1.write(linha, 0, 'Canal', fmt_cabecalho)
    ws1.write(linha, 1, 'Total de Leads', fmt_cabecalho)
    ws1.write(linha, 2, 'Matrículas', fmt_cabecalho)
    ws1.write(linha, 3, 'Taxa Conversão', fmt_cabecalho)
    ws1.write(linha, 4, 'Qualidade', fmt_cabecalho)
    linha += 1
    
    # Dados de performance (CICLO 26.1 APENAS)
    df_perf = dados['performance'].sort_values('Taxa_Conversao_%', ascending=False)
    for idx, row in df_perf.iterrows():
        taxa = row['Taxa_Conversao_%']
        
        # Define qualidade
        if taxa >= 40:
            qualidade = "⭐⭐⭐ EXCELENTE"
            fmt_qual = fmt_sucesso_verde
        elif taxa >= 10:
            qualidade = "⭐⭐ BOA"
            fmt_qual = fmt_celula
        elif taxa >= 5:
            qualidade = "⭐ FRACA"
            fmt_qual = fmt_celula
        else:
            qualidade = "❌ PÉSSIMA"
            fmt_qual = fmt_alerta_vermelho
        
        ws1.write(linha, 0, row['Canal'], fmt_celula)
        ws1.write_number(linha, 1, int(row['Total_Leads']), fmt_numero)
        ws1.write_number(linha, 2, int(row['Matriculados']), fmt_numero)
        ws1.write_number(linha, 3, taxa / 100, fmt_percentual)
        ws1.write(linha, 4, qualidade, fmt_qual)
        linha += 1
    
    linha += 1
    
    # ========================================================================
    # PERGUNTA 2: IDADE/TEMPO DOS LEADS
    # ========================================================================
    
    ws1.merge_range(linha, 0, linha, 5, '2️⃣ Há quanto tempo esses leads existem?', fmt_pergunta)
    ws1.set_row(linha, 25)
    linha += 1
    
    # Cabeçalhos
    ws1.write(linha, 0, 'Idade do Lead', fmt_cabecalho)
    ws1.write(linha, 1, 'Quantidade', fmt_cabecalho)
    ws1.write(linha, 2, '% do Total', fmt_cabecalho)
    ws1.write(linha, 3, 'Status', fmt_cabecalho)
    linha += 1
    
    # Dados de temperatura
    df_temp = dados['temperatura']
    for idx, row in df_temp.iterrows():
        faixa = row['Faixa_Tempo']
        qtd = row['Quantidade']
        perc = row['Percentual']
        
        # Define status
        if 'Ativos' in faixa or 'Recentes' in faixa:
            status = "🔥 QUENTE - ATACAR AGORA"
            fmt_status = fmt_sucesso_verde
        elif 'Mornos' in faixa:
            status = "⚠️ MORNO - AQUECER"
            fmt_status = fmt_celula
        elif 'Nunca' in faixa:
            status = "❌ NUNCA CONTATADO"
            fmt_status = fmt_alerta_vermelho
        else:
            status = "❄️ FRIO - BAIXA PRIORIDADE"
            fmt_status = fmt_celula
        
        ws1.write(linha, 0, faixa, fmt_celula)
        ws1.write_number(linha, 1, qtd, fmt_numero)
        ws1.write_number(linha, 2, perc / 100, fmt_percentual)
        ws1.write(linha, 3, status, fmt_status)
        linha += 1
    
    linha += 1
    
    # ========================================================================
    # PERGUNTA 3: ATIVIDADES/VISITAS
    # ========================================================================
    
    ws1.merge_range(linha, 0, linha, 5, '3️⃣ Quais leads tiveram atividades/visitas?', fmt_pergunta)
    ws1.set_row(linha, 25)
    linha += 1
    
    # Cabeçalhos
    ws1.write(linha, 0, 'Tipo de Atividade', fmt_cabecalho)
    ws1.write(linha, 1, 'Quantidade', fmt_cabecalho)
    ws1.write(linha, 2, '% do Total', fmt_cabecalho)
    linha += 1
    
    # Dados de atividades
    df_ativ = dados['atividades']
    for idx, row in df_ativ.iterrows():
        ws1.write(linha, 0, row['Faixa_Atividades'], fmt_celula)
        ws1.write_number(linha, 1, row['Quantidade'], fmt_numero)
        ws1.write_number(linha, 2, row['Percentual'] / 100, fmt_percentual)
        linha += 1
    
    # Ajusta larguras
    ws1.set_column(0, 0, 25)
    ws1.set_column(1, 1, 15)
    ws1.set_column(2, 2, 15)
    ws1.set_column(3, 3, 15)
    ws1.set_column(4, 4, 30)
    ws1.set_column(5, 5, 15)
    
    # ========================================================================
    # ABA 2: DECISÃO - NOVOS LEADS?
    # ========================================================================
    
    ws2 = workbook.add_worksheet('2_DECISÃO NOVOS LEADS')
    ws2.hide_gridlines(2)
    
    linha = 0
    
    # Título
    ws2.merge_range(linha, 0, linha, 4, '🎯 PRECISAMOS DE NOVOS LEADS?', fmt_titulo)
    ws2.set_row(linha, 25)
    linha += 2
    
    # ========================================================================
    # ANÁLISE DA PROJEÇÃO
    # ========================================================================
    
    ws2.merge_range(linha, 0, linha, 4, 
                   'Pergunta: O estoque atual é suficiente para fechar a meta?', 
                   fmt_pergunta)
    ws2.set_row(linha, 30)
    linha += 1
    
    # Pega dados da projeção (método Segmentação)
    df_proj = dados['projecao']
    row_seg = df_proj[df_proj['Metodo'].str.contains('Segmentacao', case=False, na=False)].iloc[0]
    
    leads_ciclo = int(row_seg['Leads_Ciclo_26_1'])
    leads_estoque = int(row_seg['Leads_Estoque_Ciclos_Anteriores'])
    leads_total = leads_ciclo + leads_estoque
    meta_faltante = int(row_seg['Meta_Faltante'])
    gap = int(row_seg['Gap'])
    
    # Cabeçalhos
    ws2.write(linha, 0, 'Métrica', fmt_cabecalho)
    ws2.write(linha, 1, 'Valor', fmt_cabecalho)
    linha += 1
    
    # Dados
    ws2.write(linha, 0, 'Leads do Ciclo 26.1 (novos)', fmt_celula)
    ws2.write_number(linha, 1, leads_ciclo, fmt_numero)
    linha += 1
    
    ws2.write(linha, 0, 'Leads de Estoque (ciclos anteriores)', fmt_celula)
    ws2.write_number(linha, 1, leads_estoque, fmt_numero)
    linha += 1
    
    ws2.write(linha, 0, 'TOTAL DE LEADS DISPONÍVEIS', fmt_cabecalho)
    ws2.write_number(linha, 1, leads_total, fmt_numero)
    linha += 1
    
    ws2.write(linha, 0, 'Meta que falta atingir', fmt_celula)
    ws2.write_number(linha, 1, meta_faltante, fmt_numero)
    linha += 2
    
    ws2.write(linha, 0, 'GAP (+ = falta, - = sobra)', fmt_cabecalho)
    if gap > 0:
        ws2.write(linha, 1, f"FALTAM {gap} matrículas", fmt_alerta_vermelho)
    else:
        ws2.write(linha, 1, f"SOBRAM {abs(gap)} matrículas", fmt_sucesso_verde)
    linha += 3
    
    # ========================================================================
    # RESPOSTA DIRETA
    # ========================================================================
    
    ws2.merge_range(linha, 0, linha, 4, '📌 RESPOSTA DIRETA:', fmt_pergunta)
    ws2.set_row(linha, 25)
    linha += 1
    
    if gap > 0:
        # Precisa de novos leads
        resposta = f"❌ SIM, PRECISAMOS DE NOVOS LEADS!\n\n" \
                  f"• Faltam {gap} matrículas para fechar a meta\n" \
                  f"• O estoque atual ({leads_total:,} leads) é INSUFICIENTE\n\n" \
                  f"AÇÕES RECOMENDADAS:\n" \
                  f"1. URGENTE: Acelerar captação de novos leads\n" \
                  f"2. Focar em canais de alta qualidade (Offline e Direto)\n" \
                  f"3. Reaquecer leads frios do estoque"
        fmt_resposta = fmt_alerta_vermelho
    else:
        # NÃO precisa de novos leads
        resposta = f"✅ NÃO, O ESTOQUE ATUAL RESOLVE!\n\n" \
                  f"• Temos margem de {abs(gap)} matrículas acima da meta\n" \
                  f"• O estoque atual ({leads_total:,} leads) é SUFICIENTE\n\n" \
                  f"AÇÕES RECOMENDADAS:\n" \
                  f"1. FOCAR: Converter os leads existentes\n" \
                  f"2. Priorizar leads quentes (0-30 dias)\n" \
                  f"3. Atacar leads de alta qualidade primeiro"
        fmt_resposta = fmt_sucesso_verde
    
    ws2.merge_range(linha, 0, linha + 7, 4, resposta, fmt_destaque)
    ws2.set_row(linha, 25)
    for i in range(1, 8):
        ws2.set_row(linha + i, 20)
    linha += 9
    
    # ========================================================================
    # PRIORIZAÇÃO DO TIME
    # ========================================================================
    
    ws2.merge_range(linha, 0, linha, 4, 
                   '🎯 Onde o time comercial deve focar?', 
                   fmt_pergunta)
    ws2.set_row(linha, 25)
    linha += 1
    
    # Cabeçalhos
    ws2.write(linha, 0, 'Prioridade', fmt_cabecalho)
    ws2.write(linha, 1, 'Quantidade', fmt_cabecalho)
    ws2.write(linha, 2, '% do Total', fmt_cabecalho)
    ws2.write(linha, 3, 'Ação', fmt_cabecalho)
    linha += 1
    
    # Dados de priorização
    df_prior = dados['priorizacao']
    acoes = {
        'ALTA': '🔥 ATACAR AGORA - SÃO OS MELHORES',
        'MEDIA': '⚠️ LIGAR NA SEQUÊNCIA',
        'BAIXA': '❄️ ÚLTIMO RECURSO OU DESCARTAR'
    }
    
    for idx, row in df_prior.iterrows():
        prioridade = row['Prioridade']
        
        if prioridade == 'ALTA':
            fmt_linha = fmt_sucesso_verde
            fmt_perc_linha = fmt_perc_alta
        elif prioridade == 'BAIXA':
            fmt_linha = fmt_alerta_vermelho
            fmt_perc_linha = fmt_perc_baixa
        else:
            fmt_linha = fmt_celula
            fmt_perc_linha = fmt_perc_media
        
        ws2.write(linha, 0, prioridade.title(), fmt_linha)
        ws2.write_number(linha, 1, row['Quantidade_Leads'], fmt_linha)
        ws2.write_number(linha, 2, row['Percentual_Total_%'] / 100, fmt_perc_linha)
        ws2.write(linha, 3, acoes.get(prioridade, ''), fmt_celula)
        linha += 1
    
    # Ajusta larguras
    ws2.set_column(0, 0, 30)
    ws2.set_column(1, 1, 20)
    ws2.set_column(2, 2, 15)
    ws2.set_column(3, 3, 35)
    ws2.set_column(4, 4, 15)
    
    # ========================================================================
    # FINALIZA
    # ========================================================================
    
    workbook.close()
    
    print("[3/3] Relatório criado com sucesso!")
    print(f"\n📁 Arquivo: {arquivo_saida.name}")
    print("\n✅ PRONTO PARA APRESENTAR AO MARCELO!")
    print("=" * 80)
    
    return arquivo_saida


# ============================================================================
# EXECUÇÃO
# ============================================================================

if __name__ == "__main__":
    try:
        arquivo = criar_relatorio()
        print(f"\n🎯 Abra o arquivo: {arquivo.name}")
        print("\n2 abas criadas:")
        print("  1. DIAGNÓSTICO DOS LEADS - Detalhes de qualidade, idade, atividades")
        print("  2. DECISÃO NOVOS LEADS - Resposta direta: precisa ou não de novos leads")
        
    except Exception as e:
        print(f"\n❌ ERRO: {e}")
        import traceback
        traceback.print_exc()
