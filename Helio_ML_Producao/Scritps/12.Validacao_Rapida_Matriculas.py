"""
==============================================================================
VALIDAÇÃO RÁPIDA - HELIO vs MATRÍCULAS
==============================================================================
Script: 12.Validacao_Rapida_Matriculas.py
Objetivo: Responder pontualmente quantos leads classificados como Nota 4 ou 5
          pelo Helio constam na base oficial de matrículas finais.
==============================================================================
"""

import pandas as pd
import os
import glob
import unicodedata

# ==============================================================================
# CONFIGURAÇÃO DE CAMINHOS
# ==============================================================================
CAMINHO_BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CAMINHO_MATRICULAS = os.path.join(CAMINHO_BASE, 'Data', 'matriculas_finais_limpo.csv')
PASTA_RELATORIOS_ML = os.path.join(CAMINHO_BASE, 'Outputs', 'Relatorios_ML')

# Pastas onde o Helio salvou os scores (Baseado no script 10)
PASTAS_SCORES = [
    os.path.join(PASTA_RELATORIOS_ML, 'Relatório 12-12-25'),
    os.path.join(PASTA_RELATORIOS_ML, 'Relatório 01-05-26'),
    os.path.join(PASTA_RELATORIOS_ML, 'start')
]

print("="*80)
print("VALIDAÇÃO RÁPIDA: HELIO (NOTA 4/5) vs MATRÍCULAS REAIS")
print("="*80)

# ==============================================================================
# 1. FUNÇÕES DE NORMALIZAÇÃO (PARA GARANTIR O MATCH)
# ==============================================================================
def normalizar_nome(nome):
    """Remove acentos, espaços extras e converte para maiúsculo."""
    if pd.isna(nome): return ""
    nome = str(nome).upper().strip()
    # Remove acentos
    nome = unicodedata.normalize('NFKD', nome)
    nome = ''.join(c for c in nome if not unicodedata.combining(c))
    # Remove espaços duplicados
    nome = ' '.join(nome.split())
    return nome

# ==============================================================================
# 2. CARREGAR BASE DE MATRÍCULAS (A VERDADE)
# ==============================================================================
print("\n[1/3] Carregando base de Matrículas Finais...")

if not os.path.exists(CAMINHO_MATRICULAS):
    print(f"[ERRO] Arquivo não encontrado: {CAMINHO_MATRICULAS}")
    exit()

try:
    # Tenta ler com separador ; ou ,
    try:
        df_mat = pd.read_csv(CAMINHO_MATRICULAS, sep=';', encoding='utf-8-sig')
    except:
        df_mat = pd.read_csv(CAMINHO_MATRICULAS, sep=',', encoding='utf-8')

    # Identificar coluna de nome
    col_nome = next((c for c in df_mat.columns if 'nome' in c.lower() and 'aluno' in c.lower()), None)
    
    if col_nome:
        df_mat['Nome_Normalizado'] = df_mat[col_nome].apply(normalizar_nome)
        set_matriculados = set(df_mat['Nome_Normalizado'].unique())
        # Remover strings vazias do set
        set_matriculados.discard("")
        print(f"   -> {len(set_matriculados)} alunos únicos identificados na base de matrículas.")
    else:
        print("   [ERRO] Coluna 'Nome do Aluno' não encontrada no CSV de matrículas.")
        exit()

except Exception as e:
    print(f"   [ERRO] Falha ao ler CSV de matrículas: {e}")
    exit()

# ==============================================================================
# 3. CARREGAR PREVISÕES DO HELIO (LEADS NOTA 4 E 5)
# ==============================================================================
print("\n[2/3] Carregando leads Nota 4 e 5 do Helio...")

leads_helio = []

for pasta in PASTAS_SCORES:
    if not os.path.exists(pasta):
        continue
        
    arquivos = glob.glob(os.path.join(pasta, '*.xlsx'))
    for arquivo in arquivos:
        try:
            xl = pd.ExcelFile(arquivo)
            # Abas de interesse (Alta Prioridade)
            abas_alta = ['2_Top500_Nota5', '3_Nota4']
            
            for aba in abas_alta:
                if aba in xl.sheet_names:
                    # Busca cabeçalho dinamicamente
                    df_preview = pd.read_excel(xl, sheet_name=aba, header=None, nrows=20)
                    header_idx = -1
                    for idx, row in df_preview.iterrows():
                        if row.astype(str).str.contains('Nome do negócio', case=False, na=False).any():
                            header_idx = idx
                            break
                    
                    if header_idx != -1:
                        df_temp = pd.read_excel(xl, sheet_name=aba, header=header_idx)
                        # Normalizar colunas
                        df_temp.columns = df_temp.columns.astype(str).str.strip()
                        
                        if 'Nome do negócio' in df_temp.columns:
                            # Extrair nomes e nota (baseado na aba)
                            nota = 5 if 'Nota5' in aba else 4
                            nomes = df_temp['Nome do negócio'].apply(normalizar_nome).tolist()
                            
                            for nome in nomes:
                                if nome: # Ignora vazios
                                    leads_helio.append({'Nome': nome, 'Nota': nota})
        except Exception as e:
            pass

df_helio = pd.DataFrame(leads_helio)

if df_helio.empty:
    print("   [AVISO] Nenhum lead Nota 4 ou 5 encontrado nos relatórios.")
    exit()

# Deduplicar (se o lead apareceu em vários relatórios, mantemos 1 registro)
df_helio_unicos = df_helio.drop_duplicates(subset=['Nome'])
print(f"   -> {len(df_helio_unicos)} leads únicos classificados como Nota 4 ou 5 pelo Helio.")

# ==============================================================================
# 4. CRUZAMENTO E RESULTADO
# ==============================================================================
print("\n[3/3] Cruzando dados...")

# Filtrar quem está na lista de matriculados
df_match = df_helio_unicos[df_helio_unicos['Nome'].isin(set_matriculados)]

print(f"\n{'='*80}")
print(f"RESPOSTA FINAL:")
print(f"Dos {len(df_helio_unicos)} leads que o Helio indicou como Alta Prioridade (Nota 4 ou 5),")
print(f"EXATAMENTE {len(df_match)} constam na base 'matriculas_finais_limpo.csv'.")
print(f"{'='*80}\n")

if not df_match.empty:
    print("LISTA DOS ALUNOS ENCONTRADOS:")
    print("-" * 40)
    df_match = df_match.sort_values('Nome')
    for idx, row in df_match.iterrows():
        print(f" > [Nota {row['Nota']}] {row['Nome']}")
    print("-" * 40)