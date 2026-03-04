import pandas as pd
import sys
from pathlib import Path
import re
import numpy as np
from datetime import datetime

# --- Constantes ---
FILE_PATH = Path("data/meta_dataset.csv")
SHEET_NAME = None # Use None para CSV. Mude para 'Meta_Completo' se for Excel.

# --- Caminhos de Saída ---
OUTPUT_DIR = Path("outputs")
OUT_EXCEL_FILE = OUTPUT_DIR / "meta_dataset_dashboard.xlsx"

# --- Função para Carregar Histórico Existente ---
def carregar_historico_existente(caminho_excel):
    """
    Carrega o dashboard existente se ele existir.
    Retorna DataFrame vazio se não existir ou houver erro.
    """
    if not caminho_excel.exists():
        print("Nenhum histórico encontrado. Será criado um novo arquivo.")
        return pd.DataFrame()
    
    try:
        print(f"Carregando histórico existente de: {caminho_excel}")
        df_historico = pd.read_excel(caminho_excel, sheet_name='Meta_Completo')
        
        # Garantir que Data_Datetime seja datetime
        if 'Data_Datetime' in df_historico.columns:
            df_historico['Data_Datetime'] = pd.to_datetime(df_historico['Data_Datetime'], errors='coerce')
        
        print(f"  {len(df_historico)} linhas de histórico carregadas")
        
        # Mostrar período do histórico
        if not df_historico.empty and 'Data_Datetime' in df_historico.columns:
            data_min = df_historico['Data_Datetime'].min()
            data_max = df_historico['Data_Datetime'].max()
            print(f"  Período do histórico: {data_min.strftime('%d/%m/%Y')} até {data_max.strftime('%d/%m/%Y')}")
        
        return df_historico
    except Exception as e:
        print(f"  Erro ao carregar histórico: {e}")
        print("  Continuando sem histórico (será criado novo arquivo)")
        return pd.DataFrame()

# --- Função para Mesclar Dados (Incremental) ---
def mesclar_dados_incremental(df_historico, df_novos, col_data='Data_Datetime'):
    """
    Mescla dados históricos com novos dados de forma inteligente:
    - Remove duplicatas por data e campanha (mantém o mais recente)
    - Preserva todo o histórico
    - Adiciona apenas novos registros
    """
    if df_historico.empty:
        print("\nSem histórico prévio. Todos os dados serão considerados novos.")
        return df_novos.copy()
    
    print("\nMesclando dados históricos com novos dados...")
    
    # Garantir que ambos têm a coluna de data como datetime
    df_historico[col_data] = pd.to_datetime(df_historico[col_data], errors='coerce')
    df_novos[col_data] = pd.to_datetime(df_novos[col_data], errors='coerce')
    
    # Estatísticas antes da mesclagem
    linhas_historico = len(df_historico)
    linhas_novas = len(df_novos)
    
    # Extrair apenas dados do histórico que são ANTERIORES à primeira data dos novos dados
    data_min_novos = df_novos[col_data].min()
    df_historico_antigos = df_historico[df_historico[col_data] < data_min_novos].copy()
    
    # Concatenar histórico antigo com novos dados (sem duplicação)
    df_merged = pd.concat([df_historico_antigos, df_novos], ignore_index=True)
    
    # Remover duplicatas mantendo o último registro (mais recente)
    df_merged = df_merged.sort_values(by=col_data)
    
    # Identificar colunas únicas para determinar duplicatas (excluindo colunas calculadas)
    colunas_originais = [col for col in df_merged.columns 
                         if col not in ['Ano', 'Mes', 'Mes_Ano', 'Data_Datetime']]
    
    df_merged = df_merged.drop_duplicates(subset=colunas_originais, keep='last')
    
    # Estatísticas após mesclagem
    linhas_finais = len(df_merged)
    linhas_antigos_mantidos = len(df_historico_antigos)
    linhas_duplicadas = linhas_historico + linhas_novas - linhas_finais
    
    print(f"  Histórico anterior: {linhas_antigos_mantidos} linhas")
    print(f"  Novos dados: {linhas_novas} linhas")
    print(f"  Duplicatas removidas: {linhas_duplicadas} linhas")
    print(f"  Total final: {linhas_finais} linhas")
    
    # Mostrar período final
    data_min = df_merged[col_data].min()
    data_max = df_merged[col_data].max()
    print(f"  Período completo: {data_min.strftime('%d/%m/%Y')} até {data_max.strftime('%d/%m/%Y')}")
    
    return df_merged

# --- Função Utilitária para Números ---
def parse_number(x):
    if pd.isna(x):
        return 0.0
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip()
    s = s.replace(" ", "")
    if "." in s and "," in s:
        # Formato brasileiro (1.234,56) -> 1234.56
        s = s.replace(".", "").replace(",", ".")
    else:
        # Tenta tratar (1,234.56) ou (1234,56)
        s = s.replace(",", ".")
    try:
        return float(s)
    except:
        return 0.0

# =====================================================================
# --- INÍCIO DO PROCESSAMENTO ---
# =====================================================================

print("="*70)
print("SISTEMA DE ATUALIZAÇÃO INCREMENTAL - META ADS")
print("="*70)

# --- 0. Carregar Histórico Existente ---
df_historico = carregar_historico_existente(OUT_EXCEL_FILE)

# --- 1. Carregar Dados Novos ---
print(f"\nCarregando dados novos de: {FILE_PATH}")
try:
    if FILE_PATH.suffix.lower() in ['.xlsx', '.xls']:
        df_novos = pd.read_excel(FILE_PATH, sheet_name=SHEET_NAME or 0)
    elif FILE_PATH.suffix.lower() == '.csv':
        # Tenta ler com engine python (mais flexível) e detectar separador
        try:
            df_novos = pd.read_csv(FILE_PATH, engine='python', sep=None)
        except Exception as e_csv:
            print(f"Aviso: Falha ao ler CSV com engine='python' ({e_csv}). Tentando engine padrão.")
            df_novos = pd.read_csv(FILE_PATH) # Tenta engine padrão
    else:
        raise ValueError(f"Formato de arquivo não suportado: {FILE_PATH.suffix}")
    
    if df_novos.empty:
        raise ValueError("O DataFrame está vazio após o carregamento.")
    print(f" {len(df_novos)} linhas carregadas dos novos dados")

except FileNotFoundError:
    print(f" ERRO: Arquivo não encontrado em: {FILE_PATH.resolve()}")
    print("   Verifique se o caminho e o nome do arquivo estão corretos.")
    sys.exit(1)
except Exception as e:
    print(f" ERRO ao carregar o arquivo: {e}")
    sys.exit(1)

# =====================================================================
# --- 2. Normalizar Colunas dos Novos Dados ---
# =====================================================================

print("\nProcessando novos dados...")
# remover espaços, normalizar caixa
df_novos.columns = df_novos.columns.map(lambda c: str(c).strip() if not pd.isna(c) else c)
# Remover aspas duplas (") que apareceram no seu log de erro
df_novos.columns = df_novos.columns.str.strip('"')
# Remover caracteres especiais (BOM, etc.)
df_novos.columns = [re.sub(r'[^\x00-\x7F]+', '', col) for col in df_novos.columns]
df_novos.columns = df_novos.columns.str.replace('\\r', '', regex=False).str.replace('\\n', '', regex=False)

print("\nColunas detectadas (normalizadas):")
print(list(df_novos.columns))

# --- 3. Encontrar a Coluna de Data (Lógica flexível) ---

possiveis_nomes = ['Dia', 'dia', 'Data', 'data', 'Date', 'date', 'Data_Datetime', 'DataFormatada']
col_data = None
for nome in possiveis_nomes:
    if nome in df_novos.columns:
        col_data = nome
        print(f" Coluna de data encontrada: '{col_data}'")
        break

if col_data is None:
    print(" Coluna de data não encontrada por nome. Tentando heurística...")
    for col in df_novos.columns:
        if df_novos[col].dtype == object:
            sample = df_novos[col].dropna().astype(str).head(20).tolist()
            if not sample: 
                continue
            n_like = sum(1 for v in sample if ('/' in v or '-' in v or v.count('/')>=1 or v.count('-')>=1 or (v.isdigit() and len(v) >= 4)))
            if n_like >= max(3, len(sample)//3):
                col_data = col
                print(f" Possível coluna de data detectada por heurística: '{col_data}'")
                break

if col_data is None:
    print(" Não foi possível localizar automaticamente uma coluna de data (esperada 'Dia' ou 'Data').")
    print("   Imprimindo amostra para inspeção manual:")
    print(df_novos.head(10).to_string(index=False))
    raise KeyError("Coluna de data 'Dia' ou 'Data' não encontrada. Verifique cabeçalho do arquivo.")

# --- 3b. Encontrar a Coluna de Investimento ---
possiveis_nomes_invest = ['Valor usado (BRL)', 'Valor', 'Investimento', 'spent', 'gasto']
col_invest = None
for nome in possiveis_nomes_invest:
    if nome in df_novos.columns:
        col_invest = nome
        print(f" Coluna de investimento encontrada: '{col_invest}'")
        break
        
if col_invest is None:
    print(" Não foi possível localizar automaticamente uma coluna de investimento.")
    raise KeyError("Coluna de investimento (ex: 'Valor usado (BRL)') não encontrada.")

# --- 4. Processar Novos Dados ---
print(f"\nUsando coluna de data: '{col_data}' -> convertendo para datetime")

# Tentar converter a data, sendo flexível com o formato
try:
    df_novos['Data_Datetime'] = pd.to_datetime(df_novos[col_data], errors='coerce')
except Exception:
    print(f"Aviso: Falha na conversão de data. Tentando formato padrão.")
    df_novos['Data_Datetime'] = pd.to_datetime(df_novos[col_data], errors='coerce')

num_na_dates = df_novos['Data_Datetime'].isna().sum()
if num_na_dates > 0:
    print(f" Atenção: {num_na_dates} linhas não puderam ser convertidas para data e serão ignoradas.")
    df_novos = df_novos.dropna(subset=['Data_Datetime'])

if df_novos.empty:
    print(" ERRO: Nenhuma linha restou após a limpeza das datas. Verifique o formato da data no arquivo.")
    sys.exit(1)

print(f"\nUsando coluna de investimento: '{col_invest}' -> convertendo para número")
df_novos[col_invest] = df_novos[col_invest].apply(parse_number)

print("\nProcessando dados... (Ex: Ano, Mês, etc.)")
df_novos['Ano'] = df_novos['Data_Datetime'].dt.year
df_novos['Mes'] = df_novos['Data_Datetime'].dt.month
df_novos['Mes_Ano'] = df_novos['Data_Datetime'].dt.to_period('M')

# =====================================================================
# --- FILTRO: EXCLUIR "BILINGUAL" DOS NOVOS DADOS ---
# =====================================================================
print("\nAplicando filtro: Excluindo registros com 'bilingual'...")
linhas_antes = len(df_novos)

# Criar máscara para identificar linhas com "bilingual" em qualquer coluna de texto
mask_bilingual = pd.Series([False] * len(df_novos), index=df_novos.index)

for col in df_novos.columns:
    if df_novos[col].dtype == 'object':  # Apenas colunas de texto
        try:
            mask_bilingual |= df_novos[col].astype(str).str.contains(
                'bilingual', 
                case=False, 
                na=False, 
                regex=False
            )
        except Exception as e:
            print(f"  Aviso: Erro ao processar coluna '{col}': {e}")
            continue

# Aplicar filtro (manter apenas linhas SEM "bilingual")
df_novos = df_novos[~mask_bilingual].copy()

linhas_removidas = linhas_antes - len(df_novos)
print(f"  Filtro aplicado: {linhas_removidas} linhas removidas (contendo 'bilingual')")
print(f"  Linhas restantes: {len(df_novos)}")

if df_novos.empty:
    print(" ERRO: Nenhuma linha restou após a exclusão de 'bilingual'. Verifique os dados.")
    sys.exit(1)

print("\nProcessamento dos novos dados concluído com sucesso!")

# =====================================================================
# --- 5. MESCLAR HISTÓRICO COM NOVOS DADOS ---
# =====================================================================

df = mesclar_dados_incremental(df_historico, df_novos, col_data='Data_Datetime')

# Recalcular campos derivados após mesclagem (caso o histórico não os tenha)
print("\nRecalculando campos derivados...")
df['Ano'] = df['Data_Datetime'].dt.year
df['Mes'] = df['Data_Datetime'].dt.month
df['Mes_Ano'] = df['Data_Datetime'].dt.to_period('M')

# Reordenar por data
df = df.sort_values(by='Data_Datetime').reset_index(drop=True)

print("Mesclagem concluída!")

# =====================================================================
# --- 6. GERAR RELATÓRIOS ---
# =====================================================================
print("\nGerando relatórios...")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# --- Relatório 1: Meta_YoY (Agregado por Dia) ---
print(f"  Processando dados para a aba 'Meta_YoY'...")
df_daily_agg = df.groupby('Data_Datetime').agg(
    Investimento=(col_invest, 'sum')
).reset_index()

# Renomear colunas para bater com o anexo
df_daily_agg = df_daily_agg.rename(columns={'Data_Datetime': 'Data'})

# Filtrar dias sem investimento para não poluir o arquivo e ordenar
df_daily_agg = df_daily_agg[df_daily_agg['Investimento'] > 0].sort_values(by='Data')


# --- Relatório 2: Abas por Ano (2023, 2024, 2025) ---
print("  Processando dados para as abas por ano...")

# Criar dicionário de dataframes por ano dinamicamente
anos_disponiveis = sorted(df['Ano'].dropna().unique().astype(int))
dfs_por_ano = {ano: df[df['Ano'] == ano] for ano in anos_disponiveis}

# Salvar tudo em um único arquivo Excel com abas
print(f"\nSalvando arquivo Excel único em: {OUT_EXCEL_FILE}")
try:
    with pd.ExcelWriter(OUT_EXCEL_FILE, engine='openpyxl') as writer:
        # Aba 1: Meta_YoY
        df_daily_agg.to_excel(writer, sheet_name='Meta_YoY', index=False, float_format='%.2f')
        print("  Aba 'Meta_YoY' salva.")
        
        # Aba 2: Meta_Completo (O dataframe 'df' original processado)
        df.to_excel(writer, sheet_name='Meta_Completo', index=False)
        print("  Aba 'Meta_Completo' salva.")

        # Abas por Ano
        for ano, df_ano in dfs_por_ano.items():
            if not df_ano.empty:
                sheet_name = f'Meta_{ano}'
                df_ano.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"  Aba '{sheet_name}' salva.")
    
    print(f"\nArquivo Excel '{OUT_EXCEL_FILE.name}' atualizado com sucesso na pasta '{OUTPUT_DIR}'!")

except ImportError:
    print("\n\nERRO: A BIBLIOTECA 'openpyxl' NÃO ESTÁ INSTALADA.")
    print("Para salvar em Excel, por favor, rode o comando no seu terminal:")
    print("pip install openpyxl")
    sys.exit(1)
except Exception as e:
    print(f"\n\nERRO AO SALVAR O EXCEL: {e}")
    print("Verifique se o arquivo não está aberto em outro programa.")
    # Não sai imediatamente para tentar salvar o resumido

# --- GERAÇÃO DO ARQUIVO RESUMIDO PARA DASHBOARD ---
RESUMIDO_DIR = OUTPUT_DIR / "visão resumida"
RESUMIDO_DIR.mkdir(parents=True, exist_ok=True)
OUT_RESUMIDO = RESUMIDO_DIR / "meta_resumido_Dash.xlsx"
print(f"\nSalvando arquivo resumido em: {OUT_RESUMIDO}")
try:
    with pd.ExcelWriter(OUT_RESUMIDO, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Meta_Completo', index=False)
    print("  Arquivo resumido salvo (Aba: Meta_Completo).")
except Exception as e:
    print(f"ERRO AO SALVAR O EXCEL RESUMIDO: {e}")


# =====================================================================
# --- 7. CONFIRMAÇÃO DE DADOS ---
# =====================================================================
print("\n" + "="*70)
print("ESTATÍSTICAS FINAIS")
print("="*70)

try:
    # Assegurar que 'Data' é datetime (já deve ser, mas para garantir)
    df_daily_agg['Data'] = pd.to_datetime(df_daily_agg['Data'])
    
    # Período completo
    data_min = df_daily_agg['Data'].min()
    data_max = df_daily_agg['Data'].max()
    total_invest = df_daily_agg['Investimento'].sum()
    
    print(f"\nPeríodo Completo: {data_min.strftime('%d/%m/%Y')} até {data_max.strftime('%d/%m/%Y')}")
    print(f"Investimento Total: R$ {total_invest:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    
    # Investimentos por ano
    print("\nInvestimento por Ano:")
    for ano in sorted(df['Ano'].unique()):
        invest_ano = df[df['Ano'] == ano][col_invest].sum()
        print(f"  {ano}: R$ {invest_ano:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    
    # Últimos meses (se existirem)
    ano_atual = datetime.now().year
    mes_atual = datetime.now().month
    
    if ano_atual in df['Ano'].values:
        # Mês atual
        invest_mes_atual = df[
            (df['Ano'] == ano_atual) & 
            (df['Mes'] == mes_atual)
        ][col_invest].sum()
        
        # Mês anterior
        mes_anterior = mes_atual - 1 if mes_atual > 1 else 12
        ano_mes_anterior = ano_atual if mes_atual > 1 else ano_atual - 1
        
        invest_mes_anterior = df[
            (df['Ano'] == ano_mes_anterior) & 
            (df['Mes'] == mes_anterior)
        ][col_invest].sum()
        
        print(f"\nÚltimos Meses:")
        print(f"  Mês Anterior ({mes_anterior:02d}/{ano_mes_anterior}): R$ {invest_mes_anterior:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        print(f"  Mês Atual ({mes_atual:02d}/{ano_atual}): R$ {invest_mes_atual:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

except Exception as e_conf:
    print(f"  Não foi possível calcular as estatísticas: {e_conf}")

print("\n" + "="*70)
print("PROCESSAMENTO CONCLUÍDO COM SUCESSO!")
print("="*70)

print("\n--- Amostra do Relatório YoY (Aba 'Meta_YoY') ---")
print(df_daily_agg.tail(10))