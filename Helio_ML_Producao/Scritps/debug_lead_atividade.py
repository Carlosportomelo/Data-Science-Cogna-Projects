"""
==============================================================================
DEBUG DE DADOS - INVESTIGAÇÃO DE VALORES NULOS (DATA ATIVIDADE)
==============================================================================
Objetivo: Verificar quantos leads possuem colunas de data vazias.
==============================================================================
"""
import pandas as pd
import os

# Configuração
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ARQUIVO_LEADS = os.path.join(BASE_DIR, 'Data', 'hubspot_leads.csv')

print(f"🔍 Investigando NULOS em colunas de Data de Atividade")
print(f" Arquivo: {ARQUIVO_LEADS}")

try:
    # Tenta ler com encoding padrão ou latin1
    try:
        df = pd.read_csv(ARQUIVO_LEADS, encoding='utf-8')
    except:
        print("⚠️ Encoding UTF-8 falhou, tentando Latin1...")
        df = pd.read_csv(ARQUIVO_LEADS, encoding='latin1')
    
    # Limpeza de colunas (remove espaços extras que impedem o match)
    df.columns = df.columns.str.strip()
    
    total_leads = len(df)
    print(f"📊 Total de Leads na base: {total_leads}")
    print("-" * 50)

    # Identificar colunas de interesse
    cols_interesse = [c for c in df.columns if 'data' in c.lower() and ('atividade' in c.lower() or 'contato' in c.lower())]
    
    if not cols_interesse:
        print("❌ Nenhuma coluna de 'Data' + 'Atividade/Contato' encontrada.")
        print("Colunas disponíveis:", df.columns.tolist())
    else:
        print(f"🔎 Colunas encontradas para análise: {cols_interesse}\n")
        
        for col in cols_interesse:
            # Contar vazios (NaN ou string vazia ou string com espaço)
            series_str = df[col].astype(str).str.strip()
            mask_vazio = series_str.isin(['nan', 'NaT', '', 'None', '<NA>']) | df[col].isna()
            
            qtd_vazios = mask_vazio.sum()
            pct_vazios = (qtd_vazios / total_leads) * 100
            
            print(f"📌 Coluna: '{col}'")
            print(f"   - Vazios: {qtd_vazios} ({pct_vazios:.2f}%)")
            print(f"   - Preenchidos: {total_leads - qtd_vazios}")
            
            if qtd_vazios > 0:
                print("   ⚠️  ALERTA: Existem valores vazios!")
                if 'Record ID' in df.columns:
                    ids_vazios = df.loc[mask_vazio, 'Record ID'].head(5).tolist()
                    print(f"   - Exemplos de IDs com '{col}' vazia: {ids_vazios}")
            else:
                print("   ✅ Coluna 100% preenchida.")
            
            if (total_leads - qtd_vazios) > 0:
                exemplo = df.loc[~mask_vazio, col].iloc[0]
                print(f"   - Exemplo de valor preenchido: '{exemplo}'")
            
            print("-" * 30)
            
except Exception as e:
    print(f"❌ Erro ao ler arquivo: {e}")
