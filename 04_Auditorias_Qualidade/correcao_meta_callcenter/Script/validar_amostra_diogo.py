import pandas as pd
import os
import sys

# Configuração de caminhos
BASE_DIR = os.getcwd()
# Ajuste caso esteja rodando de dentro da pasta Script
if os.path.basename(BASE_DIR) == 'Script':
    BASE_DIR = os.path.dirname(BASE_DIR)

OUTPUT_DIR = os.path.join(BASE_DIR, 'Output')
FILE_NAME = 'amostra_prova_de_conceito_diogo.xlsx'
FILE_PATH = os.path.join(OUTPUT_DIR, FILE_NAME)

def validar():
    print(f"🚀 INICIANDO VALIDAÇÃO DO ARQUIVO FINAL")
    print(f"📂 Arquivo alvo: {FILE_PATH}")
    
    if not os.path.exists(FILE_PATH):
        print("❌ ERRO: O arquivo não foi encontrado na pasta Output.")
        return

    try:
        # Ler o arquivo Excel (carrega metadados das abas)
        xls = pd.ExcelFile(FILE_PATH)
        print(f"✅ Arquivo aberto com sucesso. Abas encontradas: {xls.sheet_names}")
        
        # 1. Validar aba 'Diogo_Preenchido' (Aba principal)
        if 'Diogo_Preenchido' in xls.sheet_names:
            print("\n📋 --- ANÁLISE DA ABA: 'Diogo_Preenchido' ---")
            df = pd.read_excel(xls, 'Diogo_Preenchido')
            print(f"   > Total de Linhas: {len(df)}")
            
            if 'idade' in df.columns:
                n_preenchidos = df['idade'].notna().sum()
                print(f"   > Coluna 'idade' existe? SIM")
                print(f"   > Valores preenchidos na coluna 'idade': {n_preenchidos}")
                
                # Tenta identificar colunas de identificação para mostrar no print
                cols_id = [c for c in df.columns if 'mail' in c.lower() or 'nome' in c.lower()]
                cols_show = cols_id[:2] + ['idade'] if cols_id else df.columns[:3].tolist() + ['idade']
                
                print("\n🔎 Amostra das primeiras 5 linhas (Colunas principais + Idade):")
                print(df[cols_show].head(5).to_string(index=False))
            else:
                print("❌ ERRO: A coluna 'idade' NÃO foi encontrada nesta aba.")
        
        # 2. Validar aba 'Amostra' (Resumo)
        if 'Amostra' in xls.sheet_names:
            print("\n📋 --- ANÁLISE DA ABA: 'Amostra' ---")
            df_amostra = pd.read_excel(xls, 'Amostra')
            print(f"   > Total de Linhas: {len(df_amostra)}")
            print("\n🔎 Conteúdo Completo da Amostra:")
            print(df_amostra.to_string(index=False))

    except Exception as e:
        print(f"❌ Erro crítico ao ler o Excel: {e}")

if __name__ == "__main__":
    validar()