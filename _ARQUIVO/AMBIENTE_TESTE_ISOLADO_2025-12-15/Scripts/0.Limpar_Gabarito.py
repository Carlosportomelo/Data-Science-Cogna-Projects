import pandas as pd
import os

# Configuração de Caminhos
CAMINHO_BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(CAMINHO_BASE, 'Data')

print("="*80)
print("HIGIENIZADOR DE BASE DE TESTE (REMOVER GABARITO)")
print("="*80)

# 1. Encontrar o arquivo
arquivos = [f for f in os.listdir(DATA_DIR) if f.startswith('base_teste') and (f.endswith('.csv') or f.endswith('.xlsx'))]

# Filtrar arquivos temporários ou de sistema
arquivos = [f for f in arquivos if not f.startswith('~$')]

if not arquivos:
    print(f"❌ Nenhum arquivo 'base_teste' encontrado em: {DATA_DIR}")
    exit()

arquivo_nome = arquivos[0]
caminho_arquivo = os.path.join(DATA_DIR, arquivo_nome)
print(f"📂 Arquivo encontrado: {arquivo_nome}")

# 2. Carregar
try:
    if arquivo_nome.endswith('.csv'):
        try:
            df = pd.read_csv(caminho_arquivo, encoding='utf-8')
        except:
            df = pd.read_csv(caminho_arquivo, encoding='latin1')
    else:
        df = pd.read_excel(caminho_arquivo)
except Exception as e:
    print(f"❌ Erro ao ler arquivo: {e}")
    exit()

df.columns = df.columns.str.strip()

# 3. Verificar e Separar Gabarito
coluna_alvo = 'Etapa do negócio'

if coluna_alvo not in df.columns:
    print(f"⚠️  A coluna '{coluna_alvo}' NÃO existe no arquivo. Nada a fazer.")
    exit()

print(f"✂️  Removendo coluna '{coluna_alvo}'...")

# Salvar gabarito separado (para validação posterior)
caminho_gabarito = os.path.join(DATA_DIR, 'gabarito_separado.csv')
cols_gabarito = ['Record ID', coluna_alvo] if 'Record ID' in df.columns else [coluna_alvo]
df[cols_gabarito].to_csv(caminho_gabarito, index=False)
print(f"✅ Gabarito salvo separadamente em: gabarito_separado.csv")

# 4. Salvar arquivo limpo
df.drop(columns=[coluna_alvo], inplace=True)

if arquivo_nome.endswith('.csv'):
    df.to_csv(caminho_arquivo, index=False, encoding='utf-8-sig')
else:
    df.to_excel(caminho_arquivo, index=False)

print(f"✅ Arquivo original atualizado SEM a resposta final.")
print("   Agora você pode rodar o Script 3 com garantia total de cegueira.")