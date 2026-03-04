
import pandas as pd
from pathlib import Path
import glob
import sys

# --- CONFIGURAÇÕES ---
try:
    BASE_DIR = Path(__file__).resolve().parent.parent.parent
except NameError:
    BASE_DIR = Path.cwd()

OUTPUT_DIR = BASE_DIR / "outputs"

# --- 1. ENCONTRAR O ARQUIVO DE BLEND MAIS RECENTE ---
print("Procurando o arquivo de blend mais recente...")
list_of_files = glob.glob(str(OUTPUT_DIR / "meta_googleads_blend_*.xlsx"))
if not list_of_files:
    print(f"ERRO: Nenhum arquivo 'meta_googleads_blend_*.xlsx' encontrado no diretório: {OUTPUT_DIR}")
    sys.exit(1)

latest_file = max(list_of_files, key=lambda p: Path(p).stat().st_mtime)
print(f"   ... Arquivo encontrado: {Path(latest_file).name}")

# --- 2. LER E PROCESSAR OS DADOS ---
try:
    df = pd.read_excel(latest_file, sheet_name="Visao_Granular_Final")
    print("   ... Aba 'Visao_Granular_Final' lida com sucesso.")
except Exception as e:
    print(f"ERRO: Não foi possível ler a aba 'Visao_Granular_Final' do arquivo.")
    print(f"   Detalhe: {e}")
    sys.exit(1)

# --- 3. APLICAR FILTROS ---
# Converter a coluna 'Data' para datetime
df['Data'] = pd.to_datetime(df['Data'])

# Definir período de corte
start_date = pd.to_datetime("2025-10-01")
end_date = pd.to_datetime("2026-02-24")

# Aplicar os filtros conforme solicitado
df_filtrado = df[
    (df['Data'] >= start_date) &
    (df['Data'] <= end_date) &
    (df['Unidade'] == 'Sem Lead') &
    (df['Total_Negocios'] == 0) &
    (df['Midia_Paga'] > 0)
].copy()

if df_filtrado.empty:
    print("
Nenhum dia com investimento e sem leads encontrado para o período especificado.")
    sys.exit(0)

# --- 4. PREPARAR A VISUALIZAÇÃO ---
# Criar colunas separadas para investimento Meta e Google
df_filtrado['Investimento_Meta'] = df_filtrado.apply(
    lambda row: row['Midia_Paga'] if row['Origem_Principal'] == 'Social Pago' else 0,
    axis=1
)
df_filtrado['Investimento_Google'] = df_filtrado.apply(
    lambda row: row['Midia_Paga'] if row['Origem_Principal'] == 'Pesquisa Paga' else 0,
    axis=1
)

# Agrupar por data e somar os investimentos
resultado = df_filtrado.groupby('Data').agg({
    'Investimento_Meta': 'sum',
    'Investimento_Google': 'sum'
}).reset_index()

# Calcular o total
resultado['Total_Dia'] = resultado['Investimento_Meta'] + resultado['Investimento_Google']

# Ordenar por data
resultado = resultado.sort_values(by='Data')

# Formatar a data para exibição
resultado['Data'] = resultado['Data'].dt.strftime('%d/%m/%Y')


# --- 5. EXIBIR RESULTADOS ---
print("
" + "="*80)
print("DIAS COM INVESTIMENTO EM MÍDIA PAGA SEM GERAÇÃO DE LEADS")
print(f"Período de Análise: {start_date.strftime('%d/%m/%Y')} a {end_date.strftime('%d/%m/%Y')}")
print("="*80)

# Configurar o formato de exibição do pandas
pd.set_option('display.max_rows', None)
pd.set_option('display.width', 1000)
pd.options.display.float_format = 'R$ {:,.2f}'.format

# Renomear colunas para exibição
resultado.rename(columns={
    'Data': 'Data',
    'Investimento_Meta': 'Meta (Social Pago)',
    'Investimento_Google': 'Google (Pesquisa Paga)',
    'Total_Dia': 'Total Investido no Dia'
}, inplace=True)

# Imprimir o DataFrame formatado
print(resultado.to_string(index=False))
print("="*80)
