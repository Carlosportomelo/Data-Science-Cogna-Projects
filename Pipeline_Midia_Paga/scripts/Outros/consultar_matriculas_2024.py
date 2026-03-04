import pandas as pd
import glob
import os
import sys
from pathlib import Path

# Configuração de caminhos
try:
    BASE_DIR = Path(__file__).resolve().parent.parent
except NameError:
    BASE_DIR = Path.cwd()

OUTPUT_DIR = BASE_DIR / "outputs"

def find_latest_blend_file(directory):
    """Encontra o arquivo de blend mais recente."""
    pattern = str(directory / "meta_googleads_blend_*.xlsx")
    files = glob.glob(pattern)
    if not files:
        return None
    return max(files, key=os.path.getmtime)

def main():
    print("🔍 Lendo base consolidada para gerar resumo do funil...")
    
    blend_file = find_latest_blend_file(OUTPUT_DIR)
    if not blend_file:
        print("❌ Nenhum arquivo de blend encontrado em 'outputs/'. Execute o script 3 primeiro.")
        return

    print(f"📂 Arquivo: {Path(blend_file).name}")
    try:
        df = pd.read_excel(blend_file, sheet_name='Visao_Granular_Final')
    except Exception as e:
        print(f"❌ Erro ao ler arquivo: {e}")
        return

    # Garantir datetime
    df['Data'] = pd.to_datetime(df['Data'], errors='coerce')
    df['Ano_Criacao'] = df['Data'].dt.year

    # Filtrar apenas anos relevantes (ex: ignorar NaT)
    df = df.dropna(subset=['Ano_Criacao'])
    df['Ano_Criacao'] = df['Ano_Criacao'].astype(int)

    # Resumo por Status e Ano (Safra)
    print("\n" + "="*100)
    print("📊 RESUMO DO FUNIL POR ANO (SAFRA/CRIAÇÃO)")
    print("="*100)
    
    # Pivot table: Linhas=Status, Colunas=Ano
    # Mostra quantos leads criados naquele ano estão em cada status HOJE
    resumo_safra = pd.crosstab(
        df['Status_Principal'], 
        df['Ano_Criacao'], 
        margins=True, 
        margins_name='Total Geral'
    )
    
    # Configurar pandas para mostrar todas as linhas/colunas
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', 1000)
    
    print(resumo_safra)
    
    # Resumo específico de Matrículas (Safra vs Fechamento)
    print("\n" + "="*100)
    print("🎓 DETALHE DE MATRÍCULAS (Status = 8. Matrícula Realizada)")
    print("="*100)
    
    # Matrículas por Safra (Data de Criação do Lead)
    matriculas_safra = df[df['Matriculas'] == 1].groupby('Ano_Criacao').size()
    
    # Matrículas por Data de Fechamento (Data da Venda/Competência)
    df['Data_Fechamento'] = pd.to_datetime(df['Data_Fechamento'], errors='coerce')
    df['Ano_Fechamento'] = df['Data_Fechamento'].dt.year
    
    matriculas_fechamento = df[
        (df['Matriculas'] == 1) & 
        (df['Ano_Fechamento'].notna())
    ].groupby('Ano_Fechamento')['Matriculas'].count()
    
    # Combinar em um DataFrame para comparação
    df_comp = pd.DataFrame({
        'Por Safra (Criação)': matriculas_safra,
        'Por Competência (Fechamento)': matriculas_fechamento
    }).fillna(0).astype(int)
    
    print(df_comp)
    print("="*100)

if __name__ == "__main__":
    main()