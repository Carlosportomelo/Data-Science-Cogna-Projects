"""
==============================================================================
LIMPEZA DE MATRÍCULAS
==============================================================================
Objetivo: Ler o arquivo matriculas_finais, selecionar apenas as colunas
          essenciais (Nome, Status P1, Unidade, Tipo) e salvar uma versão limpa.
==============================================================================
"""

import pandas as pd
import os

# Configuração de caminhos
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, 'Data')
INPUT_FILE = os.path.join(DATA_DIR, 'matriculas_finais.csv')
OUTPUT_FILE = os.path.join(DATA_DIR, 'matriculas_finais_limpo.csv')

def clean_matriculas():
    print("="*60)
    print("LIMPEZA DE ARQUIVO DE MATRÍCULAS")
    print("="*60)

    if not os.path.exists(INPUT_FILE):
        print(f"❌ Erro: Arquivo não encontrado em {INPUT_FILE}")
        return

    print(f"📂 Lendo arquivo: {os.path.basename(INPUT_FILE)}")
    
    # Tenta ler com diferentes encodings e separadores (automático)
    try:
        df = pd.read_csv(INPUT_FILE, sep=None, engine='python', encoding='utf-8')
    except UnicodeDecodeError:
        try:
            print("   ⚠️ Encoding UTF-8 falhou, tentando Latin1...")
            df = pd.read_csv(INPUT_FILE, sep=None, engine='python', encoding='latin1')
        except Exception as e:
            print(f"   ❌ Erro crítico ao ler arquivo: {e}")
            return

    # Normalizar nomes das colunas (remove espaços e converte para minúsculo para busca)
    df.columns = df.columns.str.strip()
    
    # Dicionário para mapear colunas (Nome Final -> Palavras-chave para buscar)
    # A ordem das palavras-chave importa (prioridade)
    mapa_busca = {
        'Nome do Aluno': ['nome do aluno', 'aluno', 'nome'],
        'Status P1': ['p1', 'status', 'situacao', 'pagamento'],
        'Unidade': ['unidade', 'filial', 'escola'],
        'Tipo': ['tipo', 'categoria', 'classificacao', 'curso']
    }
    
    colunas_selecionadas = {}
    
    print("\n🔍 Identificando colunas...")
    for nome_final, keywords in mapa_busca.items():
        coluna_encontrada = None
        
        # Tenta encontrar a primeira coluna que contenha uma das palavras-chave
        for kw in keywords:
            match = next((c for c in df.columns if kw.lower() in c.lower()), None)
            if match:
                coluna_encontrada = match
                break
        
        if coluna_encontrada:
            colunas_selecionadas[coluna_encontrada] = nome_final
            print(f"   ✅ '{nome_final}' mapeado para coluna original: '{coluna_encontrada}'")
        else:
            print(f"   ⚠️  Coluna para '{nome_final}' NÃO encontrada.")

    if not colunas_selecionadas:
        print("\n❌ Nenhuma coluna alvo foi identificada. Verifique o cabeçalho do arquivo.")
        print("   Colunas disponíveis:", df.columns.tolist())
        return

    # Filtrar e renomear
    df_clean = df[list(colunas_selecionadas.keys())].rename(columns=colunas_selecionadas)
    
    # Filtrar apenas NOVOS (Remover REMA)
    if 'Tipo' in df_clean.columns:
        print("   🧹 Filtrando alunos: Removendo 'REMA' para manter apenas NOVOS...")
        qtd_antes = len(df_clean)
        df_clean = df_clean[~df_clean['Tipo'].astype(str).str.contains('REMA', case=False, na=False)]
        print(f"      -> {qtd_antes - len(df_clean)} registros removidos (REMA). Restaram: {len(df_clean)}")

    # Salvar
    df_clean.to_csv(OUTPUT_FILE, index=False, sep=';', encoding='utf-8-sig')
    print(f"\n💾 Arquivo limpo salvo em: {OUTPUT_FILE}")
    print("-" * 60)

if __name__ == "__main__":
    clean_matriculas()