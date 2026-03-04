import pandas as pd
import os

# --- CONFIGURAÇÃO ---
CAMINHO_ENTRADA = os.path.join('Data', 'hubspot_leads.csv')
if not os.path.exists(CAMINHO_ENTRADA):
    CAMINHO_ENTRADA = os.path.join('..', 'Data', 'hubspot_leads.csv')

print(f"[TESTE DE HIPÓTESE] Verificando corte de data em 01/10/2025...")

try:
    # 1. CARREGAMENTO
    try:
        df = pd.read_csv(CAMINHO_ENTRADA, sep=',', encoding='utf-8', on_bad_lines='skip')
    except:
        df = pd.read_csv(CAMINHO_ENTRADA, sep=',', encoding='latin1', on_bad_lines='skip')

    df.columns = df.columns.str.strip()
    
    # Renomeia e Converte Data
    df = df.rename(columns={'Fonte original do tráfego': 'Fonte', 'Data de criação': 'Data'})
    df['Data'] = pd.to_datetime(df['Data'], errors='coerce')

    # 2. CLASSIFICAÇÃO (A mesma da imagem)
    def classificar_dash(fonte):
        fonte = str(fonte).lower()
        if 'paga' in fonte or 'paid' in fonte or 'social pago' in fonte:
            return 'Digital - Pago'
        elif 'off-line' in fonte:
            return 'Cadastro da unidade'
        else:
            return 'Digital - Orgânico'

    df['Categoria'] = df['Fonte'].apply(classificar_dash)

    # 3. O TESTE DE CORTE (HIPÓTESE: Dash parou em 01/10/2025)
    # A Safra 'Captação Alta 25-26' geralmente começa em Out/24.
    # Então vamos pegar: 01/10/2024 até 01/10/2025
    
    data_inicio_safra = '2024-10-01'
    data_fim_dash     = '2025-10-01' # A data que aparece no topo da imagem
    
    mask = (df['Data'] >= data_inicio_safra) & (df['Data'] <= data_fim_dash)
    df_recorte = df[mask].copy()

    # 4. RESULTADO
    contagem = df_recorte['Categoria'].value_counts()
    
    print("\n" + "="*40)
    print(f"PERÍODO SIMULADO: {data_inicio_safra} até {data_fim_dash}")
    print("="*40)
    print(contagem)
    print("-" * 40)
    
    pago_encontrado = contagem.get('Digital - Pago', 0)
    diferenca = pago_encontrado - 8595
    
    print(f"Digital Pago no Script: {pago_encontrado}")
    print(f"Digital Pago na Imagem: 8.595")
    print(f"Diferença: {diferenca}")
    
    if abs(diferenca) < 200:
        print("\n[BINGO!] A diferença é mínima. O Dashboard está desatualizado desde 01/10.")
    else:
        print("\n[Ainda Diferente] O corte não é exatamente 01/10. Pode ser um filtro de Status.")

except Exception as e:
    print(e)