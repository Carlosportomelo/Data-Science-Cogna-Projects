import pandas as pd
import os
import sys

# ==============================================================================
# 1. LÓGICA CENTRAL DE CLASSIFICAÇÃO (V5 - DEFINITIVA)
# ==============================================================================
def classificar_canal_completo(row, col_map):
    """
    Versão 5.0:
    - Recupera Google/Meta Ads.
    - Agrupa CRM_UI, CONTACTS e EXTENSION como 'Offline (Manual/Balcão)'.
    - Separa Importações e Integrações.
    """
    # Converte para string e minúsculo
    fonte = str(row[col_map['fonte']]).lower()
    detalhe1 = str(row[col_map['detalhe1']]).lower()
    detalhe2 = str(row[col_map['detalhe2']]).lower()

    # --- FASE 1: CORREÇÃO DIGITAL (RECUPERAÇÃO) ---
    
    # 1.1 Google Ads (Prioridade Máxima)
    termos_ads = ['gclid', 'cpc', 'adwords', 'ppc']
    if any(termo in detalhe1 for termo in termos_ads) or any(termo in detalhe2 for termo in termos_ads):
        return 'Pesquisa paga'
    
    if fonte in ['paid search', 'pesquisa paga', 'tráfego pago']:
        return 'Pesquisa paga'

    # 1.2 Proteção Social Orgânico
    if fonte in ['organic social', 'social orgânico', 'social orgânica', 'mídias sociais']:
        return 'Social orgânico'

    # 1.3 Meta Ads (Social Pago)
    if fonte in ['paid social', 'social pago']:
        return 'Social pago'
    
    if 'social' in fonte and 'paid' in detalhe1:
        return 'Social pago'
        
    if ('facebook' in detalhe1 or 'instagram' in detalhe1) and 'social' in fonte:
         return 'Social pago'

    # 1.4 Pesquisa Orgânica
    if fonte in ['organic search', 'pesquisa orgânica']:
        return 'Pesquisa orgânica'

    # --- FASE 2: OFFLINE DETALHADO (AGRUPAMENTO INTELIGENTE) ---
    
    if 'offline' in fonte or 'off-line' in fonte:
        # TUDO ISSO É MANUAL/BALCÃO:
        # crm_ui (Web), contacts (Mobile/App), extension (Plugin Email), conversations (Chat), sales
        termos_manual = ['crm_ui', 'contacts', 'extension', 'conversations', 'sales']
        
        if any(t in detalhe1 for t in termos_manual):
            return 'Offline (Manual/Balcão)'
        
        # IMPORT = Planilhas Excel
        if 'import' in detalhe1:
            return 'Offline (Importação)'
            
        # INTEGRAÇÕES = APIs externas
        if 'integration' in detalhe1 or 'api' in detalhe1:
            return 'Offline (Integração)'
            
        # O que sobrou mesmo (deve ser quase zero agora)
        return 'Offline (Outros)'

    # --- FASE 3: OUTROS CANAIS ---
    
    if fonte in ['direct traffic', 'tráfego direto']:
        return 'Tráfego direto'
        
    if fonte in ['referrals', 'referências']:
        return 'Referências'
        
    if fonte in ['email marketing', 'e-mail marketing', 'email']:
        return 'E-mail marketing'
        
    if 'ia' in fonte or 'ai' in fonte:
        return 'Referências de IA'

    return 'Outras campanhas'

# ==============================================================================
# 2. EXECUÇÃO PRINCIPAL (MAIN)
# ==============================================================================
def main():
    # Setup de Caminhos
    diretorio_script = os.path.dirname(os.path.abspath(__file__))
    caminho_entrada = os.path.join(diretorio_script, '..', 'Data', 'hubspot_leads.csv')
    caminho_saida_excel = os.path.join(diretorio_script, '..', 'Outputs', 'Base_Auditada_Final.xlsx')
    
    arquivo_in = os.path.normpath(caminho_entrada)
    arquivo_out = os.path.normpath(caminho_saida_excel)
    
    print(f"--- [1/3] Lendo arquivo: {arquivo_in} ---")
    
    try:
        df = pd.read_csv(arquivo_in, encoding='utf-8')
    except UnicodeDecodeError:
        print("Aviso: Encoding UTF-8 falhou. Tentando Latin-1...")
        df = pd.read_csv(arquivo_in, encoding='latin1')
    except FileNotFoundError:
        print(f"ERRO: Arquivo não encontrado.")
        return

    cols = df.columns.tolist()
    col_map = {
        'fonte': next((c for c in cols if 'Fonte original' in c), None),
        'detalhe1': next((c for c in cols if 'Detalhamento' in c and '1' in c), None),
        'detalhe2': next((c for c in cols if 'Detalhamento' in c and '2' in c), None)
    }

    if not col_map['fonte']:
        print("ERRO: Colunas não encontradas.")
        return

    print("--- [2/3] Aplicando Classificação V5 (Consolidada)... ---")
    df['Canal_Auditado'] = df.apply(classificar_canal_completo, axis=1, args=(col_map,))

    # RELATÓRIO FINAL
    print("\n" + "="*60)
    print(" 📊 RESULTADO FINAL: BASE 100% AUDITADA")
    print("="*60)
    
    resultado = df['Canal_Auditado'].value_counts()
    total = len(df)
    
    print(f"{'CANAL AUDITADO':<35} | {'VOLUME':<10} | {'SHARE':<10}")
    print("-" * 60)
    for canal, volume in resultado.items():
        share = (volume / total) * 100
        print(f"{canal:<35} | {volume:<10} | {share:.2f}%")
    print("-" * 60)
    print(f"{'TOTAL':<35} | {total:<10} | 100.00%")
    print("="*60)

    # Verifica se sobrou sujeira no Offline Outros
    sobras = resultado.get('Offline (Outros)', 0)
    if sobras < 100:
        print(f"\n✅ SUCESSO! 'Offline (Outros)' reduzido para apenas {sobras} leads irrelevantes.")
    else:
        print(f"\n⚠️ AVISO: Ainda restam {sobras} leads em 'Offline (Outros)'.")

    # EXPORTAÇÃO
    print(f"\n--- [3/3] Salvando Excel final em: {arquivo_out} ---")
    try:
        os.makedirs(os.path.dirname(arquivo_out), exist_ok=True)
        df.to_excel(arquivo_out, index=False)
        print("✅ Arquivo Excel exportado com sucesso!")
    except ImportError:
        df.to_csv(arquivo_out.replace('.xlsx', '.csv'), index=False, sep=';', encoding='utf-8-sig')
        print("✅ Arquivo CSV salvo (Excel requer 'pip install openpyxl').")

if __name__ == "__main__":
    main()