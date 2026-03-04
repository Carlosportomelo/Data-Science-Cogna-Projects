import pandas as pd
import os
import glob
import warnings

warnings.filterwarnings('ignore')

# Configuração de Caminhos
CAMINHO_BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTPUTS_DIR = os.path.join(CAMINHO_BASE, 'Outputs')

# Definição das pastas criadas pelo usuário
PASTAS = {
    'MATRICULADOS': os.path.join(OUTPUTS_DIR, 'matriculados'),
    'PERDIDOS':     os.path.join(OUTPUTS_DIR, 'perdidos'),
    'MISTA':        os.path.join(OUTPUTS_DIR, 'mista')
}

print("="*80)
print("📊 RELATÓRIO CONSOLIDADO DE TESTES (POR PASTA)")
print("="*80)

for nome_cenario, caminho_pasta in PASTAS.items():
    print(f"\n📂 CENÁRIO: {nome_cenario}")
    # print(f"   Caminho: {caminho_pasta}")
    
    if not os.path.exists(caminho_pasta):
        print("   ⚠️  Pasta não encontrada. Certifique-se de criar a pasta em 'Outputs'.")
        continue

    # Buscar todos os Excels na pasta
    arquivos = glob.glob(os.path.join(caminho_pasta, "*.xlsx"))
    arquivos = [a for a in arquivos if not os.path.basename(a).startswith('~$')] # Ignorar temporários
    
    if not arquivos:
        print("   ⚠️  Nenhum arquivo Excel encontrado nesta pasta.")
        continue
        
    print(f"   📄 Arquivos analisados: {len(arquivos)}")
    
    # Consolidar dados
    df_lista = []
    for arq in arquivos:
        try:
            # Tenta ler a aba '2_Base_Pontuada' que o Script 3 gera
            df_temp = pd.read_excel(arq, sheet_name='2_Base_Pontuada')
            df_lista.append(df_temp)
        except Exception as e:
            print(f"      ❌ Erro ao ler {os.path.basename(arq)}: {e}")
    
    if not df_lista:
        continue
        
    df = pd.concat(df_lista, ignore_index=True)
    total_leads = len(df)
    print(f"   → Total de Leads neste cenário: {total_leads:,}")
    
    # Análise de Distribuição
    print("\n   [DISTRIBUIÇÃO DE NOTAS]")
    dist = df['Nota_Modelo'].value_counts(normalize=True).sort_index() * 100
    for nota in range(1, 6):
        perc = dist.get(nota, 0.0)
        bar = "█" * int(perc/5)
        print(f"     Nota {nota}: {perc:5.1f}% {bar}")
        
    # Análise de Assertividade (Se houver gabarito)
    if 'Etapa do negócio' in df.columns:
        print("\n   [VALIDAÇÃO DE PERFORMANCE]")
        
        # Identificar Sucesso vs Perdido
        df['Is_Sucesso'] = df['Etapa do negócio'].astype(str).str.contains('MATRÍCULA', case=False, na=False)
        df['Is_Perdido'] = df['Etapa do negócio'].astype(str).str.contains('PERDIDO', case=False, na=False)
        
        qtd_sucesso = df['Is_Sucesso'].sum()
        qtd_perdido = df['Is_Perdido'].sum()
        
        if qtd_sucesso > 0:
            # Assertividade: Quantos Sucessos receberam nota Alta (4 ou 5)
            acertos = len(df[(df['Is_Sucesso']) & (df['Nota_Modelo'] >= 4)])
            taxa = (acertos / qtd_sucesso) * 100
            print(f"     ✅ ASSERTIVIDADE (Matrículas): {taxa:.1f}%")
            print(f"        (O modelo identificou {acertos} de {qtd_sucesso} matrículas como Alta Prioridade)")
            
        if qtd_perdido > 0:
            # Eficiência: Quantos Perdidos receberam nota Baixa (1 ou 2)
            filtros = len(df[(df['Is_Perdido']) & (df['Nota_Modelo'] <= 2)])
            taxa = (filtros / qtd_perdido) * 100
            print(f"     🛡️  EFICIÊNCIA (Perdidos):     {taxa:.1f}%")
            print(f"        (O modelo descartou {filtros} de {qtd_perdido} leads ruins automaticamente)")
            
    # Tradução para o Negócio (Qualificação vs Pausa)
    print("\n   [TRADUÇÃO PARA O NEGÓCIO]")
    qtd_qualif = len(df[df['Nota_Modelo'] >= 4])
    qtd_pausa = len(df[df['Nota_Modelo'] <= 2])
    qtd_obs = len(df[df['Nota_Modelo'] == 3])
    
    print(f"     🚀 QUALIFICAÇÃO (Notas 4-5): {qtd_qualif/total_leads:.1%} dos leads -> Foco Total (Vendas)")
    print(f"     ⏸️  EM PAUSA (Notas 1-2):     {qtd_pausa/total_leads:.1%} dos leads -> Nutrição (Marketing)")
    if qtd_obs > 0:
        print(f"     👀 OBSERVAÇÃO (Nota 3):      {qtd_obs/total_leads:.1%} dos leads -> Análise Humana")

    print("-" * 80)

print("\n🏁 Análise concluída.")