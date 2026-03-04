import os
import shutil
from datetime import datetime

# Configuração
CAMINHO_BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PASTA_SCRIPTS = os.path.join(CAMINHO_BASE, 'Scritps') # Mantendo o nome da pasta original
DATA_HOJE = datetime.now().strftime('%Y-%m-%d')

# DEFININDO DIRETÓRIO ISOLADO (FORA DO PROJETO)
CAMINHO_PAI = os.path.dirname(CAMINHO_BASE) # Sobe um nível (sai da pasta do projeto)
NOME_PASTA_PROD = f'AMBIENTE_TESTE_ISOLADO_{DATA_HOJE}'
CAMINHO_PROD = os.path.join(CAMINHO_PAI, NOME_PASTA_PROD)

# Scripts oficiais que vão para produção
SCRIPTS_OFICIAIS = [
    '15.Score_External_File.py'
]

print("="*60)
print(f"CRIANDO AMBIENTE ISOLADO: {NOME_PASTA_PROD}")
print(f"LOCAL: {CAMINHO_PROD}")
print("="*60)

# 1. Criar pasta principal
os.makedirs(CAMINHO_PROD, exist_ok=True)
print(f"✅ Pasta criada: {CAMINHO_PROD}")

# 2. Criar subpastas da estrutura
for sub in ['Data', 'Outputs', 'models', 'Scripts']:
    os.makedirs(os.path.join(CAMINHO_PROD, sub), exist_ok=True)
print("✅ Estrutura de pastas (Data, Outputs, models, Scripts) criada.")

# 3. Copiar Scripts Oficiais
print("\n[COPIANDO SCRIPTS]")
for script in SCRIPTS_OFICIAIS:
    src = os.path.join(PASTA_SCRIPTS, script)
    dst = os.path.join(CAMINHO_PROD, 'Scripts', script)
    if os.path.exists(src):
        shutil.copy2(src, dst)
        print(f"  -> Copiado: {script}")
    else:
        print(f"  ❌ ERRO: {script} não encontrado na origem!")

# 4. Copiar Modelo Aprovado (O Cérebro)
print("\n[COPIANDO MODELO APROVADO]")
src_model = os.path.join(CAMINHO_BASE, 'models', 'lead_scoring_model.pkl')
dst_model = os.path.join(CAMINHO_PROD, 'models', 'lead_scoring_model.pkl')

if os.path.exists(src_model):
    shutil.copy2(src_model, dst_model)
    print(f"  ✅ Modelo copiado: lead_scoring_model.pkl")
    print(f"     (Agora você pode rodar o teste SEM precisar da base completa)")
else:
    print(f"  ⚠️  Modelo não encontrado na origem! Você precisará rodar o script 11 na nova pasta.")

# 5. Criar README com instruções
readme_content = f"""
=========================================================
AMBIENTE DE TESTE ISOLADO - LEAD SCORING (IA)
Data de Geração: {DATA_HOJE}
=========================================================

ESTRUTURA DE PASTAS:
/Scripts  -> Contém os códigos Python (O Cérebro)
/Data     -> Onde os dados de entrada devem ser colocados:
             - 'base_teste.csv': Arquivo novo/desconhecido para o modelo classificar.
/Outputs  -> Onde os relatórios Excel serão gerados
/models   -> Onde o arquivo de inteligência fica salvo (.pkl)

---------------------------------------------------------
REQUISITOS DE EXPORTAÇÃO (HUBSPOT):
---------------------------------------------------------
O arquivo 'hubspot_leads.csv' PRECISA ter estas colunas visíveis na exportação:
(Apenas se for retreinar o modelo no futuro. Para teste, basta o base_teste.csv)

---------------------------------------------------------
COMO USAR:
---------------------------------------------------------

1. PROVA DE FOGO (TESTE DOS GESTORES) - **PRIORIDADE**:
   - O modelo aprovado (.pkl) JÁ FOI COPIADO para a pasta /models.
   - Coloque APENAS o arquivo 'base_teste.csv' na pasta /Data.
   - Execute: python Scripts/15.Score_External_File.py
   - O resultado sairá na pasta /Outputs.
"""

with open(os.path.join(CAMINHO_PROD, 'LEIA_ME.txt'), 'w', encoding='utf-8') as f:
    f.write(readme_content)
print("✅ Arquivo LEIA_ME.txt criado com instruções.")

print("\n" + "="*60)
print("CONCLUÍDO! AMBIENTE ISOLADO CRIADO COM SUCESSO.")
print("="*60)