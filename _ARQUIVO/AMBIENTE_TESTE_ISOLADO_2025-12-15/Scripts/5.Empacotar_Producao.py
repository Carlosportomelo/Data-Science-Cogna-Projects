import os
import shutil
from datetime import datetime

# Configuração de Caminhos
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTPUT_DIR = os.path.join(BASE_DIR, 'Entrega_Final_Producao')
MODEL_SRC = os.path.join(BASE_DIR, 'models', 'lead_scoring_model.pkl')
SCRIPT_SRC = os.path.join(BASE_DIR, 'Scripts', '3.Score_External_File.py')

# Criar diretório de entrega
if os.path.exists(OUTPUT_DIR):
    shutil.rmtree(OUTPUT_DIR)
os.makedirs(OUTPUT_DIR)

print(f"📦 GERANDO PACOTE DE ENTREGA EM: {OUTPUT_DIR}")
print("-" * 60)

# 1. Copiar Modelo Validado
if os.path.exists(MODEL_SRC):
    shutil.copy2(MODEL_SRC, os.path.join(OUTPUT_DIR, 'lead_scoring_model.pkl'))
    print("   ✅ [MODELO] lead_scoring_model.pkl copiado com sucesso.")
else:
    print("   ❌ [ERRO] Modelo .pkl não encontrado!")

# 2. Copiar Script Validado (Renomeando para Pipeline)
if os.path.exists(SCRIPT_SRC):
    dest_script = os.path.join(OUTPUT_DIR, 'pipeline_scoring_validado.py')
    shutil.copy2(SCRIPT_SRC, dest_script)
    print("   ✅ [LÓGICA] Script de Scoring copiado como 'pipeline_scoring_validado.py'.")
else:
    print("   ❌ [ERRO] Script de Scoring não encontrado!")

# 3. Criar Instruções de Migração
readme_content = f"""PACOTE DE MIGRAÇÃO - LEAD SCORING
Data de Validação: {datetime.now().strftime('%Y-%m-%d')}
Status: APROVADO EM TESTE CEGO
=========================================================

1. lead_scoring_model.pkl:
   Substitua o arquivo antigo na pasta 'models' do projeto original.

2. pipeline_scoring_validado.py:
   Este código contém a lógica validada (Blindagem + Explicabilidade).
   Use-o para atualizar seu script de produção.
"""
with open(os.path.join(OUTPUT_DIR, 'LEIA_ME_MIGRACAO.txt'), 'w', encoding='utf-8') as f:
    f.write(readme_content)

print("-" * 60)
print("🏁 Pacote pronto! Copie a pasta 'Entrega_Final_Producao' para o projeto original.")