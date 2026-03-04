import os
import shutil

# Configuração
CAMINHO_SCRIPTS = os.path.dirname(os.path.abspath(__file__))
PASTA_ARCHIVE = os.path.join(CAMINHO_SCRIPTS, 'Archive')

# Scripts para MANTER na raiz (O Trio de Ouro + Este script)
KEEP_FILES = [
    '11.ML_Lead_Scoring.py',
    '12.Blind_Test_Accuracy.py',
    '15.Score_External_File.py',
    '99.Arquivar_Scripts_Antigos.py'
]

print("="*60)
print(f"🧹 LIMPEZA DE AMBIENTE - MANTENDO APENAS O ESSENCIAL")
print("="*60)
print(f"📂 Diretório: {CAMINHO_SCRIPTS}")
print(f"🛡️  Arquivos Protegidos: {KEEP_FILES}")

# Criar pasta Archive se não existir
if not os.path.exists(PASTA_ARCHIVE):
    os.makedirs(PASTA_ARCHIVE)
    print(f"✅ Pasta 'Archive' criada.")

count = 0
for item in os.listdir(CAMINHO_SCRIPTS):
    caminho_item = os.path.join(CAMINHO_SCRIPTS, item)
    
    # Ignorar pastas e arquivos que não são .py (mantém .bat ou .txt se houver)
    if os.path.isdir(caminho_item) or not item.endswith('.py'):
        continue

    # Se não estiver na lista de protegidos, move para Archive
    if item not in KEEP_FILES:
        dst = os.path.join(PASTA_ARCHIVE, item)
        try:
            shutil.move(caminho_item, dst)
            print(f"   -> 📦 Arquivado: {item}")
            count += 1
        except Exception as e:
            print(f"   ❌ [ERRO] Ao mover {item}: {e}")

print("-" * 60)
print(f"🏁 Limpeza concluída! {count} scripts movidos para a pasta 'Archive'.")
print("   Sua pasta agora contém apenas os scripts operacionais (11, 12, 15).")