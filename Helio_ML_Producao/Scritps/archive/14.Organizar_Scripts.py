import os
import time

# Caminhos
path_scripts = r'c:\Users\a483650\analise_recuperação_alta26.1\Scritps'
script_antigo = os.path.join(path_scripts, '12.Blind_Test_Accuracy.py')
script_novo = os.path.join(path_scripts, '13_Blind_Test_Ajustado.py')

print("=== ORGANIZANDO SCRIPTS DE VALIDAÇÃO ===")

# 1. Remover o script 12 antigo (se existir)
if os.path.exists(script_antigo):
    try:
        os.remove(script_antigo)
        print(f"✅ [REMOVIDO] {script_antigo}")
    except Exception as e:
        print(f"❌ [ERRO] Não foi possível remover o arquivo antigo: {e}")

# 2. Renomear o script 13 para 12 (para manter a sequência 11 -> 12)
if os.path.exists(script_novo):
    try:
        # Aguarda um pouco para garantir que o sistema liberou o arquivo
        time.sleep(1)
        os.rename(script_novo, script_antigo)
        print(f"✅ [RENOMEADO] {script_novo} \n            -> {script_antigo}")
        print("\nAgora o script oficial de validação é o '12.Blind_Test_Accuracy.py' (com o código atualizado).")
    except Exception as e:
        print(f"❌ [ERRO] Não foi possível renomear o arquivo novo: {e}")
else:
    print(f"⚠️ [AVISO] O script 13 não foi encontrado. Verifique se já foi renomeado.")

print("="*40)