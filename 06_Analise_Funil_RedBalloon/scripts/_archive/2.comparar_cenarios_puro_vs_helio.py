#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
🎈 COMPARAÇÃO DE CENÁRIOS - ANÁLISE PURA vs HELIO
===================================================

OBJETIVO:
    Gerar DOIS relatórios de funil para comparação:
    1. Relatório PURO: Dados brutos (apenas remove erro de data)
    2. Relatório HELIO: Aplica regras de ML do Helio
    
REGRAS DO HELIO (Aprendidas do ML Lead Scoring):
    ✅ Remove matrículas com < 3 atividades comerciais
       Motivo: Atendimento presencial não registrado distorce volumetria
       (indica que "quanto menos falar, melhor" - quando na verdade não é)
    
    ✅ Corrige datas de atividade vazias
       Preenche com Data de Criação quando não há registro

SAÍDA:
    - funil_redballoon_PURO_[timestamp].xlsx
    - funil_redballoon_HELIO_[timestamp].xlsx

Autor: Sistema de Data Science  
Data: Fevereiro 2026
"""

import subprocess
import sys
from pathlib import Path
from datetime import datetime

# Configuração
BASE_DIR = Path(__file__).parent.parent
SCRIPT_ORIGINAL = BASE_DIR / "scripts" / "1.analise_funil_redballoon.py"

print("="*80)
print("   COMPARAÇÃO DE CENÁRIOS - PURO vs HELIO")
print("="*80)
print(f"\nInicio: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")

print("📊 Este script irá gerar DOIS relatórios para comparação:\n")
print("  1️⃣  PURO: Dados brutos (apenas remove erro de data)")
print("       -> Mostra a realidade dos dados do HubSpot\n")
print("  2️⃣  HELIO: Aplica regras de ML (remove matrículas com < 3 atividades)")
print("       -> Análise refinada seguindo critérios do modelo de IA\n")

print("⏳ Gerando relatório PURO...")
print("-"*80)

# Executar análise PURA
cmd_puro = [sys.executable, str(SCRIPT_ORIGINAL), "--modo", "puro"]
result_puro = subprocess.run(cmd_puro, capture_output=False, text=True)

if result_puro.returncode != 0:
    print("\n❌ ERRO ao gerar relatório PURO")
    sys.exit(1)

print("\n" + "="*80)
print("⏳ Gerando relatório HELIO...")
print("-"*80)

# Executar análise HELIO
cmd_helio = [sys.executable, str(SCRIPT_ORIGINAL), "--modo", "helio"]
result_helio = subprocess.run(cmd_helio, capture_output=False, text=True)

if result_helio.returncode != 0:
    print("\n❌ ERRO ao gerar relatório HELIO")
    sys.exit(1)

print("\n" + "="*80)
print("   ✅ COMPARAÇÃO CONCLUÍDA COM SUCESSO!")
print("="*80)

print("\n📁 Arquivos gerados em: outputs/")
print("\n💡 PRÓXIMOS PASSOS:")
print("   1. Abra os dois arquivos Excel")
print("   2. Compare as métricas (principalmente conversão e atividades)")
print("   3. Use o relatório HELIO para decisões estratégicas")
print("   4. Use o relatório PURO para auditar qualidade dos dados\n")

print(f"🕐 Conclusão: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
