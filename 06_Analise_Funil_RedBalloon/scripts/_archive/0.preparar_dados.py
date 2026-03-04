#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Script de Preparação de Dados - Red Balloon
Filtra os dados do HubSpot completo para incluir apenas leads da Red Balloon
"""

import pandas as pd
from pathlib import Path
import sys

# Caminhos
BASE_DIR = Path(__file__).parent.parent
CENTRAL_DATA = BASE_DIR.parent / "_DADOS_CENTRALIZADOS" / "hubspot" / "hubspot_leads_ATUAL.csv"
OUTPUT_FILE = BASE_DIR / "data" / "hubspot_leads_redballoon.csv"

print("="*80)
print("   🎈 PREPARAÇÃO DE DADOS - RED BALLOON")
print("="*80)
print(f"\n📥 Origem: {CENTRAL_DATA}")
print(f"📤 Destino: {OUTPUT_FILE}\n")

# Verificar se arquivo fonte existe
if not CENTRAL_DATA.exists():
    print(f"❌ ERRO: Arquivo não encontrado: {CENTRAL_DATA}")
    print("\n💡 Execute primeiro:")
    print("   python ../_SCRIPTS_COMPARTILHADOS/sincronizar_bases.py")
    sys.exit(1)

# Carregar dados
print("📂 Carregando dados do HubSpot...")
df = pd.read_csv(CENTRAL_DATA, encoding='utf-8')
print(f"   ✅ {len(df):,} registros carregados")

# Filtrar Red Balloon
print("\n🔍 Filtrando leads da Red Balloon...")
df_redballoon = df[
    df['Pipeline'].str.contains('Red Balloon', case=False, na=False) |
    df['Etapa do negócio'].str.contains('Red Balloon', case=False, na=False)
].copy()

print(f"   ✅ {len(df_redballoon):,} leads da Red Balloon identificados")
print(f"   📊 {len(df_redballoon)/len(df)*100:.1f}% do total")

# Salvar
print("\n💾 Salvando arquivo filtrado...")
OUTPUT_FILE.parent.mkdir(exist_ok=True)
df_redballoon.to_csv(OUTPUT_FILE, index=False, encoding='utf-8')
print(f"   ✅ Arquivo salvo: {OUTPUT_FILE}")

print("\n" + "="*80)
print("   ✅ PREPARAÇÃO CONCLUÍDA!")
print("="*80 + "\n")
