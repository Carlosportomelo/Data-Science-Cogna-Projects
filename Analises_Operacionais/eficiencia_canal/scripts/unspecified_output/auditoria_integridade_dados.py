import pandas as pd
import numpy as np
from datetime import datetime
import os

# Caminhos dos arquivos
base_atual = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\data\hubspot_leads.csv'
base_backup = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\data\backup\hubspot_leads.csv'

print("=" * 80)
print("AUDITORIA DE INTEGRIDADE DE DADOS - HubSpot Leads")
print("=" * 80)
print(f"Data da auditoria: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
print()

# Carregar dados
print("Carregando arquivos...")
df_atual = pd.read_csv(base_atual)
df_backup = pd.read_csv(base_backup)

print(f"✓ Base atual carregada: {len(df_atual)} registros")
print(f"✓ Base backup carregada: {len(df_backup)} registros")
print()

# Converter datas
df_atual['Data de criação'] = pd.to_datetime(df_atual['Data de criação'], format='%Y-%m-%d %H:%M')
df_backup['Data de criação'] = pd.to_datetime(df_backup['Data de criação'], format='%Y-%m-%d %H:%M')

# Filtrar período: outubro/24 a fevereiro/25
data_inicio = pd.Timestamp('2024-10-01')
data_fim = pd.Timestamp('2025-02-28')

df_atual_filtrado = df_atual[(df_atual['Data de criação'] >= data_inicio) & 
                              (df_atual['Data de criação'] <= data_fim)].copy()
df_backup_filtrado = df_backup[(df_backup['Data de criação'] >= data_inicio) & 
                                (df_backup['Data de criação'] <= data_fim)].copy()

print("FILTRO POR PERÍODO: outubro/24 a fevereiro/25")
print("-" * 80)
print(f"Registros base atual neste período: {len(df_atual_filtrado)}")
print(f"Registros base backup neste período: {len(df_backup_filtrado)}")
print()

# Comparação por Record ID
record_ids_atual = set(df_atual_filtrado['Record ID'].astype(str))
record_ids_backup = set(df_backup_filtrado['Record ID'].astype(str))

ids_novos = record_ids_atual - record_ids_backup
ids_removidos = record_ids_backup - record_ids_atual
ids_comuns = record_ids_atual & record_ids_backup

print("ANÁLISE DE RECORD IDs:")
print("-" * 80)
print(f"IDs novos (apenas na base atual): {len(ids_novos)}")
print(f"IDs removidos (apenas na base backup): {len(ids_removidos)}")
print(f"IDs em comum: {len(ids_comuns)}")
print()

# Comparar registros comuns
if ids_comuns:
    print("ANÁLISE DE MODIFICAÇÕES NOS REGISTROS COMUNS:")
    print("-" * 80)
    
    df_atual_comuns = df_atual_filtrado[df_atual_filtrado['Record ID'].astype(str).isin(ids_comuns)].set_index('Record ID')
    df_backup_comuns = df_backup_filtrado[df_backup_filtrado['Record ID'].astype(str).isin(ids_comuns)].set_index('Record ID')
    
    # Comparar coluna por coluna
    colunas = df_atual_comuns.columns
    modificacoes_por_coluna = {}
    registros_modificados = set()
    
    for col in colunas:
        # Comparar valores
        diferentes = df_atual_comuns[col] != df_backup_comuns[col]
        
        # Lidar com NaN
        mask_nan_atual = df_atual_comuns[col].isna()
        mask_nan_backup = df_backup_comuns[col].isna()
        diferentes = diferentes & ~(mask_nan_atual & mask_nan_backup)
        
        qtd_diferentes = diferentes.sum()
        if qtd_diferentes > 0:
            modificacoes_por_coluna[col] = qtd_diferentes
            registros_modificados.update(diferentes[diferentes].index.astype(str))
    
    print(f"Total de registros modificados: {len(registros_modificados)}")
    print()
    
    if modificacoes_por_coluna:
        print("Campos modificados:")
        for col, qtd in sorted(modificacoes_por_coluna.items(), key=lambda x: x[1], reverse=True):
            print(f"  • {col}: {qtd} registros")
        print()
        
        # Mostar detalhes das modificações
        print("DETALHES DAS MODIFICAÇÕES:")
        print("-" * 80)
        
        for record_id in sorted(registros_modificados)[:20]:  # Mostrar até 20
            print(f"\nRecord ID: {record_id}")
            for col in colunas:
                val_atual = df_atual_comuns.loc[int(record_id), col]
                val_backup = df_backup_comuns.loc[int(record_id), col]
                
                # Comparar
                if pd.isna(val_atual) and pd.isna(val_backup):
                    continue
                elif str(val_atual) != str(val_backup):
                    print(f"  {col}:")
                    print(f"    Base atual   : {val_atual}")
                    print(f"    Base backup  : {val_backup}")
        
        if len(registros_modificados) > 20:
            print(f"\n... e mais {len(registros_modificados) - 20} registros modificados")
    else:
        print("Nenhuma modificação detectada nos registros comuns!")
else:
    print("Nenhum registro em comum para comparar!")

print()
print("=" * 80)
print("RESUMO EXECUTIVO:")
print("=" * 80)
print(f"Novos leads: {len(ids_novos)}")
print(f"Leads removidos: {len(ids_removidos)}")
print(f"Leads modificados: {len(registros_modificados)}")
print(f"Integridade: {'✓ OK' if len(ids_removidos) == 0 and len(registros_modificados) == 0 else '✗ ALTERAÇÕES DETECTADAS'}")
print("=" * 80)

# Exportar relatório
output_file = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\outputs\relatorio_auditoria_dados\auditoria_integridade_2026-02-06.txt'
os.makedirs(os.path.dirname(output_file), exist_ok=True)

with open(output_file, 'w', encoding='utf-8') as f:
    f.write("AUDITORIA DE INTEGRIDADE DE DADOS - HubSpot Leads\n")
    f.write("=" * 80 + "\n")
    f.write(f"Data: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
    f.write(f"Período analisado: outubro/24 a fevereiro/25\n\n")
    
    f.write(f"Base atual: {len(df_atual_filtrado)} registros\n")
    f.write(f"Base backup: {len(df_backup_filtrado)} registros\n\n")
    
    f.write(f"IDs novos: {len(ids_novos)}\n")
    f.write(f"IDs removidos: {len(ids_removidos)}\n")
    f.write(f"IDs modificados: {len(registros_modificados)}\n")

print(f"\n✓ Relatório exportado para: {output_file}")
