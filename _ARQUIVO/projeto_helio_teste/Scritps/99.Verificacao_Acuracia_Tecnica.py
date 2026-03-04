"""
==============================================================================
VERIFICAÇÃO TÉCNICA DE ACURÁCIA (FACT-CHECK)
==============================================================================
Script: 99.Verificacao_Acuracia_Tecnica.py
Objetivo: Calcular explicitamente a métrica de ACURÁCIA (Accuracy Score)
          para validar o número de 99,47% mencionado na apresentação.
==============================================================================
"""

import pandas as pd
import numpy as np
import os
from datetime import datetime
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import accuracy_score, confusion_matrix
from sklearn.preprocessing import LabelEncoder
import warnings
warnings.filterwarnings('ignore')

# CONFIGURAÇÃO
CAMINHO_BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CAMINHO_LEADS = os.path.join(CAMINHO_BASE, 'Data', 'hubspot_leads.csv')
CAMINHO_PERDIDOS = os.path.join(CAMINHO_BASE, 'Data', 'hubspot_negocios_perdidos.csv')

print("="*80)
print("VERIFICAÇÃO DE ACURÁCIA (FACT-CHECK)")
print("="*80)

# 1. CARREGAR
print("Carregando dados...")
try:
    df = pd.read_csv(CAMINHO_LEADS, encoding='utf-8')
except:
    df = pd.read_csv(CAMINHO_LEADS, encoding='latin1')
df.columns = df.columns.str.strip()

# Carregar perdidos para enriquecer (igual ao script original)
if os.path.exists(CAMINHO_PERDIDOS):
    try:
        df_perdidos = pd.read_csv(CAMINHO_PERDIDOS, encoding='utf-8')
    except:
        df_perdidos = pd.read_csv(CAMINHO_PERDIDOS, encoding='latin1')
    df_perdidos.columns = df_perdidos.columns.str.strip()
    
    # Padronizar ID
    if 'Record ID' in df.columns:
        df['Record ID'] = df['Record ID'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    if 'Record ID' in df_perdidos.columns:
        df_perdidos['Record ID'] = df_perdidos['Record ID'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        df_perdidos = df_perdidos.drop_duplicates('Record ID')
    
    cols_enrich = ['Record ID']
    motivo_col = [c for c in df_perdidos.columns if 'motivo' in c.lower()]
    if motivo_col: cols_enrich.append(motivo_col[0])
    
    df_enrich = df_perdidos[cols_enrich].copy()
    if len(cols_enrich) > 1: df_enrich.columns = ['Record ID', 'Motivo_Perda']
    df_enrich['In_Lost_Base'] = True
    
    df = df.merge(df_enrich, on='Record ID', how='left')
    df['In_Lost_Base'] = df['In_Lost_Base'].fillna(False)

# 2. FEATURES
print("Preparando features...")
df['Data de criação'] = pd.to_datetime(df['Data de criação'], errors='coerce')
df['Data de fechamento'] = pd.to_datetime(df['Data de fechamento'], errors='coerce')

data_referencia = df['Data de fechamento'].fillna(datetime.now())
df['Dias_Desde_Criacao'] = (data_referencia - df['Data de criação']).dt.days.clip(lower=0)

df['Ano'] = df['Data de criação'].dt.year
df['Mes'] = df['Data de criação'].dt.month
df['DiaSemana'] = df['Data de criação'].dt.dayofweek
df['Trimestre'] = df['Data de criação'].dt.quarter

df['Fonte original do tráfego'] = df['Fonte original do tráfego'].fillna('(Sem fonte)')
df['Detalhamento da fonte original do tráfego 1'] = df['Detalhamento da fonte original do tráfego 1'].fillna('(Sem detalhe)')
df['Número de atividades de vendas'] = pd.to_numeric(df['Número de atividades de vendas'], errors='coerce').fillna(0)
df['Unidade Desejada'] = df['Unidade Desejada'].fillna('(Sem unidade)')

for col in ['Ano', 'Mes', 'DiaSemana', 'Trimestre', 'Dias_Desde_Criacao']:
    df[col] = df[col].fillna(0)

df['Densidade_Atividade'] = df['Número de atividades de vendas'] / (df['Dias_Desde_Criacao'] + 1)

def classificar_atividade_comercial(n):
    if n < 3: return '0_Outlier_Sem_Registro'
    elif n <= 6: return '1_Inicial_3_a_6'
    elif n <= 12: return '2_Engajamento_7_a_12'
    elif n <= 20: return '3_Persistencia_13_a_20'
    elif n <= 35: return '4_Alta_Persistencia_21_a_35'
    elif n <= 60: return '5_Recuperacao_Longa_36_a_60'
    else: return '6_Saturacao_Extrema_60+'
df['Faixa_Atividade'] = df['Número de atividades de vendas'].apply(classificar_atividade_comercial)

def categorizar_fonte(fonte):
    fonte = str(fonte).lower()
    if 'off-line' in fonte or 'offline' in fonte: return 'Offline'
    elif 'social pago' in fonte: return 'Meta_Ads'
    elif 'pesquisa paga' in fonte: return 'Google_Ads'
    elif 'pesquisa orgânica' in fonte: return 'Google_Organico'
    elif 'social orgânico' in fonte: return 'Social_Organico'
    elif 'tráfego direto' in fonte: return 'Direto'
    else: return 'Outros'
df['Fonte_Cat'] = df['Fonte original do tráfego'].apply(categorizar_fonte)

def definir_segmento_ml(fonte_cat):
    if fonte_cat == 'Offline': return 'Offline'
    elif fonte_cat in ['Google_Ads', 'Google_Organico']: return 'Digital_Busca'
    elif fonte_cat in ['Meta_Ads', 'Social_Organico']: return 'Digital_Social'
    else: return 'Digital_Direto'
df['Segmento_ML'] = df['Fonte_Cat'].apply(definir_segmento_ml)

def categorizar_detalhe_offline(row):
    if row['Fonte_Cat'] != 'Offline': return 'N/A'
    detalhe = str(row['Detalhamento da fonte original do tráfego 1']).upper()
    if 'CRM_UI' in detalhe: return 'CRM_UI'
    elif 'CONTACTS' in detalhe: return 'CONTACTS'
    elif 'IMPORT' in detalhe: return 'IMPORT'
    else: return 'Outros_Offline'
df['Detalhe_Offline'] = df.apply(categorizar_detalhe_offline, axis=1)

# 3. TARGET
def classificar_etapa(row):
    etapa = str(row['Etapa do negócio']).lower()
    in_lost = row.get('In_Lost_Base', False)
    if 'matrícula' in etapa: return 'SUCESSO'
    elif 'perdido' in etapa: return 'PERDIDO'
    elif in_lost: return 'PERDIDO'
    else: return 'IGNORE'
df['Etapa_Class'] = df.apply(classificar_etapa, axis=1)

df_treino = df[df['Etapa_Class'].isin(['SUCESSO', 'PERDIDO'])].copy()
df_treino['Target'] = (df_treino['Etapa_Class'] == 'SUCESSO').astype(int)

# 4. ENCODING
FEATURES_CAT = ['Fonte_Cat', 'Detalhe_Offline', 'Unidade Desejada', 'Faixa_Atividade']
FEATURES_NUM_BASE = ['Mes', 'DiaSemana', 'Trimestre', 'Dias_Desde_Criacao', 'Densidade_Atividade']

for col in FEATURES_CAT:
    le = LabelEncoder()
    le.fit(df[col].astype(str))
    df_treino[f'{col}_enc'] = le.transform(df_treino[col].astype(str))

# 5. TREINAMENTO E VERIFICAÇÃO
print("\n" + "="*80)
print("RESULTADOS DA VERIFICAÇÃO DE ACURÁCIA")
print("="*80)

segmentos = df_treino['Segmento_ML'].unique()

for seg in segmentos:
    if seg == 'Offline':
        cols_cat = ['Detalhe_Offline', 'Unidade Desejada', 'Faixa_Atividade']
        cols_cat_enc = [f'{c}_enc' for c in cols_cat]
        df_seg = df_treino[(df_treino['Segmento_ML'] == seg) & (df_treino['Número de atividades de vendas'] >= 3)].copy()
        features = cols_cat_enc + FEATURES_NUM_BASE + ['Número de atividades de vendas']
    else:
        cols_cat = ['Fonte_Cat', 'Unidade Desejada', 'Faixa_Atividade']
        cols_cat_enc = [f'{c}_enc' for c in cols_cat]
        df_seg = df_treino[df_treino['Segmento_ML'] == seg].copy()
        features = cols_cat_enc + FEATURES_NUM_BASE

    if len(df_seg) < 50: continue

    X = df_seg[features].values
    y = df_seg['Target'].values

    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42, stratify=y)

    clf = RandomForestClassifier(n_estimators=100, max_depth=12, random_state=42, class_weight='balanced')
    clf.fit(X_train, y_train)

    # Acurácia
    acc_train = accuracy_score(y_train, clf.predict(X_train))
    acc_test = accuracy_score(y_test, clf.predict(X_test))
    
    print(f"\nSEGMENTO: {seg}")
    print(f"  - Amostras: {len(df_seg)}")
    print(f"  - Acurácia TREINO: {acc_train:.2%}")
    print(f"  - Acurácia TESTE:  {acc_test:.2%}")
    
    if acc_train > 0.99 and acc_test < 0.90:
        print("  ⚠️  ALERTA: Possível Overfitting (Modelo decorou os dados de treino)")
    elif acc_test > 0.99:
        print("  ✅ ALTA PRECISÃO CONFIRMADA (Verificar se não há desbalanceamento extremo)")

print("\n" + "="*80)