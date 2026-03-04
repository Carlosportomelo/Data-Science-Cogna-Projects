"""
==============================================================================
HÉLIO AI - INTERFACE GRÁFICA (STREAMLIT)
==============================================================================
Para rodar: streamlit run Helio_App.py
==============================================================================
"""
import streamlit as st
import pandas as pd
import os
import subprocess
import sys
import glob
import time
from datetime import datetime
import socket

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Hélio | Red Balloon AI",
    page_icon="🎈",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- PROTEÇÃO POR SENHA ---
if "senha_ok" not in st.session_state:
    st.session_state["senha_ok"] = False

if not st.session_state["senha_ok"]:
    st.markdown("## 🔒 Acesso Restrito | Hélio Red Balloon")
    senha = st.text_input("Digite a senha de acesso:", type="password")
    if st.button("Entrar"):
        if senha == "redballoon": # <--- SUA SENHA AQUI
            st.session_state["senha_ok"] = True
            st.rerun()
        else:
            st.error("Senha incorreta.")
    st.stop()

# --- ESTILO CUSTOMIZADO (CSS) ---
st.markdown("""
    <style>
    .main {
        background-color: #F0F2F6;
    }
    .stButton>button {
        width: 100%;
        background-color: #CE0E2D;
        color: white;
        font-weight: 600;
        height: 3em;
        border-radius: 8px;
        border: none;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #990000;
        color: #e0e0e0;
        transform: translateY(-2px);
        box-shadow: 0 6px 8px rgba(0,0,0,0.15);
    }
    h1, h2, h3 {
        color: #CE0E2D;
        font-family: 'Segoe UI', sans-serif;
    }
    .success-box {
        padding: 1rem;
        background-color: #d4edda;
        color: #155724;
        border-radius: 5px;
        margin-bottom: 1rem;
    }
    /* Estilo de Card para Métricas */
    div[data-testid="stMetric"] {
        background-color: white;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        border-left: 5px solid #CE0E2D;
    }
    /* Container Branco */
    .block-container {
        padding-top: 2rem;
    }
    </style>
    """, unsafe_allow_html=True)

# --- CAMINHOS DOS SCRIPTS ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# Ajuste os nomes aqui se necessário para corresponder exatamente aos seus arquivos
SCRIPT_ML = os.path.join(BASE_DIR, "1.ML_Lead_Scoring.py")
SCRIPT_UNIDADES = os.path.join(BASE_DIR, "4.Analise_Unidades.py") 
SCRIPT_HISTORICO = os.path.join(BASE_DIR, "5.Consolidador_Historico.py")

PASTA_OUTPUT = os.path.join(os.path.dirname(BASE_DIR), "Outputs")
PASTA_RELATORIOS_ML = os.path.join(PASTA_OUTPUT, "Relatorios_ML")
PASTA_RELATORIOS_UNIDADES = os.path.join(PASTA_OUTPUT, "Relatorios_Unidades")
PASTA_DADOS = os.path.join(PASTA_OUTPUT, "Dados_Scored")

# --- FUNção PARA RODAR SCRIPTS ---
# (conteúdo omitido para backup)
