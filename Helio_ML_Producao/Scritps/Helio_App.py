"""
==============================================================================
HÉLIO GROWTH IA - PLATAFORMA INTEGRADA (STREAMLIT + DESIGN SYSTEM)
Versão 2.0 - Corrigida e Otimizada
==============================================================================
"""
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import logging

# ==============================================================================
# 0. CONFIGURAÇÃO E LOGGING
# ==============================================================================
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Verifica dependências críticas
try:
    import plotly.express as px
    import plotly.graph_objects as go
except ImportError:
    st.error("❌ Plotly não instalado. Execute: pip install plotly")
    st.stop()

# ==============================================================================
# 1. CONFIGURAÇÃO DA PÁGINA (ANTES DE QUALQUER OUTRO COMANDO STREAMLIT)
# ==============================================================================
st.set_page_config(
    page_title="Helio - Growth IA RedBalloon",
    page_icon="🎈",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==============================================================================
# 2. INICIALIZAÇÃO DO SESSION STATE
# ==============================================================================
if "senha_ok" not in st.session_state:
    st.session_state["senha_ok"] = False
if "show_simulador" not in st.session_state:
    st.session_state["show_simulador"] = False
if "menu_atual" not in st.session_state:
    st.session_state["menu_atual"] = "Visão Geral (Rede)"

# ==============================================================================
# 3. CSS INJECT (STYLES OTIMIZADOS E CORRIGIDOS)
# ==============================================================================
st.markdown("""
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    
    <style>
    /* --- VARIÁVEIS DE COR (DESIGN SYSTEM) --- */
    :root {
        --red-50: #fef2f2;
        --red-100: #fee2e2;
        --red-500: #ef4444;
        --red-600: #dc2626;
        --red-700: #b91c1c;
        --red-900: #7f1d1d;
        --helio-light: #f5f3ff;
        --helio-main: #7c3aed;
        --slate-50: #f8fafc;
        --slate-100: #f1f5f9;
        --slate-200: #e2e8f0;
        --slate-800: #1e293b;
        --slate-900: #0f172a;
        --green-500: #10b981;
        --amber-400: #fbbf24;
        --orange-500: #f97316;
    }

    /* --- FUNDO GERAL --- */
    .stApp { 
        background-color: var(--slate-50); 
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Helvetica Neue', sans-serif;
    }

    /* --- SIDEBAR CUSTOMIZADA --- */
    [data-testid="stSidebar"] { 
        background-color: var(--slate-900) !important;
        border-right: 1px solid #334155;
    }
    
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"],
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3,
    [data-testid="stSidebar"] span,
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] button {
        color: white !important;
    }

    [data-testid="stSidebar"] [data-baseweb="select"] > div {
        background-color: #1e293b !important;
        border: 1px solid #334155 !important;
    }

    [data-testid="stSidebar"] [data-baseweb="radio"] {
        background-color: transparent !important;
    }

    /* --- BANNER DE BOAS VINDAS --- */
    .helio-banner {
        background: linear-gradient(135deg, var(--red-700), var(--red-900));
        border-radius: 0.75rem;
        padding: 1.5rem;
        color: white;
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1);
        position: relative;
        overflow: hidden;
        margin-bottom: 1.5rem;
    }
    
    .helio-banner h2 {
        color: white !important;
        margin: 0 !important;
        font-weight: bold !important;
    }
    
    .helio-banner p {
        color: rgba(255, 255, 255, 0.9) !important;
        margin-top: 0.5rem !important;
        margin-bottom: 0 !important;
    }

    /* --- CARDS DE KPI --- */
    [data-testid="metric-container"] {
        background-color: white;
        border: 1px solid var(--slate-200);
        border-radius: 0.75rem;
        padding: 1.25rem;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    
    [data-testid="metric-container"]:hover { 
        transform: translateY(-2px);
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1);
    }

    /* --- KANBAN STYLING --- */
    .kanban-col-header {
        padding: 0.75rem;
        background-color: var(--slate-100);
        border-radius: 0.5rem 0.5rem 0 0;
        border-bottom: 1px solid var(--slate-200);
        font-weight: 600;
        color: var(--slate-800);
        display: flex;
        justify-content: space-between;
        align-items: center;
        font-size: 14px;
    }

    .kanban-card {
        background: white;
        padding: 0.75rem;
        border-radius: 0.5rem;
        margin-bottom: 0.5rem;
        border: 1px solid var(--slate-200);
        box-shadow: 0 1px 2px rgba(0, 0, 0, 0.05);
        transition: all 0.2s ease;
    }

    .kanban-card:hover { 
        transform: translateY(-2px);
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }

    .kanban-card-title {
        font-weight: bold;
        color: var(--slate-800);
        font-size: 14px;
        margin-top: 0.5rem;
    }

    .kanban-card-footer {
        display: flex;
        justify-content: space-between;
        margin-top: 0.5rem;
        padding-top: 0.5rem;
        border-top: 1px solid var(--slate-50);
        font-weight: bold;
        font-size: 12px;
        color: #475569;
    }

    /* --- DESTAQUE NOTA 5 --- */
    .helio-score-5 {
        border-left: 4px solid var(--red-600) !important;
        box-shadow: 0 0 12px rgba(220, 38, 38, 0.25) !important;
        background: linear-gradient(90deg, rgba(252, 165, 165, 0.05), transparent) !important;
    }

    /* --- BOTÕES --- */
    .stButton > button {
        background-color: var(--red-600);
        color: white;
        border-radius: 0.5rem;
        border: none;
        font-weight: 600;
        padding: 0.5rem 1rem;
        transition: background-color 0.2s ease;
    }

    .stButton > button:hover { 
        background-color: var(--red-700);
    }

    .stButton > button:active {
        background-color: var(--red-900);
    }

    /* --- ANIMAÇÃO BALÃO --- */
    @keyframes float { 
        0% { transform: translateY(0px); } 
        50% { transform: translateY(-3px); } 
        100% { transform: translateY(0px); } 
    }
    .balloon-float { 
        animation: float 3s ease-in-out infinite; 
        display: inline-block; 
    }

    /* --- EXPANDER CUSTOMIZADO --- */
    [data-testid="stExpander"] {
        border-radius: 0.5rem;
    }

    /* --- TITLE E HEADERS --- */
    h1, h2, h3 {
        color: var(--slate-800) !important;
    }

    /* --- INPUT FIELDS --- */
    [data-baseweb="input"] input,
    [data-baseweb="select"] > div {
        border-radius: 0.5rem !important;
    }

    /* --- DATATABLE --- */
    [data-testid="stDataFrame"] {
        border-radius: 0.5rem;
    }

    /* --- PLOTLY CHARTS --- */
    .plotly {
        border-radius: 0.5rem;
    }

    /* --- LOGIN PAGE --- */
    .login-container {
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        min-height: 100vh;
        background: linear-gradient(135deg, var(--red-50), var(--helio-light));
    }

    .login-box {
        background: white;
        border-radius: 1rem;
        padding: 2rem;
        box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.1);
        width: 100%;
        max-width: 400px;
    }

    .login-icon {
        width: 80px;
        height: 80px;
        background: white;
        border-radius: 1.25rem;
        display: flex;
        align-items: center;
        justify-content: center;
        margin: 0 auto 1rem;
        box-shadow: 0 10px 25px -5px rgba(220, 38, 38, 0.3);
        font-size: 2.5rem;
    }
    </style>
""", unsafe_allow_html=True)

# ==============================================================================
# 4. FUNÇÕES AUXILIARES
# ==============================================================================

@st.cache_data
def get_kanban_data():
    """Retorna dados do kanban com tratamento de erro"""
    try:
        return [
            {"nome": "Pedro Alencar", "val": "3.800", "status": 0, "score": 5},
            {"nome": "Maria Clara", "val": "4.200", "status": 1, "score": 4},
            {"nome": "João Pedro", "val": "3.500", "status": 0, "score": 2},
            {"nome": "Ana Beatriz", "val": "3.800", "status": 2, "score": 5},
            {"nome": "Lucas Silva", "val": "4.500", "status": 3, "score": 5},
            {"nome": "Carolina Costa", "val": "3.900", "status": 1, "score": 3},
            {"nome": "Rafael Santos", "val": "4.100", "status": 2, "score": 4},
        ]
    except Exception as e:
        logger.error(f"Erro ao carregar dados kanban: {str(e)}")
        return []

@st.cache_data
def get_units_data():
    """Retorna dados das unidades com tratamento de erro"""
    try:
        return pd.DataFrame([
            {"Unidade": "Vila Leopoldina", "Tipo": "Própria", "Score Médio": 4.8, "Conversão": 0.22, "Receita": 142500},
            {"Unidade": "Moema", "Tipo": "Própria", "Score Médio": 4.6, "Conversão": 0.19, "Receita": 118200},
            {"Unidade": "Bauru", "Tipo": "Franquia", "Score Médio": 3.2, "Conversão": 0.12, "Receita": 45100},
            {"Unidade": "Santos", "Tipo": "Franquia", "Score Médio": 3.5, "Conversão": 0.15, "Receita": 67800},
            {"Unidade": "Campinas", "Tipo": "Própria", "Score Médio": 4.4, "Conversão": 0.18, "Receita": 98500},
        ])
    except Exception as e:
        logger.error(f"Erro ao carregar dados de unidades: {str(e)}")
        return pd.DataFrame()

@st.cache_data
def get_regional_data():
    """Retorna dados regionais com tratamento de erro"""
    try:
        return pd.DataFrame({
            'Regional': ['SP Capital', 'SP Interior', 'Rio de Janeiro', 'Minas', 'Sul'],
            'Realizado': [120, 80, 45, 30, 15],
            'Meta': [100, 85, 50, 25, 20]
        })
    except Exception as e:
        logger.error(f"Erro ao carregar dados regionais: {str(e)}")
        return pd.DataFrame()

def render_kanban_card(card: dict) -> str:
    """Renderiza um card do kanban com validação"""
    try:
        if not all(key in card for key in ['nome', 'val', 'score']):
            return ""
        
        score_cls = "helio-score-5" if card['score'] == 5 else ""
        
        # Gera estrelas dinamicamente
        stars_html = ""
        for i in range(5):
            if i < card['score']:
                stars_html += '<i class="fa-solid fa-star" style="color:#fbbf24; font-size:11px; margin-right:2px;"></i>'
            else:
                stars_html += '<i class="fa-solid fa-star" style="color:#e2e8f0; font-size:11px; margin-right:2px;"></i>'
        
        return f"""
            <div class="kanban-card {score_cls}">
                <div style="display:flex; justify-content:flex-start; gap:2px;">
                    {stars_html}
                </div>
                <div class="kanban-card-title">{card['nome']}</div>
                <div class="kanban-card-footer">
                    <span>R$ {card['val']}</span>
                </div>
            </div>
        """
    except Exception as e:
        logger.error(f"Erro ao renderizar card: {str(e)}")
        return ""

def render_banner() -> str:
    """Renderiza o banner de boas-vindas"""
    return """
        <div class="helio-banner">
            <div style="display:flex; justify-content:space-between; align-items:center;">
                <div style="flex:1;">
                    <h2>Olá, Time de Growth! 👋</h2>
                    <p>
                        O <strong>Helio</strong> processou <strong>1.420 leads</strong> hoje. 
                        <span style="background:rgba(255,255,255,0.2); padding:2px 8px; border-radius:4px; font-weight:bold;">127 oportunidades "Nota 5"</span>.
                    </p>
                </div>
                <div style="font-size:48px; opacity:0.2; margin-left:20px;">
                    <i class="fa-solid fa-balloon"></i>
                </div>
            </div>
        </div>
    """

# ==============================================================================
# 5. TELA DE LOGIN
# ==============================================================================
if not st.session_state["senha_ok"]:
    st.markdown("""
        <div style="display:flex; align-items:center; justify-content:center; min-height:100vh; background:linear-gradient(135deg, var(--red-50), var(--helio-light));">
            <div style="background:white; border-radius:1rem; padding:2rem; box-shadow:0 20px 25px -5px rgba(0,0,0,0.1); width:100%; max-width:400px; text-align:center;">
                <div style="width:80px; height:80px; background:white; border-radius:1.25rem; display:flex; align-items:center; justify-content:center; margin:0 auto 1rem; box-shadow:0 10px 25px -5px rgba(220,38,38,0.3);" class="balloon-float">
                    <i class="fa-solid fa-balloon" style="color:#dc2626; font-size:2.5rem;"></i>
                </div>
                <h1 style="color:#1e293b; font-size:24px; margin:0 0 0.5rem 0; font-weight:bold;">Helio Growth IA</h1>
                <p style="color:#64748b; margin:0 0 1.5rem 0;">Plataforma de Growth para Red Balloon</p>
            </div>
        </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        senha = st.text_input("🔐 Chave de Acesso", type="password", label_visibility="collapsed")
        
        if st.button("ACESSAR SISTEMA", use_container_width=True, key="login_btn"):
            if senha == "redballoon":
                st.session_state["senha_ok"] = True
                st.rerun()
            elif senha:  # Só mostra erro se digitou algo
                st.error("❌ Chave de acesso incorreta. Tente novamente.")
    
    st.stop()

# ==============================================================================
# 6. SIDEBAR NAVIGATION
# ==============================================================================
with st.sidebar:
    # Logo e título
    st.markdown("""
        <div style="display:flex; align-items:center; gap:12px; padding:10px 0 20px 0; border-bottom:1px solid #334155;">
            <div style="width:40px; height:40px; background:white; border-radius:10px; display:flex; align-items:center; justify-content:center;" class="balloon-float">
                <i class="fa-solid fa-balloon" style="color:#dc2626; font-size:20px;"></i>
            </div>
            <div>
                <div style="font-weight:bold; font-size:18px; color:white;">Helio</div>
                <div style="font-size:10px; color:#94a3b8; text-transform:uppercase; letter-spacing:1px;">Growth IA</div>
            </div>
        </div>
    """, unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Menu de navegação
    menu_options = ["Visão Geral (Rede)", "Pipeline de Matrículas", "Gestão de Unidades", "Processamento"]
    selected_menu = st.radio("MENU", menu_options, label_visibility="collapsed", key="menu_radio")
    st.session_state["menu_atual"] = selected_menu
    
    st.markdown("---")
    
    # Botão simulador
    col_sim1, col_sim2 = st.columns([1, 1])
    with col_sim1:
        if st.button("🧪 Simulador", use_container_width=True, key="simulador_btn"):
            st.session_state["show_simulador"] = not st.session_state.get("show_simulador", False)
            st.rerun()
    
    st.markdown("<br><br><br><br><br><br>")
    
    # Perfil do usuário (fixado na parte inferior)
    st.markdown("""
        <div style="border-top:1px solid #334155; padding-top:15px;">
            <div style="display:flex; align-items:center; gap:10px;">
                <img src="https://i.pravatar.cc/150?u=redballoon_admin" style="width:35px; height:35px; border-radius:50%; border:2px solid #dc2626;">
                <div>
                    <div style="font-size:13px; font-weight:500; color:white;">Gestor Growth</div>
                    <div style="font-size:11px; color:#94a3b8;">Matriz SP</div>
                </div>
            </div>
        </div>
    """, unsafe_allow_html=True)

# ==============================================================================
# 7. SIMULADOR MODAL
# ==============================================================================
if st.session_state.get("show_simulador"):
    with st.expander("🧪 Simulador de Probabilidade (Helio)", expanded=True):
        st.info("A Growth IA analisa a origem da campanha e interações do CRM para calcular a probabilidade de matrícula.")
        
        sc1, sc2, sc3 = st.columns([2, 1, 1])
        
        with sc1:
            origem = st.selectbox("Origem do Lead", ["Indicação", "Google Ads", "Social Ads", "Lista Fria"])
        
        with sc2:
            ativ = st.number_input("Interações CRM", min_value=0, value=1, step=1)
        
        with sc3:
            st.markdown("<br>", unsafe_allow_html=True)
            calc = st.button("Calcular", use_container_width=True, key="calc_score")
        
        if calc:
            try:
                # Lógica de scoring
                score = 1
                if origem in ["Indicação", "Google Ads"]:
                    score += 2
                if ativ >= 2:
                    score += 1
                if ativ >= 5:
                    score += 1
                
                # Garante que o score fica entre 1 e 5
                score = min(max(score, 1), 5)
                
                # Probabilidade associada ao score
                prob_map = {1: 15, 2: 35, 3: 55, 4: 78, 5: 92}
                prob = prob_map.get(score, 50)
                
                # Define cor baseado no score
                if score >= 4:
                    bg_color = "#10b981"
                    icon = "✅"
                elif score >= 3:
                    bg_color = "#f59e0b"
                    icon = "⚠️"
                else:
                    bg_color = "#ef4444"
                    icon = "❌"
                
                st.markdown(f"""
                    <div style="background:{bg_color}; color:white; padding:20px; border-radius:10px; text-align:center; margin-top:10px; box-shadow:0 4px 6px -1px rgba(0,0,0,0.1);">
                        <h3 style="margin:0; color:white; font-size:18px;">{icon} Score: {score}/5</h3>
                        <p style="margin:0.5rem 0 0 0; font-weight:bold; opacity:0.95; font-size:16px;">{prob}% Probabilidade de Matrícula</p>
                    </div>
                """, unsafe_allow_html=True)
                
            except Exception as e:
                logger.error(f"Erro no cálculo do score: {str(e)}")
                st.error("Erro ao calcular o score. Tente novamente.")

# ==============================================================================
# 8. CONTEÚDO PRINCIPAL POR MENU
# ==============================================================================

menu = st.session_state["menu_atual"]

# --- TELA 1: DASHBOARD ---
if menu == "Visão Geral (Rede)":
    # Header
    h1, h2 = st.columns([3, 1])
    h1.title("📊 Visão de Growth")
    h2.selectbox(
        "Filtro Unidade:",
        ["Rede Completa", "Próprias", "Franquias", "SP - Vila Leopoldina"],
        key="filter_unit"
    )
    
    # Banner
    st.markdown(render_banner(), unsafe_allow_html=True)
    
    # KPIs
    k1, k2, k3, k4 = st.columns(4)
    
    try:
        with k1:
            st.metric("📈 Matrículas (Mês)", "284", "+12%", delta_color="normal")
        
        with k2:
            st.metric("⭐ Qualidade (Helio)", "4.2", "Alta", delta_color="normal")
        
        with k3:
            st.metric("🎯 Pipeline Aberto", "R$ 4.2M", "1.4k Leads")
        
        with k4:
            st.metric("💹 Conversão", "18.4%", "-1.2%", delta_color="inverse")
    
    except Exception as e:
        logger.error(f"Erro ao renderizar KPIs: {str(e)}")
        st.error("Erro ao carregar os KPIs")
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Gráficos
    try:
        g1, g2 = st.columns([2, 1])
        
        with g1:
            st.markdown("##### 📊 Matrículas por Regional")
            data_reg = get_regional_data()
            
            if not data_reg.empty:
                fig_bar = go.Figure(data=[
                    go.Bar(
                        name='Realizado',
                        x=data_reg['Regional'],
                        y=data_reg['Realizado'],
                        marker_color='#ef4444',
                        marker_line_color='rgba(0,0,0,0)',
                        hovertemplate='<b>%{x}</b><br>Realizado: %{y}<extra></extra>'
                    ),
                    go.Bar(
                        name='Meta',
                        x=data_reg['Regional'],
                        y=data_reg['Meta'],
                        marker_color='#e2e8f0',
                        marker_line_color='rgba(0,0,0,0)',
                        hovertemplate='<b>%{x}</b><br>Meta: %{y}<extra></extra>'
                    )
                ])
                
                fig_bar.update_layout(
                    barmode='group',
                    plot_bgcolor='white',
                    paper_bgcolor='white',
                    height=300,
                    margin=dict(l=20, r=20, t=20, b=20),
                    showlegend=True,
                    hovermode='x unified',
                    font=dict(family="Arial, sans-serif", size=12),
                    xaxis=dict(showgrid=False),
                    yaxis=dict(showgrid=True, gridwidth=1, gridcolor='#e2e8f0')
                )
                
                st.plotly_chart(fig_bar, use_container_width=True, key="chart_regional")
        
        with g2:
            st.markdown("##### 🎯 Distribuição de Leads")
            data_pie = pd.DataFrame({
                'Label': ['Nota 5 (Hot)', 'Nota 4 (Warm)', 'Frios'],
                'Value': [15, 35, 50]
            })
            
            fig_pie = px.pie(
                data_pie,
                values='Value',
                names='Label',
                hole=0.7,
                color_discrete_sequence=['#ef4444', '#fbbf24', '#cbd5e1']
            )
            
            fig_pie.update_layout(
                showlegend=False,
                height=300,
                margin=dict(l=20, r=20, t=20, b=20),
                paper_bgcolor='white',
                font=dict(family="Arial, sans-serif", size=12)
            )
            
            fig_pie.add_annotation(
                text="<b>1.4k</b><br>Leads",
                x=0.5,
                y=0.5,
                showarrow=False,
                font_size=16,
                font_color='#1e293b'
            )
            
            st.plotly_chart(fig_pie, use_container_width=True, key="chart_pie")
    
    except Exception as e:
        logger.error(f"Erro ao renderizar gráficos: {str(e)}")
        st.error("Erro ao carregar os gráficos")

# --- TELA 2: PIPELINE (KANBAN) ---
elif menu == "Pipeline de Matrículas":
    st.markdown("### 🎯 Pipeline de Matrículas")
    st.caption("Visualização em tempo real do funil de vendas")
    
    try:
        col_names = [("Interesse", "#94a3b8"), ("Visita", "#fbbf24"), ("Negociação", "#f97316"), ("Matriculado", "#10b981")]
        cols = st.columns(len(col_names))
        
        cards = get_kanban_data()
        
        if not cards:
            st.warning("⚠️ Nenhum dado disponível para o Kanban")
        else:
            for idx, (name, color) in enumerate(col_names):
                with cols[idx]:
                    # Header da coluna
                    col_cards = [c for c in cards if c['status'] == idx]
                    
                    st.markdown(f"""
                        <div class="kanban-col-header" style="border-top: 3px solid {color};">
                            <span>{name}</span>
                            <span style="background:white; padding:2px 6px; border-radius:4px; font-size:11px; color:#1e293b;">{len(col_cards)}</span>
                        </div>
                        <div style="background:#f1f5f9; padding:10px; min-height:450px; border-radius:0 0 8px 8px;">
                    """, unsafe_allow_html=True)
                    
                    # Cards
                    for card in col_cards:
                        card_html = render_kanban_card(card)
                        st.markdown(card_html, unsafe_allow_html=True)
                    
                    st.markdown("</div>", unsafe_allow_html=True)
    
    except Exception as e:
        logger.error(f"Erro ao renderizar kanban: {str(e)}")
        st.error("Erro ao carregar o pipeline")

# --- TELA 3: GESTÃO DE UNIDADES ---
elif menu == "Gestão de Unidades":
    st.markdown("### 🏢 Ranking de Performance (YTD)")
    
    try:
        df_units = get_units_data()
        
        if df_units.empty:
            st.warning("⚠️ Nenhum dado disponível")
        else:
            st.dataframe(
                df_units.sort_values('Score Médio', ascending=False),
                column_config={
                    "Score Médio": st.column_config.ProgressColumn(
                        "Helio Score",
                        min_value=0,
                        max_value=5,
                        format="%.1f"
                    ),
                    "Conversão": st.column_config.NumberColumn(
                        "Taxa Matrícula",
                        format="%.0%"
                    ),
                    "Receita": st.column_config.NumberColumn(
                        "Resultado (R$)",
                        format="R$ %.0f"
                    )
                },
                use_container_width=True,
                hide_index=True,
                key="table_units"
            )
            
            # Gráfico de performance
            st.markdown("---")
            st.markdown("### 📈 Evolução Mensal")
            
            fig_perf = px.bar(
                df_units.sort_values('Receita', ascending=True),
                x='Receita',
                y='Unidade',
                orientation='h',
                color='Score Médio',
                color_continuous_scale=['#ef4444', '#f59e0b', '#10b981'],
                title='Receita por Unidade (Colorido por Score)'
            )
            
            fig_perf.update_layout(
                plot_bgcolor='white',
                paper_bgcolor='white',
                height=350,
                margin=dict(l=20, r=20, t=40, b=20),
                showlegend=True,
                font=dict(family="Arial, sans-serif", size=12)
            )
            
            st.plotly_chart(fig_perf, use_container_width=True, key="chart_perf")
    
    except Exception as e:
        logger.error(f"Erro ao renderizar gestão de unidades: {str(e)}")
        st.error("Erro ao carregar dados das unidades")

# --- TELA 4: PROCESSAMENTO ---
elif menu == "Processamento":
    st.title("⚙️ Scripts de Processamento")
    st.info("🔧 Gerencie e execute os scripts Python do backend aqui.")
    
    st.markdown("---")
    
    col1, col2 = st.columns(2)
    
    try:
        with col1:
            st.markdown("### 🤖 Lead Scoring")
            st.caption("Atualiza scores de todos os leads em tempo real")
            
            if st.button("▶️ Rodar Processamento", use_container_width=True, key="btn_scoring"):
                with st.spinner("⏳ Processando 1.420 leads..."):
                    # Simula processamento
                    import time
                    time.sleep(2)
                
                st.success("✅ Scoring atualizado para 1.420 leads.")
                st.metric("Leads Processados", "1.420", "+45 leads")
        
        with col2:
            st.markdown("### 📄 Relatório C-Level")
            st.caption("Gera e envia relatório executivo por e-mail")
            
            if st.button("▶️ Gerar Relatório", use_container_width=True, key="btn_relatorio"):
                with st.spinner("📧 Gerando PDF..."):
                    # Simula geração
                    import time
                    time.sleep(1.5)
                
                st.success("✅ Relatório enviado para redballoon@email.com")
                st.markdown("""
                    **Relatório incluiu:**
                    - 📊 Dashboard de Matrículas
                    - 🎯 Análise de Pipeline
                    - 💰 Projeção de Receita
                    - 📈 Ranking de Unidades
                """)
    
    except Exception as e:
        logger.error(f"Erro na tela de processamento: {str(e)}")
        st.error("Erro ao executar operação")

# ==============================================================================
# 9. RODAPÉ
# ==============================================================================
st.markdown("---")
st.markdown("""
    <div style="text-align:center; padding:20px; color:#64748b; font-size:12px;">
        <p>
            🎈 <strong>Helio Growth IA</strong> v2.0 | Plataforma Integrada de Growth<br>
            Desenvolvido para <strong>Red Balloon</strong> | Última atualização: """ + datetime.now().strftime("%d/%m/%Y às %H:%M") + """
        </p>
    </div>
""", unsafe_allow_html=True)