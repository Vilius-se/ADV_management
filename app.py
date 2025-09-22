import streamlit as st
import importlib

# --- BOM modulio validacija ---
bom_available = True
stage3_bom = None
try:
    stage3_bom = importlib.import_module("stage3_bom")
except Exception as e:
    bom_available = False

st.set_page_config(
    page_title="Advansor Wireset Helper",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# --- CUSTOM CSS ---
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
.stApp {background: linear-gradient(135deg, #0f1419 0%, #1 50%, #0f1419 100%);}
.main .block-container {padding-top: 2rem; padding-bottom: 2rem;}
.stMarkdown, p {color: #e2e8f0;}
.main-title {font-family: 'Inter', sans-serif; font-size: 3.5rem; font-weight: 700; text-align: center; margin-bottom: 0.5rem; background: linear-gradient(135deg, #00d4aa 0%, #00a693 30%, #0ea5e9 70%, #0284c7 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text; text-shadow: 0 4px 20px rgba(0, 212, 170, 0.3);}
.subtitle {font-family: 'Inter', sans-serif; text-align: center; color: #94a3b8; font-size: 1.3rem; font-weight: 400; margin-bottom: 3rem;}
.electric-line {height: 2px; background: linear-gradient(90deg, transparent 0%, #00d4aa 20%, #0ea5e9 50%, #00d4aa 80%, transparent 100%); margin: 1rem auto 2rem auto; width: 60%; box-shadow: 0 0 10px rgba(0, 212, 170, 0.5);}
.upload-container {border: 2px dashed #334155; border-radius: 16px; padding: 3rem 2rem; text-align: center; background: linear-gradient(135deg, rgba(15, 23, 42, 0.8) 0%, rgba(30, 41, 59, 0.6) 100%); margin: 2rem 0; backdrop-filter: blur(10px); transition: all 0.3s ease;}
.upload-container:hover {border-color: #00d4aa;}
.status-success {background: linear-gradient(135deg, #00d4aa 0%, #059669 100%); color: white; padding: 1rem; border-radius: 12px;}
.status-info {background: linear-gradient(135deg, #0ea5e9 0%, #0284c7 100%); color: white; padding: 1rem; border-radius: 12px;}
.status-warning {background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%); color: white; padding: 1rem; border-radius: 12px;}
.stMetric {background: linear-gradient(135deg, rgba(30,41,59,0.8) 0%, rgba(51,65,85,0.6) 100%); padding: 1rem; border-radius: 8px;}
.stButton > button {background: linear-gradient(135deg, #00d4aa 0%, #0ea5e9 100%); color: white; border-radius: 12px; padding: 0.75rem 2rem; font-weight: 600; font-family: 'Inter', sans-serif; transition: all 0.3s;}
.stButton > button:hover {transform: translateY(-2px);}
.success-message {color: #22c55e; font-weight: 600; font-size: 0.9rem;}
.blank-cell-highlight {background-color: #fef3c7 !important; border: 2px solid #f59e0b !important;}
#MainMenu, footer, header {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# --- HEADER ---
st.markdown('<h1 class="main-title">⚡ Advansor Wireset Helper</h1>', unsafe_allow_html=True)
st.markdown('<div class="electric-line"></div>', unsafe_allow_html=True)
st.markdown(
    '<p class="subtitle">Intelligent Excel Processing • Sustainable Data Solutions • The Future is Electric</p>',
    unsafe_allow_html=True
)

# --- NAVIGATION ---
if "stage" not in st.session_state:
    st.session_state.stage = None

st.markdown("<div style='text-align:center; margin-bottom:2rem;'>", unsafe_allow_html=True)

col_left, col_center, col_right = st.columns([4, 2, 4])
with col_center:
    if st.button("🚀 Convert for EPLAN", key="btn_eplan", use_container_width=True):
        st.session_state.stage = "eplan"
    st.write("")
    if st.button("🔧 Convert for KOMAX", key="btn_komax", use_container_width=True):
        st.session_state.stage = "komax"
    st.write("")
    if bom_available:
        if st.button("📦 BOM", key="btn_bom", use_container_width=True):
            st.session_state.stage = "bom"
    else:
        st.button("📦 BOM (Not avalable)", key="btn_bom_disabled", use_container_width=True, disabled=True)

st.markdown("---")

# --- ROUTER ---
if st.session_state.stage == "eplan":
    st.header("Stage 1: Convert for EPLAN")
    st.info("⚙️ Čia bus EPLAN logika (stage1 pipelines).")

elif st.session_state.stage == "komax":
    st.header("Stage 2: Convert for KOMAX")
    st.info("⚙️ Čia bus KOMAX logika (stage2 pipelines).")

elif st.session_state.stage == "bom" and bom_available:
    try:
        stage3_bom.render()
    except Exception as e:
        st.error("❌ BOM modulio klaida: " + str(e))
        st.session_state.stage = None

# --- FOOTER ---
st.markdown("---")
st.markdown("""
<div style="text-align:center; padding:1rem 0; color:#64748b;">
  🌱 Sustainable Data Solutions • ⚡ The Future is Electric
</div>
""", unsafe_allow_html=True)
