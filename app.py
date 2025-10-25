# ------------------------------------------------------------
# app.py  –  Advansor Project Preparation Tool (main interface)
# ------------------------------------------------------------
import streamlit as st
import importlib
import pandas as pd

# --- SAFE IMPORTS ---
stage1_ok = True
stage2_ok = True
bom_ok = True
stage1_err = ""
stage2_err = ""
bom_err = ""

try:
    stage1 = importlib.import_module("stage1_to_eplan")
except Exception as e:
    stage1_ok = False
    stage1_err = str(e)

try:
    stage2 = importlib.import_module("stage2_komax")
except Exception as e:
    stage2_ok = False
    stage2_err = str(e)

try:
    stage3 = importlib.import_module("stage3_bom")
except Exception as e:
    bom_ok = False
    bom_err = str(e)

# --- PAGE CONFIG ---
st.set_page_config(
    page_title="Advansor Project Preparation Tool",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# --- CUSTOM CSS ---
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');

.stApp {
  background-color: #081d19;
  background-image:
    radial-gradient(circle at 10% 20%, rgba(0, 255, 204, 0.03) 0%, transparent 80%),
    radial-gradient(circle at 90% 80%, rgba(0, 255, 204, 0.02) 0%, transparent 80%);
  background-repeat: no-repeat;
  background-size: cover;
  background-position: center center;
  font-family: 'Inter', sans-serif;
}

/* Floating logo */
.elcor-block {
  position: absolute;
  top: 18px;
  left: 30px;
  display: flex;
  flex-direction: column;
  align-items: flex-start;
}

#elcor-logo {
  font-size: 4.5rem;
  font-weight: 700;
  color: #00d4aa;
  text-shadow: 0 0 14px rgba(0, 212, 170, 0.6);
  margin-bottom: 10px;
  animation: glow 3s ease-in-out infinite alternate;
}

@keyframes glow {
  from { text-shadow: 0 0 10px #00ffcc; }
  to   { text-shadow: 0 0 30px #00b4d8; }
}

.stButton > button {
  background: linear-gradient(135deg, #064e3b 0%, #047857 100%);
  color: white;
  border: none;
  border-radius: 12px;
  padding: 0.8rem 2rem;
  font-size: 1.1rem;
  font-weight: 700;
  transition: all 0.3s ease;
}
.stButton > button:hover {
  transform: translateY(-2px);
  box-shadow: 0 0 12px rgba(0, 212, 170, 0.4);
}
.stButton > button:disabled {
  opacity: 0.5 !important;
  cursor: not-allowed !important;
}

#MainMenu, footer, header {visibility: hidden;}
</style>

<div class="elcor-block">
  <div id="elcor-logo">elcor.</div>
  <img id="advansor-logo" src="https://raw.githubusercontent.com/Vilius-se/ADV_management/main/logo_Advansor.png" width="160">
</div>
""", unsafe_allow_html=True)

# --- HEADER ---
st.markdown('<h1 style="text-align:center; color:#00ffcc;">Project Preparation Tool</h1>', unsafe_allow_html=True)
st.markdown('<p style="text-align:center; color:#b6cfc8;">Intelligent Excel Processing • Sustainable Data Solutions • The Future is Electric</p>', unsafe_allow_html=True)
st.markdown("<hr>", unsafe_allow_html=True)

# --- SESSION STATE ---
if "stage" not in st.session_state:
    st.session_state.stage = None

# --- MAIN BUTTONS ---
col1, col2, col3 = st.columns([3,2,3])
with col2:
    if stage1_ok:
        if st.button("🚀 Convert for EPLAN", use_container_width=True):
            st.session_state.stage = "eplan"
    else:
        st.button("❌ EPLAN module error", disabled=True, use_container_width=True)

    st.write("")
    if stage2_ok:
        if st.button("🔧 Convert for KOMAX", use_container_width=True):
            st.session_state.stage = "komax"
    else:
        st.button("❌ KOMAX module error", disabled=True, use_container_width=True)

    st.write("")
    if bom_ok:
        if st.button("📦 BOM Generator", use_container_width=True):
            st.session_state.stage = "bom"
    else:
        st.button("❌ BOM module error", disabled=True, use_container_width=True)

st.markdown("---")

# --- ROUTING ---
if st.session_state.stage == "eplan" and stage1_ok:
    stage1.render()
elif st.session_state.stage == "komax" and stage2_ok:
    stage2.render()
elif st.session_state.stage == "bom" and bom_ok:
    stage3.render()

# --- FOOTER ---
st.markdown("---")
st.markdown("""
<div style="text-align:center; padding:1rem; color:#7fa59a;">
  🌱 Sustainable Data Solutions • ⚡ The Future is Electric
</div>
""", unsafe_allow_html=True)
