import streamlit as st
import importlib
import pandas as pd
import time
import math
from io import BytesIO

# --- SAFE IMPORTS (fuse checks) ---
processing_ok = True
bom_ok = True
processing_err = ""
bom_err = ""

try:
    import processing
except Exception as e:
    processing_ok = False
    processing_err = str(e)

stage3_bom = None
try:
    stage3_bom = importlib.import_module("stage3_bom")
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

# --- LOAD LOGO ---
logo_path = "A_vector_logo_showcases_an_elephant-like_creature_.png"

# --- CUSTOM CSS ---
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');

.stApp {{
  background-color: #081d19;
  background-image:
    radial-gradient(circle at 10% 20%, rgba(0, 255, 204, 0.03) 0%, transparent 80%),
    radial-gradient(circle at 90% 80%, rgba(0, 255, 204, 0.02) 0%, transparent 80%);
  background-repeat: no-repeat;
  background-size: cover;
  font-family: 'Inter', sans-serif;
}}

#elcor-logo {{
  position: absolute;
  top: 20px;
  left: 30px;
  width: 90px;
  height: 90px;
  animation: glow 3s ease-in-out infinite alternate;
  filter: drop-shadow(0 0 10px rgba(0,255,204,0.6));
}}

@keyframes glow {{
  0% {{ filter: drop-shadow(0 0 5px rgba(0,255,204,0.2)); }}
  100% {{ filter: drop-shadow(0 0 18px rgba(0,255,204,0.8)); }}
}}

.main-title {{
  font-family: 'Inter', sans-serif;
  font-size: 3.4rem;
  font-weight: 700;
  text-align: center;
  background: linear-gradient(135deg, #00d4aa 0%, #00b4d8 30%, #00a693 60%, #00ffcc 100%);
  -webkit-background-clip: text;
  -webkit-text-fill-color: transparent;
  letter-spacing: 0.04em;
  margin-top: 1.5rem;
  text-shadow: 0 0 20px rgba(0, 212, 170, 0.25);
}}
.electric-line {{
  height: 3px;
  width: 65%;
  margin: 1rem auto 2.5rem auto;
  background: linear-gradient(90deg, transparent, #00d4aa, #00ffcc, #00d4aa, transparent);
  background-size: 300% 100%;
  animation: moveLine 4s ease-in-out infinite alternate;
  box-shadow: 0 0 20px rgba(0, 255, 204, 0.7), 0 4px 10px rgba(0, 255, 204, 0.2);
  border-radius: 3px;
  opacity: 0.95;
}}
@keyframes moveLine {{
  0% {{ background-position: 0% 50%; }}
  50% {{ background-position: 100% 50%; }}
  100% {{ background-position: 0% 50%; }}
}}
.subtitle {{
  font-family: 'Inter', sans-serif;
  text-align: center;
  color: #b6cfc8;
  font-size: 1.2rem;
  font-weight: 400;
  margin-bottom: 2.5rem;
}}
.stButton > button {{
  background: linear-gradient(135deg, #064e3b 0%, #047857 100%);
  color: white;
  border: none;
  border-radius: 12px;
  padding: 0.8rem 2rem;
  font-size: 1.1rem;
  font-weight: 700;
  transition: all 0.3s ease;
}}
.stButton > button:hover {{
  transform: translateY(-2px);
  box-shadow: 0 0 12px rgba(0, 212, 170, 0.4);
}}
#MainMenu, footer, header {{visibility: hidden;}}
</style>

<!-- ELCOR LOGO -->
<img id="elcor-logo" src="data:image/png;base64,{st.image(logo_path, output_format='PNG').data}" />
""", unsafe_allow_html=True)

# --- HEADER ---
st.markdown('<h1 class="main-title">Advansor Project Preparation Tool</h1>', unsafe_allow_html=True)
st.markdown('<div class="electric-line"></div>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">Intelligent Excel Processing • Sustainable Data Solutions • The Future is Electric</p>', unsafe_allow_html=True)

# --- STAGE NAVIGATION ---
if "stage" not in st.session_state:
    st.session_state.stage = None

st.markdown("<div style='text-align:center; margin-bottom:2rem;'>", unsafe_allow_html=True)
col_left, col_center, col_right = st.columns([4, 2, 4])
with col_center:
    if processing_ok:
        if st.button("🚀 Convert for EPLAN", key="btn_eplan", use_container_width=True):
            st.session_state.stage = "eplan"
    else:
        st.button("🚫 EPLAN (module error)", key="btn_eplan_disabled", use_container_width=True, disabled=True)

    st.write("")

    if processing_ok:
        if st.button("🔧 Convert for KOMAX", key="btn_komax", use_container_width=True):
            st.session_state.stage = "komax"
    else:
        st.button("🚫 KOMAX (module error)", key="btn_komax_disabled", use_container_width=True, disabled=True)

    st.write("")

    if bom_ok:
        if st.button("📦 BOM", key="btn_bom", use_container_width=True):
            st.session_state.stage = "bom"
    else:
        st.button("🚫 BOM (module error)", key="btn_bom_disabled", use_container_width=True, disabled=True)

st.markdown("---")

# --- ROUTING ---
if st.session_state.stage == "eplan" and processing_ok:
    st.header("Stage 1: Convert for EPLAN")
    st.info("⚙️ EPLAN transformation logic goes here (stage1 pipelines).")

elif st.session_state.stage == "komax" and processing_ok:
    st.header("Stage 2: Convert for KOMAX")
    st.info("⚙️ KOMAX CSV transformation logic goes here (stage2 pipelines).")

elif st.session_state.stage == "bom" and bom_ok:
    try:
        stage3_bom.render()
    except Exception as e:
        st.error("❌ BOM module error: " + str(e))
        st.session_state.stage = None

# --- FOOTER ---
st.markdown("---")
st.markdown("""
<div style="text-align:center; padding:1rem 0; color:#7fa59a;">
  🌱 Sustainable Data Solutions • ⚡ The Future is Electric
</div>
""", unsafe_allow_html=True)
