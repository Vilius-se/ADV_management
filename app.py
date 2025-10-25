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

# --- CUSTOM CSS + FLOATING LOGOS ---
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');

.stApp {
  background-color: #081d19;
  background-image:
    radial-gradient(circle at 10% 20%, rgba(0, 255, 204, 0.03) 0%, transparent 80%),
    radial-gradient(circle at 90% 80%, rgba(0, 255, 204, 0.02) 0%, transparent 80%),
    url("data:image/svg+xml;utf8,\
    <svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 1600 900'>\
      <defs>\
        <linearGradient id='wireGrad' x1='0%' y1='0%' x2='100%' y2='0%'>\
          <stop offset='0%' stop-color='%23007360'/>\
          <stop offset='100%' stop-color='%23006284'/>\
        </linearGradient>\
        <filter id='softGlow'>\
          <feGaussianBlur stdDeviation='1.4' result='blur'/>\
          <feMerge><feMergeNode in='blur'/><feMergeNode in='SourceGraphic'/></feMerge>\
        </filter>\
      </defs>\
      <g filter='url(%23softGlow)' stroke='url(%23wireGrad)' stroke-width='1' fill='none' opacity='0.08'>\
        <path d='M0 850 Q400 600 800 850 T1600 850'/>\
        <path d='M0 700 Q400 500 800 700 T1600 700'/>\
        <path d='M0 550 Q400 400 800 550 T1600 550'/>\
        <path d='M0 400 Q400 300 800 400 T1600 400'/>\
        <path d='M0 250 Q400 200 800 250 T1600 250'/>\
        <path d='M0 100 Q400 150 800 100 T1600 100'/>\
        <path d='M200 0 Q600 300 1000 600 T1600 800'/>\
        <path d='M0 0 Q300 200 600 500 T1200 900'/>\
      </g>\
      <g fill='%2300ffcc' opacity='0.04'>\
        <circle cx='200' cy='750' r='2'/>\
        <circle cx='450' cy='580' r='2'/>\
        <circle cx='650' cy='420' r='2'/>\
        <circle cx='950' cy='250' r='2'/>\
        <circle cx='1200' cy='450' r='2'/>\
        <circle cx='1400' cy='650' r='2'/>\
        <circle cx='1550' cy='300' r='2'/>\
        <circle cx='800' cy='800' r='1.5'/>\
        <circle cx='1000' cy='100' r='1.5'/>\
        <circle cx='300' cy='200' r='1.5'/>\
      </g>\
    </svg>");
  background-repeat: no-repeat;
  background-size: cover;
  background-position: center center;
  font-family: 'Inter', sans-serif;
}

/* FLOATING LOGO BLOCK */
.elcor-block {
  position: absolute;
  top: 18px;
  left: 30px;
  display: flex;
  flex-direction: column;
  align-items: flex-start;
}

/* ELCOR text */
#elcor-logo {
  font-family: 'Inter', sans-serif;
  font-size: 4.6rem;
  font-weight: 700;
  color: #00d4aa;
  animation: floatUpDown 5s ease-in-out infinite, elcorGlow 3s ease-in-out infinite;
  text-shadow: 0 0 14px rgba(0, 212, 170, 0.6);
  letter-spacing: -0.02em;
  margin-bottom: 10px;
}

/* ADVANSOR image */
#advansor-logo {
  width: 150px;
  animation: floatUpDown 5s ease-in-out infinite;
  filter: drop-shadow(0 0 10px rgba(0,255,204,0.3));
  position: relative;
}

/* Reflection under Advansor logo */
#advansor-logo::after {
  content: "";
  position: absolute;
  bottom: -70px;
  left: 0;
  right: 0;
  height: 60px;
  background: url("https://raw.githubusercontent.com/Vilius-se/Advansor-Tool/main/logo_Advansor.png") no-repeat center;
  background-size: contain;
  opacity: 0.25;
  transform: scaleY(-1);
  mask-image: linear-gradient(to bottom, rgba(255,255,255,0.8), transparent);
  -webkit-mask-image: linear-gradient(to bottom, rgba(255,255,255,0.8), transparent);
}

/* Animations */
@keyframes floatUpDown {
  0%   { transform: translateY(0px); }
  50%  { transform: translateY(-10px); }
  100% { transform: translateY(0px); }
}

@keyframes elcorGlow {
  0%   { text-shadow: 0 0 6px rgba(0,255,204,0.2); opacity: 0.9; }
  50%  { text-shadow: 0 0 26px rgba(0,255,204,1); opacity: 1; }
  100% { text-shadow: 0 0 6px rgba(0,255,204,0.2); opacity: 0.9; }
}

/* Main title */
.main-title {
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
}

/* Glowing line */
.electric-line {
  height: 3px;
  width: 65%;
  margin: 1rem auto 2.5rem auto;
  background: linear-gradient(90deg, transparent, #00d4aa, #00ffcc, #00d4aa, transparent);
  background-size: 300% 100%;
  animation: moveLine 4s ease-in-out infinite alternate;
  box-shadow: 0 0 20px rgba(0, 255, 204, 0.7), 0 4px 10px rgba(0, 255, 204, 0.2);
  border-radius: 3px;
  opacity: 0.95;
}

@keyframes moveLine {
  0% { background-position: 0% 50%; }
  50% { background-position: 100% 50%; }
  100% { background-position: 0% 50%; }
}

/* Subtitle */
.subtitle {
  text-align: center;
  color: #b6cfc8;
  font-size: 1.2rem;
  font-weight: 400;
  margin-bottom: 2.5rem;
}

/* Buttons */
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

/* Hide menu/footer */
#MainMenu, footer, header {visibility: hidden;}
</style>

<!-- FLOATING BLOCK -->
<div class="elcor-block">
  <div id="elcor-logo">elcor.</div>
  <img id="advansor-logo" src="https://raw.githubusercontent.com/Vilius-se/Advansor-Tool/main/logo_Advansor.png" alt="Advansor Logo">
</div>
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
