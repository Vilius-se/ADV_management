import streamlit as st
import pandas as pd
import re

def validate_excel(uploaded_file, required_columns):
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"⚠️ Cannot open file: {e}")
        return None
    missing = [col for col in required_columns if col not in df.columns]
    if missing:
        st.error(f"⚠️ Invalid file structure. Missing columns: {missing}")
        return None
    return df

def render():
    st.header("Stage 3: BOM Management")

    # --- USER INPUT ---
    st.subheader("🔢 Project Information")

    project_number = st.text_input("Project number (format: 1234-567)")
    if project_number and not re.match(r"^\d{4}-\d{3}$", project_number):
        st.error("⚠️ Project number must be in format: 1234-567")

    panel_type = st.selectbox("Panel type", ["Type A", "Type B", "Type C"])
    grounding_type = st.selectbox("Grounding type", ["Type 1", "Type 2", "Type 3"])
    main_switch = st.selectbox("Main switch", ["Switch A", "Switch B", "Switch C"])

    swing_frame = st.checkbox("Swing frame?")
    ups = st.checkbox("UPS?")
    rittal = st.checkbox("Rittal?")

    # --- FILE UPLOADS ---
    st.subheader("📂 Upload Required Files")

    dfs = {}

    if not rittal:
        cubic_bom = st.file_uploader("Insert CUBIC BOM (Excel)", type=["xls", "xlsx"])
        if cubic_bom:
            dfs["cubic_bom"] = validate_excel(cubic_bom, ["Item No."])

    bom = st.file_uploader("Insert BOM (Excel)", type=["xls", "xlsx"])
    if bom:
        dfs["bom"] = validate_excel(bom, ["Part No."])

    data_file = st.file_uploader("Insert DATA (Excel)", type=["xls", "xlsx"])
    if data_file:
        dfs["data"] = validate_excel(data_file, ["Line-Function"])

    ks_file = st.file_uploader("Insert Kaunas Stock (Excel)", type=["xls", "xlsx"])
    if ks_file:
        dfs["ks"] = validate_excel(ks_file, ["Manufacturer", "Qty"])

    # --- SAVE TO SESSION ---
    if project_number and re.match(r"^\d{4}-\d{3}$", project_number) and all(v is not None for v in dfs.values()):
        st.success("✅ All inputs and files are valid!")
        st.session_state["stage3"] = {
            "project_number": project_number,
            "panel_type": panel_type,
            "grounding_type": grounding_type,
            "main_switch": main_switch,
            "swing_frame": swing_frame,
            "ups": ups,
            "rittal": rittal,
            "files": dfs,
        }
    else:
        st.info("ℹ️ Please complete all fields and upload valid files to continue.")
