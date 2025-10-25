import streamlit as st
import pandas as pd
from io import BytesIO
from processing import stage2_pipeline_1, stage2_pipeline_2, stage2_pipeline_4

def render():
    st.header("Stage 2: Convert for KOMAX")
    uploaded_csv = st.file_uploader("📁 Upload KOMAX CSV", type=["csv"], key="komax_csv")

    if uploaded_csv:
        try:
            df_stage2 = stage2_pipeline_1(uploaded_csv)
            df_stage2 = stage2_pipeline_2(df_stage2)
            df_stage2 = stage2_pipeline_4(df_stage2)
        except Exception as e:
            st.error(f"❌ Error processing: {e}")
            st.stop()

        st.success("✅ KOMAX CSV processed successfully!")
        st.dataframe(df_stage2.head(10), use_container_width=True)

        buf = BytesIO()
        df_stage2.to_csv(buf, index=False)
        base = uploaded_csv.name[:8]
        st.download_button(
            "📥 Download KOMAX Output",
            buf.getvalue(),
            file_name=f"{base}_ADV_DLW_IMPORT.csv",
            mime="text/csv"
        )
    else:
        st.info("Upload a KOMAX CSV file to continue.")
