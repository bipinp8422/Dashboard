import io
import tempfile
import os
import streamlit as st
from make_dashboard_enhanced import generate_html

st.set_page_config(
    page_title="Denave x Canon CPP Dashboard Generator",
    page_icon="📊",
    layout="centered",
    initial_sidebar_state="collapsed"
)

st.title("📊 Denave x Canon CPP -- Dashboard Generator")
st.write(
    "Upload the **Target vs Achievement Report** workbook (.xlsm/.xlsx) and click "
    "**Generate** to build a downloadable HTML performance dashboard with pivot tables and advanced analytics."
)

uploaded = st.file_uploader("Upload workbook", type=["xlsm", "xlsx"])

if uploaded is not None:
    st.success(f"✅ Loaded: {uploaded.name} ({uploaded.size/1_000_000:.1f} MB)")
    if st.button("🚀 Generate Dashboard", type="primary", use_container_width=True):
        with st.spinner("📈 Crunching numbers and building your dashboard..."):
            with tempfile.TemporaryDirectory() as tmp:
                in_path = os.path.join(tmp, uploaded.name)
                with open(in_path, "wb") as f:
                    f.write(uploaded.getbuffer())
                out_path = os.path.join(tmp, "dashboard.html")
                try:
                    generate_html(in_path, out_path)
                except Exception as e:
                    st.error(f"❌ Failed to generate dashboard: {e}")
                    st.stop()
                with open(out_path, "rb") as f:
                    html_bytes = f.read()

        st.success("✨ Dashboard generated successfully!")
        out_name = os.path.splitext(uploaded.name)[0] + "-dashboard.html"
        st.download_button(
            "⬇️ Download Dashboard HTML",
            data=html_bytes,
            file_name=out_name,
            mime="text/html",
            use_container_width=True
        )
        st.divider()
        st.subheader("📋 Dashboard Preview")
        st.components.v1.html(html_bytes.decode("utf-8"), height=1200, scrolling=True)
else:
    st.info("⏳ Waiting for a file upload...")
