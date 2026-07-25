import io
import tempfile
import os
import streamlit as st
from make_dashboard import generate_html

st.set_page_config(page_title="Denave x Canon CPP Dashboard Generator", page_icon="\U0001F4CA", layout="centered")

st.title("\U0001F4CA Denave x Canon CPP -- Dashboard Generator")
st.write(
    "Upload the **Target vs Achievement Report** workbook (.xlsm/.xlsx) and click "
    "**Generate** to build a downloadable HTML performance dashboard."
)

uploaded = st.file_uploader("Upload workbook", type=["xlsm", "xlsx"])

if uploaded is not None:
    st.success(f"Loaded: {uploaded.name} ({uploaded.size/1_000_000:.1f} MB)")
    if st.button("Generate Dashboard", type="primary"):
        with st.spinner("Crunching numbers and building your dashboard..."):
            with tempfile.TemporaryDirectory() as tmp:
                in_path = os.path.join(tmp, uploaded.name)
                with open(in_path, "wb") as f:
                    f.write(uploaded.getbuffer())
                out_path = os.path.join(tmp, "dashboard.html")
                try:
                    generate_html(in_path, out_path)
                except Exception as e:
                    st.error(f"Failed to generate dashboard: {e}")
                    st.stop()
                with open(out_path, "rb") as f:
                    html_bytes = f.read()

        st.success("Dashboard generated!")
        out_name = os.path.splitext(uploaded.name)[0] + "-dashboard.html"
        st.download_button(
            "\u2b07\ufe0f Download Dashboard HTML",
            data=html_bytes,
            file_name=out_name,
            mime="text/html",
        )
        st.components.v1.html(html_bytes.decode("utf-8"), height=900, scrolling=True)
else:
    st.info("Waiting for a file upload...")
