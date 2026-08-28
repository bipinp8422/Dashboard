import streamlit as st
import os
import tempfile
from make_dashboard import generate_html

st.set_page_config(page_title="Denave CPP Dashboard Generator", layout="wide")

st.markdown("""
<style>
.main-title {
    font-size: 2.5em;
    font-weight: 700;
    margin-bottom: 10px;
}
.subtitle {
    font-size: 1.1em;
    color: #666;
    margin-bottom: 20px;
}
</style>
<div class="main-title">📊 Denave × Canon CPP Dashboard</div>
<div class="subtitle">Region-wise Performance Cockpit Generator</div>
""", unsafe_allow_html=True)

col1, col2 = st.columns(2)

with col1:
    st.info("📋 **How it works:**\n1. Upload your Excel file\n2. Generate separate dashboards\n3. Download North & South dashboards")

with col2:
    st.success("✨ **Features:**\n- Executive summary\n- KPI cards\n- Interactive charts\n- Rep rankings\n- Product details")

st.divider()

uploaded_file = st.file_uploader("Upload Excel (.xlsm/.xlsx)", type=["xlsm", "xlsx"])

if uploaded_file:
    with tempfile.TemporaryDirectory() as tmpdir:
        # Save uploaded file
        input_path = os.path.join(tmpdir, uploaded_file.name)
        with open(input_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        # Generate dashboards
        if st.button("🚀 Generate Dashboards", use_container_width=True):
            st.info("Generating region-wise dashboards...")
            
            try:
                results = generate_html(input_path, tmpdir)
                
                st.success("✅ Dashboards generated successfully!")
                st.divider()
                
                col1, col2 = st.columns(2)
                
                for region, path_or_error, success in results:
                    if success:
                        with open(path_or_error, "rb") as f:
                            if region == "North":
                                with col1:
                                    st.markdown(f"### 🔵 {region} Region")
                                    st.download_button(
                                        label=f"📥 Download {region} Dashboard",
                                        data=f.read(),
                                        file_name=f"dashboard_{region}.html",
                                        mime="text/html",
                                        use_container_width=True
                                    )
                            else:
                                with col2:
                                    st.markdown(f"### 🟡 {region} Region")
                                    st.download_button(
                                        label=f"📥 Download {region} Dashboard",
                                        data=f.read(),
                                        file_name=f"dashboard_{region}.html",
                                        mime="text/html",
                                        use_container_width=True
                                    )
                        
                        with st.expander(f"ℹ️ {region} Details"):
                            st.caption(f"✓ File: {path_or_error}")
                    else:
                        st.error(f"✗ {region}: {path_or_error}")
                
            except Exception as e:
                st.error(f"❌ Error generating dashboards: {str(e)}")

st.divider()
st.markdown("""
### 📌 Requirements
- **Excel format**: .xlsm or .xlsx
- **Sheets needed**: 
  - Target vs Achievement
  - Raw Data
  - Product Description (optional)
""")

st.caption("💡 Tip: Keep the Excel file open to make updates, then re-upload to refresh dashboards")
