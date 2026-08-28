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
    st.info("📋 **How it works:**\n1. Upload Excel file\n2. Click Generate\n3. Download dashboards")

with col2:
    st.success("✨ **Features:**\n- Executive summary\n- BM filtering\n- Interactive charts\n- Product details\n- Mobile ready")

st.divider()

uploaded_file = st.file_uploader("Upload Excel (.xlsm/.xlsx)", type=["xlsm", "xlsx"])

if uploaded_file:
    st.success(f"✅ File ready: {uploaded_file.name}")
    
    if st.button("🚀 Generate Dashboards", use_container_width=True):
        try:
            # Create temporary directory for processing
            temp_dir = tempfile.mkdtemp()
            st.info("🔄 Processing your file...")
            
            # Save uploaded file temporarily
            input_path = os.path.join(temp_dir, uploaded_file.name)
            with open(input_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            # Generate dashboards
            st.info("📊 Generating dashboards...")
            results = generate_html(input_path, temp_dir)
            
            # Check if generation was successful
            successful = [r for r in results if r[2]]
            
            if len(successful) == 2:
                st.success("✅ Dashboards generated successfully!")
                st.divider()
                
                col1, col2 = st.columns(2)
                
                for region, file_path, success in results:
                    try:
                        if success and os.path.isfile(file_path):
                            with open(file_path, "rb") as f:
                                file_data = f.read()
                            
                            if region == "North":
                                with col1:
                                    st.markdown("### 🔵 North Region Dashboard")
                                    st.download_button(
                                        label="📥 Download North Dashboard",
                                        data=file_data,
                                        file_name="dashboard_North.html",
                                        mime="text/html",
                                        use_container_width=True
                                    )
                                    st.caption(f"Size: {len(file_data)/1024:.1f} KB")
                            
                            elif region == "South":
                                with col2:
                                    st.markdown("### 🟡 South Region Dashboard")
                                    st.download_button(
                                        label="📥 Download South Dashboard",
                                        data=file_data,
                                        file_name="dashboard_South.html",
                                        mime="text/html",
                                        use_container_width=True
                                    )
                                    st.caption(f"Size: {len(file_data)/1024:.1f} KB")
                    except Exception as e:
                        st.error(f"❌ Error reading {region} dashboard: {str(e)}")
            else:
                st.error("❌ Failed to generate all dashboards")
                for region, error_msg, success in results:
                    if not success:
                        st.error(f"**{region}**: {error_msg}")
        
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
            st.error("Please check your Excel file format and try again")

st.divider()

st.markdown("""
### 📋 Excel File Requirements

**Sheets needed:**
- `Target vs Achievement` — Main sales data
- `Raw Data` — Transaction details  
- `Product Description` — Optional

**Key columns:**
- Region (North/South)
- BM (Business Manager)
- Name (Rep name)
- Revenue Target
- Revenue Achived
- Units Sold
- Achievement in %

### ✨ Dashboard Features
- 5 tabs per region (Executive Summary, Overview, Sales Rep, Products, Analysis)
- BM filtering on Overview & Sales Rep tabs
- Interactive charts
- Search functionality
- Mobile responsive
- Works offline
""")

st.caption("💡 Tip: After download, open HTML files in any web browser — no internet needed!")
