import io
import os
import numpy as np
import pandas as pd
import streamlit as st

# -----------------------------------------------------------------------------
# PAGE CONFIGURATION
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Upper Verification Limit (UVL) Calculator",
    page_icon="🧪",
    layout="wide"
)

st.title("🧪 Upper Verification Limit (UVL) Calculator")
st.write(
    "Calculate Repeatability (UVLR) and Within-Laboratory (UVLWL) Upper Verification Limits "
    "from QC data using standard statistical lookup tables."
)

# -----------------------------------------------------------------------------
# 1. HARDCODED LOOKUP TABLES
# -----------------------------------------------------------------------------
# Table 6 data: maps p (rounded to 2 decimals) -> df_WL
table_6_data = {
    2.74: 5, 2.06: 6, 1.78: 7, 1.62: 8, 1.51: 9, 1.43: 10,
    1.37: 11, 1.32: 12, 1.28: 13, 1.24: 14, 1.21: 15, 1.19: 16,
    1.16: 17, 1.14: 18, 1.12: 19, 1.10: 20, 1.08: 21, 1.05: 22,
    1.03: 23, 1.00: 24
}

def lookup_df_wl(p_val):
    closest_p = min(table_6_data.keys(), key=lambda x: abs(x - p_val))
    return table_6_data[closest_p]

# Table 7 data: maps DF -> list of values for Number of Samples [1, 2, 3, 4, 5, 6]
table_7_data = {
    5:  [1.49, 1.60, 1.66, 1.71, 1.74, 1.76],
    6:  [1.45, 1.55, 1.61, 1.65, 1.67, 1.70],
    7:  [1.42, 1.51, 1.56, 1.60, 1.62, 1.65],
    8:  [1.39, 1.48, 1.53, 1.56, 1.58, 1.60],
    9:  [1.37, 1.45, 1.50, 1.53, 1.55, 1.57],
    10: [1.35, 1.43, 1.47, 1.50, 1.52, 1.54],
    11: [1.34, 1.41, 1.45, 1.48, 1.50, 1.52],
    12: [1.32, 1.39, 1.43, 1.46, 1.48, 1.49],
    13: [1.31, 1.38, 1.42, 1.44, 1.46, 1.47],
    14: [1.30, 1.37, 1.40, 1.42, 1.44, 1.46],
    15: [1.29, 1.35, 1.39, 1.41, 1.43, 1.44],
    16: [1.28, 1.34, 1.38, 1.40, 1.41, 1.43],
    17: [1.27, 1.33, 1.36, 1.39, 1.40, 1.41],
    18: [1.27, 1.32, 1.35, 1.37, 1.39, 1.40],
    19: [1.26, 1.31, 1.34, 1.36, 1.38, 1.39],
    20: [1.25, 1.31, 1.34, 1.36, 1.37, 1.38],
    21: [1.25, 1.30, 1.33, 1.35, 1.36, 1.37],
    22: [1.24, 1.29, 1.32, 1.34, 1.35, 1.36],
    23: [1.24, 1.29, 1.31, 1.33, 1.35, 1.36],
    24: [1.23, 1.28, 1.31, 1.32, 1.34, 1.35],
    25: [1.23, 1.28, 1.30, 1.32, 1.33, 1.34],
    26: [1.22, 1.27, 1.30, 1.31, 1.32, 1.34],
    27: [1.22, 1.26, 1.29, 1.31, 1.32, 1.33],
    28: [1.22, 1.26, 1.28, 1.30, 1.31, 1.32],
    29: [1.21, 1.26, 1.28, 1.30, 1.31, 1.32],
    30: [1.21, 1.25, 1.27, 1.29, 1.30, 1.31],
    31: [1.20, 1.25, 1.27, 1.29, 1.30, 1.31],
    32: [1.20, 1.24, 1.27, 1.28, 1.29, 1.30],
    33: [1.20, 1.24, 1.26, 1.28, 1.29, 1.30],
    34: [1.20, 1.24, 1.26, 1.27, 1.28, 1.29]
}

def lookup_table_7(df, num_samples):
    df_clamped = max(5, min(34, int(df)))
    ns_clamped = max(1, min(6, int(num_samples)))
    return table_7_data[df_clamped][ns_clamped - 1]

# -----------------------------------------------------------------------------
# 2. SIDEBAR INPUT CONTROLS
# -----------------------------------------------------------------------------
st.sidebar.header("⚙️ Analysis Parameters")

cvr_input = st.sidebar.number_input("Calculated CVr (%)", value=0.0, format="%.4f")
num_samples_input = st.sidebar.slider("Number of Samples", min_value=1, max_value=6, value=1)
cvr_mfg_input = st.sidebar.number_input("CVr Manufacturer Limit (%)", value=0.0, format="%.4f")
cvwl_mfg_input = st.sidebar.number_input("CVwl Manufacturer Limit (%)", value=0.0, format="%.4f")

# -----------------------------------------------------------------------------
# 3. FILE UPLOAD & PROCESSING
# -----------------------------------------------------------------------------
uploaded_file = st.file_uploader(
    "Upload Excel file containing QC data:",
    type=["xlsx", "xls"]
)

if uploaded_file is not None:
    file_name = uploaded_file.name
    
    try:
        df_excel = pd.read_excel(uploaded_file, engine="openpyxl")
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        st.stop()

    st.subheader(f"📄 Processing: `{file_name}`")
    with st.expander("👀 Preview Raw Excel Data"):
        st.dataframe(df_excel)

    # Calculate overall mean from numeric values
    overall_mean = df_excel.select_dtypes(include=[np.number]).mean().mean()

    # Fixed Parameters
    k = 5
    N = 25
    dfr = N - k  # 20

    # Verification Calculations (Repeatability)
    f_for_r = lookup_table_7(dfr, num_samples_input)
    uvlr = f_for_r * cvr_mfg_input
    result_r = "PASS" if cvr_input < uvlr else "FAIL"

    # Within Lab (WL) Calculations
    p = cvwl_mfg_input / cvr_mfg_input if cvr_mfg_input != 0 else 0
    dfwl = lookup_df_wl(p)
    f_for_wl = lookup_table_7(dfwl, num_samples_input)
    uvlwl = f_for_wl * cvwl_mfg_input
    result_wl = "PASS" if cvr_input < uvlwl else "FAIL"

    # -------------------------------------------------------------------------
    # 4. DISPLAY RESULTS
    # -------------------------------------------------------------------------
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Overall Mean", f"{overall_mean:.4f}")
    m2.metric("UVLR Target", f"{uvlr:.4f}")
    m3.metric("UVLWL Target", f"{uvlwl:.4f}")
    
    if result_r == "PASS" and result_wl == "PASS":
        m4.success("OVERALL: PASS")
    else:
        m4.error("OVERALL: FAIL")

    summary_data = {
        "Parameter": [
            "Overall Mean of Data", "CVr (Input)", "Number of Samples", "CVr of Manufacturer", 
            "k (Days)", "N (Total Runs)", "DFR", "F for R", "UVLR", "CVR Result",
            "CVWL of Manufacturer", "p ratio", "DFWL", "F for WL", "UVLWL", "CVWL Result"
        ],
        "Value": [
            f"{overall_mean:.4f}", f"{cvr_input:.4f}", num_samples_input, f"{cvr_mfg_input:.4f}",
            k, N, dfr, f"{f_for_r:.2f}", f"{uvlr:.4f}", result_r,
            f"{cvwl_mfg_input:.4f}", f"{p:.4f}", dfwl, f"{f_for_wl:.2f}", f"{uvlwl:.4f}", result_wl
        ]
    }

    summary_df = pd.DataFrame(summary_data)

    st.write("### 📊 Calculation Summary")
    
    def highlight_results(val):
        if val == "FAIL":
            return "background-color: #ffcdd2; color: #b71c1c; font-weight: bold;"
        elif val == "PASS":
            return "background-color: #c8e6c9; color: #1b5e20; font-weight: bold;"
        return ""

    st.dataframe(
        summary_df.style.map(highlight_results, subset=["Value"]),
        use_container_width=True
    )

    # -------------------------------------------------------------------------
    # 5. EXCEL DOWNLOAD GENERATION
    # -------------------------------------------------------------------------
    base_name = os.path.splitext(file_name)[0]
    output_filename = f"{base_name}_UVL_Results.xlsx"

    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
        df_excel.to_excel(writer, sheet_name="Original Data", index=False)
        summary_df.to_excel(writer, sheet_name="UVL Calculation", index=False)
    
    output_buffer.seek(0)

    st.download_button(
        label=f"📥 Download Processed Excel ({output_filename})",
        data=output_buffer,
        file_name=output_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("👋 Please upload an Excel file using the uploader above to run the UVL calculations.")