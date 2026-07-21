import io
import os
import numpy as np
import pandas as pd
import streamlit as st

# -----------------------------------------------------------------------------
# PAGE CONFIGURATION
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="pSMILE New Reagent Lot Verification Tool",
    page_icon="🧪",
    layout="wide"
)

st.title("🧪 pSMILE New Reagent Lot Verification Tool")
st.write(
    "Upload your reagent verification Excel file containing replicates for current "
    "and candidate lots to evaluate the acceptability ratio against your target CV%."
)

# -----------------------------------------------------------------------------
# SIDEBAR CONTROLS
# -----------------------------------------------------------------------------
st.sidebar.header("⚙️ Target Parameters")

target_cv = st.sidebar.number_input(
    "Target CV (%)",
    min_value=0.01,
    max_value=100.0,
    value=7.86,
    step=0.01,
    format="%.2f",
    help="Target Coefficient of Variation percentage."
)

# -----------------------------------------------------------------------------
# FILE UPLOAD & PROCESSING
# -----------------------------------------------------------------------------
uploaded_file = st.file_uploader(
    "Upload Excel file (.xlsx):",
    type=["xlsx", "xls"]
)

if uploaded_file is not None:
    filename = uploaded_file.name
    input_base_name = os.path.splitext(filename)[0]

    try:
        df_data = pd.read_excel(uploaded_file, engine="openpyxl")
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        st.stop()

    # Column Validation
    required_cols = ['Sample', 'Current_R1', 'Current_R2', 'Candidate_R1', 'Candidate_R2']
    if not all(col in df_data.columns for col in required_cols):
        st.error(
            f"❌ Columns mismatch! The uploaded sheet must contain exact headers: `{required_cols}`"
        )
        st.stop()

    st.subheader(f"📄 Processing: `{filename}` ({len(df_data)} samples found)")

    # -------------------------------------------------------------------------
    # 1. PHASE 1 CALCULATIONS
    # -------------------------------------------------------------------------
    df_data['Current_Mean'] = df_data[['Current_R1', 'Current_R2']].mean(axis=1)
    df_data['Candidate_Mean'] = df_data[['Candidate_R1', 'Candidate_R2']].mean(axis=1)

    avg_mean_Xc = df_data['Current_Mean'].mean()
    avg_mean_Xn = df_data['Candidate_Mean'].mean()
    grand_mean_Xb = (avg_mean_Xc + avg_mean_Xn) / 2

    # -------------------------------------------------------------------------
    # 2. PHASE 2 CALCULATIONS (ACCEPTABILITY RATIO)
    # -------------------------------------------------------------------------
    numerator = abs(avg_mean_Xc - avg_mean_Xn)
    denominator = (target_cv / 100.0) * grand_mean_Xb
    
    ratio_val = numerator / denominator if denominator != 0 else np.nan

    # Determine universal acceptability status
    if ratio_val <= 1.0:
        status_text = "Acceptable, Reagent fit for use"
        is_acceptable = True
    else:
        status_text = "Unacceptable"
        is_acceptable = False

    # Store metrics in DataFrame
    df_data['Average_Mean_Xc'] = avg_mean_Xc
    df_data['Average_Mean_Xn'] = avg_mean_Xn
    df_data['Grand_Mean_Xb'] = grand_mean_Xb
    df_data['Target_CV_Pct'] = target_cv
    df_data['Acceptability_Ratio'] = ratio_val
    df_data['Final_Status'] = status_text

    # -------------------------------------------------------------------------
    # 3. DISPLAY SUMMARY METRICS
    # -------------------------------------------------------------------------
    st.write("### 📊 Summary Lot Metrics")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Current Lot Mean (Xc)", f"{avg_mean_Xc:.4f}")
    m2.metric("New Lot Mean (Xn)", f"{avg_mean_Xn:.4f}")
    m3.metric("Grand Mean (Xb)", f"{grand_mean_Xb:.4f}")
    m4.metric("Acceptability Ratio", f"{ratio_val:.4f}")

    if is_acceptable:
        st.success(f"✅ **FINAL STATUS:** {status_text} (Ratio ≤ 1.0)")
    else:
        st.error(f"❌ **FINAL STATUS:** {status_text} (Ratio > 1.0)")

    # -------------------------------------------------------------------------
    # 4. DATA TABLES
    # -------------------------------------------------------------------------
    tab1, tab2 = st.tabs(["📋 Sample Breakdown", "📑 Complete Analysis Report"])

    with tab1:
        st.dataframe(
            df_data[[
                'Sample', 'Current_R1', 'Current_R2', 'Current_Mean', 
                'Candidate_R1', 'Candidate_R2', 'Candidate_Mean'
            ]].round(4),
            use_container_width=True
        )

    with tab2:
        st.dataframe(df_data.round(4), use_container_width=True)

    # -------------------------------------------------------------------------
    # 5. EXCEL EXPORT GENERATION
    # -------------------------------------------------------------------------
    output_filename = f"{input_base_name}_pSMILE_results.xlsx"
    output_buffer = io.BytesIO()

    try:
        with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
            df_data.to_excel(writer, sheet_name="Reagent Verification", index=False)
        
        output_buffer.seek(0)

        st.download_button(
            label=f"📥 Download Processed Report ({output_filename})",
            data=output_buffer,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Failed to create Excel download file: {e}")

else:
    st.info("👋 Please upload a Reagent Verification Excel file to begin calculations.")