import io
import os
import numpy as np
import pandas as pd
import scipy.stats as stats
import streamlit as st

# -----------------------------------------------------------------------------
# PAGE CONFIGURATION
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Serological Precision & ANOVA Analyzer",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Serological Precision & ANOVA Analyzer")
st.write(
    "Upload your QC matrix data (e.g., 5 days × 5 replicates) to perform One-Way ANOVA, "
    "calculate repeatability (CVR) and within-lab (CVWL) precision, and evaluate against manufacturer limits."
)

# -----------------------------------------------------------------------------
# SIDEBAR: MANUFACTURER LIMITS
# -----------------------------------------------------------------------------
st.sidebar.header("⚙️ Manufacturer Specifications")
cvr_limit = st.sidebar.number_input(
    "CVR Manufacturer Limit (%)",
    min_value=0.0,
    max_value=100.0,
    value=3.0,
    step=0.1,
    help="Repeatability Coefficient of Variation limit."
)

cvwl_limit = st.sidebar.number_input(
    "CVWL Manufacturer Limit (%)",
    min_value=0.0,
    max_value=100.0,
    value=5.0,
    step=0.1,
    help="Within-Laboratory Coefficient of Variation limit."
)

# -----------------------------------------------------------------------------
# FILE UPLOAD
# -----------------------------------------------------------------------------
uploaded_files = st.file_uploader(
    "Upload your Excel or CSV data file(s):",
    type=["xlsx", "xls", "csv"],
    accept_multiple_files=True
)

if uploaded_files:
    for uploaded_file in uploaded_files:
        filename = uploaded_file.name
        st.divider()
        st.subheader(f"📄 File: `{filename}`")

        # Load file
        try:
            if filename.endswith(".xlsx") or filename.endswith(".xls"):
                df_input = pd.read_excel(uploaded_file, index_col=0, engine="openpyxl")
            else:
                df_input = pd.read_csv(uploaded_file, index_col=0)
        except Exception as e:
            st.error(f"Error reading `{filename}`: {e}. Skipping file.")
            continue

        # Data preview foldout
        with st.expander("👀 View Raw Input Data"):
            st.dataframe(df_input)

        # ---------------------------------------------------------------------
        # 1. MATRIX VALIDATION & PREPROCESSING
        # ---------------------------------------------------------------------
        data_matrix = df_input.values
        n_days, n_replicates = data_matrix.shape

        if n_days != 5 or n_replicates != 5:
            st.warning(
                f"⚠️ Note: Expected a 5×5 data matrix, but received {n_days} days × {n_replicates} replicates. "
                "ANOVA will calculate based on detected dimensions."
            )

        # ---------------------------------------------------------------------
        # 2. STATISTICAL CALCULATIONS
        # ---------------------------------------------------------------------
        overall_mean = np.mean(data_matrix)
        anova_results = stats.f_oneway(*data_matrix)

        # Between-group calculations (Days)
        day_means = np.mean(data_matrix, axis=1)
        ss_between = n_replicates * np.sum((day_means - overall_mean) ** 2)
        df_between = n_days - 1
        ms_between = ss_between / df_between if df_between > 0 else 0.0

        # Within-group calculations (Error)
        ss_within = np.sum((data_matrix - day_means[:, np.newaxis]) ** 2)
        df_within = n_days * (n_replicates - 1)
        ms_within = ss_within / df_within if df_within > 0 else 0.0

        # Total calculations
        ss_total = ss_between + ss_within
        df_total = df_between + df_within

        # 3. Custom Metrics (Variances & Standard Deviations)
        vb = (ms_between - ms_within) / n_replicates
        vw = ms_within
        sdr = np.sqrt(vw) if vw >= 0 else np.nan
        
        # SDWL safeguard for negative variance sums
        var_sum = vb + vw
        sdwl = np.sqrt(var_sum) if var_sum >= 0 else np.nan

        cvr = (sdr / overall_mean) * 100 if overall_mean != 0 else 0.0
        cvwl = (sdwl / overall_mean) * 100 if overall_mean != 0 else 0.0

        # 4. Status Evaluation
        cvr_status = "PASS" if cvr <= cvr_limit else "FAIL"
        cvwl_status = "PASS" if cvwl <= cvwl_limit else "FAIL"

        # ---------------------------------------------------------------------
        # 3. SUMMARY CARDS & DISPLAY TABLES
        # ---------------------------------------------------------------------
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Overall Mean", f"{overall_mean:.4f}")
        col2.metric("Repeatability CV (CVR)", f"{cvr:.2f}%")
        col3.metric("Within-Lab CV (CVWL)", f"{cvwl:.2f}%")
        
        if cvr_status == "PASS" and cvwl_status == "PASS":
            col4.success("EVALUATION: PASS")
        else:
            col4.error("EVALUATION: FAIL")

        # Performance Evaluation Table
        st.write("### 🎯 Performance Evaluation")
        comparison_table = pd.DataFrame(
            {
                "Parameter": ["CVR (Repeatability CV %)", "CVWL (Within-laboratory CV %)"],
                "Calculated Value (%)": [f"{cvr:.2f}%", f"{cvwl:.2f}%"],
                "Manufacturer Limit (%)": [f"{cvr_limit:.2f}%", f"{cvwl_limit:.2f}%"],
                "Status": [cvr_status, cvwl_status],
            }
        )

        def highlight_status(val):
            if val == "FAIL":
                return "background-color: #ffcdd2; color: #b71c1c; font-weight: bold;"
            elif val == "PASS":
                return "background-color: #c8e6c9; color: #1b5e20; font-weight: bold;"
            return ""

        st.dataframe(
            comparison_table.style.map(highlight_status, subset=["Status"]),
            use_container_width=True
        )

        # ANOVA & Custom Metrics in Tabs
        tab1, tab2 = st.tabs(["📊 One-Way ANOVA Table", "📐 Detailed Precision Metrics"])

        with tab1:
            anova_table = pd.DataFrame(
                {
                    "Source of Variation": ["Between Days (Groups)", "Within Days (Error)", "Total"],
                    "SS": [ss_between, ss_within, ss_total],
                    "df": [df_between, df_within, df_total],
                    "MS": [ms_between, ms_within, None],
                    "F": [anova_results.statistic, None, None],
                    "P-value": [anova_results.pvalue, None, None],
                }
            )
            st.dataframe(anova_table, use_container_width=True)

        with tab2:
            custom_metrics_table = pd.DataFrame(
                {
                    "Metric": [
                        "Overall Mean",
                        "VB (Between-day Variance)",
                        "VW (Within-day Variance)",
                        "SDR (Repeatability SD)",
                        "SDWL (Within-laboratory SD)",
                        "CVR (Repeatability CV %)",
                        "CVWL (Within-laboratory CV %)",
                    ],
                    "Value": [
                        f"{overall_mean:.4f}",
                        f"{vb:.6f}",
                        f"{vw:.6f}",
                        f"{sdr:.4f}",
                        f"{sdwl:.4f}" if not np.isnan(sdwl) else "N/A",
                        f"{cvr:.2f}%",
                        f"{cvwl:.2f}%",
                    ],
                }
            )
            st.dataframe(custom_metrics_table, use_container_width=True)

        # ---------------------------------------------------------------------
        # 4. EXCEL EXPORT GENERATION
        # ---------------------------------------------------------------------
        file_base, file_ext = os.path.splitext(filename)
        output_filename = f"{file_base}_analyzed.xlsx"

        output_buffer = io.BytesIO()
        try:
            with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
                comparison_table.to_excel(writer, sheet_name="Evaluation Summary", index=False)
                anova_table.to_excel(writer, sheet_name="ANOVA", index=False)
                custom_metrics_table.to_excel(writer, sheet_name="Custom Metrics", index=False)
            
            output_buffer.seek(0)

            st.download_button(
                label=f"📥 Download Processed Excel ({output_filename})",
                data=output_buffer,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_{filename}"
            )
        except Exception as e:
            st.error(f"Failed to create Excel output: {e}")

    st.toast("ANOVA evaluation complete!", icon="🎉")