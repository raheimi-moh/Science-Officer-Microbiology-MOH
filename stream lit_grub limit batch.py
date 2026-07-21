import io
import os
import numpy as np
import pandas as pd
import streamlit as st

# -----------------------------------------------------------------------------
# PAGE CONFIGURATION
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Serological QC Data Analyzer",
    page_icon="🧪",
    layout="wide"
)

st.title("🧪 Serological QC Data Analyzer")
st.write(
    "Upload your QC data files (Excel or CSV) to calculate summary statistics, "
    "evaluate Grubbs limits, and export formatted reports."
)

# -----------------------------------------------------------------------------
# FILE UPLOAD
# -----------------------------------------------------------------------------
uploaded_files = st.file_uploader(
    "Please select and upload your QC data file(s):",
    type=["csv", "xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files:
    for uploaded_file in uploaded_files:
        filename = uploaded_file.name
        
        st.divider()
        st.subheader(f"📄 Processing File: `{filename}`")
        
        # Read the file based on extension
        try:
            if filename.endswith(".xlsx") or filename.endswith(".xls"):
                # Requires openpyxl engine
                df_input = pd.read_excel(uploaded_file, engine="openpyxl")
            else:
                df_input = pd.read_csv(uploaded_file)
        except Exception as e:
            st.error(f"Error reading `{filename}`: {e}. Skipping file.")
            continue

        # ---------------------------------------------------------------------
        # 1. IDENTIFY COLUMNS & EXTRACT DATA
        # ---------------------------------------------------------------------
        day_col = df_input.columns[0]
        replicate_cols = df_input.columns[1:6]

        # Flatten all 25 replicate data points for overall statistical calculations
        all_values = df_input[replicate_cols].values.flatten()

        # ---------------------------------------------------------------------
        # 2. PERFORM STATISTICAL CALCULATIONS
        # ---------------------------------------------------------------------
        total_sum = np.sum(all_values)
        mean_val = np.mean(all_values)  # Average
        sd_val = np.std(all_values, ddof=1)  # Sample standard deviation
        cv_percentage = (sd_val / mean_val) * 100 if mean_val != 0 else 0.0

        # Grubbs Limits based on formula: mean +/- (3.135 * sd)
        upper_limit = mean_val + (3.135 * sd_val)
        lower_limit = mean_val - (3.135 * sd_val)

        # ---------------------------------------------------------------------
        # 3. EVALUATE INDIVIDUAL CELLS & CHECK FOR FAILURES
        # ---------------------------------------------------------------------
        df_result = df_input.copy()
        any_failed = False  # Tracks if any single data point fails
        
        for row_idx in range(len(df_input)):
            for col in replicate_cols:
                val = df_input.at[row_idx, col]
                
                # Check individual value against limits
                if lower_limit <= val <= upper_limit:
                    status = "PASS"
                else:
                    status = "FAIL"
                    any_failed = True
                
                # Format cell text display
                df_result.at[row_idx, col] = f"{val} [{status}]"

        # Overall QC Status
        overall_status = "FAIL!" if any_failed else "PASS"

        # ---------------------------------------------------------------------
        # 4. DISPLAY SUMMARY METRICS & STATUS
        # ---------------------------------------------------------------------
        m_col1, m_col2, m_col3, m_col4 = st.columns(4)
        m_col1.metric("Total Sum", f"{total_sum:.4f}")
        m_col2.metric("Overall Mean", f"{mean_val:.4f}")
        m_col3.metric("Std Deviation", f"{sd_val:.4f}")
        m_col4.metric("CV (%)", f"{cv_percentage:.2f}%")

        l_col1, l_col2, status_col = st.columns([1, 1, 1])
        l_col1.metric("Grubbs Lower Limit", f"{lower_limit:.4f}")
        l_col2.metric("Grubbs Upper Limit", f"{upper_limit:.4f}")

        if overall_status == "FAIL!":
            status_col.error(f"OVERALL QC STATUS: {overall_status}")
            st.warning("⚠️ Notice: One or more raw data points exceeded the Grubbs limits.")
        else:
            status_col.success(f"OVERALL QC STATUS: {overall_status}")
            st.success("✅ Success: All 25 raw data points are perfectly within limits.")

        # ---------------------------------------------------------------------
        # 5. HIGHLIGHT & DISPLAY RESULT TABLE
        # ---------------------------------------------------------------------
        st.write("### QC Evaluation Results Table")

        # Function to highlight [FAIL] cells in soft red
        def highlight_failures(val):
            if isinstance(val, str) and "[FAIL]" in val:
                return "background-color: #ffcdd2; color: #b71c1c; font-weight: bold;"
            elif isinstance(val, str) and "[PASS]" in val:
                return "background-color: #c8e6c9; color: #1b5e20;"
            return ""

        styled_df = df_result.style.applymap(
            highlight_failures, 
            subset=replicate_cols
        )
        
        st.dataframe(styled_df, use_container_width=True)

        # ---------------------------------------------------------------------
        # 6. CREATE EXCEL EXPORT (IN-MEMORY BUFFER)
        # ---------------------------------------------------------------------
        summary_data = {
            "Metric": [
                "Total Sum", 
                "Overall Mean", 
                "Standard Deviation", 
                "CV (%)", 
                "Grubbs Upper Limit", 
                "Grubbs Lower Limit", 
                "Overall QC Status"
            ],
            "Value": [
                total_sum, 
                mean_val, 
                sd_val, 
                f"{cv_percentage:.2f}%", 
                upper_limit, 
                lower_limit, 
                overall_status
            ]
        }
        df_summary = pd.DataFrame(summary_data)

        # Build Excel file into BytesIO buffer
        output_buffer = io.BytesIO()
        try:
            with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
                df_result.to_excel(writer, sheet_name="Replicate Results", index=False)
                df_summary.to_excel(writer, sheet_name="Overall Summary", index=False)
            output_buffer.seek(0)

            # Generate dynamic filename
            base_name, _ = os.path.splitext(filename)
            output_filename = f"{base_name}_Grubbs_Results.xlsx"

            # Download Button
            st.download_button(
                label=f"📥 Download Processed Excel ({output_filename})",
                data=output_buffer,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=filename
            )
        except Exception as e:
            st.error(f"Failed to generate Excel file for download: {e}")

    st.toast("Processing complete for all files!", icon="🎉")