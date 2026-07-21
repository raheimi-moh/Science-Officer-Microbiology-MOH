import io
import os
import numpy as np
import pandas as pd
import streamlit as st

# Page setup
st.set_page_config(page_title="Serological QC Data Analyzer", page_icon="🧪", layout="wide")

st.title("🧪 Serological QC Data Analyzer")
st.write("Upload your QC data files (Excel or CSV) to calculate summary statistics, evaluate Grubbs limits, and export formatted reports.")

# 1. Upload multiple files via Web Interface
uploaded_files = st.file_uploader(
    "Please select and upload your QC data file(s):",
    type=["csv", "xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files:
    # 2. Loop through every uploaded file
    for uploaded_file in uploaded_files:
        filename = uploaded_file.name
        
        st.divider()
        st.subheader(f"📄 Processing File: `{filename}`")
        
        # Read the data based on file type
        try:
            if filename.endswith(".xlsx") or filename.endswith(".xls"):
                df_input = pd.read_excel(uploaded_file)
            else:
                df_input = pd.read_csv(uploaded_file)
        except Exception as e:
            st.error(f"Error reading {filename}: {e}. Skipping file.")
            continue

        # 3. Identify columns (First column is Day, next 5 are Replicates)
        day_col = df_input.columns[0]
        replicate_cols = df_input.columns[1:6]

        # Flatten all data points for overall statistical calculations
        all_values = df_input[replicate_cols].values.flatten()

        # 4. Perform Statistical Calculations
        total_sum = np.sum(all_values)
        mean_val = np.mean(all_values)  # Average
        sd_val = np.std(all_values, ddof=1)  # Sample standard deviation
        cv_percentage = (sd_val / mean_val) * 100 if mean_val != 0 else 0

        # Grubbs Limits based on formula: mean +/- (3.135 * sd)
        upper_limit = mean_val + (3.135 * sd_val)
        lower_limit = mean_val - (3.135 * sd_val)

        # 5. Evaluate Individual Cells (Replicates) and check for any failures
        df_result = df_input.copy()
        any_failed = False  # Track if any single data point fails
        
        for row_idx in range(len(df_input)):
            for col in replicate_cols:
                val = df_input.at[row_idx, col]
                
                # Check individual value against limits
                if lower_limit <= val <= upper_limit:
                    status = "PASS"
                else:
                    status = "FAIL"
                    any_failed = True  # At least one data point exceeded limits
                
                # Format display string
                df_result.at[row_idx, col] = f"{val} [{status}]"

        # 6. Strict Overall Status Check
        overall_status = "FAIL!" if any_failed else "PASS"

        # Display Summary Cards / Metrics in Streamlit
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Sum", f"{total_sum:.4f}")
        col2.metric("Overall Mean", f"{mean_val:.4f}")
        col3.metric("Std Deviation", f"{sd_val:.4f}")
        col4.metric("CV (%)", f"{cv_percentage:.2f}%")

        col_l1, col_l2, col_status = st.columns([1, 1, 1])
        col_l1.metric("Grubbs Lower Limit", f"{lower_limit:.4f}")
        col_l2.metric("Grubbs Upper Limit", f"{upper_limit:.4f}")

        # Display overall status banner
        if overall_status == "FAIL!":
            col_status.error(f"OVERALL STATUS: {overall_status}")
            st.warning("⚠️ Notice: One or more raw data points exceeded the Grubbs limits.")
        else:
            col_status.success(f"OVERALL STATUS: {overall_status}")
            st.success("✅ Success: All raw data points are within limits.")

        # Display Result Table on web page
        st.write("### QC Evaluation Table")
        st.dataframe(df_result, use_container_width=True)

        # 7. Create Excel Export File in Memory
        summary_data = {
            "Metric": ["Total Sum", "Overall Mean", "Standard Deviation", "CV (%)", "Grubbs Upper Limit", "Grubbs Lower Limit", "Overall QC Status"],
            "Value": [total_sum, mean_val, sd_val, f"{cv_percentage:.2f}%", upper_limit, lower_limit, overall_status]
        }
        df_summary = pd.DataFrame(summary_data)

        # Write multi-sheet Excel into an in-memory buffer
        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
            df_result.to_excel(writer, sheet_name="Replicate Results", index=False)
            df_summary.to_excel(writer, sheet_name="Overall Summary", index=False)
        output_buffer.seek(0)

        # File naming
        base_name, _ = os.path.splitext(filename)
        output_filename = f"{base_name}_Grubbs_Results.xlsx"

        # Streamlit Download Button
        st.download_button(
            label=f"📥 Download {output_filename}",
            data=output_buffer,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=filename  # Unique key for each file loop
        )

    st.toast("All files processed successfully!", icon="🎉")