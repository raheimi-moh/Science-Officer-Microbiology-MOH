import io
import os
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import streamlit as st

# -----------------------------------------------------------------------------
# PAGE CONFIGURATION
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Levey-Jennings (L-J) Chart Generator",
    page_icon="📈",
    layout="wide"
)

st.title("📈 Levey-Jennings (L-J) Chart Generator")
st.write(
    "Upload your QC measurement data to calculate summary statistics, generate "
    "a Levey-Jennings control chart, and export the analysis to Excel."
)

# -----------------------------------------------------------------------------
# SIDEBAR: METADATA INPUTS
# -----------------------------------------------------------------------------
st.sidebar.header("📝 Run Metadata")

kit_name = st.sidebar.text_input("Kit Name", value="Elisa Kit A")
lot_no = st.sidebar.text_input("Lot No", value="LOT-12345")
expiry_date = st.sidebar.text_input("Expiry Date", value="2026-12-31")
equipment = st.sidebar.text_input("Equipment", value="Analyzer 01")
operator = st.sidebar.text_input("Operator", value="Tech A")

# -----------------------------------------------------------------------------
# FILE UPLOAD & DATA PROCESSING
# -----------------------------------------------------------------------------
uploaded_file = st.file_uploader(
    "Upload Excel file containing QC measurements:",
    type=["xlsx", "xls", "csv"]
)

if uploaded_file is not None:
    filename = uploaded_file.name
    base_name = os.path.splitext(filename)[0]

    try:
        if filename.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file, engine="openpyxl")
    except Exception as e:
        st.error(f"Error reading file: {e}")
        st.stop()

    # Clean column names
    df.columns = df.columns.astype(str).str.strip()

    # Column Validation
    required_cols = ["NO OF ASSAY", "MEASUREMENT"]
    if not all(col in df.columns for col in required_cols):
        st.error(
            f"❌ Header Mismatch! File must contain these exact headers: `{required_cols}`"
        )
        st.stop()

    # Clean data drop NA
    df = df.dropna(subset=["MEASUREMENT"])
    measurements = df["MEASUREMENT"].values

    if len(measurements) < 2:
        st.error("❌ At least 2 valid numeric measurements are required to calculate Standard Deviation.")
        st.stop()

    # -------------------------------------------------------------------------
    # 1. CALCULATE QC STATISTICS
    # -------------------------------------------------------------------------
    mean_val = np.mean(measurements)
    sd_val = np.std(measurements, ddof=1)  # Sample Standard Deviation
    cv_val = (sd_val / mean_val) * 100 if mean_val != 0 else 0.0

    sd_lines = {
        "+3SD": mean_val + 3 * sd_val,
        "+2SD": mean_val + 2 * sd_val,
        "+1SD": mean_val + 1 * sd_val,
        "MEAN": mean_val,
        "-1SD": mean_val - 1 * sd_val,
        "-2SD": mean_val - 2 * sd_val,
        "-3SD": mean_val - 3 * sd_val,
    }

    # Display Metrics
    st.write("### 📊 Calculated Statistics")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Count (N)", len(measurements))
    m2.metric("Mean", f"{mean_val:.4f}")
    m3.metric("Std Deviation (SD)", f"{sd_val:.4f}")
    m4.metric("CV (%)", f"{cv_val:.2f}%")

    # -------------------------------------------------------------------------
    # 2. GENERATE LEVEY-JENNINGS PLOT
    # -------------------------------------------------------------------------
    st.write("### 📉 Levey-Jennings Control Chart")

    fig, ax = plt.subplots(figsize=(12, 6), dpi=100)
    x_axis = df["NO OF ASSAY"].astype(str).values

    # Plot measurement line
    ax.plot(
        x_axis,
        measurements,
        marker="o",
        color="black",
        linewidth=1.5,
        label="Measurement",
        zorder=5
    )

    # SD Line Configs
    colors = {
        "+3SD": "red",
        "+2SD": "orange",
        "+1SD": "goldenrod",
        "MEAN": "green",
        "-1SD": "goldenrod",
        "-2SD": "orange",
        "-3SD": "red",
    }
    linestyles = {
        "+3SD": "--",
        "+2SD": "--",
        "+1SD": ":",
        "MEAN": "-",
        "-1SD": ":",
        "-2SD": "--",
        "-3SD": "--",
    }

    for label, val in sd_lines.items():
        ax.axhline(
            y=val,
            color=colors[label],
            linestyle=linestyles[label],
            linewidth=1.2,
            label=f"{label} ({val:.2f})"
        )

    title_text = (
        f"Levey-Jennings Chart\n"
        f"Kit: {kit_name} | Lot: {lot_no} | Exp: {expiry_date} | Equip: {equipment} | Op: {operator}"
    )
    ax.set_title(title_text, fontsize=11, fontweight="bold", pad=15)
    ax.set_xlabel("No. of Assay", fontweight="bold")
    ax.set_ylabel("Measurement", fontweight="bold")
    plt.xticks(rotation=45)
    ax.grid(True, linestyle=":", alpha=0.6)
    ax.legend(bbox_to_anchor=(1.02, 1), loc="upper left", borderaxespad=0)
    plt.tight_layout()

    st.pyplot(fig)

    # -------------------------------------------------------------------------
    # 3. EXCEL EXPORT GENERATION
    # -------------------------------------------------------------------------
    stats_df = pd.DataFrame(
        [
            ["Kit Name", kit_name],
            ["Lot No", lot_no],
            ["Expiry Date", expiry_date],
            ["Equipment", equipment],
            ["Operator", operator],
            ["", ""],
            ["Count (N)", len(measurements)],
            ["Mean", mean_val],
            ["SD", sd_val],
            ["CV%", f"{cv_val:.2f}%"],
            ["+3SD", sd_lines["+3SD"]],
            ["+2SD", sd_lines["+2SD"]],
            ["+1SD", sd_lines["+1SD"]],
            ["MEAN Line", sd_lines["MEAN"]],
            ["-1SD", sd_lines["-1SD"]],
            ["-2SD", sd_lines["-2SD"]],
            ["-3SD", sd_lines["-3SD"]],
        ],
        columns=["QC Parameter", "Value"]
    )

    output_filename = f"LJ_Result_{base_name}.xlsx"
    output_buffer = io.BytesIO()

    try:
        with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Assay Data", index=False)
            stats_df.to_excel(writer, sheet_name="QC Summary", index=False)

        output_buffer.seek(0)

        st.download_button(
            label=f"📥 Download Processed Excel ({output_filename})",
            data=output_buffer,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Failed to build Excel download file: {e}")

else:
    st.info("👋 Please upload an Excel or CSV file containing 'NO OF ASSAY' and 'MEASUREMENT' columns.")