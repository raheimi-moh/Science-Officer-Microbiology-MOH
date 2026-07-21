import io
import os
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from scipy import stats
import streamlit as st

# -----------------------------------------------------------------------------
# PAGE CONFIGURATION
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Method Comparison: Bland-Altman & Regression",
    page_icon="📈",
    layout="wide"
)

st.title("📈 Method Comparison Analyzer")
st.write(
    "Upload your Excel or CSV file containing `CURRENT METHOD` and `NEW METHOD` "
    "columns to run Linear Regression, Bland-Altman analysis, and generate plots."
)

# -----------------------------------------------------------------------------
# FILE UPLOAD
# -----------------------------------------------------------------------------
uploaded_file = st.file_uploader(
    "Upload your QC data file:",
    type=["xlsx", "xls", "csv"]
)

if uploaded_file is not None:
    filename = uploaded_file.name
    st.subheader(f"📄 Processing File: `{filename}`")

    # Read file
    try:
        if filename.endswith(".xlsx") or filename.endswith(".xls"):
            df = pd.read_excel(uploaded_file, engine="openpyxl")
        else:
            df = pd.read_csv(uploaded_file)
    except Exception as e:
        st.error(f"Error loading file: {e}")
        st.stop()

    # Clean column names
    df.columns = df.columns.astype(str).str.strip()

    # Column Validation
    required_cols = ["CURRENT METHOD", "NEW METHOD"]
    if not all(col in df.columns for col in required_cols):
        st.error(
            f"❌ Missing required columns! Your file must contain these exact headers: {required_cols}"
        )
        st.stop()

    with st.expander("👀 View Raw Input Data"):
        st.dataframe(df)

    # -------------------------------------------------------------------------
    # 1. CALCULATIONS (BLAND-ALTMAN & REGRESSION)
    # -------------------------------------------------------------------------
    df["Mean X"] = (df["CURRENT METHOD"] + df["NEW METHOD"]) / 2
    df["Differences Y"] = df["NEW METHOD"] - df["CURRENT METHOD"]
    bias = df["Differences Y"].mean()
    sd_diff = df["Differences Y"].std()
    upper_loa = bias + (1.96 * sd_diff)
    lower_loa = bias - (1.96 * sd_diff)

    df["BIAS"] = bias
    df["SD of Differences"] = sd_diff
    df["UPPER LOA"] = upper_loa
    df["LOWER LOA"] = lower_loa
    df["Outlier"] = (df["Differences Y"] > upper_loa) | (df["Differences Y"] < lower_loa)

    # Linear Regression Calculations (Least Squares)
    x_reg = df["CURRENT METHOD"]
    y_reg = df["NEW METHOD"]
    slope, intercept, r_value, p_value, std_err = stats.linregress(x_reg, y_reg)
    r_squared = r_value**2

    df["Regression Slope"] = slope
    df["Regression Intercept"] = intercept
    df["Correlation Coefficient (r)"] = r_value
    df["Coefficient of Determination (r2)"] = r_squared

    # Determine Correlation Threshold Status
    if r_value >= 0.975 or r_squared >= 0.95:
        validation_text = "Data range is adequate and strong correlation (Passed Criteria)"
        passed_validation = True
    else:
        validation_text = "Correlation does not meet the specified strong threshold"
        passed_validation = False

    df["Validation Status"] = validation_text

    # -------------------------------------------------------------------------
    # 2. DISPLAY SUMMARY STATS & CARDS
    # -------------------------------------------------------------------------
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Bland-Altman Bias", f"{bias:.4f}")
    m2.metric("Upper LoA (+1.96 SD)", f"{upper_loa:.4f}")
    m3.metric("Lower LoA (-1.96 SD)", f"{lower_loa:.4f}")
    m4.metric("Correlation (r)", f"{r_value:.4f}")

    if passed_validation:
        st.success(f"✅ **Conclusion:** {validation_text}")
    else:
        st.warning(f"⚠️ **Conclusion:** {validation_text}")

    # -------------------------------------------------------------------------
    # 3. PLOTTING SIDE-BY-SIDE CHARTS
    # -------------------------------------------------------------------------
    st.write("### 📊 Diagnostic Analysis Plots")
    
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 6), dpi=100)

    # --- PLOT 1: SCATTER PLOT & REGRESSION ---
    ax1.scatter(
        x_reg, y_reg,
        color="#2b5c8f", alpha=0.7, edgecolors="k", label="Data Points"
    )

    x_line = np.linspace(x_reg.min(), x_reg.max(), 100)
    y_line = slope * x_line + intercept
    ax1.plot(
        x_line, y_line,
        color="#d9534f", linestyle="-", linewidth=2,
        label=f"Regression Line\nY = {slope:.3f}X + {intercept:.3f}"
    )

    ax1.plot(
        [x_reg.min(), x_reg.max()], [x_reg.min(), x_reg.max()],
        color="gray", linestyle=":", alpha=0.7, label="Identity Line (Y=X)"
    )

    stats_box = f"r = {r_value:.4f}\nr² = {r_squared:.4f}\n\n{validation_text}"
    ax1.text(
        0.05, 0.95, stats_box,
        transform=ax1.transAxes, fontsize=9, verticalalignment="top",
        bbox=dict(boxstyle="round,pad=0.5", facecolor="#f9f9f9", alpha=0.9, edgecolor="#ccc")
    )

    ax1.set_title("Regression Analysis: Least Squares", fontsize=12, weight="bold")
    ax1.set_xlabel("CURRENT METHOD", fontsize=10)
    ax1.set_ylabel("NEW METHOD", fontsize=10)
    ax1.grid(True, linestyle=":", alpha=0.6)
    ax1.legend(loc="lower right")

    # --- PLOT 2: BLAND-ALTMAN PLOT ---
    normal_data = df[~df["Outlier"]]
    outlier_data = df[df["Outlier"]]

    ax2.scatter(
        normal_data["Mean X"], normal_data["Differences Y"],
        color="#2b5c8f", alpha=0.7, edgecolors="k", label="Data Points"
    )
    if not outlier_data.empty:
        ax2.scatter(
            outlier_data["Mean X"], outlier_data["Differences Y"],
            color="#d9534f", alpha=0.9, edgecolors="k", s=60, marker="X",
            label="Outliers (>1.96 SD)"
        )

    ax2.axhline(bias, color="#222222", linestyle="-", linewidth=1.5)
    ax2.axhline(upper_loa, color="#d9534f", linestyle="--", linewidth=1.2)
    ax2.axhline(lower_loa, color="#d9534f", linestyle="--", linewidth=1.2)

    x_pos = df["Mean X"].max()
    ax2.text(x_pos, bias, f" Bias: {bias:.4f}", va="bottom", ha="left", fontweight="bold")
    ax2.text(x_pos, upper_loa, f" Upper LoA: {upper_loa:.4f}", va="bottom", ha="left", color="#d9534f")
    ax2.text(x_pos, lower_loa, f" Lower LoA: {lower_loa:.4f}", va="top", ha="left", color="#d9534f")

    ax2.set_title("Bland-Altman Difference Plot", fontsize=12, weight="bold")
    ax2.set_xlabel("Mean of Methods: (Current + New) / 2", fontsize=10)
    ax2.set_ylabel("Difference: New - Current", fontsize=10)
    ax2.grid(True, linestyle=":", alpha=0.6)
    ax2.legend(loc="upper left")

    plt.tight_layout()
    st.pyplot(fig)

    # -------------------------------------------------------------------------
    # 4. EXCEL EXPORT GENERATION
    # -------------------------------------------------------------------------
    base_name, _ = os.path.splitext(filename)
    output_filename = f"Analyzed_{base_name}.xlsx"

    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Bland Altman Analysis", index=False)

    output_buffer.seek(0)

    st.download_button(
        label=f"📥 Download Analyzed Excel File ({output_filename})",
        data=output_buffer,
        file_name=output_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("👋 Please upload an Excel or CSV file to begin analysis.")