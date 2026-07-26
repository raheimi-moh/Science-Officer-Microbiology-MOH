import io
import pandas as pd
import streamlit as st

# -----------------------------------------------------------------------------
# PAGE CONFIGURATION
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Cohen's Kappa Calculator",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Cohen's Kappa Inter-Rater Reliability Calculator")
st.write(
    "Calculate **Cohen's Kappa ($\kappa$)** to assess agreement between two raters "
    "evaluating categorical items, accounting for chance agreement."
)

# -----------------------------------------------------------------------------
# 1. INPUT CONTROLS
# -----------------------------------------------------------------------------
st.sidebar.header("📥 Contingency Table Values")

a = st.sidebar.number_input("A (Pos / Pos)", min_value=0, value=45, step=1)
b = st.sidebar.number_input("B (Pos / Neg)", min_value=0, value=5, step=1)
c = st.sidebar.number_input("C (Neg / Pos)", min_value=0, value=10, step=1)
d = st.sidebar.number_input("D (Neg / Neg)", min_value=0, value=40, step=1)

n_total = a + b + c + d

if n_total <= 0:
    st.error("⚠️ Total observations (N) must be greater than 0.")
    st.stop()

# -----------------------------------------------------------------------------
# 2. CORE CALCULATIONS
# -----------------------------------------------------------------------------
# 1. Observed Agreement (P_o)
p_o = (a + d) / n_total

# 2. Expected Agreement (P_e)
p_i_dot = ((a + b) / n_total) * ((a + c) / n_total)
p_dot_i = ((c + d) / n_total) * ((b + d) / n_total)
p_e = p_i_dot + p_dot_i

# 3. Cohen's Kappa (k)
kappa = 1.0 if p_e == 1 else (p_o - p_e) / (1.0 - p_e)

# Interpretation mapping
if kappa < 0.21:
    level, reliability = "Poor agreement", "0 - 4%"
elif 0.21 <= kappa <= 0.40:
    level, reliability = "Fair agreement", "4 - 15%"
elif 0.41 <= kappa <= 0.60:
    level, reliability = "Moderate agreement", "35 - 63%"
elif 0.61 <= kappa <= 0.80:
    level, reliability = "Good agreement", "64 - 81%"
else:
    level, reliability = "Almost perfect", "82 - 100%"

# -----------------------------------------------------------------------------
# 3. DISPLAY RESULTS
# -----------------------------------------------------------------------------
st.write("### 📈 Key Metrics")
m1, m2, m3, m4 = st.columns(4)
m1.metric("Total Count (N)", n_total)
m2.metric("Observed Agreement (P_o)", f"{p_o:.4f}")
m3.metric("Expected Agreement (P_e)", f"{p_e:.4f}")
m4.metric("Cohen's Kappa (κ)", f"{kappa:.4f}")

# Qualitative Status Card
if kappa >= 0.61:
    st.success(f"✅ **Level of Agreement:** {level} (Reliability: {reliability})")
elif kappa >= 0.41:
    st.info(f"ℹ️ **Level of Agreement:** {level} (Reliability: {reliability})")
else:
    st.warning(f"⚠️ **Level of Agreement:** {level} (Reliability: {reliability})")

st.divider()

col_left, col_right = st.columns(2)

with col_left:
    st.write("### 📋 2×2 Contingency Matrix")
    matrix_data = {
        "Rater 1 / Rater 2": ["Positive", "Negative", "Total"],
        "Positive": [a, c, a + c],
        "Negative": [b, d, b + d],
        "Total": [a + b, c + d, n_total]
    }
    df_matrix = pd.DataFrame(matrix_data)
    st.dataframe(df_matrix, use_container_width=True)

with col_right:
    st.write("### 📝 Reliability Benchmark Table")
    benchmarks = {
        "Kappa Statistic (κ)": ["< 0.21", "0.21 – 0.40", "0.41 – 0.60", "0.61 – 0.80", "0.81 – 1.00"],
        "Level of Agreement": ["Poor agreement", "Fair agreement", "Moderate agreement", "Good agreement", "Almost perfect"],
        "Reliability of Data": ["0 - 4%", "4 - 15%", "35 - 63%", "64 - 81%", "82 - 100%"]
    }
    df_benchmarks = pd.DataFrame(benchmarks)
    st.dataframe(df_benchmarks, use_container_width=True)

# -----------------------------------------------------------------------------
# 4. EXCEL EXPORT GENERATION
# -----------------------------------------------------------------------------
output_filename = "Cohens_Kappa_Analysis_Report.xlsx"

summary_data = {
    "Parameter": [
        "Cell A (Pos/Pos)", "Cell B (Pos/Neg)", "Cell C (Neg/Pos)", "Cell D (Neg/Neg)",
        "Total Count (N)", "Observed Agreement (P_o)", "Expected Agreement (P_e)",
        "Cohen's Kappa (k)", "Level of Agreement", "Reliability Range"
    ],
    "Value": [
        a, b, c, d,
        n_total, round(p_o, 4), round(p_e, 4),
        round(kappa, 4), level, reliability
    ]
}
df_summary = pd.DataFrame(summary_data)

output_buffer = io.BytesIO()

try:
    with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
        df_matrix.to_excel(writer, sheet_name="Contingency Matrix", index=False)
        df_summary.to_excel(writer, sheet_name="Kappa Analysis", index=False)

    output_buffer.seek(0)

    st.download_button(
        label=f"📥 Download Report ({output_filename})",
        data=output_buffer,
        file_name=output_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
except Exception as e:
    st.error(f"Failed to generate Excel download file: {e}")