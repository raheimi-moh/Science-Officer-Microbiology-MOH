import io
import pandas as pd
import streamlit as st

# -----------------------------------------------------------------------------
# PAGE CONFIGURATION
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Cumulative Uncertainty (CU%) Calculator",
    page_icon="📐",
    layout="wide"
)

st.title("📐 Cumulative Uncertainty (CU%) Analysis Tool")
st.write(
    "Calculate the **Cumulative Uncertainty (CU%)** for diagnostic methods across "
    "Positive ($CU_1$) and Negative ($CU_2$) groups against the $\le 10\%$ acceptance threshold."
)

# -----------------------------------------------------------------------------
# 1. INPUT CONTROLS
# -----------------------------------------------------------------------------
st.sidebar.header("📥 Contingency Matrix Inputs")

a = st.sidebar.number_input("a (True Positives - TP)", min_value=0, value=45, step=1)
b = st.sidebar.number_input("b (False Positives - FP)", min_value=0, value=5, step=1)
c = st.sidebar.number_input("c (False Negatives - FN)", min_value=0, value=5, step=1)
d = st.sidebar.number_input("d (True Negatives - TN)", min_value=0, value=45, step=1)

# Input Validation for denominators
pos_total = a + c
neg_total = b + d

if pos_total == 0 or neg_total == 0:
    st.error("⚠️ Validation Error: Both $(a + c)$ and $(b + d)$ must be greater than 0.")
    st.stop()

# -----------------------------------------------------------------------------
# 2. CORE CALCULATIONS
# -----------------------------------------------------------------------------
# Section 1: Positive Group CU% (CU1)
p1_1 = a / pos_total
p2_1 = 1.0 - p1_1
cu1 = 2.0 * p1_1 * p2_1 * 100.0

# Section 2: Negative Group CU% (CU2)
p1_2 = d / neg_total
p2_2 = 1.0 - p1_2
cu2 = 2.0 * p1_2 * p2_2 * 100.0

# Section 3: Average CU%
avg_cu = (cu1 + cu2) / 2.0

# Acceptance Criteria Check
passed = avg_cu <= 10.0
status_text = "ACCEPTED (≤ 10%)" if passed else "REJECTED (> 10%)"

# -----------------------------------------------------------------------------
# 3. DISPLAY SUMMARY METRICS & STATUS
# -----------------------------------------------------------------------------
st.write("### 📈 Key Results Summary")
m1, m2, m3 = st.columns(3)
m1.metric("Positive Group (CU1%)", f"{cu1:.2f}%")
m2.metric("Negative Group (CU2%)", f"{cu2:.2f}%")
m3.metric("Average CU%", f"{avg_cu:.2f}%")

if passed:
    st.success(f"✅ **Evaluation Status:** {status_text}")
else:
    st.error(f"❌ **Evaluation Status:** {status_text}")

st.divider()

col_left, col_right = st.columns(2)

with col_left:
    st.write("### 📋 2×2 Input Matrix")
    matrix_data = {
        "Condition / Test": ["Positive Result", "Negative Result", "Total Group Size"],
        "Target Positive (a + c)": [a, c, pos_total],
        "Target Negative (b + d)": [b, d, neg_total]
    }
    df_matrix = pd.DataFrame(matrix_data)
    st.dataframe(df_matrix, use_container_width=True)

with col_right:
    st.write("### 📊 Calculated Probability Parameters")
    param_data = {
        "Group Parameter": ["Positive Group (CU1)", "Negative Group (CU2)"],
        "p1 (Proportion Correct)": [f"{p1_1:.4f}", f"{p1_2:.4f}"],
        "p2 (1 - p1 Error)": [f"{p2_1:.4f}", f"{p2_2:.4f}"],
        "Group CU (%)": [f"{cu1:.2f}%", f"{cu2:.2f}%"]
    }
    df_params = pd.DataFrame(param_data)
    st.dataframe(df_params, use_container_width=True)

# -----------------------------------------------------------------------------
# 4. EXCEL EXPORT GENERATION
# -----------------------------------------------------------------------------
output_filename = "Cumulative_Uncertainty_Report.xlsx"

summary_data = {
    "Parameter": [
        "a (True Positives)", "b (False Positives)", "c (False Negatives)", "d (True Negatives)",
        "Positive Group Size (a + c)", "Negative Group Size (b + d)",
        "Positive Group p1", "Positive Group p2", "Positive Group CU1 (%)",
        "Negative Group p1", "Negative Group p2", "Negative Group CU2 (%)",
        "Average CU (%)", "Evaluation Status"
    ],
    "Value": [
        a, b, c, d,
        pos_total, neg_total,
        round(p1_1, 4), round(p2_1, 4), round(cu1, 2),
        round(p1_2, 4), round(p2_2, 4), round(cu2, 2),
        round(avg_cu, 2), status_text
    ]
}
df_summary = pd.DataFrame(summary_data)

output_buffer = io.BytesIO()

try:
    with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
        df_matrix.to_excel(writer, sheet_name="Input Matrix", index=False)
        df_summary.to_excel(writer, sheet_name="CU Analysis Summary", index=False)

    output_buffer.seek(0)

    st.download_button(
        label=f"📥 Download Processed Report ({output_filename})",
        data=output_buffer,
        file_name=output_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
except Exception as e:
    st.error(f"Failed to generate Excel download file: {e}")