import io
import math
import pandas as pd
import streamlit as st

# -----------------------------------------------------------------------------
# PAGE CONFIGURATION
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Method Comparison Calculator",
    page_icon="🧪",
    layout="wide"
)

st.title("🧪 Diagnostic Method Comparison Calculator")
st.write(
    "Enter your $2 \\times 2$ contingency table values to compute PPA, "
    "NPA, OPA, and 95% Score Confidence Intervals."
)

# -----------------------------------------------------------------------------
# 1. INPUT CONTROLS
# -----------------------------------------------------------------------------
st.sidebar.header("📊 Diagnostic Input Values")

tp = st.sidebar.number_input("True Positives (TP)", min_value=0, value=80, step=1)
fp = st.sidebar.number_input("False Positives (FP)", min_value=0, value=10, step=1)
fn = st.sidebar.number_input("False Negatives (FN)", min_value=0, value=5, step=1)
tn = st.sidebar.number_input("True Negatives (TN)", min_value=0, value=105, step=1)

# -----------------------------------------------------------------------------
# 2. CORE CALCULATIONS
# -----------------------------------------------------------------------------
total_pos_method = tp + fp
total_neg_method = fn + tn
total_pos_criteria = tp + fn
total_neg_criteria = fp + tn
n_total = tp + fp + fn + tn

# Core Metrics (%)
ppa = round((tp / total_pos_criteria) * 100, 2) if total_pos_criteria > 0 else 0.0
npa = round((tn / total_neg_criteria) * 100, 2) if total_neg_criteria > 0 else 0.0
opa = round(((tp + tn) / n_total) * 100, 2) if n_total > 0 else 0.0

# Constant for 95% Confidence Interval
z_sq = 3.8416

# --- 95% CONFIDENCE INTERVAL FOR PPA (Sensitivity) ---
ppa_q1 = round((2 * tp) + z_sq, 2)
if total_pos_criteria > 0:
    ppa_q2 = round(1.96 * math.sqrt(z_sq + ((4 * tp * fn) / total_pos_criteria)), 2)
    ppa_q3 = round(2 * (total_pos_criteria + z_sq), 2)
    ppa_upper = round(((ppa_q1 + ppa_q2) / ppa_q3) * 100, 2)
    ppa_lower = round(((ppa_q1 - ppa_q2) / ppa_q3) * 100, 2)
else:
    ppa_q2, ppa_q3, ppa_upper, ppa_lower = 0.0, 0.0, 0.0, 0.0

# --- 95% CONFIDENCE INTERVAL FOR NPA (Specificity) ---
npa_q1 = round((2 * tn) + z_sq, 2)
if total_neg_criteria > 0:
    npa_q2 = round(1.96 * math.sqrt(z_sq + ((4 * fp * tn) / total_neg_criteria)), 2)
    npa_q3 = round(2 * (total_neg_criteria + z_sq), 2)
    npa_upper = round(((npa_q1 + npa_q2) / npa_q3) * 100, 2)
    npa_lower = round(((npa_q1 - npa_q2) / npa_q3) * 100, 2)
else:
    npa_q2, npa_q3, npa_upper, npa_lower = 0.0, 0.0, 0.0, 0.0

# --- 95% CONFIDENCE INTERVAL FOR OPA (Accuracy) ---
opa_q1 = round(2 * (tp + tn) + z_sq, 2)
if n_total > 0:
    opa_q2 = round(1.96 * math.sqrt(z_sq + ((4 * (tp + tn) * (fp + fn)) / n_total)), 2)
    opa_q3 = round(2 * (n_total + z_sq), 2)
    opa_upper = round(((opa_q1 + opa_q2) / opa_q3) * 100, 2)
    opa_lower = round(((opa_q1 - opa_q2) / opa_q3) * 100, 2)
else:
    opa_q2, opa_q3, opa_upper, opa_lower = 0.0, 0.0, 0.0, 0.0

# -----------------------------------------------------------------------------
# 3. DISPLAY SUMMARY METRICS & CONTINGENCY TABLE
# -----------------------------------------------------------------------------
m1, m2, m3 = st.columns(3)
m1.metric("PPA", f"{ppa:.2f}%")
m2.metric("NPA", f"{npa:.2f}%")
m3.metric("OPA", f"{opa:.2f}%")

st.divider()

col_left, col_right = st.columns([1, 1])

with col_left:
    st.write("### 📋 2×2 Contingency Table")
    contingency_data = {
        "Candidate Method": ["Positive", "Negative", "Total"],
        "Known Target Condition (Pos)": [tp, fn, total_pos_criteria],
        "Known Target Condition (Neg)": [fp, tn, total_neg_criteria],
        "Total": [total_pos_method, total_neg_method, n_total]
    }
    df_matrix = pd.DataFrame(contingency_data)
    st.dataframe(df_matrix, use_container_width=True)

with col_right:
    st.write("### 📐 95% Score Confidence Intervals")
    interval_data = {
        "Parameter": ["Q1", "Q2", "Q3", "Upper Limit (%)", "Lower Limit (%)"],
        "PPA CI": [ppa_q1, ppa_q2, ppa_q3, ppa_upper, ppa_lower],
        "NPA CI": [npa_q1, npa_q2, npa_q3, npa_upper, npa_lower],
        "OPA CI": [opa_q1, opa_q2, opa_q3, opa_upper, opa_lower]
    }
    df_intervals = pd.DataFrame(interval_data)
    st.dataframe(df_intervals, use_container_width=True)

# Performance Metrics Table
st.write("### 🎯 Performance Metrics Summary")
metrics_data = {
    "Metric": ["PPA (%)", "NPA (%)", "OPA (%)"],
    "Value (%)": [ppa, npa, opa],
    "95% CI Range": [
        f"{ppa_lower:.2f}% – {ppa_upper:.2f}%",
        f"{npa_lower:.2f}% – {npa_upper:.2f}%",
        f"{opa_lower:.2f}% – {opa_upper:.2f}%"
    ]
}
df_metrics = pd.DataFrame(metrics_data)
st.dataframe(df_metrics, use_container_width=True)

# -----------------------------------------------------------------------------
# 4. EXCEL EXPORT GENERATION
# -----------------------------------------------------------------------------
output_filename = "Diagnostic_Agreement_Report.xlsx"
output_buffer = io.BytesIO()

try:
    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
        df_matrix.to_excel(writer, sheet_name='Contingency Table', index=False)
        df_metrics.to_excel(writer, sheet_name='Agreement Metrics', index=False)
        df_intervals.to_excel(writer, sheet_name='Confidence Intervals', index=False)
        
    output_buffer.seek(0)

    st.download_button(
        label=f"📥 Download Report ({output_filename})",
        data=output_buffer,
        file_name=output_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
except Exception as e:
    st.error(f"Failed to create Excel file: {e}")