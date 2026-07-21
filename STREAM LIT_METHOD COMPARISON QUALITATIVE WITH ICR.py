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
    "Enter your $2 \\times 2$ contingency table values to compute Sensitivity, "
    "Specificity, PPV, NPV, Accuracy, and 95% Score Confidence Intervals."
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
sensitivity = round((tp / total_pos_criteria) * 100, 2) if total_pos_criteria > 0 else 0.0
specificity = round((tn / total_neg_criteria) * 100, 2) if total_neg_criteria > 0 else 0.0
ppv = round((tp / total_pos_method) * 100, 2) if total_pos_method > 0 else 0.0
npv = round((tn / total_neg_method) * 100, 2) if total_neg_method > 0 else 0.0
accuracy = round(((tp + tn) / n_total) * 100, 2) if n_total > 0 else 0.0

# Constant for 95% Confidence Interval
z_sq = 3.8416

# --- 95% CONFIDENCE INTERVAL FOR SENSITIVITY ---
sens_q1 = round((2 * tp) + z_sq, 2)
if total_pos_criteria > 0:
    sens_q2 = round(1.96 * math.sqrt(z_sq + ((4 * tp * fn) / total_pos_criteria)), 2)
    sens_q3 = round(2 * (total_pos_criteria + z_sq), 2)
    sens_upper = round(((sens_q1 + sens_q2) / sens_q3) * 100, 2)
    sens_lower = round(((sens_q1 - sens_q2) / sens_q3) * 100, 2)
else:
    sens_q2, sens_q3, sens_upper, sens_lower = 0.0, 0.0, 0.0, 0.0

# --- 95% CONFIDENCE INTERVAL FOR SPECIFICITY ---
spec_q1 = round((2 * tn) + z_sq, 2)
if total_neg_criteria > 0:
    spec_q2 = round(1.96 * math.sqrt(z_sq + ((4 * fp * tn) / total_neg_criteria)), 2)
    spec_q3 = round(2 * (total_neg_criteria + z_sq), 2)
    spec_upper = round(((spec_q1 + spec_q2) / spec_q3) * 100, 2)
    spec_lower = round(((spec_q1 - spec_q2) / spec_q3) * 100, 2)
else:
    spec_q2, spec_q3, spec_upper, spec_lower = 0.0, 0.0, 0.0, 0.0

# -----------------------------------------------------------------------------
# 3. DISPLAY SUMMARY METRICS & CONTINGENCY TABLE
# -----------------------------------------------------------------------------
m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("Sensitivity", f"{sensitivity:.2f}%")
m2.metric("Specificity", f"{specificity:.2f}%")
m3.metric("PPV", f"{ppv:.2f}%")
m4.metric("NPV", f"{npv:.2f}%")
m5.metric("Accuracy", f"{accuracy:.2f}%")

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
        "Sensitivity CI": [sens_q1, sens_q2, sens_q3, sens_upper, sens_lower],
        "Specificity CI": [spec_q1, spec_q2, spec_q3, spec_upper, spec_lower]
    }
    df_intervals = pd.DataFrame(interval_data)
    st.dataframe(df_intervals, use_container_width=True)

# Performance Metrics Table
st.write("### 🎯 Performance Metrics Summary")
metrics_data = {
    "Metric": ["Sensitivity (%)", "Specificity (%)", "PPV (%)", "NPV (%)", "Accuracy (%)"],
    "Value (%)": [sensitivity, specificity, ppv, npv, accuracy],
    "95% CI Range": [
        f"{sens_lower:.2f}% – {sens_upper:.2f}%",
        f"{spec_lower:.2f}% – {spec_upper:.2f}%",
        "N/A",
        "N/A",
        "N/A"
    ]
}
df_metrics = pd.DataFrame(metrics_data)
st.dataframe(df_metrics, use_container_width=True)

# -----------------------------------------------------------------------------
# 4. EXCEL EXPORT GENERATION
# -----------------------------------------------------------------------------
output_filename = "Diagnostic_Accuracy_Report.xlsx"
output_buffer = io.BytesIO()

try:
    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
        df_matrix.to_excel(writer, sheet_name='Contingency Table', index=False)
        df_metrics.to_excel(writer, sheet_name='Accuracy Metrics', index=False)
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