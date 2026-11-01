import io
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
import streamlit as st

st.set_page_config(
    page_title="AST Accuracy (CA) Workbench",
    page_icon="🔬",
    layout="wide"
)
sns.set_theme(style="whitegrid", font="sans-serif")


# ==============================================================================
# 1. CORE CALCULATION ENGINE
# ==============================================================================

def classify_ast_error(new_res, ref_res):
    if not new_res or not ref_res:
        return "None"
    n, r = str(new_res).strip().upper(), str(ref_res).strip().upper()
    if n == r:
        return "None"
    if n == 'S' and r == 'R':
        return "Very Major"
    if n == 'R' and r == 'S':
        return "Major"
    if (n == 'I' and r in ['S', 'R']) or (r == 'I' and n in ['S', 'R']):
        return "Minor"
    return "None"


def calculate_ca_summary(df_ca):
    clean = df_ca.dropna(subset=['New Method', 'Reference Method']).copy()
    clean = clean[(clean['New Method'] != '') & (clean['Reference Method'] != '')]
    total = len(clean)
    if total == 0:
        return {
            "total": 0, "matches": 0, "ca_rate": 0.0, "status": "No data", "df": clean,
            "me_count": 0, "ref_s": 0, "me_rate": 0.0, "vme_count": 0, "ref_r": 0,
            "vme_rate": 0.0, "min_count": 0, "min_rate": 0.0
        }
    clean['Agreement'] = clean.apply(
        lambda r: 'Yes' if str(r['New Method']).strip().upper() == str(r['Reference Method']).strip().upper() else 'No',
        axis=1
    )
    clean['Error'] = clean.apply(lambda r: classify_ast_error(r['New Method'], r['Reference Method']), axis=1)

    matches = (clean['Agreement'] == 'Yes').sum()
    ca_rate = matches / total
    ref_s = (clean['Reference Method'].astype(str).str.upper() == 'S').sum()
    ref_r = (clean['Reference Method'].astype(str).str.upper() == 'R').sum()
    me_count = (clean['Error'] == 'Major').sum()
    vme_count = (clean['Error'] == 'Very Major').sum()
    min_count = (clean['Error'] == 'Minor').sum()

    return {
        "df": clean, "total": total, "matches": matches, "ca_rate": ca_rate,
        "status": "Meets 90% benchmark*" if ca_rate >= 0.90 else "Below 90% benchmark*",
        "me_count": me_count, "ref_s": ref_s, "me_rate": (me_count / ref_s) if ref_s > 0 else 0.0,
        "vme_count": vme_count, "ref_r": ref_r, "vme_rate": (vme_count / ref_r) if ref_r > 0 else 0.0,
        "min_count": min_count, "min_rate": min_count / total
    }


# ==============================================================================
# 2. PLOTTING
# ==============================================================================

def plot_accuracy_visuals(res):
    df = res['df']
    if len(df) == 0:
        return None
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(13, 4.5))

    # Heatmap
    cats = ['S', 'I', 'R']
    cm = pd.crosstab(df['Reference Method'], df['New Method']).reindex(index=cats, columns=cats, fill_value=0)
    sns.heatmap(cm, annot=True, fmt='d', cmap='Blues', cbar=False, ax=ax1, annot_kws={"size": 13, "weight": "bold"})
    ax1.set_title(f"Categorical Agreement: {res['ca_rate']*100:.1f}%\n({res['status']})", fontsize=12, fontweight='bold', color='#1F497D')
    ax1.set_xlabel("New Method Interpretation", fontweight='bold')
    ax1.set_ylabel("Reference Method Interpretation", fontweight='bold')

    # Error bar chart
    err_labels = ['Major (ME)', 'Minor (mE)', 'Very Major (VME)']
    rates = [res['me_rate'] * 100, res['min_rate'] * 100, res['vme_rate'] * 100]
    colors = ['#f39c12', '#3498db', '#e74c3c']
    bars = ax2.bar(err_labels, rates, color=colors, width=0.5, edgecolor='black', linewidth=0.8)
    ax2.axhline(3.0, color='red', linestyle='--', linewidth=1.5, label='Benchmark Max (< 3%)')
    ax2.set_ylabel("Discrepancy Rate (%)", fontweight='bold')
    ax2.set_title("Error Rates vs Benchmark (<3%)", fontsize=12, fontweight='bold', color='#1F497D')
    ax2.set_ylim(0, max(max(rates) + 2, 5))
    ax2.legend(loc='upper right')

    for bar in bars:
        h = bar.get_height()
        ax2.text(bar.get_x() + bar.get_width()/2., h + 0.15, f"{h:.2f}%", ha='center', va='bottom', fontweight='bold', fontsize=11)

    plt.tight_layout()
    return fig


# ==============================================================================
# 3. EXCEL EXPORTER
# ==============================================================================

def export_accuracy_excel(ca_df):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ACCURACY (CA)"
    b_font = Font(name="Calibri", size=11, bold=True)
    ws.views.sheetView[0].showGridLines = True

    ws['B2'] = "Laboratory"
    ws['D2'] = "Microbiology Lab"
    ws['B7'] = "SUMMARY ACCURACY REPORT"
    ws['B7'].font = b_font

    ca = calculate_ca_summary(ca_df)
    ws['B8'] = "Total tested"; ws['D8'] = ca['total']
    ws['E8'] = "Categorical matches"; ws['F8'] = ca['matches']
    ws['G8'] = "Categorical agreement (CA)"; ws['H8'] = f"{ca['ca_rate']*100:.1f}%"
    ws['B9'] = "Discordant results"; ws['D9'] = ca['total'] - ca['matches']
    ws['G9'] = "Interpretation"; ws['H9'] = ca['status']

    headers = [('B12', "No."), ('C12', "Organism / Isolate ID"), ('E12', "New Method"), 
               ('F12', "Reference Method"), ('G12', "Agreement"), ('I12', "Error classification")]
    for cell_ref, text in headers:
        ws[cell_ref] = text
        ws[cell_ref].font = b_font

    r_idx = 14
    for _, row in ca_df.iterrows():
        ws[f'B{r_idx}'] = row.get('No.', r_idx - 13)
        ws[f'C{r_idx}'] = row.get('Organism', '')
        ws[f'E{r_idx}'] = row.get('New Method', '')
        ws[f'F{r_idx}'] = row.get('Reference Method', '')
        ws[f'G{r_idx}'] = "Yes" if str(row.get('New Method')).upper() == str(row.get('Reference Method')).upper() else "No"
        ws[f'I{r_idx}'] = classify_ast_error(row.get('New Method'), row.get('Reference Method'))
        r_idx += 1

    summary_start = max(r_idx + 2, 46)
    ws[f'B{summary_start}'] = "DISCREPANCY / ERROR SUMMARY"
    ws[f'B{summary_start}'].font = b_font

    ws[f'B{summary_start+2}'] = "Major error (ME)"
    ws[f'C{summary_start+2}'] = ca['me_count']
    ws[f'D{summary_start+2}'] = ca['ref_s']
    ws[f'E{summary_start+2}'] = f"{ca['me_rate']*100:.2f}%"
    ws[f'F{summary_start+2}'] = "Below 3%" if ca['me_rate'] < 0.03 else "At/above 3%"

    ws[f'B{summary_start+3}'] = "Minor error (mE)"
    ws[f'C{summary_start+3}'] = ca['min_count']
    ws[f'D{summary_start+3}'] = ca['total']
    ws[f'E{summary_start+3}'] = f"{ca['min_rate']*100:.2f}%"
    ws[f'F{summary_start+3}'] = "Below 3%" if ca['min_rate'] < 0.03 else "At/above 3%"

    ws[f'B{summary_start+4}'] = "Very major error (VME)"
    ws[f'C{summary_start+4}'] = ca['vme_count']
    ws[f'D{summary_start+4}'] = ca['ref_r']
    ws[f'E{summary_start+4}'] = f"{ca['vme_rate']*100:.2f}%"
    ws[f'F{summary_start+4}'] = "Below 3%" if ca['vme_rate'] < 0.03 else "At/above 3%"

    for col in ws.columns:
        m_len = max(len(str(c.value or '')) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = max(m_len + 3, 12)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


# ==============================================================================
# 4. STREAMLIT UI
# ==============================================================================

st.title("🔬 AST Categorical Agreement (CA) Calculator")
st.markdown("Evaluate agreement and discrepancy rates (VME, ME, mE) between your New AST Method and Reference Method.")

input_method = st.radio("Choose Input Mode", ["Manual Entry / Paste TSV", "Upload Excel File"], horizontal=True)

ca_df_to_analyze = None

if input_method == "Upload Excel File":
    uploaded_file = st.file_uploader("Upload an Excel file (.xlsx, .xls)", type=['xlsx', 'xls'])
    if uploaded_file is not None:
        xls = pd.ExcelFile(uploaded_file)
        target_sheet = None
        for s in xls.sheet_names:
            if "ACCURACY" in s.upper():
                target_sheet = s
                break
        
        if target_sheet:
            raw = pd.read_excel(xls, sheet_name=target_sheet, skiprows=12)
            ca_df_to_analyze = pd.DataFrame({
                'No.': raw.iloc[:, 1],
                'Organism': raw.iloc[:, 2],
                'New Method': raw.iloc[:, 4],
                'Reference Method': raw.iloc[:, 5]
            }).dropna(how='all')
            st.success(f"Loaded '{target_sheet}' successfully.")
        else:
            raw = pd.read_excel(xls)
            st.info("No formatted 'ACCURACY (CA)' sheet detected. Using first sheet columns.")
            st.dataframe(raw.head(3))
            cols = list(raw.columns)
            c1, c2, c3 = st.columns(3)
            org_c = c1.selectbox("Organism / Isolate column", cols, index=0)
            new_c = c2.selectbox("New Method column", cols, index=min(1, len(cols)-1))
            ref_c = c3.selectbox("Reference Method column", cols, index=min(2, len(cols)-1))
            
            ca_df_to_analyze = pd.DataFrame({
                'No.': range(1, len(raw) + 1),
                'Organism': raw[org_c],
                'New Method': raw[new_c],
                'Reference Method': raw[ref_c]
            })

else:
    default_data = pd.DataFrame([
        {'Organism': 'E.coli', 'New Method': 'S', 'Reference Method': 'S'},
        {'Klebsiella pneumoniae': 'S', 'New Method': 'S', 'Reference Method': 'S'},
        {'Organism': 'Providencia sp.', 'New Method': 'R', 'Reference Method': 'R'},
        {'Organism': 'Proteus mirabilis', 'New Method': 'I', 'Reference Method': 'I'},
        {'Organism': 'Proteus vulgaris', 'New Method': 'S', 'Reference Method': 'S'},
        {'Organism': 'Klebsiella oxytoca', 'New Method': 'S', 'Reference Method': 'S'},
        {'Organism': 'Enterobacter cloacae', 'New Method': 'S', 'Reference Method': 'S'},
        {'Organism': 'Citrobacter freundii', 'New Method': 'R', 'Reference Method': 'R'}
    ])

    with st.expander("Paste Data (Optional tab-separated input)"):
        raw_text = st.text_area(
            "Paste TSV lines (Organism, New Method, Reference Method):",
            placeholder="E.coli\tS\tS\nKlebsiella pneumoniae\tS\tS"
        )
        if st.button("Load Pasted Text into Grid"):
            lines = [ln.split('\t') if '\t' in ln else ln.split(',') for ln in raw_text.strip().splitlines() if ln.strip()]
            parsed = [{'Organism': r[0].strip(), 'New Method': r[1].strip() if len(r)>1 else '', 'Reference Method': r[2].strip() if len(r)>2 else ''} for r in lines if len(r)>0]
            if parsed:
                st.session_state['ca_grid'] = pd.DataFrame(parsed)

    if 'ca_grid' not in st.session_state:
        st.session_state['ca_grid'] = default_data

    st.markdown("#### Interactive Accuracy Data Editor")
    edited_df = st.data_editor(st.session_state['ca_grid'], num_rows="dynamic", use_container_width=True)
    
    ca_df_to_analyze = edited_df.copy()
    ca_df_to_analyze.insert(0, 'No.', range(1, len(ca_df_to_analyze) + 1))

# Run Calculation & Render Results
if ca_df_to_analyze is not None:
    if st.button("Calculate & Generate Report", type="primary"):
        res = calculate_ca_summary(ca_df_to_analyze)

        if res['total'] > 0:
            st.divider()
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Categorical Agreement", f"{res['ca_rate']*100:.1f}%", res['status'])
            col2.metric("Major Errors (ME)", f"{res['me_count']} ({res['me_rate']*100:.2f}%)")
            col3.metric("Minor Errors (mE)", f"{res['min_count']} ({res['min_rate']*100:.2f}%)")
            col4.metric("Very Major Errors (VME)", f"{res['vme_count']} ({res['vme_rate']*100:.2f}%)")

            fig = plot_accuracy_visuals(res)
            if fig:
                st.pyplot(fig)

            st.markdown("#### Evaluated Data Table")
            st.dataframe(res['df'], use_container_width=True)

            excel_data = export_accuracy_excel(res['df'])
            st.download_button(
                label="📥 Download Accuracy Verification Report (.xlsx)",
                data=excel_data,
                file_name="ACCURACY_CA_REPORT.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No valid S/I/R data detected. Please ensure 'New Method' and 'Reference Method' columns are filled.")