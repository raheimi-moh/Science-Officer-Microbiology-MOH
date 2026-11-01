import io
from collections import Counter
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
import streamlit as st

# Configure page settings
st.set_page_config(
    page_title="AST Accuracy & DD Precision Workbench",
    layout="wide",
    initial_sidebar_state="expanded"
)
sns.set_theme(style="whitegrid", font="sans-serif")


# ==============================================================================
# 1. CORE CALCULATION ENGINES
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


def calculate_dd_precision(df_dd):
    results = []
    for iso_id, group in df_dd.groupby('Isolate ID', sort=False):
        valid = [str(v).strip().upper() for v in group['Interpretation'].dropna() if str(v).strip().upper() in ['S', 'I', 'R']]
        n_valid = len(valid)
        org = group['Organism'].iloc[0] if 'Organism' in group and len(group['Organism']) > 0 else ''
        if n_valid == 0:
            results.append({
                "Isolate ID": iso_id, "Organism": org, "Valid (n)": 0,
                "Majority": "No data", "Matches (n)": 0, "PCA (%)": 0.0, "Review": "No data"
            })
            continue
        counts = Counter(valid)
        max_f = max(counts.values())
        top = [k for k, v in counts.items() if v == max_f]
        if len(top) > 1:
            maj, matches, pca, rev = "TIE", 0, 0.0, "Review tie"
        else:
            maj = top[0]
            matches = counts[maj]
            pca = (matches / n_valid) * 100
            rev = "Concordant (100%)" if pca == 100.0 else "Review outlier(s)"
        results.append({
            "Isolate ID": iso_id, "Organism": org, "Valid (n)": n_valid,
            "Majority": maj, "Matches (n)": matches, "PCA (%)": pca, "Review": rev
        })
    df_res = pd.DataFrame(results)
    tot_m = df_res['Matches (n)'].sum() if not df_res.empty else 0
    tot_v = df_res['Valid (n)'].sum() if not df_res.empty else 0
    overall_pca = (tot_m / tot_v * 100) if tot_v > 0 else 0.0
    return df_res, overall_pca


# ==============================================================================
# 2. PLOTTING FUNCTIONS
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


def plot_dd_precision_visuals(dd_df, dd_sum, ov_pca):
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 4.5))

    dd_valid = dd_df.copy()
    dd_valid['Zone diameter (mm)'] = pd.to_numeric(dd_valid['Zone diameter (mm)'], errors='coerce')
    dd_valid = dd_valid.dropna(subset=['Zone diameter (mm)'])

    if len(dd_valid) > 0 and dd_valid['Zone diameter (mm)'].nunique() > 1:
        sns.boxplot(data=dd_valid, x='Isolate ID', y='Zone diameter (mm)', ax=ax1, palette='Blues_r', width=0.45)
        sns.stripplot(data=dd_valid, x='Isolate ID', y='Zone diameter (mm)', ax=ax1, color='#e74c3c', size=6, jitter=0.2)
        ax1.set_title("Zone Diameter (mm) Replicate Distribution & Outliers", fontsize=12, fontweight='bold', color='#1F497D')
    else:
        sns.countplot(data=dd_df, x='Isolate ID', hue='Interpretation', ax=ax1, palette='Set2')
        ax1.set_title("Interpretation Replicate Distribution", fontsize=12, fontweight='bold', color='#1F497D')

    bars = ax2.bar(dd_sum['Isolate ID'], dd_sum['PCA (%)'], color='#217346', width=0.45, edgecolor='black', linewidth=0.8)
    ax2.axhline(100.0, color='green', linestyle=':', label='100% Concordance')
    ax2.set_ylabel("PCA Agreement (%)", fontweight='bold')
    ax2.set_ylim(0, 115)
    ax2.set_title(f"Polychotomous Concordance (Overall PCA: {ov_pca:.1f}%)", fontsize=12, fontweight='bold', color='#1F497D')
    ax2.legend(loc='lower right')

    for bar in bars:
        h = bar.get_height()
        ax2.text(bar.get_x() + bar.get_width()/2., h + 1.2, f"{h:.1f}%", ha='center', va='bottom', fontweight='bold', fontsize=10)

    plt.tight_layout()
    return fig


# ==============================================================================
# 3. EXCEL EXPORTER
# ==============================================================================

def build_excel_bytes(ca_df=None, dd_df=None):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    b_font = Font(name="Calibri", size=11, bold=True)

    # 1. ACCURACY (CA) SHEET
    if ca_df is not None and len(ca_df) > 0:
        ws = wb.create_sheet(title="ACCURACY (CA)")
        ws.views.sheetView[0].showGridLines = True
        ws['B2'] = "Laboratory"; ws['D2'] = "Microbiology Lab"
        ws['B7'] = "SUMMARY ACCURACY REPORT"; ws['B7'].font = b_font
        ca = calculate_ca_summary(ca_df)
        ws['B8'] = "Total tested"; ws['D8'] = ca['total']
        ws['E8'] = "Categorical matches"; ws['F8'] = ca['matches']
        ws['G8'] = "Categorical agreement (CA)"; ws['H8'] = f"{ca['ca_rate']*100:.1f}%"
        ws['B9'] = "Discordant results"; ws['D9'] = ca['total'] - ca['matches']
        ws['G9'] = "Interpretation"; ws['H9'] = ca['status']

        ws['B12'] = "No."; ws['C12'] = "Organism / Isolate ID"; ws['E12'] = "New Method"; ws['F12'] = "Reference Method"; ws['G12'] = "Agreement"; ws['I12'] = "Error classification"
        for c in ['B', 'C', 'E', 'F', 'G', 'I']: ws[f'{c}12'].font = b_font
        r_idx = 14
        for _, row in ca_df.iterrows():
            ws[f'B{r_idx}'] = row.get('No.', r_idx - 13)
            ws[f'C{r_idx}'] = row.get('Organism', '')
            ws[f'E{r_idx}'] = row.get('New Method', '')
            ws[f'F{r_idx}'] = row.get('Reference Method', '')
            ws[f'G{r_idx}'] = "Yes" if str(row.get('New Method')).upper() == str(row.get('Reference Method')).upper() else "No"
            ws[f'I{r_idx}'] = classify_ast_error(row.get('New Method'), row.get('Reference Method'))
            r_idx += 1

        ws['B46'] = "DISCREPANCY / ERROR SUMMARY"; ws['B46'].font = b_font
        ws['B48'] = "Major error (ME)"; ws['C48'] = ca['me_count']; ws['D48'] = ca['ref_s']; ws['E48'] = f"{ca['me_rate']*100:.2f}%"; ws['F48'] = "Below 3%" if ca['me_rate'] < 0.03 else "At/above 3%"
        ws['B49'] = "Minor error (mE)"; ws['C49'] = ca['min_count']; ws['D49'] = ca['total']; ws['E49'] = f"{ca['min_rate']*100:.2f}%"; ws['F49'] = "Below 3%" if ca['min_rate'] < 0.03 else "At/above 3%"
        ws['B50'] = "Very major error (VME)"; ws['C50'] = ca['vme_count']; ws['D50'] = ca['ref_r']; ws['E50'] = f"{ca['vme_rate']*100:.2f}%"; ws['F50'] = "Below 3%" if ca['vme_rate'] < 0.03 else "At/above 3%"

    # 2. DD PRECISION SHEET
    if dd_df is not None and len(dd_df) > 0:
        ws = wb.create_sheet(title="DD Precision")
        ws.views.sheetView[0].showGridLines = True
        ws['A1'] = "DISC DIFFUSION (DD) PRECISION VERIFICATION"; ws['A1'].font = b_font
        cols = ['Isolate ID', 'Organism', 'Day', 'Replicate', 'Zone diameter (mm)', 'Interpretation']
        for ci, h in enumerate(cols, 1): ws.cell(10, ci, h).font = b_font
        for ri, r in dd_df.reset_index(drop=True).iterrows():
            for ci, h in enumerate(cols, 1): ws.cell(11 + ri, ci, r.get(h, ''))

        df_sum, overall_pca = calculate_dd_precision(dd_df)
        sr = 11 + len(dd_df) + 2
        ws.cell(sr, 1, "PRECISION SUMMARY – AUTOMATIC CALCULATION").font = b_font
        scols = ['Isolate ID', 'Organism', 'Valid zone results (n)', 'Majority interpretation', 'Matching results (n)', 'PCA (%)', 'Review']
        for ci, h in enumerate(scols, 1): ws.cell(sr + 2, ci, h).font = b_font
        for i, r in df_sum.iterrows():
            cr = sr + 3 + i
            ws.cell(cr, 1, r['Isolate ID']); ws.cell(cr, 2, r['Organism']); ws.cell(cr, 3, r['Valid (n)'])
            ws.cell(cr, 4, r['Majority']); ws.cell(cr, 5, r['Matches (n)']); ws.cell(cr, 6, f"{r['PCA (%)']:.1f}%"); ws.cell(cr, 7, r['Review'])
        ws.cell(cr + 2, 1, "Overall PCA").font = b_font
        ws.cell(cr + 2, 2, f"{overall_pca:.1f}%").font = b_font

    for sheet in wb.worksheets:
        for col in sheet.columns:
            m_len = max(len(str(c.value or '')) for c in col)
            sheet.column_dimensions[get_column_letter(col[0].column)].width = max(m_len + 3, 12)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


# ==============================================================================
# 4. STREAMLIT USER INTERFACE
# ==============================================================================

st.title("🔬 AST Accuracy & DD Precision Analysis System")
st.markdown("Analyze antimicrobial susceptibility testing accuracy (Categorical Agreement) and disc diffusion replicate precision.")

# Data input mode selection
app_mode = st.sidebar.radio("Choose Input Mode", ["Manual Entry / Paste Grid", "Upload Existing Excel File"])

if app_mode == "Upload Existing Excel File":
    st.subheader("Upload Existing AST Excel File")
    uploaded_file = st.file_uploader("Select an .xlsx or .xls file", type=['xlsx', 'xls'])

    if uploaded_file is not None:
        xls = pd.ExcelFile(uploaded_file)
        st.info(f"Available Sheets: {', '.join(xls.sheet_names)}")

        ca_df, dd_df = None, None

        if 'ACCURACY (CA)' in xls.sheet_names:
            raw = pd.read_excel(xls, sheet_name='ACCURACY (CA)', skiprows=12)
            ca_df = pd.DataFrame({
                'No.': raw.iloc[:, 1],
                'Organism': raw.iloc[:, 2],
                'New Method': raw.iloc[:, 4],
                'Reference Method': raw.iloc[:, 5]
            })
            ca_res = calculate_ca_summary(ca_df)
            st.markdown(f"### Accuracy (CA): **{ca_res['ca_rate']*100:.1f}%** ({ca_res['status']})")
            fig_ca = plot_accuracy_visuals(ca_res)
            if fig_ca:
                st.pyplot(fig_ca)

        if 'DD Precision' in xls.sheet_names:
            raw = pd.read_excel(xls, sheet_name='DD Precision', skiprows=9)
            dd_df = raw.iloc[:, 0:6]
            dd_df.columns = ['Isolate ID', 'Organism', 'Day', 'Replicate', 'Zone diameter (mm)', 'Interpretation']
            dd_sum, dd_overall = calculate_dd_precision(dd_df)
            st.markdown(f"### DD Precision: Overall PCA = **{dd_overall:.1f}%**")
            fig_dd = plot_dd_precision_visuals(dd_df, dd_sum, dd_overall)
            st.pyplot(fig_dd)
            st.dataframe(dd_sum, use_container_width=True)

        # Export file
        excel_bytes = build_excel_bytes(ca_df, dd_df)
        st.download_button(
            label="📥 Download Calculated Report (.xlsx)",
            data=excel_bytes,
            file_name="CALCULATED_AST_REPORT.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    target_sheet = st.sidebar.selectbox("Target Analysis Type", ['ACCURACY (CA)', 'DD Precision'])

    # Sample defaults
    if target_sheet == 'ACCURACY (CA)':
        default_data = pd.DataFrame([
            {'Organism': 'E.coli', 'New Method': 'S', 'Reference Method': 'S'},
            {'Organism': 'Klebsiella pneumoniae', 'New Method': 'S', 'Reference Method': 'S'},
            {'Organism': 'Providencia sp.', 'New Method': 'R', 'Reference Method': 'R'},
            {'Organism': 'Proteus mirabilis', 'New Method': 'I', 'Reference Method': 'I'},
            {'Organism': 'Proteus vulgaris', 'New Method': 'S', 'Reference Method': 'S'},
            {'Organism': 'Klebsiella oxytoca', 'New Method': 'S', 'Reference Method': 'S'},
            {'Organism': 'Enterobacter cloacae', 'New Method': 'S', 'Reference Method': 'S'},
            {'Organism': 'Citrobacter freundii', 'New Method': 'R', 'Reference Method': 'R'}
        ])

        with st.expander("Paste Data (Optional tab-separated input)"):
            raw_text = st.text_area(
                "Paste TSV data (Organism, New Method, Reference Method):",
                placeholder="E.coli\tS\tS\nKlebsiella pneumoniae\tS\tS"
            )
            if st.button("Load Pasted Text into Grid"):
                lines = [line.split('\t') if '\t' in line else line.split(',') for line in raw_text.strip().splitlines() if line.strip()]
                parsed_rows = []
                for row in lines:
                    parsed_rows.append({
                        'Organism': row[0].strip() if len(row) > 0 else '',
                        'New Method': row[1].strip() if len(row) > 1 else '',
                        'Reference Method': row[2].strip() if len(row) > 2 else ''
                    })
                if parsed_rows:
                    st.session_state['ca_grid_data'] = pd.DataFrame(parsed_rows)

        if 'ca_grid_data' not in st.session_state:
            st.session_state['ca_grid_data'] = default_data

        st.markdown("#### Interactive Data Grid")
        st.caption("You can edit cells, select rows, and add new rows below:")
        edited_ca_df = st.data_editor(st.session_state['ca_grid_data'], num_rows="dynamic", use_container_width=True)

        if st.button("Calculate & Visualize Accuracy", type="primary"):
            ca_df_calc = edited_ca_df.copy()
            ca_df_calc.insert(0, 'No.', range(1, len(ca_df_calc) + 1))
            res = calculate_ca_summary(ca_df_calc)

            if res['total'] > 0:
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("Categorical Agreement", f"{res['ca_rate']*100:.1f}%", res['status'])
                col2.metric("Major Errors (ME)", f"{res['me_count']} ({res['me_rate']*100:.1f}%)")
                col3.metric("Minor Errors (mE)", f"{res['min_count']} ({res['min_rate']*100:.1f}%)")
                col4.metric("Very Major Errors (VME)", f"{res['vme_count']} ({res['vme_rate']*100:.1f}%)")

                fig = plot_accuracy_visuals(res)
                if fig:
                    st.pyplot(fig)

                out_excel = build_excel_bytes(ca_df=ca_df_calc)
                st.download_button(
                    label="📥 Download Accuracy Report (.xlsx)",
                    data=out_excel,
                    file_name="CALCULATED_ACCURACY_REPORT.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("Please ensure there is valid data in 'New Method' and 'Reference Method' columns.")

    elif target_sheet == 'DD Precision':
        rep_cols = [f"R{i}" for i in range(1, 16)]
        columns = ['Isolate ID', 'Organism', 'Zone (mm)'] + rep_cols

        default_dd = pd.DataFrame([
            ['ISO-01', 'E.coli 25922', 25] + ['S']*14 + ['I'],
            ['ISO-02', 'S.aureus 25923', 18] + ['S']*15,
            ['ISO-03', 'P.aeruginosa 27853', 22] + ['S']*15
        ], columns=columns)

        if 'dd_grid_data' not in st.session_state:
            st.session_state['dd_grid_data'] = default_dd

        st.markdown("#### Interactive Replicate Grid")
        st.caption("Enter 15 replicate determinations (S/I/R) per isolate:")
        edited_dd_df = st.data_editor(st.session_state['dd_grid_data'], num_rows="dynamic", use_container_width=True)

        if st.button("Calculate & Visualize DD Precision", type="primary"):
            long_rows = []
            for _, r in edited_dd_df.iterrows():
                iso = str(r.get('Isolate ID', 'ISO')).strip()
                org = str(r.get('Organism', '')).strip()
                zone = r.get('Zone (mm)', 20)
                replicate_vals = [r.get(col) for col in rep_cols]
                for ri, cat in enumerate(replicate_vals):
                    val = str(cat).strip().upper() if pd.notna(cat) else ''
                    if val in ['S', 'I', 'R']:
                        long_rows.append({
                            'Isolate ID': iso,
                            'Organism': org,
                            'Day': (ri // 3) + 1,
                            'Replicate': (ri % 3) + 1,
                            'Zone diameter (mm)': zone,
                            'Interpretation': val
                        })

            if long_rows:
                dd_df_calc = pd.DataFrame(long_rows)
                dd_sum, ov_pca = calculate_dd_precision(dd_df_calc)

                st.metric("Overall Polychotomous Concordance (PCA)", f"{ov_pca:.1f}%")
                fig = plot_dd_precision_visuals(dd_df_calc, dd_sum, ov_pca)
                st.pyplot(fig)
                st.dataframe(dd_sum, use_container_width=True)

                out_excel = build_excel_bytes(dd_df=dd_df_calc)
                st.download_button(
                    label="📥 Download DD Precision Report (.xlsx)",
                    data=out_excel,
                    file_name="CALCULATED_DD_PRECISION.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("No valid S/I/R replicates found in the table.")