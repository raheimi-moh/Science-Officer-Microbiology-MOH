import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
import streamlit as st

st.set_page_config(
    page_title="DD Precision Verification",
    page_icon="🧫",
    layout="wide"
)
sns.set_theme(style="whitegrid", font="sans-serif")


# ==============================================================================
# 1. CORE CALCULATION ENGINE
# ==============================================================================

def calculate_precision(df_input):
    """Computes majority interpretation, match status, and precision summaries."""
    df = df_input.copy()

    # Standardize column naming
    df.columns = [str(c).strip() for c in df.columns]

    # Normalize interpretations and zone numbers
    if 'Interpretation (S/I/R)' in df.columns:
        df['Interpretation (S/I/R)'] = (
            df['Interpretation (S/I/R)'].astype(str).str.strip().str.upper()
        )
    else:
        df['Interpretation (S/I/R)'] = ''

    if 'Zone diameter (mm)' in df.columns:
        df['Zone diameter (mm)'] = pd.to_numeric(
            df['Zone diameter (mm)'], errors='coerce'
        )
    else:
        df['Zone diameter (mm)'] = np.nan

    # Compute within-isolate majority interpretation
    majority_dict = {}
    for isolate, group in df.groupby('Isolate ID', sort=False):
        valid_interps = group['Interpretation (S/I/R)'].dropna()
        valid_interps = valid_interps[valid_interps.isin(['S', 'I', 'R'])]
        counts = valid_interps.value_counts()

        if counts.empty:
            majority_dict[isolate] = ''
        else:
            max_val = counts.max()
            top_cats = counts[counts == max_val].index.tolist()
            majority_dict[isolate] = top_cats[0] if len(top_cats) == 1 else 'TIE'

    df['Majority interpretation*'] = df['Isolate ID'].map(majority_dict)

    def match_func(row):
        interp = row['Interpretation (S/I/R)']
        maj = row['Majority interpretation*']
        if pd.isna(interp) or interp not in ['S', 'I', 'R']:
            return ''
        if maj == 'TIE':
            return 'Review tie'
        return 'Yes' if interp == maj else 'No'

    df['Match majority?'] = df.apply(match_func, axis=1)

    # Generate Precision Summary Table
    summary_records = []
    for isolate, group in df.groupby('Isolate ID', sort=False):
        org = group['Organism'].iloc[0] if 'Organism' in group.columns else ''
        valid_zones = group['Zone diameter (mm)'].dropna().count()
        valid_interps = group['Interpretation (S/I/R)'][
            group['Interpretation (S/I/R)'].isin(['S', 'I', 'R'])
        ]
        tot_valid = len(valid_interps)
        maj = majority_dict.get(isolate, '')

        if maj in ['S', 'I', 'R']:
            match_n = (valid_interps == maj).sum()
        else:
            match_n = 0

        pca = (match_n / tot_valid * 100) if tot_valid > 0 else 0.0

        if tot_valid == 0:
            review = 'No data'
        elif maj == 'TIE':
            review = 'Review tie'
        elif pca == 100.0:
            review = 'Concordant (100%)'
        else:
            review = 'Review outlier(s)'

        summary_records.append({
            'Isolate ID': isolate,
            'Organism': org,
            'Valid zone results (n)': int(valid_zones),
            'Majority interpretation': maj,
            'Matching results (n)': int(match_n),
            'PCA (%)': round(pca, 1),
            'Review': review,
        })

    summary_df = pd.DataFrame(summary_records)

    total_matches = summary_df['Matching results (n)'].sum() if not summary_df.empty else 0
    total_evals = df['Interpretation (S/I/R)'].isin(['S', 'I', 'R']).sum()
    overall_pca = (
        round((total_matches / total_evals) * 100, 1) if total_evals > 0 else 0.0
    )

    return df, summary_df, overall_pca


# ==============================================================================
# 2. EXCEL GENERATOR (IN-MEMORY BYTESIO)
# ==============================================================================

def export_to_excel_buffer(df_calc, df_sum, overall_pca):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'DD Precision'
    ws.views.sheetView[0].showGridLines = True

    header_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
    section_fill = PatternFill(start_color='2C3E50', end_color='2C3E50', fill_type='solid')
    summary_hdr_fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
    row_stripe = PatternFill(start_color='F9FAFB', end_color='F9FAFB', fill_type='solid')
    bold_font = Font(name='Segoe UI', size=10, bold=True, color='000000')
    regular_font = Font(name='Segoe UI', size=10, color='000000')

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9'),
    )

    # Title Header
    ws.merge_cells('A1:H1')
    ws['A1'] = 'DISC DIFFUSION (DD) PRECISION VERIFICATION – ONE ANTIMICROBIAL'
    ws['A1'].font = Font(name='Segoe UI', size=13, bold=True, color='FFFFFF')
    ws['A1'].fill = header_fill
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 28

    ws.merge_cells('A2:H2')
    ws['A2'] = 'Record zone diameter (mm) and categorical interpretation (S/I/R) for 3 replicates per day over 5 days.'
    ws['A2'].font = Font(name='Segoe UI', size=9, italic=True, color='595959')
    ws['A2'].alignment = Alignment(horizontal='center', vertical='center')

    # Metadata section
    meta_fields = [
        ('Laboratory', 'Organism group'),
        ('Study date', 'Antimicrobial agent'),
        ('AST system / method', 'Medium / disk potency'),
    ]
    for idx, (f1, f2) in enumerate(meta_fields, start=4):
        ws[f'A{idx}'] = f1
        ws[f'A{idx}'].font = bold_font
        ws[f'E{idx}'] = f2
        ws[f'E{idx}'].font = bold_font

    # Data Entry Section Header
    ws.merge_cells('A8:H8')
    ws['A8'] = 'DATA ENTRY: Enter 5 isolates. For each isolate, enter 15 results (Day 1–5; 3 replicates/day).'
    ws['A8'].font = Font(name='Segoe UI', size=10, bold=True, color='FFFFFF')
    ws['A8'].fill = section_fill
    ws['A8'].alignment = Alignment(vertical='center')
    ws.row_dimensions[8].height = 22

    headers = [
        'Isolate ID', 'Organism', 'Day', 'Replicate',
        'Zone diameter (mm)', 'Interpretation (S/I/R)',
        'Majority interpretation*', 'Match majority?'
    ]
    for col_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=10, column=col_idx, value=h)
        cell.font = bold_font
        cell.fill = summary_hdr_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border
    ws.row_dimensions[10].height = 24

    current_row = 11
    for _, row in df_calc.iterrows():
        vals = [
            row.get('Isolate ID', ''),
            row.get('Organism', ''),
            row.get('Day', ''),
            row.get('Replicate', ''),
            row.get('Zone diameter (mm)', ''),
            row.get('Interpretation (S/I/R)', ''),
            row.get('Majority interpretation*', ''),
            row.get('Match majority?', '')
        ]
        for col_idx, val in enumerate(vals, start=1):
            cell = ws.cell(row=current_row, column=col_idx, value=val)
            cell.font = regular_font
            cell.border = thin_border
            if col_idx in [3, 4, 5]:
                cell.alignment = Alignment(horizontal='right', vertical='center')
            elif col_idx in [1, 6, 7, 8]:
                cell.alignment = Alignment(horizontal='center', vertical='center')

            if col_idx == 8 and str(val).lower() == 'no':
                cell.fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
                cell.font = Font(name='Segoe UI', size=10, color='9C0006', bold=True)
            elif current_row % 2 == 0:
                cell.fill = row_stripe

        current_row += 1

    # Precision Summary Section
    current_row += 2
    ws.merge_cells(f'A{current_row}:H{current_row}')
    ws[f'A{current_row}'] = 'PRECISION SUMMARY – AUTOMATIC CALCULATION'
    ws[f'A{current_row}'].font = Font(name='Segoe UI', size=11, bold=True, color='FFFFFF')
    ws[f'A{current_row}'].fill = section_fill
    ws.row_dimensions[current_row].height = 22

    current_row += 1
    sum_headers = [
        'Isolate ID', 'Organism', 'Valid zone results (n)',
        'Majority interpretation', 'Matching results (n)', 'PCA (%)', 'Review'
    ]
    for col_idx, h in enumerate(sum_headers, start=1):
        cell = ws.cell(row=current_row, column=col_idx, value=h)
        cell.font = bold_font
        cell.fill = summary_hdr_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border
    ws.row_dimensions[current_row].height = 22

    current_row += 1
    for _, row in df_sum.iterrows():
        vals = [
            row['Isolate ID'],
            row['Organism'],
            row['Valid zone results (n)'],
            row['Majority interpretation'],
            row['Matching results (n)'],
            f"{row['PCA (%)']}%",
            row['Review']
        ]
        for col_idx, val in enumerate(vals, start=1):
            cell = ws.cell(row=current_row, column=col_idx, value=val)
            cell.font = regular_font
            cell.border = thin_border
            if col_idx in [3, 5]:
                cell.alignment = Alignment(horizontal='right', vertical='center')
            elif col_idx in [1, 4, 6]:
                cell.alignment = Alignment(horizontal='center', vertical='center')

            if col_idx == 7:
                if '100%' in str(val):
                    cell.font = Font(name='Segoe UI', size=10, color='0E6251', bold=True)
                else:
                    cell.font = Font(name='Segoe UI', size=10, color='C0392B', bold=True)
        current_row += 1

    ws.cell(row=current_row, column=1, value='Overall PCA').font = bold_font
    overall_cell = ws.cell(row=current_row, column=2, value=f'{round(overall_pca, 1)}%')
    overall_cell.font = Font(name='Segoe UI', size=11, bold=True, color='1F4E78')
    ws.row_dimensions[current_row].height = 22

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 14)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ==============================================================================
# 3. VISUALIZATION ENGINE
# ==============================================================================

def generate_visualizations(df_calc, df_sum, overall_pca):
    fig, axes = plt.subplots(1, 2, figsize=(15, 5))

    # Bar Chart: PCA (%)
    palette = [
        '#2ECC71' if p == 100 else '#E67E22' if p >= 90 else '#E74C3C'
        for p in df_sum['PCA (%)']
    ]
    bars = axes[0].bar(
        df_sum['Isolate ID'],
        df_sum['PCA (%)'],
        color=palette,
        edgecolor='black',
        linewidth=0.8,
        width=0.55,
    )
    axes[0].axhline(
        overall_pca,
        color='#2980B9',
        linestyle='--',
        linewidth=1.8,
        label=f'Overall PCA: {overall_pca}%',
    )
    axes[0].set_ylim(0, 115)
    axes[0].set_title('Precision Categorical Agreement (PCA %)', fontsize=12, fontweight='bold', color='#1F4E78')
    axes[0].set_ylabel('Agreement with Majority (%)', fontsize=10, fontweight='bold')
    axes[0].set_xlabel('Isolate ID', fontsize=10, fontweight='bold')
    axes[0].legend(loc='lower right')

    for bar in bars:
        h = bar.get_height()
        axes[0].text(
            bar.get_x() + bar.get_width() / 2.0,
            h + 2,
            f'{h:.1f}%',
            ha='center',
            va='bottom',
            fontsize=9,
            fontweight='bold',
        )

    # Boxplot: Zone Diameter Distributions
    valid_zones = df_calc.dropna(subset=['Zone diameter (mm)'])
    if not valid_zones.empty and valid_zones['Zone diameter (mm)'].nunique() > 1:
        sns.boxplot(
            data=valid_zones,
            x='Isolate ID',
            y='Zone diameter (mm)',
            ax=axes[1],
            palette='Blues_r',
            showmeans=True,
            meanprops={'marker': 'o', 'markerfacecolor': 'red', 'markeredgecolor': 'black'},
        )
        sns.stripplot(
            data=valid_zones,
            x='Isolate ID',
            y='Zone diameter (mm)',
            ax=axes[1],
            color='black',
            alpha=0.45,
            jitter=0.2,
        )
        axes[1].set_title('Zone Diameter Dispersion (15 Replicates)', fontsize=12, fontweight='bold', color='#1F4E78')
    else:
        sns.countplot(data=df_calc, x='Isolate ID', hue='Interpretation (S/I/R)', ax=axes[1], palette='Set2')
        axes[1].set_title('Interpretation Replicate Distribution', fontsize=12, fontweight='bold', color='#1F4E78')

    axes[1].set_ylabel('Zone Diameter (mm)', fontsize=10, fontweight='bold')
    axes[1].set_xlabel('Isolate ID', fontsize=10, fontweight='bold')

    plt.tight_layout()
    return fig


# ==============================================================================
# 4. STREAMLIT APPLICATION LAYOUT
# ==============================================================================

st.title("🧫 Disc Diffusion (DD) Precision Verification Tool")
st.caption("Verify precision, compute within-isolate majority concordances (PCA %), inspect zone variance, and generate validated reports.")

# Default dataset
SAMPLE_DATA = """Isolate ID\tOrganism\tDay\tReplicate\tZone diameter (mm)\tInterpretation (S/I/R)
ISO-01\tE.coli 25922\t1\t1\t23\tS
ISO-01\tE.coli 25922\t1\t2\t25\tS
ISO-01\tE.coli 25922\t1\t3\t20\tS
ISO-01\tE.coli 25922\t2\t1\t25\tS
ISO-01\tE.coli 25922\t2\t2\t25\tS
ISO-01\tE.coli 25922\t2\t3\t25\tS
ISO-01\tE.coli 25922\t3\t1\t26\tS
ISO-01\tE.coli 25922\t3\t2\t25\tS
ISO-01\tE.coli 25922\t3\t3\t25\tS
ISO-01\tE.coli 25922\t4\t1\t32\tS
ISO-01\tE.coli 25922\t4\t2\t25\tS
ISO-01\tE.coli 25922\t4\t3\t25\tS
ISO-01\tE.coli 25922\t5\t1\t25\tS
ISO-01\tE.coli 25922\t5\t2\t30\tS
ISO-01\tE.coli 25922\t5\t3\t25\tS
ISO-02\tStaph. aureus 25923\t1\t1\t18\tS
ISO-02\tStaph. aureus 25923\t1\t2\t18\tS
ISO-02\tStaph. aureus 25923\t1\t3\t18\tS
ISO-02\tStaph. aureus 25923\t2\t1\t18\tS
ISO-02\tStaph. aureus 25923\t2\t2\t18\tS
ISO-02\tStaph. aureus 25923\t2\t3\t18\tS
ISO-02\tStaph. aureus 25923\t3\t1\t18\tS
ISO-02\tStaph. aureus 25923\t3\t2\t22\tS
ISO-02\tStaph. aureus 25923\t3\t3\t18\tS
ISO-02\tStaph. aureus 25923\t4\t1\t18\tS
ISO-02\tStaph. aureus 25923\t4\t2\t18\tS
ISO-02\tStaph. aureus 25923\t4\t3\t17\tS
ISO-02\tStaph. aureus 25923\t5\t1\t18\tS
ISO-02\tStaph. aureus 25923\t5\t2\t18\tS
ISO-02\tStaph. aureus 25923\t5\t3\t18\tS
ISO-03\tE.coli 35218\t1\t1\t6\tR
ISO-03\tE.coli 35218\t1\t2\t6\tR
ISO-03\tE.coli 35218\t1\t3\t6\tR
ISO-03\tE.coli 35218\t2\t1\t6\tR
ISO-03\tE.coli 35218\t2\t2\t6\tR
ISO-03\tE.coli 35218\t2\t3\t6\tR
ISO-03\tE.coli 35218\t3\t1\t11\tI
ISO-03\tE.coli 35218\t3\t2\t6\tR
ISO-03\tE.coli 35218\t3\t3\t6\tR
ISO-03\tE.coli 35218\t4\t1\t6\tR
ISO-03\tE.coli 35218\t4\t2\t6\tR
ISO-03\tE.coli 35218\t4\t3\t6\tR
ISO-03\tE.coli 35218\t5\t1\t6\tR
ISO-03\tE.coli 35218\t5\t2\t6\tR
ISO-03\tE.coli 35218\t5\t3\t6\tR
ISO-04\tKleb. pneumoniae\t1\t1\t16\tI
ISO-04\tKleb. pneumoniae\t1\t2\t16\tI
ISO-04\tKleb. pneumoniae\t1\t3\t16\tI
ISO-04\tKleb. pneumoniae\t2\t1\t16\tI
ISO-04\tKleb. pneumoniae\t2\t2\t16\tI
ISO-04\tKleb. pneumoniae\t2\t3\t16\tI
ISO-04\tKleb. pneumoniae\t3\t1\t16\tI
ISO-04\tKleb. pneumoniae\t3\t2\t16\tI
ISO-04\tKleb. pneumoniae\t3\t3\t16\tI
ISO-04\tKleb. pneumoniae\t4\t1\t20\tS
ISO-04\tKleb. pneumoniae\t4\t2\t16\tI
ISO-04\tKleb. pneumoniae\t4\t3\t16\tI
ISO-04\tKleb. pneumoniae\t5\t1\t16\tI
ISO-04\tKleb. pneumoniae\t5\t2\t16\tI
ISO-04\tKleb. pneumoniae\t5\t3\t16\tI
ISO-05\tE.coli\t1\t1\t30\tS
ISO-05\tE.coli\t1\t2\t30\tS
ISO-05\tE.coli\t1\t3\t30\tS
ISO-05\tE.coli\t2\t1\t30\tS
ISO-05\tE.coli\t2\t2\t30\tS
ISO-05\tE.coli\t2\t3\t30\tS
ISO-05\tE.coli\t3\t1\t30\tS
ISO-05\tE.coli\t3\t2\t30\tS
ISO-05\tE.coli\t3\t3\t30\tS
ISO-05\tE.coli\t4\t1\t30\tS
ISO-05\tE.coli\t4\t2\t30\tS
ISO-05\tE.coli\t4\t3\t30\tS
ISO-05\tE.coli\t5\t1\t30\tS
ISO-05\tE.coli\t5\t2\t30\tS
ISO-05\tE.coli\t5\t3\t30\tS"""

# Input selection
input_mode = st.radio(
    "Choose Input Method:",
    options=["Interactive Grid / Paste Text", "Upload Excel (.xlsx) File"],
    horizontal=True
)

df_loaded = None

if input_mode == "Upload Excel (.xlsx) File":
    uploaded_file = st.file_uploader("Upload .xlsx file", type=['xlsx', 'xls'])
    if uploaded_file is not None:
        try:
            content = uploaded_file.read()
            excel_data = pd.read_excel(io.BytesIO(content), header=None)

            # Detect template header row if present
            header_row_idx = None
            for i, row in excel_data.iterrows():
                row_str = ' '.join(row.astype(str).values).lower()
                if 'isolate id' in row_str and 'zone diameter' in row_str:
                    header_row_idx = i
                    break

            if header_row_idx is not None:
                df_loaded = pd.read_excel(io.BytesIO(content), skiprows=header_row_idx).iloc[:75]
                st.success(f"Detected table header at row {header_row_idx + 1}")
            else:
                df_loaded = pd.read_excel(io.BytesIO(content))

            df_loaded.columns = [str(c).strip() for c in df_loaded.columns]
        except Exception as e:
            st.error(f"Error reading Excel: {e}")

else:
    # Initialize session state for the editable table
    if 'raw_grid_df' not in st.session_state:
        st.session_state['raw_grid_df'] = pd.read_csv(io.StringIO(SAMPLE_DATA), sep='\t')

    with st.expander("Paste Data (TSV / CSV)"):
        pasted_text = st.text_area(
            "Paste lines with headers:",
            value="",
            height=130,
            placeholder="Isolate ID\tOrganism\tDay\tReplicate\tZone diameter (mm)\tInterpretation (S/I/R)..."
        )
        if st.button("Load Pasted Text into Grid"):
            if pasted_text.strip():
                sep = '\t' if '\t' in pasted_text.splitlines()[0] else ','
                parsed_df = pd.read_csv(io.StringIO(pasted_text.strip()), sep=sep)
                st.session_state['raw_grid_df'] = parsed_df
                st.success("Loaded pasted data into editor below.")

    st.markdown("##### Edit or Verify Replicate Data")
    df_loaded = st.data_editor(
        st.session_state['raw_grid_df'],
        num_rows="dynamic",
        use_container_width=True
    )

# Execution
if df_loaded is not None and not df_loaded.empty:
    if st.button("Run Precision Calculation", type="primary"):
        df_calc, df_sum, overall_pca = calculate_precision(df_loaded)

        st.divider()

        # Display Summary
        st.subheader("Disc Diffusion Precision Summary")
        col1, col2 = st.columns([1, 3])
        with col1:
            st.metric(label="Overall PCA", value=f"{overall_pca}%")
        with col2:
            st.dataframe(df_sum, use_container_width=True, hide_index=True)

        # Visualizations
        st.subheader("Visual Analytics")
        fig = generate_visualizations(df_calc, df_sum, overall_pca)
        st.pyplot(fig)

        # Excel Export
        excel_buf = export_to_excel_buffer(df_calc, df_sum, overall_pca)
        st.download_button(
            label="📥 Download Formatted Excel Report",
            data=excel_buf,
            file_name="DD_Precision_Verification_Results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

        with st.expander("View Full Evaluated Replicates Table"):
            st.dataframe(df_calc, use_container_width=True, hide_index=True)