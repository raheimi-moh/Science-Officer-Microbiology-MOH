import io
import matplotlib.pyplot as plt
import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import seaborn as sns
import streamlit as st

st.set_page_config(
    page_title="MIC Precision (6 Antimicrobials)",
    page_icon="💊",
    layout="wide"
)
sns.set_theme(style="whitegrid", font="sans-serif")


# ==============================================================================
# 1. CORE CALCULATION ENGINE
# ==============================================================================

def calculate_mic6_precision(df_input):
    """Computes modal MIC, ±1 dilution agreement, per-isolate PEA (%),
    and overall PEA (%) across up to 6 antimicrobials.
    """
    df = df_input.copy()

    # Identify metadata vs antimicrobial columns
    fixed_cols = {}
    anti_cols = []
    for col in df.columns:
        c_clean = str(col).strip()
        c_lower = c_clean.lower()
        if 'isolate' in c_lower:
            fixed_cols[col] = 'Isolate ID'
        elif 'organism' in c_lower:
            fixed_cols[col] = 'Organism'
        elif 'day' in c_lower:
            fixed_cols[col] = 'Day'
        elif 'replicate' in c_lower:
            fixed_cols[col] = 'Replicate'
        elif not any(k in c_lower for k in ['mode', 'within', 'pea', 'review', 'unnamed', 'match']):
            anti_cols.append(c_clean)

    df = df.rename(columns=fixed_cols)
    for col in ['Isolate ID', 'Organism', 'Day', 'Replicate']:
        if col not in df.columns:
            df[col] = ''

    if not anti_cols:
        anti_cols = [f'Antimicrobial {i}' for i in range(1, 7) if f'Antimicrobial {i}' in df.columns]
    anti_cols = anti_cols[:6]

    for ac in anti_cols:
        df[ac] = pd.to_numeric(df[ac], errors='coerce')

    # Compute modal MIC per Isolate x Antimicrobial
    modes_dict = {}
    for isolate, group in df.groupby('Isolate ID', sort=False):
        for drug in anti_cols:
            valid_series = group[drug].dropna()
            if len(valid_series) == 0:
                modes_dict[(isolate, drug)] = (np.nan, 'No data')
            else:
                counts = valid_series.value_counts()
                max_count = counts.max()
                top_modes = counts[counts == max_count].index.tolist()
                if len(top_modes) == 1:
                    modes_dict[(isolate, drug)] = (top_modes[0], 'Unique')
                else:
                    modes_dict[(isolate, drug)] = (
                        top_modes[0],
                        f"Tie ({', '.join(str(m) for m in sorted(top_modes))})"
                    )

    # Replicate-level mode columns
    for idx, drug in enumerate(anti_cols, start=1):
        df[f'Mode A{idx}'] = df['Isolate ID'].apply(
            lambda iso: modes_dict.get((iso, drug), (np.nan, ''))[0]
        )

    # Precision summary table by Isolate and Antimicrobial
    summary_records = []
    for isolate, group in df.groupby('Isolate ID', sort=False):
        org = group['Organism'].iloc[0] if 'Organism' in group.columns else ''
        for drug in anti_cols:
            valid_mics = group[drug].dropna()
            n_valid = len(valid_mics)
            mode_val, mode_status = modes_dict.get((isolate, drug), (np.nan, 'No data'))

            if n_valid == 0:
                summary_records.append({
                    'Isolate ID': isolate,
                    'Organism': org,
                    'Antimicrobial': drug,
                    'Valid MIC (n)': 0,
                    'Modal MIC': '',
                    'Within ±1 dilution (n)': 0,
                    'Outside ±1 dilution (n)': 0,
                    'PEA (%)': np.nan,
                    'Review': 'No data',
                })
                continue

            within_n = (
                (valid_mics >= (mode_val / 2.0)) & (valid_mics <= (mode_val * 2.0))
            ).sum()
            outside_n = n_valid - within_n
            pea = (within_n / n_valid) * 100.0

            if n_valid < 15:
                review = 'Incomplete'
            elif 'Tie' in mode_status:
                review = f'Review: {mode_status}'
            elif outside_n > 0:
                review = 'Review outlier(s)'
            else:
                review = 'All within ±1 dilution'

            summary_records.append({
                'Isolate ID': isolate,
                'Organism': org,
                'Antimicrobial': drug,
                'Valid MIC (n)': int(n_valid),
                'Modal MIC': mode_val if 'Tie' not in mode_status else f'{mode_status}',
                'Within ±1 dilution (n)': int(within_n),
                'Outside ±1 dilution (n)': int(outside_n),
                'PEA (%)': round(pea, 1),
                'Review': review,
            })

    summary_df = pd.DataFrame(summary_records)

    # Overall summary by Antimicrobial
    overall_records = []
    for drug in anti_cols:
        drug_summary = summary_df[summary_df['Antimicrobial'] == drug]
        tot_valid = drug_summary['Valid MIC (n)'].sum()
        tot_within = drug_summary['Within ±1 dilution (n)'].sum()
        tot_outside = drug_summary['Outside ±1 dilution (n)'].sum()
        overall_pea = round((tot_within / tot_valid) * 100.0, 1) if tot_valid > 0 else np.nan

        overall_records.append({
            'Antimicrobial': drug,
            'Valid MIC results (n)': int(tot_valid),
            'Within ±1 dilution (n)': int(tot_within),
            'Outside ±1 dilution (n)': int(tot_outside),
            'Overall PEA (%)': overall_pea,
        })

    overall_df = pd.DataFrame(overall_records)

    return df, summary_df, overall_df, anti_cols


# ==============================================================================
# 2. EXCEL BUILDER (IN-MEMORY BYTESIO)
# ==============================================================================

def export_mic6_to_excel_buffer(df_calc, df_sum, overall_df, anti_cols):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'MIC6 Precision'
    ws.views.sheetView[0].showGridLines = True

    header_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
    section_fill = PatternFill(start_color='2C3E50', end_color='2C3E50', fill_type='solid')
    table_hdr_fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
    row_stripe = PatternFill(start_color='F9FAFB', end_color='F9FAFB', fill_type='solid')

    white_title_font = Font(name='Segoe UI', size=13, bold=True, color='FFFFFF')
    bold_font = Font(name='Segoe UI', size=10, bold=True, color='000000')
    regular_font = Font(name='Segoe UI', size=10, color='000000')

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9'),
    )

    # Title Banner
    ws.merge_cells('A1:P1')
    ws['A1'] = 'MIC PRECISION VERIFICATION – SIX ANTIMICROBIALS'
    ws['A1'].font = white_title_font
    ws['A1'].fill = header_fill
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 28

    ws.merge_cells('A2:P2')
    ws['A2'] = 'Enter MIC results for 5 isolates, 3 replicates daily over 5 days (15 replicates per isolate). Enter one numeric MIC value in each antimicrobial column.'
    ws['A2'].font = Font(name='Segoe UI', size=9, italic=True, color='595959')
    ws['A2'].alignment = Alignment(horizontal='center', vertical='center')

    # Metadata
    meta = [
        ('Laboratory / Institution', 'Study date', 'AST system / method'),
        ('Organism group', 'Antimicrobial agent(s)', 'Medium / inoculum details'),
        ('Performed by', 'Reviewed by', ''),
    ]
    for idx, (f1, f2, f3) in enumerate(meta, start=4):
        ws[f'A{idx}'] = f1
        ws[f'A{idx}'].font = bold_font
        ws[f'D{idx}'] = f2
        ws[f'D{idx}'].font = bold_font
        if f3:
            ws[f'F{idx}'] = f3
            ws[f'F{idx}'].font = bold_font

    # Replicate Data Header
    ws.merge_cells('A8:P8')
    ws['A8'] = 'DATA ENTRY: Rename the six antimicrobial columns. Enter numeric MIC values (e.g. 0.25, 0.5, 1, 2, 4).'
    ws['A8'].font = Font(name='Segoe UI', size=10, bold=True, color='FFFFFF')
    ws['A8'].fill = section_fill
    ws['A8'].alignment = Alignment(vertical='center')
    ws.row_dimensions[8].height = 22

    headers = ['Isolate ID', 'Organism', 'Day', 'Replicate'] + anti_cols
    mode_headers = [f'Mode A{i}' for i in range(1, len(anti_cols) + 1)]
    all_headers = headers + mode_headers

    for col_idx, h in enumerate(all_headers, start=1):
        cell = ws.cell(row=10, column=col_idx, value=h)
        cell.font = bold_font
        cell.fill = table_hdr_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border
    ws.row_dimensions[10].height = 24

    curr_row = 11
    for _, row in df_calc.iterrows():
        vals = [
            row['Isolate ID'],
            row['Organism'],
            row['Day'],
            row['Replicate'],
        ]
        for ac in anti_cols:
            vals.append(row.get(ac, ''))
        for mh in mode_headers:
            vals.append(row.get(mh, ''))

        for col_idx, val in enumerate(vals, start=1):
            cell = ws.cell(row=curr_row, column=col_idx, value=val)
            cell.font = regular_font
            cell.border = thin_border
            if col_idx in [3, 4] or (5 <= col_idx <= 10):
                cell.alignment = Alignment(horizontal='right', vertical='center')
            else:
                cell.alignment = Alignment(horizontal='center', vertical='center')
            if curr_row % 2 == 0:
                cell.fill = row_stripe
        curr_row += 1

    # Precision Summary Section
    curr_row += 2
    ws.merge_cells(f'A{curr_row}:I{curr_row}')
    ws[f'A{curr_row}'] = 'PRECISION SUMMARY – AUTOMATIC PEA CALCULATION (BY ISOLATE AND ANTIMICROBIAL)'
    ws[f'A{curr_row}'].font = Font(name='Segoe UI', size=11, bold=True, color='FFFFFF')
    ws[f'A{curr_row}'].fill = section_fill
    ws.row_dimensions[curr_row].height = 22

    curr_row += 1
    ws.merge_cells(f'A{curr_row}:I{curr_row}')
    ws[f'A{curr_row}'] = 'PEA (%) = number of MIC results within ±1 twofold dilution of the modal MIC ÷ total valid MIC results × 100.'
    ws[f'A{curr_row}'].font = Font(name='Segoe UI', size=9, italic=True, color='595959')

    curr_row += 1
    sum_headers = [
        'Isolate ID', 'Organism', 'Antimicrobial', 'Valid MIC (n)',
        'Modal MIC', 'Within ±1 dilution (n)', 'Outside ±1 dilution (n)',
        'PEA (%)', 'Review'
    ]
    for col_idx, h in enumerate(sum_headers, start=1):
        cell = ws.cell(row=curr_row, column=col_idx, value=h)
        cell.font = bold_font
        cell.fill = table_hdr_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border
    ws.row_dimensions[curr_row].height = 22

    curr_row += 1
    for _, row in df_sum.iterrows():
        vals = [
            row['Isolate ID'],
            row['Organism'],
            row['Antimicrobial'],
            row['Valid MIC (n)'],
            row['Modal MIC'],
            row['Within ±1 dilution (n)'],
            row['Outside ±1 dilution (n)'],
            f"{row['PEA (%)']}%" if pd.notna(row['PEA (%)']) else '',
            row['Review'],
        ]
        for col_idx, val in enumerate(vals, start=1):
            cell = ws.cell(row=curr_row, column=col_idx, value=val)
            cell.font = regular_font
            cell.border = thin_border
            if col_idx in [4, 5, 6, 7]:
                cell.alignment = Alignment(horizontal='right', vertical='center')
            elif col_idx in [1, 8]:
                cell.alignment = Alignment(horizontal='center', vertical='center')

            if col_idx == 9:
                if 'All within' in str(val):
                    cell.font = Font(name='Segoe UI', size=10, color='0E6251', bold=True)
                elif 'Review' in str(val):
                    cell.font = Font(name='Segoe UI', size=10, color='C0392B', bold=True)
        curr_row += 1

    # Overall Summary Section
    curr_row += 2
    ws.merge_cells(f'A{curr_row}:E{curr_row}')
    ws[f'A{curr_row}'] = 'OVERALL PEA BY ANTIMICROBIAL (DESCRIPTIVE)'
    ws[f'A{curr_row}'].font = Font(name='Segoe UI', size=11, bold=True, color='FFFFFF')
    ws[f'A{curr_row}'].fill = section_fill
    ws.row_dimensions[curr_row].height = 22

    curr_row += 1
    ov_headers = ['Antimicrobial', 'Valid MIC results (n)', 'Within ±1 dilution (n)', 'Outside ±1 dilution (n)', 'Overall PEA (%)']
    for col_idx, h in enumerate(ov_headers, start=1):
        cell = ws.cell(row=curr_row, column=col_idx, value=h)
        cell.font = bold_font
        cell.fill = table_hdr_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border
    ws.row_dimensions[curr_row].height = 22

    curr_row += 1
    for _, row in overall_df.iterrows():
        vals = [
            row['Antimicrobial'],
            row['Valid MIC results (n)'],
            row['Within ±1 dilution (n)'],
            row['Outside ±1 dilution (n)'],
            f"{row['Overall PEA (%)']}%" if pd.notna(row['Overall PEA (%)']) else '',
        ]
        for col_idx, val in enumerate(vals, start=1):
            cell = ws.cell(row=curr_row, column=col_idx, value=val)
            cell.font = regular_font
            cell.border = thin_border
            if col_idx in [2, 3, 4]:
                cell.alignment = Alignment(horizontal='right', vertical='center')
            elif col_idx == 5:
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.font = Font(name='Segoe UI', size=10, bold=True, color='1F4E78')
        curr_row += 1

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

def plot_mic6_precision(df_calc, df_sum, overall_df, anti_cols):
    valid_sum = df_sum[df_sum['Valid MIC (n)'] > 0].copy()
    if valid_sum.empty:
        return None

    fig = plt.figure(figsize=(16, 10))

    # 1. Heatmap: Isolate vs Antimicrobial PEA (%)
    ax1 = plt.subplot2grid((2, 2), (0, 0))
    pivot_pea = valid_sum.pivot(index='Isolate ID', columns='Antimicrobial', values='PEA (%)')
    sns.heatmap(
        pivot_pea,
        annot=True,
        fmt='.1f',
        cmap='RdYlGn',
        vmin=70,
        vmax=100,
        cbar_kws={'label': 'PEA (%)'},
        ax=ax1,
        linewidths=1,
        linecolor='white',
    )
    ax1.set_title('PEA (%) Heatmap (Isolates × Antimicrobials)', fontsize=12, fontweight='bold', color='#1F4E78')

    # 2. Bar Chart: Overall PEA (%)
    ax2 = plt.subplot2grid((2, 2), (0, 1))
    valid_overall = overall_df.dropna(subset=['Overall PEA (%)'])
    bars = ax2.bar(
        valid_overall['Antimicrobial'],
        valid_overall['Overall PEA (%)'],
        color='#3498DB',
        edgecolor='black',
        width=0.5,
    )
    ax2.set_ylim(0, 115)
    ax2.set_title('Overall PEA (%) by Antimicrobial', fontsize=12, fontweight='bold', color='#1F4E78')
    ax2.set_ylabel('Pooled PEA (%)', fontsize=10, fontweight='bold')
    ax2.tick_params(axis='x', rotation=20)

    for bar in bars:
        h = bar.get_height()
        ax2.text(
            bar.get_x() + bar.get_width() / 2.0,
            h + 2,
            f'{h:.1f}%',
            ha='center',
            va='bottom',
            fontsize=9,
            fontweight='bold',
        )

    # 3. Strip/Box Plot: MIC Distributions (Log2)
    ax3 = plt.subplot2grid((2, 2), (1, 0), colspan=2)
    melted_df = df_calc.melt(
        id_vars=['Isolate ID', 'Day', 'Replicate'],
        value_vars=anti_cols,
        var_name='Antimicrobial',
        value_name='MIC',
    ).dropna(subset=['MIC'])

    if not melted_df.empty:
        sns.boxplot(
            data=melted_df,
            x='Antimicrobial',
            y='MIC',
            hue='Isolate ID',
            ax=ax3,
            palette='Set2',
        )
        ax3.set_yscale('log', base=2)
        ax3.set_title('Replicate MIC Distributions (Log2 Scale)', fontsize=12, fontweight='bold', color='#1F4E78')
        ax3.set_ylabel('MIC (µg/mL)', fontsize=10, fontweight='bold')
        ax3.legend(bbox_to_anchor=(1.01, 1), loc='upper left')

    plt.tight_layout()
    return fig


# ==============================================================================
# 4. STREAMLIT APPLICATION LAYOUT
# ==============================================================================

st.title("💊 MIC Precision Verification (6 Antimicrobials)")
st.caption("Verify multi-antimicrobial MIC precision across replicates, generate isolate × drug PEA heatmaps, and export validated Excel workbooks.")

SAMPLE_DATA = """Isolate ID\tOrganism\tDay\tReplicate\tAmpicillin\tCefotaxime\tCiprofloxacin\tGentamicin\tMeropenem\tVancomycin
ISO-01\tE. coli ATCC 25922\t1\t1\t4\t0.06\t0.015\t0.5\t0.03\t16
ISO-01\tE. coli ATCC 25922\t1\t2\t4\t0.06\t0.015\t0.5\t0.03\t16
ISO-01\tE. coli ATCC 25922\t1\t3\t4\t0.06\t0.015\t0.5\t0.03\t16
ISO-01\tE. coli ATCC 25922\t2\t1\t4\t0.06\t0.015\t0.5\t0.03\t16
ISO-01\tE. coli ATCC 25922\t2\t2\t4\t0.06\t0.015\t0.5\t0.03\t16
ISO-01\tE. coli ATCC 25922\t2\t3\t4\t0.06\t0.015\t0.5\t0.03\t16
ISO-01\tE. coli ATCC 25922\t3\t1\t4\t0.06\t0.015\t0.5\t0.03\t16
ISO-01\tE. coli ATCC 25922\t3\t2\t4\t0.06\t0.015\t1.0\t0.03\t16
ISO-01\tE. coli ATCC 25922\t3\t3\t8\t0.06\t0.015\t0.5\t0.03\t16
ISO-01\tE. coli ATCC 25922\t4\t1\t4\t0.06\t0.015\t0.5\t0.03\t16
ISO-01\tE. coli ATCC 25922\t4\t2\t4\t0.06\t0.015\t0.5\t0.03\t16
ISO-01\tE. coli ATCC 25922\t4\t3\t4\t0.06\t0.015\t0.5\t0.03\t16
ISO-01\tE. coli ATCC 25922\t5\t1\t4\t0.06\t0.015\t0.5\t0.03\t16
ISO-01\tE. coli ATCC 25922\t5\t2\t4\t0.06\t0.015\t0.5\t0.03\t16
ISO-01\tE. coli ATCC 25922\t5\t3\t4\t0.06\t0.015\t0.5\t0.03\t16
ISO-02\tP. aeruginosa 27853\t1\t1\t32\t8\t0.5\t2\t0.5\t64
ISO-02\tP. aeruginosa 27853\t1\t2\t32\t8\t0.5\t2\t0.5\t64
ISO-02\tP. aeruginosa 27853\t1\t3\t32\t8\t0.5\t2\t0.5\t64
ISO-02\tP. aeruginosa 27853\t2\t1\t32\t8\t0.5\t2\t0.5\t64
ISO-02\tP. aeruginosa 27853\t2\t2\t32\t8\t0.5\t2\t0.5\t64
ISO-02\tP. aeruginosa 27853\t2\t3\t32\t8\t0.5\t2\t0.5\t64
ISO-02\tP. aeruginosa 27853\t3\t1\t32\t8\t0.5\t2\t0.5\t64
ISO-02\tP. aeruginosa 27853\t3\t2\t32\t8\t0.5\t2\t0.5\t64
ISO-02\tP. aeruginosa 27853\t3\t3\t32\t8\t0.5\t2\t0.5\t64
ISO-02\tP. aeruginosa 27853\t4\t1\t32\t8\t0.5\t2\t0.5\t64
ISO-02\tP. aeruginosa 27853\t4\t2\t32\t8\t0.5\t2\t0.5\t64
ISO-02\tP. aeruginosa 27853\t4\t3\t32\t8\t0.5\t2\t0.5\t64
ISO-02\tP. aeruginosa 27853\t5\t1\t32\t8\t0.5\t2\t0.5\t64
ISO-02\tP. aeruginosa 27853\t5\t2\t32\t8\t0.5\t2\t0.5\t64
ISO-02\tP. aeruginosa 27853\t5\t3\t32\t8\t0.5\t2\t0.5\t64"""

input_mode = st.radio(
    "Choose Input Method:",
    options=["Interactive Grid / Paste Text", "Upload Excel (.xlsx) File"],
    horizontal=True
)

df_loaded = None

if input_mode == "Upload Excel (.xlsx) File":
    uploaded_file = st.file_uploader("Select an .xlsx or .xls file", type=['xlsx', 'xls'])
    if uploaded_file is not None:
        try:
            content = uploaded_file.read()
            preview = pd.read_excel(io.BytesIO(content), header=None)

            h_idx = None
            for i, row in preview.iterrows():
                r_text = ' '.join(row.astype(str).values).lower()
                if 'isolate id' in r_text and ('antimicrobial' in r_text or 'ampicillin' in r_text):
                    h_idx = i
                    break

            if h_idx is not None:
                df_loaded = pd.read_excel(io.BytesIO(content), skiprows=h_idx).iloc[:75]
                st.success(f"Detected table headers at row {h_idx + 1}")
            else:
                df_loaded = pd.read_excel(io.BytesIO(content))

            df_loaded.columns = [str(c).strip() for c in df_loaded.columns]
        except Exception as e:
            st.error(f"Error parsing Excel file: {e}")

else:
    if 'mic6_grid_df' not in st.session_state:
        st.session_state['mic6_grid_df'] = pd.read_csv(io.StringIO(SAMPLE_DATA), sep='\t')

    with st.expander("Paste Data (TSV / CSV)"):
        pasted_text = st.text_area(
            "Paste rows with column headers:",
            value="",
            height=130,
            placeholder="Isolate ID\tOrganism\tDay\tReplicate\tAmpicillin\tCefotaxime..."
        )
        if st.button("Load Pasted Text into Grid"):
            if pasted_text.strip():
                sep = '\t' if '\t' in pasted_text.splitlines()[0] else ','
                parsed_df = pd.read_csv(io.StringIO(pasted_text.strip()), sep=sep)
                st.session_state['mic6_grid_df'] = parsed_df
                st.success("Loaded pasted data into editor below.")

    st.markdown("##### Edit or Verify Replicate Data")
    df_loaded = st.data_editor(
        st.session_state['mic6_grid_df'],
        num_rows="dynamic",
        use_container_width=True
    )

if df_loaded is not None and not df_loaded.empty:
    if st.button("Run 6-Antimicrobial Precision Calculation", type="primary"):
        df_calc, df_sum, overall_df, anti_cols = calculate_mic6_precision(df_loaded)

        st.divider()

        # Summary Display
        st.subheader("Overall PEA (%) by Antimicrobial")
        st.dataframe(overall_df, use_container_width=True, hide_index=True)

        with st.expander("View Detailed Summary by Isolate and Antimicrobial"):
            st.dataframe(df_sum, use_container_width=True, hide_index=True)

        # Visual Analytics
        st.subheader("Visual Analytics")
        fig = plot_mic6_precision(df_calc, df_sum, overall_df, anti_cols)
        if fig:
            st.pyplot(fig)

        # Excel Export
        excel_buf = export_mic6_to_excel_buffer(df_calc, df_sum, overall_df, anti_cols)
        st.download_button(
            label="📥 Download Formatted Excel Report",
            data=excel_buf,
            file_name="MIC6_Precision_Verification_Results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

        with st.expander("View Full Raw Calculation Table with Modal Assignments"):
            st.dataframe(df_calc, use_container_width=True, hide_index=True)