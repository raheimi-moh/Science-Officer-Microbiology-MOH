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
    page_title="MIC Precision Verification",
    page_icon="🧪",
    layout="wide"
)
sns.set_theme(style="whitegrid", font="sans-serif")


# ==============================================================================
# 1. CALCULATION ENGINE
# ==============================================================================

def calculate_mic_precision(df_input):
    """Computes modal MIC, ±1 dilution agreement, and PEA (%) precision metrics."""
    df = df_input.copy()

    # Standardize column naming
    col_map = {}
    for c in df.columns:
        c_lower = str(c).strip().lower()
        if 'isolate' in c_lower:
            col_map[c] = 'Isolate ID'
        elif 'organism' in c_lower:
            col_map[c] = 'Organism'
        elif 'day' in c_lower:
            col_map[c] = 'Day'
        elif 'replicate' in c_lower:
            col_map[c] = 'Replicate'
        elif 'mic' in c_lower and ('result' in c_lower or 'value' in c_lower or c_lower == 'mic'):
            col_map[c] = 'MIC result'

    df = df.rename(columns=col_map)
    required = ['Isolate ID', 'Organism', 'Day', 'Replicate', 'MIC result']
    for col in required:
        if col not in df.columns:
            df[col] = ''

    df['MIC result'] = pd.to_numeric(df['MIC result'], errors='coerce')

    modal_dict = {}
    multi_mode_dict = {}

    for isolate, group in df.groupby('Isolate ID', sort=False):
        valid_mics = group['MIC result'].dropna()
        if len(valid_mics) == 0:
            modal_dict[isolate] = np.nan
            multi_mode_dict[isolate] = 'No data'
        else:
            val_counts = valid_mics.value_counts()
            max_count = val_counts.max()
            top_modes = val_counts[val_counts == max_count].index.tolist()

            if len(top_modes) == 1:
                modal_dict[isolate] = top_modes[0]
                multi_mode_dict[isolate] = 'Unique'
            else:
                modal_dict[isolate] = top_modes[0]  # First mode as calculation baseline
                multi_mode_dict[isolate] = f"Tie ({', '.join(str(m) for m in sorted(top_modes))})"

    df['Modal MIC*'] = df['Isolate ID'].map(modal_dict)

    def check_within_dilution(row):
        mic = row['MIC result']
        mode = row['Modal MIC*']
        if pd.isna(mic) or pd.isna(mode):
            return ''
        if (mode / 2.0) <= mic <= (mode * 2.0):
            return 'Yes'
        return 'No'

    df['Within ±1 dilution?'] = df.apply(check_within_dilution, axis=1)

    summary_records = []
    for isolate, group in df.groupby('Isolate ID', sort=False):
        org = group['Organism'].iloc[0] if 'Organism' in group.columns else ''
        valid_mics = group['MIC result'].dropna()
        n_valid = len(valid_mics)
        mode_status = multi_mode_dict.get(isolate, 'No data')
        mode_val = modal_dict.get(isolate, np.nan)

        if n_valid == 0:
            summary_records.append({
                'Isolate ID': isolate,
                'Organism': org,
                'Valid MIC results (n)': 0,
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
            'Valid MIC results (n)': int(n_valid),
            'Modal MIC': mode_val if 'Tie' not in mode_status else f'{mode_status}',
            'Within ±1 dilution (n)': int(within_n),
            'Outside ±1 dilution (n)': int(outside_n),
            'PEA (%)': round(pea, 1),
            'Review': review,
        })

    summary_df = pd.DataFrame(summary_records)

    tot_valid = summary_df['Valid MIC results (n)'].sum()
    tot_within = summary_df['Within ±1 dilution (n)'].sum()
    overall_pea = round((tot_within / tot_valid) * 100.0, 1) if tot_valid > 0 else 0.0

    return df, summary_df, overall_pea


# ==============================================================================
# 2. EXCEL BUILDER (IN-MEMORY BYTESIO)
# ==============================================================================

def export_mic_to_excel_buffer(df_calc, df_sum, overall_pea):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'MIC1 Precision'
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
    ws.merge_cells('A1:I1')
    ws['A1'] = 'MIC PRECISION VERIFICATION – ONE ANTIMICROBIAL'
    ws['A1'].font = white_title_font
    ws['A1'].fill = header_fill
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 28

    ws.merge_cells('A2:I2')
    ws['A2'] = (
        'Enter numeric MIC values for 5 isolates, tested 3 times daily over 5'
        ' days (15 replicates per isolate). PEA is calculated against the modal'
        ' MIC for each isolate, within ±1 twofold dilution.'
    )
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

    # Replicate Section Header
    ws.merge_cells('A8:I8')
    ws['A8'] = (
        'DATA ENTRY: Enter one MIC result per replicate. Use numeric MIC values'
        ' (e.g. 0.25, 0.5, 1, 2, 4).'
    )
    ws['A8'].font = Font(name='Segoe UI', size=10, bold=True, color='FFFFFF')
    ws['A8'].fill = section_fill
    ws['A8'].alignment = Alignment(vertical='center')
    ws.row_dimensions[8].height = 22

    headers = [
        'Isolate ID', 'Organism', 'Day', 'Replicate',
        'MIC result', 'Modal MIC*', 'Within ±1 dilution?'
    ]
    for col_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=10, column=col_idx, value=h)
        cell.font = bold_font
        cell.fill = table_hdr_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border
    ws.row_dimensions[10].height = 24

    curr_row = 11
    for _, row in df_calc.iterrows():
        vals = [
            row.get('Isolate ID', ''),
            row.get('Organism', ''),
            row.get('Day', ''),
            row.get('Replicate', ''),
            row.get('MIC result', ''),
            row.get('Modal MIC*', ''),
            row.get('Within ±1 dilution?', ''),
        ]
        for col_idx, val in enumerate(vals, start=1):
            cell = ws.cell(row=curr_row, column=col_idx, value=val)
            cell.font = regular_font
            cell.border = thin_border
            if col_idx in [3, 4, 5, 6]:
                cell.alignment = Alignment(horizontal='right', vertical='center')
            else:
                cell.alignment = Alignment(horizontal='center', vertical='center')

            if col_idx == 7 and str(val).lower() == 'no':
                cell.fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
                cell.font = Font(name='Segoe UI', size=10, color='9C0006', bold=True)
            elif curr_row % 2 == 0:
                cell.fill = row_stripe
        curr_row += 1

    # Precision Summary Header
    curr_row += 2
    ws.merge_cells(f'A{curr_row}:H{curr_row}')
    ws[f'A{curr_row}'] = 'PRECISION SUMMARY – AUTOMATIC PEA CALCULATION'
    ws[f'A{curr_row}'].font = Font(name='Segoe UI', size=11, bold=True, color='FFFFFF')
    ws[f'A{curr_row}'].fill = section_fill
    ws.row_dimensions[curr_row].height = 22

    curr_row += 1
    ws.merge_cells(f'A{curr_row}:H{curr_row}')
    ws[f'A{curr_row}'] = (
        'PEA (%) = number of MIC results within ±1 twofold dilution of the'
        ' modal MIC ÷ total valid MIC results × 100.'
    )
    ws[f'A{curr_row}'].font = Font(name='Segoe UI', size=9, italic=True, color='595959')

    curr_row += 1
    sum_headers = [
        'Isolate ID', 'Organism', 'Valid MIC results (n)',
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
            row['Valid MIC results (n)'],
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
            if col_idx in [3, 4, 5, 6]:
                cell.alignment = Alignment(horizontal='right', vertical='center')
            elif col_idx in [1, 7]:
                cell.alignment = Alignment(horizontal='center', vertical='center')

            if col_idx == 8:
                if 'All within' in str(val):
                    cell.font = Font(name='Segoe UI', size=10, color='0E6251', bold=True)
                elif 'Review' in str(val):
                    cell.font = Font(name='Segoe UI', size=10, color='C0392B', bold=True)
        curr_row += 1

    ws.cell(row=curr_row, column=1, value='Overall PEA').font = bold_font
    res_cell = ws.cell(row=curr_row, column=2, value=f'{overall_pea}%')
    res_cell.font = Font(name='Segoe UI', size=11, bold=True, color='1F4E78')
    ws.row_dimensions[curr_row].height = 22

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 13)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ==============================================================================
# 3. VISUALIZATION
# ==============================================================================

def plot_mic_precision(df_calc, df_sum, overall_pea):
    valid_sum = df_sum[df_sum['Valid MIC results (n)'] > 0].copy()
    if valid_sum.empty:
        return None

    fig, axes = plt.subplots(1, 2, figsize=(15, 5))

    # Bar chart: PEA (%)
    colors = [
        '#2ECC71' if p == 100 else '#E67E22' if p >= 90 else '#E74C3C'
        for p in valid_sum['PEA (%)'].fillna(0)
    ]
    bars = axes[0].bar(
        valid_sum['Isolate ID'],
        valid_sum['PEA (%)'].fillna(0),
        color=colors,
        edgecolor='black',
        linewidth=0.8,
        width=0.5,
    )
    axes[0].axhline(
        overall_pea,
        color='#2980B9',
        linestyle='--',
        linewidth=1.8,
        label=f'Overall PEA: {overall_pea}%',
    )
    axes[0].set_ylim(0, 115)
    axes[0].set_title('Precision Essential Agreement (PEA %)', fontsize=12, fontweight='bold', color='#1F4E78')
    axes[0].set_ylabel('Within ±1 Dilution of Mode (%)', fontsize=10, fontweight='bold')
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

    # Boxplot / Stripplot: Log2 scale
    valid_calc = df_calc[df_calc['MIC result'].notna()].copy()
    if not valid_calc.empty:
        sns.boxplot(
            data=valid_calc,
            x='Isolate ID',
            y='MIC result',
            ax=axes[1],
            palette='Purples_r',
            showmeans=True,
            meanprops={'marker': 'o', 'markerfacecolor': 'red', 'markeredgecolor': 'black'},
        )
        sns.stripplot(
            data=valid_calc,
            x='Isolate ID',
            y='MIC result',
            ax=axes[1],
            color='black',
            size=6,
            alpha=0.5,
            jitter=0.2,
        )
        axes[1].set_yscale('log', base=2)
        axes[1].set_title('Replicate MIC Dispersion (Log2 Scale)', fontsize=12, fontweight='bold', color='#1F4E78')
        axes[1].set_ylabel('MIC (µg/mL or mg/L)', fontsize=10, fontweight='bold')
        axes[1].set_xlabel('Isolate ID', fontsize=10, fontweight='bold')

    plt.tight_layout()
    return fig


# ==============================================================================
# 4. STREAMLIT APPLICATION
# ==============================================================================

st.title("🧪 MIC Precision Verification (1 Antimicrobial)")
st.caption("Calculate modal MIC, evaluate replicate concordance within ±1 twofold dilution, compute Precision Essential Agreement (PEA %), and download styled Excel workbooks.")

SAMPLE_DATA = """Isolate ID\tOrganism\tDay\tReplicate\tMIC result
ISO-01\tQC Strain 1\t1\t1\t16
ISO-01\tQC Strain 1\t1\t2\t8
ISO-01\tQC Strain 1\t1\t3\t16
ISO-01\tQC Strain 1\t2\t1\t16
ISO-01\tQC Strain 1\t2\t2\t16
ISO-01\tQC Strain 1\t2\t3\t16
ISO-01\tQC Strain 1\t3\t1\t16
ISO-01\tQC Strain 1\t3\t2\t32
ISO-01\tQC Strain 1\t3\t3\t32
ISO-01\tQC Strain 1\t4\t1\t32
ISO-01\tQC Strain 1\t4\t2\t32
ISO-01\tQC Strain 1\t4\t3\t16
ISO-01\tQC Strain 1\t5\t1\t32
ISO-01\tQC Strain 1\t5\t2\t32
ISO-01\tQC Strain 1\t5\t3\t32
ISO-02\tQC Strain 2\t1\t1\t16
ISO-02\tQC Strain 2\t1\t2\t16
ISO-02\tQC Strain 2\t1\t3\t16
ISO-02\tQC Strain 2\t2\t1\t16
ISO-02\tQC Strain 2\t2\t2\t16
ISO-02\tQC Strain 2\t2\t3\t16
ISO-02\tQC Strain 2\t3\t1\t16
ISO-02\tQC Strain 2\t3\t2\t16
ISO-02\tQC Strain 2\t3\t3\t16
ISO-02\tQC Strain 2\t4\t1\t16
ISO-02\tQC Strain 2\t4\t2\t16
ISO-02\tQC Strain 2\t4\t3\t16
ISO-02\tQC Strain 2\t5\t1\t16
ISO-02\tQC Strain 2\t5\t2\t16
ISO-02\tQC Strain 2\t5\t3\t16"""

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
            preview_df = pd.read_excel(io.BytesIO(content), header=None)

            header_idx = None
            for i, row in preview_df.iterrows():
                row_text = ' '.join(row.astype(str).values).lower()
                if 'isolate id' in row_text and ('mic result' in row_text or 'mic' in row_text):
                    header_idx = i
                    break

            if header_idx is not None:
                df_loaded = pd.read_excel(io.BytesIO(content), skiprows=header_idx).iloc[:75]
                st.success(f"Detected table headers at row {header_idx + 1}")
            else:
                df_loaded = pd.read_excel(io.BytesIO(content))

            df_loaded.columns = [str(c).strip() for c in df_loaded.columns]
        except Exception as e:
            st.error(f"Error parsing Excel file: {e}")

else:
    if 'mic_grid_df' not in st.session_state:
        st.session_state['mic_grid_df'] = pd.read_csv(io.StringIO(SAMPLE_DATA), sep='\t')

    with st.expander("Paste Data (TSV / CSV)"):
        pasted_text = st.text_area(
            "Paste rows with headers:",
            value="",
            height=130,
            placeholder="Isolate ID\tOrganism\tDay\tReplicate\tMIC result..."
        )
        if st.button("Load Pasted Text into Grid"):
            if pasted_text.strip():
                sep = '\t' if '\t' in pasted_text.splitlines()[0] else ','
                parsed_df = pd.read_csv(io.StringIO(pasted_text.strip()), sep=sep)
                st.session_state['mic_grid_df'] = parsed_df
                st.success("Loaded pasted data into the interactive editor below.")

    st.markdown("##### Edit or Verify Replicate Data")
    df_loaded = st.data_editor(
        st.session_state['mic_grid_df'],
        num_rows="dynamic",
        use_container_width=True
    )

if df_loaded is not None and not df_loaded.empty:
    if st.button("Run MIC Precision Calculation", type="primary"):
        df_calc, df_sum, overall_pea = calculate_mic_precision(df_loaded)

        st.divider()

        st.subheader("MIC Precision Summary (Automatic PEA Calculation)")
        col1, col2 = st.columns([1, 3])
        with col1:
            st.metric(label="Overall PEA", value=f"{overall_pea}%")
        with col2:
            st.dataframe(df_sum, use_container_width=True, hide_index=True)

        st.subheader("Visual Analytics")
        fig = plot_mic_precision(df_calc, df_sum, overall_pea)
        if fig:
            st.pyplot(fig)

        excel_buf = export_mic_to_excel_buffer(df_calc, df_sum, overall_pea)
        st.download_button(
            label="📥 Download Formatted Excel Report",
            data=excel_buf,
            file_name="MIC_Precision_Verification_Results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

        with st.expander("View Full Replicate Table with Modal Comparison"):
            st.dataframe(
                df_calc[['Isolate ID', 'Organism', 'Day', 'Replicate', 'MIC result', 'Modal MIC*', 'Within ±1 dilution?']],
                use_container_width=True,
                hide_index=True
            )