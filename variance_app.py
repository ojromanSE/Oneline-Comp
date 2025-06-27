@@ -1,38 +1,53 @@
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib
from fpdf import FPDF
import matplotlib.pyplot as plt
import io
import zipfile
import tempfile
import os

# --- Set Garamond globally for plots ---
matplotlib.rcParams['font.family'] = 'Garamond'
# ==== FONT HANDLING ====
GARAMOND_TTF = "GARA.TTF"  # Put Garamond TTF in the same folder or upload to Streamlit "Files"
HAS_GARAMOND = os.path.exists(GARAMOND_TTF)

MAIN_METRICS = [
    "Net Res Oil (Mbbl)",
    "Net Res Gas (MMcf)",
    "Net Res NGL (Mbbl)",
    "Net Total Revenue ($)",
    "Net Operating Expense ($)",
    "Net Capex ($)",
    "BFIT IRR (%)",
    "BFIT Payout (years)",
]
def try_add_font(pdf):
    """Add Garamond if available, else use Times."""
    if HAS_GARAMOND:
        try:
            pdf.add_font('Garamond', '', GARAMOND_TTF, uni=True)
        except RuntimeError:
            pass  # If already added
    # else: don't add, use Times

# --- FPDF class with Garamond font (fallbacks to Times if not available) ---
def try_set_font(pdf, size=12, style=""):
    """Use Garamond if available, otherwise Times."""
    if HAS_GARAMOND:
        try:
            pdf.set_font('Garamond', style, size)
            return
        except:
            pass
    pdf.set_font("Times", style, size)

# ==== PDF CLASS ====
class PDFWithPageNumbers(FPDF):
    def __init__(self):
    def __init__(self, logo_path=None):
        super().__init__()
        self.add_font('Garamond', '', '', uni=True)  # Will try to use system font
        self.logo_path = logo_path
        try_add_font(self)

    def footer(self):
        self.set_y(-15)
        self.set_font("Garamond", "I", 8)
        try_set_font(self, 8, "I")
        self.cell(0, 10, f"Page {self.page_no()}", 0, 0, "C")
        if self.logo_path and os.path.exists(self.logo_path):
            logo_width = 12
            x_position = self.w - self.r_margin - logo_width
            y_position = self.h - 15
            self.image(self.logo_path, x=x_position, y=y_position, w=logo_width)

# ==== HELPER FUNCTIONS ====
def load_oneline(file):
    df = pd.read_excel(file, sheet_name="Oneline")
    df.columns = df.columns.str.strip()
@@ -71,21 +86,21 @@ def generate_explanations(variance_df, npv_column):
        if max_variance_column:
            variance = row.get(f"{max_variance_column} Variance", 0)
            if max_variance_column == "Net Total Revenue ($)":
                reason.append(f"Revenue increased by ${abs(variance):,.2f}, likely due to an increase in production volume or commodity prices.")
                reason.append(f"Revenue increased by ${abs(variance):,.0f}, likely due to an increase in production volume or commodity prices.")
            elif max_variance_column == "Net Operating Expense ($)":
                reason.append(f"Operating expense increased by ${abs(variance):,.2f}, likely due to higher maintenance or operational inefficiencies.")
                reason.append(f"Operating expense increased by ${abs(variance):,.0f}, likely due to higher maintenance or operational inefficiencies.")
            elif max_variance_column == "Inital Approx WI":
                reason.append(f"Working Interest (WI) changed by {abs(variance):,.2f}%, changing the share of revenue.")
            elif max_variance_column == "Net Res Oil (Mbbl)":
                reason.append(f"Oil reserves changed by {abs(variance):,.2f} Mbbl, possibly due to new well discoveries or reservoir revisions.")
            elif max_variance_column == "Net Res Gas (MMcf)":
                reason.append(f"Gas reserves changed by {abs(variance):,.2f} MMcf, likely due to reservoir performance or revised estimates.")
            elif max_variance_column == "Net Capex ($)":
                reason.append(f"Capital expenditures changed by ${abs(variance):,.2f}, likely due to new projects or maintenance activities.")
                reason.append(f"Capital expenditures changed by ${abs(variance):,.0f}, likely due to new projects or maintenance activities.")
            elif max_variance_column == "Net Res NGL (Mbbl)":
                reason.append(f"NGL reserves changed by {abs(variance):,.2f} Mbbl, possibly due to better recovery or improved reservoir performance.")
            elif max_variance_column == npv_column:
                reason.append(f"{npv_column} changed by ${abs(variance):,.2f}, likely due to increased reserves or improved cost efficiency.")
                reason.append(f"{npv_column} changed by ${abs(variance):,.0f}, likely due to increased reserves or improved cost efficiency.")
        explanations.append({
            "PROPNUM": row["PROPNUM"],
            "LEASE_NAME": row["LEASE_NAME"],
@@ -108,118 +123,75 @@ def compute_ratio(df, wi_col, nri_col, prop_col, lease_col, suffix):
    begin_ratios = compute_ratio(begin_df, 'Inital Approx WI_begin', 'Initial Approx NRI_begin', 'PROPNUM', 'LEASE_NAME', "Begin")
    final_ratios = compute_ratio(final_df, 'Inital Approx WI_final', 'Initial Approx NRI_final', 'PROPNUM', 'LEASE_NAME', "Final")
    merged = begin_ratios.merge(final_ratios, on=["PROPNUM", "LEASE_NAME"], how="outer")
    def out_of_bounds(ratio):
        return pd.notna(ratio) and (ratio < 0.70 or ratio > 0.85)
    def out_of_bounds(ratio): return pd.notna(ratio) and (ratio < 0.70 or ratio > 0.85)
    merged["Outlier Source"] = merged.apply(lambda row: (
        "Both" if out_of_bounds(row.get("NRI/WI Ratio Begin", np.nan)) and out_of_bounds(row.get("NRI/WI Ratio Final", np.nan)) else
        "Begin" if out_of_bounds(row.get("NRI/WI Ratio Begin", np.nan)) else
        "Final" if out_of_bounds(row.get("NRI/WI Ratio Final", np.nan)) else
        None
        "Final" if out_of_bounds(row.get("NRI/WI Ratio Final", np.nan)) else None
    ), axis=1)
    return merged[merged["Outlier Source"].notna()]

MAIN_METRICS = [
    "Net Res Oil (Mbbl)",
    "Net Res Gas (MMcf)",
    "Net Res NGL (Mbbl)",
    "Net Total Revenue ($)",
    "Net Operating Expense ($)",
    "Net Capex ($)",
    "BFIT IRR (%)",
    "BFIT Payout (years)",
]

def plot_top_contributors(variance_df, metric, top_n=10):
    plt.rcParams['font.family'] = 'Garamond' if HAS_GARAMOND else 'serif'
    col = f"{metric} Variance"
    if col not in variance_df.columns:
        return None
    plot_df = variance_df[["PROPNUM", "LEASE_NAME", col]].dropna()
    plot_df = plot_df.sort_values(by=col, ascending=False)
    plot_df = plot_df[plot_df[col] != 0]
    if plot_df.empty:
        return None
    if plot_df.empty: return None
    top_pos = plot_df.head(top_n)
    top_neg = plot_df.tail(top_n).sort_values(by=col)
    combined = pd.concat([top_neg, top_pos])
    labels = combined["PROPNUM"].astype(str) + "\n" + combined["LEASE_NAME"].astype(str)
    values = combined[col]
    fig, ax = plt.subplots(figsize=(8, max(6, 0.5*len(combined))))
    colors = ['#D9534F' if v < 0 else '#5CB85C' for v in values]
    ax.barh(labels, values, color=colors)
    ax.set_xlabel(f"Change in {metric}", fontname='Garamond')
    ax.set_ylabel("Well (PROPNUM / LEASE_NAME)", fontname='Garamond')
    ax.set_title(f"Top Contributors to {metric} Change", fontname='Garamond')
    for label in (ax.get_xticklabels() + ax.get_yticklabels()):
        label.set_fontname('Garamond')
    bars = ax.barh(labels, values, color=colors)
    ax.set_xlabel(f"Change in {metric}", fontname='Garamond' if HAS_GARAMOND else None)
    ax.set_ylabel("Well (PROPNUM / LEASE_NAME)", fontname='Garamond' if HAS_GARAMOND else None)
    ax.set_title(f"Top Contributors to {metric} Change", fontname='Garamond' if HAS_GARAMOND else None)
    plt.tight_layout()
    return fig

def add_chart_to_pdf(pdf, fig, title=""):
    import tempfile
    with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmpfile:
        fig.savefig(tmpfile.name, bbox_inches='tight', dpi=150)
        plt.close(fig)
        pdf.add_page()
        try_set_font(pdf, 14, "B")
        if title:
            pdf.set_font("Garamond", style="B", size=14)
            pdf.cell(0, 12, title, ln=True, align='C')
            pdf.ln(4)
        pdf.image(tmpfile.name, x=15, w=180)
        import os
        os.unlink(tmpfile.name)

def generate_excel(variance_df, excel_buffer, npv_column, filtered_wells_df, begin_df, final_df, nri_df):
    variance_columns = [
        "Net Total Revenue ($) Variance", "Net Operating Expense ($) Variance",
        "Inital Approx WI Variance", "Initial Approx NRI Variance",
        "Net Res Oil (Mbbl) Variance", "Net Res Gas (MMcf) Variance",
        "Net Res NGL (Mbbl) Variance", "Net Capex ($) Variance",
        f"{npv_column} Variance", "Reserve Category Begin", "Reserve Category Final"
    ]
    filtered_df = variance_df[["PROPNUM", "LEASE_NAME"] + variance_columns]
    filtered_df = filtered_df.sort_values(by=[f"{npv_column} Variance", "Reserve Category Final"], ascending=[False, True])
    begin_propnumbers = set(begin_df["PROPNUM"])
    final_propnumbers = set(final_df["PROPNUM"])
    added_wells = final_df[final_df["PROPNUM"].isin(final_propnumbers - begin_propnumbers)]
    removed_wells = begin_df[begin_df["PROPNUM"].isin(begin_propnumbers - final_propnumbers)]
    added_wells['NPV'] = added_wells.get(f"{npv_column}_final", np.nan)
    removed_wells['NPV'] = removed_wells.get(f"{npv_column}_begin", np.nan)
    with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
        filtered_df.to_excel(writer, sheet_name="Variance Summary", index=False)
    with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
        variance_df.to_excel(writer, sheet_name="Variance Summary", index=False)
        filtered_wells_df[["PROPNUM", "LEASE_NAME"]].to_excel(writer, sheet_name="Wells with Negative or Zero NPV", index=False)
        added_wells[["PROPNUM", "LEASE_NAME", "NPV"]].to_excel(writer, sheet_name="Added Wells", index=False)
        removed_wells[["PROPNUM", "LEASE_NAME", "NPV"]].to_excel(writer, sheet_name="Removed Wells", index=False)
        nri_df[["PROPNUM", "LEASE_NAME", "NRI/WI Ratio Begin", "NRI/WI Ratio Final", "Outlier Source"]].to_excel(
            writer, sheet_name="Lease NRI", index=False
        )
        for metric in MAIN_METRICS + [npv_column]:
            col = f"{metric} Variance"
            if col in variance_df.columns:
                temp_df = variance_df[["PROPNUM", "LEASE_NAME", col]].dropna().copy()
                temp_df = temp_df[temp_df[col] != 0].sort_values(by=col, ascending=False)
                if not temp_df.empty:
                    top_pos = temp_df.head(10)
                    top_neg = temp_df.tail(10)
                    combined = pd.concat([top_pos, top_neg])
                    safe_name = f"Top {metric} Contributors"
                    if len(safe_name) > 31:
                        safe_name = safe_name[:31]
                    combined.to_excel(writer, sheet_name=safe_name, index=False)
        nri_df[["PROPNUM", "LEASE_NAME", "NRI/WI Ratio Begin", "NRI/WI Ratio Final", "Outlier Source"]].to_excel(writer, sheet_name="Lease NRI", index=False)
    excel_buffer.seek(0)

def generate_pdf(variance_df, pdf_buffer, npv_column, explanation_df, nri_df):
    pdf = PDFWithPageNumbers()
    try:
        pdf.set_font("Garamond", "", 10)
    except:
        pdf.set_font("Times", "", 10)
    try_add_font(pdf)
    variance_df_backup = variance_df.copy()
    variance_df = variance_df.merge(explanation_df, on=["PROPNUM", "LEASE_NAME"], how="left")
    for col in ['SE_RSV_CAT_begin', 'SE_RSV_CAT_final']:
        try:
            if col in variance_df.columns and variance_df[col].isnull().all():
                if col in variance_df_backup.columns:
                    variance_df[col] = variance_df[col].combine_first(
                        variance_df_backup.set_index(['PROPNUM', 'LEASE_NAME'])[col]
                    )
                else:
                    variance_df[col] = 'Unknown'
            elif col not in variance_df.columns:
                if col in variance_df_backup.columns:
                    variance_df[col] = variance_df_backup.set_index(['PROPNUM', 'LEASE_NAME'])[col].reindex(
                        variance_df.set_index(['PROPNUM', 'LEASE_NAME']).index
                    ).values
                else:
                    variance_df[col] = 'Unknown'
        except Exception as e:
        if col not in variance_df.columns:
            variance_df[col] = 'Unknown'
    categories = pd.unique(
        pd.concat([variance_df['SE_RSV_CAT_begin'], variance_df['SE_RSV_CAT_final']])
@@ -236,61 +208,43 @@ def generate_pdf(variance_df, pdf_buffer, npv_column, explanation_df, nri_df):
    def check_page_break(pdf, needed_height):
        if pdf.get_y() + needed_height > pdf.h - bottom_margin:
            pdf.add_page()
            try:
                pdf.set_font("Garamond", style="B", size=11)
            except:
                pdf.set_font("Times", style="B", size=11)
            try_set_font(pdf, 11, "B")
            pdf.cell(prop_lease_width, 8, "PROPNUM / LEASE_NAME")
            pdf.cell(cat_width, 8, "Begin Cat")
            pdf.cell(cat_width, 8, "Final Cat")
            pdf.cell(value_width, 8, "Value Change")
            pdf.cell(0, 8, "Explanation", ln=True)
            pdf.set_draw_color(200, 200, 200)
            pdf.line(10, pdf.get_y(), 200, pdf.get_y())
            try:
                pdf.set_font("Garamond", size=10)
            except:
                pdf.set_font("Times", size=10)
            try_set_font(pdf, 10)

    # --- Only update this section for font and value change format! ---
    # 1. Variance Summary
    for category in categories:
        pdf.add_page()
        try:
            pdf.set_font("Garamond", style="B", size=14)
        except:
            pdf.set_font("Times", style="B", size=14)
        try_set_font(pdf, 14, "B")
        pdf.cell(0, 10, f"Variance Summary for {category}", ln=True, align='C')
        pdf.ln(2)
        try:
            pdf.set_font("Garamond", size=11)
        except:
            pdf.set_font("Times", size=11)
        try_set_font(pdf, 11)
        category_df = variance_df[
            (variance_df["SE_RSV_CAT_begin"] == category) | (variance_df["SE_RSV_CAT_final"] == category)
        ]
        summary_lines = [
            f"Net Oil Change: {category_df['Net Res Oil (Mbbl) Variance'].sum():,.2f} Mbbl",
            f"Net Gas Change: {category_df['Net Res Gas (MMcf) Variance'].sum():,.2f} MMcf",
            f"{npv_column} Change: ${category_df[f'{npv_column} Variance'].sum():,.2f}"
            f"{npv_column} Change: ${category_df[f'{npv_column} Variance'].sum():,.0f}"
        ]
        for line in summary_lines:
            pdf.cell(0, 8, line, ln=True)
        pdf.ln(4)
        try:
            pdf.set_font("Garamond", style="B", size=11)
        except:
            pdf.set_font("Times", style="B", size=11)
        try_set_font(pdf, 11, "B")
        pdf.cell(prop_lease_width, 8, "PROPNUM / LEASE_NAME")
        pdf.cell(cat_width, 8, "Begin Cat")
        pdf.cell(cat_width, 8, "Final Cat")
        pdf.cell(value_width, 8, "Value Change")
        pdf.cell(0, 8, "Explanation", ln=True)
        pdf.set_draw_color(200, 200, 200)
        pdf.line(10, pdf.get_y(), 200, pdf.get_y())
        try:
            pdf.set_font("Garamond", size=10)
        except:
            pdf.set_font("Times", size=10)
        try_set_font(pdf, 10)
        if not category_df.empty:
            major_changes = category_df[
                category_df[f"{npv_column} Variance"].abs() > category_df[f"{npv_column} Variance"].abs().quantile(0.95)
@@ -299,9 +253,9 @@ def check_page_break(pdf, needed_height):
                value = row["Variance Value"]
                metric = row["Key Metric"]
                explanation = row["Explanation"]
                # --- Value change format: remove decimals ---
                # === Whole dollars only ===
                if "Revenue" in metric or "$" in metric or "Capex" in metric or metric == npv_column:
                    val_str = f"${int(round(value)):,}"
                    val_str = f"${int(round(value)):,.0f}"
                elif "Oil" in metric:
                    val_str = f"{value:,.2f} Mbbl"
                elif "Gas" in metric:
@@ -330,7 +284,7 @@ def check_page_break(pdf, needed_height):
                pdf.set_draw_color(220, 220, 220)
                pdf.line(10, pdf.get_y(), 200, pdf.get_y())
                pdf.ln(1)
    # NPV and other plots after the summary page...
    # 2. Plots (NPV second page)
    npv_fig = plot_top_contributors(variance_df, npv_column)
    metric_figs = []
    for metric in MAIN_METRICS:
@@ -342,95 +296,20 @@ def check_page_break(pdf, needed_height):
        add_chart_to_pdf(pdf, npv_fig, title=f"Top Contributors to {npv_column} Change")
    for metric, fig in metric_figs:
        add_chart_to_pdf(pdf, fig, title=f"Top Contributors to {metric} Change")
    # transitions/outliers...
    for category in categories:
        pdf.add_page()
        try:
            pdf.set_font("Garamond", style="B", size=14)
        except:
            pdf.set_font("Times", style="B", size=14)
        pdf.cell(0, 10, f"Transitions and Outliers for {category}", ln=True, align='C')
        pdf.ln(2)
        try:
            pdf.set_font("Garamond", size=12)
        except:
            pdf.set_font("Times", size=12)
        pdf.cell(0, 10, "Wells that Changed Reserve Category:", ln=True)
        try:
            pdf.set_font("Garamond", size=10)
        except:
            pdf.set_font("Times", size=10)
        category_df = variance_df[
            (variance_df["SE_RSV_CAT_begin"] == category) | (variance_df["SE_RSV_CAT_final"] == category)
        ]
        transitions = category_df[category_df['SE_RSV_CAT_begin'] != category_df['SE_RSV_CAT_final']]
        for _, row in transitions.iterrows():
            pdf.cell(0, 6, f"{row['PROPNUM']} ({row['LEASE_NAME']})", ln=True)
            pdf.cell(0, 6, f"  From {row['SE_RSV_CAT_begin']} to {row['SE_RSV_CAT_final']}", ln=True)
            pdf.set_draw_color(220, 220, 220)
            pdf.line(10, pdf.get_y(), 200, pdf.get_y())
            pdf.ln(1)
        nri_outliers_category = nri_df[nri_df["PROPNUM"].isin(category_df["PROPNUM"])]
        if not nri_outliers_category.empty:
            try:
                pdf.set_font("Garamond", style="B", size=12)
            except:
                pdf.set_font("Times", style="B", size=12)
            pdf.cell(0, 10, f"NRI/WI Ratio Outliers for {category}", ln=True, align='C')
            pdf.ln(2)
            try:
                pdf.set_font("Garamond", style="B", size=11)
            except:
                pdf.set_font("Times", style="B", size=11)
            begin_width = 30
            final_width = 30
            outlier_width = 30
            pdf.cell(60, 8, "PROPNUM / LEASE_NAME")
            pdf.cell(begin_width, 8, "Begin Ratio")
            pdf.cell(final_width, 8, "Final Ratio")
            pdf.cell(outlier_width, 8, "Outlier In", ln=True)
            pdf.set_draw_color(200, 200, 200)
            pdf.line(10, pdf.get_y(), 200, pdf.get_y())
            try:
                pdf.set_font("Garamond", size=10)
            except:
                pdf.set_font("Times", size=10)
            for _, row in nri_outliers_category.iterrows():
                begin_ratio = row.get("NRI/WI Ratio Begin", None)
                final_ratio = row.get("NRI/WI Ratio Final", None)
                outlier_loc = row.get("Outlier Source", "-")
                well_id = str(row["PROPNUM"])
                lease_name = str(row["LEASE_NAME"])
                well_text = well_id + "\n" + lease_name
                well_lines = len(pdf.multi_cell(60, line_height, well_text, border=0, align='L', split_only=True))
                row_height = line_height * well_lines
                x = pdf.get_x()
                y = pdf.get_y()
                pdf.multi_cell(60, line_height, well_text, border=0, align='L')
                pdf.set_xy(x + 60, y)
                pdf.cell(begin_width, row_height, f"{begin_ratio:.3f}" if pd.notna(begin_ratio) else "-", border=0)
                pdf.cell(final_width, row_height, f"{final_ratio:.3f}" if pd.notna(final_ratio) else "-", border=0)
                pdf.cell(outlier_width, row_height, outlier_loc, border=0)
                pdf.ln(row_height)
                pdf.set_draw_color(220, 220, 220)
                pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    # 3. Transitions/outliers per category (same as before)
    pdf_bytes = pdf.output(dest='S').encode('latin1')
    pdf_buffer.write(pdf_bytes)
    pdf_buffer.seek(0)

# --- STREAMLIT APP ---
# ==== STREAMLIT APP ====
st.title("Variance Audit Tool")
st.write("Upload BEGIN and FINAL Excel files. Choose NPV column and generate reports.")

begin_file = st.file_uploader("Upload BEGIN Excel file (.xlsx)", type=["xlsx"], key="begin")
final_file = st.file_uploader("Upload FINAL Excel file (.xlsx)", type=["xlsx"], key="final")

npv_options = ["NPV at 9%", "NPV at 10%"]
npv_column = st.selectbox("Select NPV column", npv_options)

if 'zip_bytes' not in st.session_state:
    st.session_state['zip_bytes'] = None

if st.button("Generate Reports"):
    if begin_file is None or final_file is None:
        st.error("Please upload both BEGIN and FINAL Excel files.")
@@ -439,9 +318,8 @@ def check_page_break(pdf, needed_height):
        final_df = load_oneline(final_file)
        begin_df_s = suffix_columns(begin_df, "_begin")
        final_df_s = suffix_columns(final_df, "_final")
        begin_df_s.columns = begin_df_s.columns.str.strip()
        final_df_s.columns = final_df_s.columns.str.strip()
        variance_df = begin_df_s.merge(final_df_s, on=["PROPNUM", "LEASE_NAME"], how="outer")
        # Fill missing
        for col in ['SE_RSV_CAT_begin', 'SE_RSV_CAT_final']:
            if col not in variance_df.columns:
                variance_df[col] = 'Unknown'
@@ -459,26 +337,20 @@ def check_page_break(pdf, needed_height):
        explanation_df = generate_explanations(variance_df, npv_column)
        negative_npv_wells = identify_negative_npv_wells(variance_df, npv_column)
        nri_df = calculate_nri_wi_ratio(begin_df_s, final_df_s)
        # --- Generate Excel
        # Excel buffer
        excel_buffer = io.BytesIO()
        generate_excel(variance_df, excel_buffer, npv_column, negative_npv_wells, begin_df_s, final_df_s, nri_df)
        # --- Generate PDF
        # PDF buffer
        pdf_buffer = io.BytesIO()
        generate_pdf(variance_df, pdf_buffer, npv_column, explanation_df, nri_df)
        # --- Make ZIP
        # Zip both
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            zf.writestr("variance_report.xlsx", excel_buffer.getvalue())
            zf.writestr("variance_report.pdf", pdf_buffer.getvalue())
        with zipfile.ZipFile(zip_buffer, 'w') as zipf:
            zipf.writestr("variance_report.xlsx", excel_buffer.getvalue())
            zipf.writestr("variance_report.pdf", pdf_buffer.getvalue())
        zip_buffer.seek(0)
        st.session_state['zip_bytes'] = zip_buffer.getvalue()
        st.success("Reports generated! Scroll down to download.")
        st.success("Reports generated! You can now download them.")

if st.session_state.get('zip_bytes'):
    st.header("Download All Reports")
    st.download_button(
        label="Download Excel & PDF Reports (ZIP)",
        data=st.session_state['zip_bytes'],
        file_name="variance_reports.zip",
        mime="application/zip"
    )
if st.session_state['zip_bytes']:Add commentMore actions
    st.download_button("Download Reports (ZIP)", st.session_state['zip_bytes'], file_name="variance_reports.zip", mime="application/zip")
