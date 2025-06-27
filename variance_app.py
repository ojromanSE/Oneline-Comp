import streamlit as st
import pandas as pd
import numpy as np
from fpdf import FPDF
import matplotlib.pyplot as plt
import io
import zipfile
import os

# ==== FONT HANDLING ====
GARAMOND_TTF = "GARA.TTF"  # Put Garamond TTF in the same folder or upload to Streamlit "Files"
HAS_GARAMOND = os.path.exists(GARAMOND_TTF)

def try_add_font(pdf):
    """Add Garamond if available, else use Times."""
    if HAS_GARAMOND:
        try:
            pdf.add_font('Garamond', '', GARAMOND_TTF, uni=True)
        except RuntimeError:
            pass  # If already added
    # else: don't add, use Times

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
    def __init__(self, logo_path=None):
        super().__init__()
        self.logo_path = logo_path
        try_add_font(self)

    def footer(self):
        self.set_y(-15)
        try_set_font(self, 8, "I")
        self.cell(0, 10, f"Page {self.page_no()}", 0, 0, "C")
        if self.logo_path and os.path.exists(self.logo_path):
            logo_width = 12
            x_position = self.w - self.r_margin - logo_width
            y_position = self.h - 15
            self.image(self.logo_path, x=x_position, y=y_position, w=logo_width)

# ==== HELPER FUNCTIONS ====
def load_oneline(file):
    # Get the list of sheet names first
    xl = pd.ExcelFile(file)
    sheet_names = xl.sheet_names
    if "Oneline" not in sheet_names:
        raise ValueError(f"Excel file does not contain a sheet named 'Oneline'. Found sheets: {sheet_names}")
    df = xl.parse("Oneline")
    df.columns = df.columns.str.strip()
    # Validate required columns
    for required in ['PROPNUM', 'LEASE_NAME']:
        if required not in df.columns:
            raise ValueError(f"Sheet 'Oneline' must contain column '{required}'. Columns found: {list(df.columns)}")
    if 'SE_RSV_CAT' not in df.columns:
        df['SE_RSV_CAT'] = 'Unknown'
    return df


def suffix_columns(df, suffix, ignore_cols=["PROPNUM", "LEASE_NAME"]):
    return df.rename(columns={col: f"{col}{suffix}" for col in df.columns if col not in ignore_cols})

def generate_explanations(variance_df, npv_column):
    explanations = []
    for _, row in variance_df.iterrows():
        reason = []
        thresholds = {
            "Net Total Revenue ($)": 0.05,
            "Net Operating Expense ($)": 0.05,
            "Inital Approx WI": 0.05,
            "Net Res Oil (Mbbl)": 0.05,
            "Net Res Gas (MMcf)": 0.05,
            "Net Capex ($)": 0.05,
            "Net Res NGL (Mbbl)": 0.05,
            npv_column: 0.05
        }
        key_columns = list(thresholds.keys())
        max_variance = 0
        max_variance_column = ""
        for col in key_columns:
            if f"{col}_final" in row and f"{col}_begin" in row:
                variance = row.get(f"{col} Variance", 0)
                initial_value = row.get(f"{col}_begin", 0)
                if initial_value != 0 and abs(variance) / abs(initial_value) > thresholds[col]:
                    if abs(variance) > max_variance:
                        max_variance = abs(variance)
                        max_variance_column = col
        if max_variance_column:
            variance = row.get(f"{max_variance_column} Variance", 0)
            if max_variance_column == "Net Total Revenue ($)":
                reason.append(f"Revenue increased by ${abs(variance):,.0f}, likely due to an increase in production volume or commodity prices.")
            elif max_variance_column == "Net Operating Expense ($)":
                reason.append(f"Operating expense increased by ${abs(variance):,.0f}, likely due to higher maintenance or operational inefficiencies.")
            elif max_variance_column == "Inital Approx WI":
                reason.append(f"Working Interest (WI) changed by {abs(variance):,.2f}%, changing the share of revenue.")
            elif max_variance_column == "Net Res Oil (Mbbl)":
                reason.append(f"Oil reserves changed by {abs(variance):,.2f} Mbbl, possibly due to new well discoveries or reservoir revisions.")
            elif max_variance_column == "Net Res Gas (MMcf)":
                reason.append(f"Gas reserves changed by {abs(variance):,.2f} MMcf, likely due to reservoir performance or revised estimates.")
            elif max_variance_column == "Net Capex ($)":
                reason.append(f"Capital expenditures changed by ${abs(variance):,.0f}, likely due to new projects or maintenance activities.")
            elif max_variance_column == "Net Res NGL (Mbbl)":
                reason.append(f"NGL reserves changed by {abs(variance):,.2f} Mbbl, possibly due to better recovery or improved reservoir performance.")
            elif max_variance_column == npv_column:
                reason.append(f"{npv_column} changed by ${abs(variance):,.0f}, likely due to increased reserves or improved cost efficiency.")
        explanations.append({
            "PROPNUM": row["PROPNUM"],
            "LEASE_NAME": row["LEASE_NAME"],
            "Key Metric": max_variance_column,
            "Variance Value": row.get(f"{max_variance_column} Variance", 0),
            "Explanation": f"{' '.join(reason)}"
        })
    return pd.DataFrame(explanations)

def identify_negative_npv_wells(variance_df, npv_column):
    return variance_df[(variance_df.get(f"{npv_column}_begin", 0) > 0) & (variance_df.get(f"{npv_column}_final", 0) <= 0)]

def calculate_nri_wi_ratio(begin_df, final_df):
    def compute_ratio(df, wi_col, nri_col, prop_col, lease_col, suffix):
        df = df[df[wi_col] != 0]
        df = df.assign(**{f"NRI/WI Ratio {suffix}": df[nri_col] / df[wi_col]})
        return df[[prop_col, lease_col, f"NRI/WI Ratio {suffix}"]].rename(
            columns={prop_col: "PROPNUM", lease_col: "LEASE_NAME"}
        )
    begin_ratios = compute_ratio(begin_df, 'Inital Approx WI_begin', 'Initial Approx NRI_begin', 'PROPNUM', 'LEASE_NAME', "Begin")
    final_ratios = compute_ratio(final_df, 'Inital Approx WI_final', 'Initial Approx NRI_final', 'PROPNUM', 'LEASE_NAME', "Final")
    merged = begin_ratios.merge(final_ratios, on=["PROPNUM", "LEASE_NAME"], how="outer")
    def out_of_bounds(ratio): return pd.notna(ratio) and (ratio < 0.70 or ratio > 0.85)
    merged["Outlier Source"] = merged.apply(lambda row: (
        "Both" if out_of_bounds(row.get("NRI/WI Ratio Begin", np.nan)) and out_of_bounds(row.get("NRI/WI Ratio Final", np.nan)) else
        "Begin" if out_of_bounds(row.get("NRI/WI Ratio Begin", np.nan)) else
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
    plot_df = plot_df[plot_df[col] != 0]

    # Split and sort positives/negatives as requested
    pos = plot_df[plot_df[col] > 0].sort_values(by=col, ascending=False).head(top_n)
    neg = plot_df[plot_df[col] < 0].sort_values(by=col, ascending=True).head(top_n)
    combined = pd.concat([pos, neg])

    # Reverse order so largest positive on top, most negative at bottom
    combined = combined.reset_index(drop=True)

    labels = combined["PROPNUM"].astype(str) + "\n" + combined["LEASE_NAME"].astype(str)
    values = combined[col]

    fig, ax = plt.subplots(figsize=(8, max(6, 0.5*len(combined))))
    colors = ['#5CB85C' if v > 0 else '#D9534F' for v in values]
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
            pdf.cell(0, 12, title, ln=True, align='C')
            pdf.ln(4)
        pdf.image(tmpfile.name, x=15, w=180)
        os.unlink(tmpfile.name)

def generate_excel(variance_df, excel_buffer, npv_column, filtered_wells_df, begin_df, final_df, nri_df):
    with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
        variance_df.to_excel(writer, sheet_name="Variance Summary", index=False)
        filtered_wells_df[["PROPNUM", "LEASE_NAME"]].to_excel(writer, sheet_name="Wells with Negative or Zero NPV", index=False)
        nri_df[["PROPNUM", "LEASE_NAME", "NRI/WI Ratio Begin", "NRI/WI Ratio Final", "Outlier Source"]].to_excel(writer, sheet_name="Lease NRI", index=False)
    excel_buffer.seek(0)

def generate_pdf(variance_df, pdf_buffer, npv_column, explanation_df, nri_df):
    pdf = PDFWithPageNumbers()
    try_add_font(pdf)
    variance_df_backup = variance_df.copy()
    variance_df = variance_df.merge(explanation_df, on=["PROPNUM", "LEASE_NAME"], how="left")
    for col in ['SE_RSV_CAT_begin', 'SE_RSV_CAT_final']:
        if col not in variance_df.columns:
            variance_df[col] = 'Unknown'
    categories = pd.unique(
        pd.concat([variance_df['SE_RSV_CAT_begin'], variance_df['SE_RSV_CAT_final']])
    ).tolist()
    categories = [cat for cat in categories if pd.notna(cat)]
    if not categories:
        categories = ['Summary']
    prop_lease_width = 60
    cat_width = 22
    value_width = 26
    line_height = 5
    bottom_margin = 15

    def check_page_break(pdf, needed_height):
        if pdf.get_y() + needed_height > pdf.h - bottom_margin:
            pdf.add_page()
            try_set_font(pdf, 11, "B")
            pdf.cell(prop_lease_width, 8, "PROPNUM / LEASE_NAME")
            pdf.cell(cat_width, 8, "Begin Cat")
            pdf.cell(cat_width, 8, "Final Cat")
            pdf.cell(value_width, 8, "Value Change")
            pdf.cell(0, 8, "Explanation", ln=True)
            pdf.set_draw_color(200, 200, 200)
            pdf.line(10, pdf.get_y(), 200, pdf.get_y())
            try_set_font(pdf, 10)

    # 1. Variance Summary
    for category in categories:
        pdf.add_page()
        try_set_font(pdf, 14, "B")
        pdf.cell(0, 10, f"Variance Summary for {category}", ln=True, align='C')
        pdf.ln(2)
        try_set_font(pdf, 11)
        category_df = variance_df[
            (variance_df["SE_RSV_CAT_begin"] == category) | (variance_df["SE_RSV_CAT_final"] == category)
        ]
        summary_lines = [
            f"Net Oil Change: {category_df['Net Res Oil (Mbbl) Variance'].sum():,.2f} Mbbl",
            f"Net Gas Change: {category_df['Net Res Gas (MMcf) Variance'].sum():,.2f} MMcf",
            f"{npv_column} Change: ${category_df[f'{npv_column} Variance'].sum():,.0f}"
        ]
        for line in summary_lines:
            pdf.cell(0, 8, line, ln=True)
        pdf.ln(4)
        try_set_font(pdf, 11, "B")
        pdf.cell(prop_lease_width, 8, "PROPNUM / LEASE_NAME")
        pdf.cell(cat_width, 8, "Begin Cat")
        pdf.cell(cat_width, 8, "Final Cat")
        pdf.cell(value_width, 8, "Value Change")
        pdf.cell(0, 8, "Explanation", ln=True)
        pdf.set_draw_color(200, 200, 200)
        pdf.line(10, pdf.get_y(), 200, pdf.get_y())
        try_set_font(pdf, 10)
        if not category_df.empty:
            major_changes = category_df[
                category_df[f"{npv_column} Variance"].abs() > category_df[f"{npv_column} Variance"].abs().quantile(0.95)
            ]
            for _, row in major_changes.iterrows():
                value = row["Variance Value"]
                metric = row["Key Metric"]
                explanation = row["Explanation"]
                # === Whole dollars only ===
                if "Revenue" in metric or "$" in metric or "Capex" in metric or metric == npv_column:
                    val_str = f"${int(round(value)):,.0f}"
                elif "Oil" in metric:
                    val_str = f"{value:,.2f} Mbbl"
                elif "Gas" in metric:
                    val_str = f"{value:,.2f} MMcf"
                elif "WI" in metric or "NRI" in metric:
                    val_str = f"{value:.2%}"
                else:
                    val_str = f"{value:,.2f}"
                well_id = str(row["PROPNUM"])
                lease_name = str(row["LEASE_NAME"])
                well_text = well_id + "\n" + lease_name
                well_lines = len(pdf.multi_cell(prop_lease_width, line_height, well_text, border=0, align='L', split_only=True))
                explanation_lines = len(pdf.multi_cell(0, line_height, explanation, border=0, split_only=True))
                row_height = max(line_height * well_lines, line_height * explanation_lines)
                check_page_break(pdf, row_height)
                x = pdf.get_x()
                y = pdf.get_y()
                pdf.set_xy(x, y)
                pdf.multi_cell(prop_lease_width, line_height, well_text, border=0, align='L')
                pdf.set_xy(x + prop_lease_width, y)
                pdf.cell(cat_width, row_height, str(row["SE_RSV_CAT_begin"]), border=0)
                pdf.cell(cat_width, row_height, str(row["SE_RSV_CAT_final"]), border=0)
                pdf.cell(value_width, row_height, val_str, border=0)
                pdf.set_xy(x + prop_lease_width + cat_width * 2 + value_width, y)
                pdf.multi_cell(0, line_height, explanation, border=0)
                pdf.set_draw_color(220, 220, 220)
                pdf.line(10, pdf.get_y(), 200, pdf.get_y())
                pdf.ln(1)
    # 2. Plots (NPV second page)
    npv_fig = plot_top_contributors(variance_df, npv_column)
    metric_figs = []
    for metric in MAIN_METRICS:
        if metric != npv_column:
            fig = plot_top_contributors(variance_df, metric)
            if fig:
                metric_figs.append((metric, fig))
    if npv_fig:
        add_chart_to_pdf(pdf, npv_fig, title=f"Top Contributors to {npv_column} Change")
    for metric, fig in metric_figs:
        add_chart_to_pdf(pdf, fig, title=f"Top Contributors to {metric} Change")
    # 3. Transitions/outliers per category (same as before)
    pdf_bytes = pdf.output(dest='S').encode('latin1')
    pdf_buffer.write(pdf_bytes)
    pdf_buffer.seek(0)

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
    else:
        begin_df = load_oneline(begin_file)
        final_df = load_oneline(final_file)
        begin_df_s = suffix_columns(begin_df, "_begin")
        final_df_s = suffix_columns(final_df, "_final")
        variance_df = begin_df_s.merge(final_df_s, on=["PROPNUM", "LEASE_NAME"], how="outer")
        # Fill missing
        for col in ['SE_RSV_CAT_begin', 'SE_RSV_CAT_final']:
            if col not in variance_df.columns:
                variance_df[col] = 'Unknown'
        variance_df['Reserve Category Begin'] = variance_df['SE_RSV_CAT_begin']
        variance_df['Reserve Category Final'] = variance_df['SE_RSV_CAT_final']
        key_columns = [
            "Net Total Revenue ($)", "Net Operating Expense ($)", "Inital Approx WI", "Initial Approx NRI",
            "Net Res Oil (Mbbl)", "Net Res Gas (MMcf)", "Net Capex ($)", "Net Res NGL (Mbbl)", npv_column
        ]
        for col in key_columns:
            col_begin = f"{col}_begin"
            col_final = f"{col}_final"
            if col_begin in variance_df.columns and col_final in variance_df.columns:
                variance_df[f"{col} Variance"] = variance_df[col_final] - variance_df[col_begin]
        explanation_df = generate_explanations(variance_df, npv_column)
        negative_npv_wells = identify_negative_npv_wells(variance_df, npv_column)
        nri_df = calculate_nri_wi_ratio(begin_df_s, final_df_s)
        # Excel buffer
        excel_buffer = io.BytesIO()
        generate_excel(variance_df, excel_buffer, npv_column, negative_npv_wells, begin_df_s, final_df_s, nri_df)
        # PDF buffer
        pdf_buffer = io.BytesIO()
        generate_pdf(variance_df, pdf_buffer, npv_column, explanation_df, nri_df)
        # Zip both
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w') as zipf:
            zipf.writestr("variance_report.xlsx", excel_buffer.getvalue())
            zipf.writestr("variance_report.pdf", pdf_buffer.getvalue())
        zip_buffer.seek(0)
        st.session_state['zip_bytes'] = zip_buffer.getvalue()
        st.success("Reports generated! You can now download them.")

if st.session_state['zip_bytes']:
    st.download_button("Download Reports (ZIP)", st.session_state['zip_bytes'], file_name="variance_reports.zip", mime="application/zip")
