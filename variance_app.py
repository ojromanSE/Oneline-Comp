import streamlit as st
import pandas as pd
import numpy as np
from fpdf import FPDF
import matplotlib.pyplot as plt
import io
import tempfile
from datetime import datetime
import base64

# ====== FONT SETUP FOR PDF ======
HAS_GARAMOND = False
try:
    from matplotlib import font_manager
    available_fonts = set(f.name for f in font_manager.fontManager.ttflist)
    HAS_GARAMOND = 'Garamond' in available_fonts
except Exception:
    HAS_GARAMOND = False

DEFAULT_FONT = "Garamond" if HAS_GARAMOND else "Times"

# ====== PDF CLASS ======
class PDFWithPageNumbers(FPDF):
    def __init__(self):
        super().__init__()
        try:
            if HAS_GARAMOND:
                self.add_font('Garamond', '', '', uni=True)
        except Exception:
            pass

    def footer(self):
        self.set_y(-15)
        self.set_font(DEFAULT_FONT, "I", 8)
        self.cell(0, 10, f"Page {self.page_no()}", 0, 0, "C")

# ====== FUNCTIONS ======

def load_oneline(file):
    df = pd.read_excel(file, sheet_name="Oneline")
    df.columns = df.columns.str.strip()
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

    def out_of_bounds(ratio):
        return pd.notna(ratio) and (ratio < 0.70 or ratio > 0.85)

    merged["Outlier Source"] = merged.apply(lambda row: (
        "Both" if out_of_bounds(row.get("NRI/WI Ratio Begin", np.nan)) and out_of_bounds(row.get("NRI/WI Ratio Final", np.nan)) else
        "Begin" if out_of_bounds(row.get("NRI/WI Ratio Begin", np.nan)) else
        "Final" if out_of_bounds(row.get("NRI/WI Ratio Final", np.nan)) else
        None
    ), axis=1)
    return merged[merged["Outlier Source"].notna()]

# ===== BAR CHART, ORDERED POSITIVE TO NEGATIVE =====
def plot_top_contributors(variance_df, metric, top_n=10):
    col = f"{metric} Variance"
    if col not in variance_df.columns:
        return None
    plot_df = variance_df[["PROPNUM", "LEASE_NAME", col]].dropna()
    plot_df = plot_df[plot_df[col] != 0]
    if plot_df.empty:
        return None
    plot_df = plot_df.sort_values(by=col, ascending=False)
    # Take top N positive, top N negative, then concat and flip so biggest positive is at the top
    top_pos = plot_df.head(top_n)
    top_neg = plot_df.tail(top_n)
    combined = pd.concat([top_pos, top_neg])
    combined = combined.sort_values(by=col, ascending=False).iloc[::-1]  # flip order for barh

    labels = combined["PROPNUM"].astype(str) + "\n" + combined["LEASE_NAME"].astype(str)
    values = combined[col]
    fig, ax = plt.subplots(figsize=(8, max(6, 0.5*len(combined))))
    if HAS_GARAMOND:
        plt.rcParams['font.family'] = 'Garamond'
    else:
        plt.rcParams['font.family'] = 'serif'
    colors = ['#5CB85C' if v >= 0 else '#D9534F' for v in values]
    ax.barh(labels, values, color=colors)
    ax.set_xlabel(f"Change in {metric}", fontname=DEFAULT_FONT)
    ax.set_ylabel("Well (PROPNUM / LEASE_NAME)", fontname=DEFAULT_FONT)
    ax.set_title(f"Top Contributors to {metric} Change", fontname=DEFAULT_FONT)
    plt.tight_layout()
    return fig

def add_chart_to_pdf(pdf, fig, title=""):
    with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmpfile:
        fig.savefig(tmpfile.name, bbox_inches='tight', dpi=150)
        plt.close(fig)
        pdf.add_page()
        pdf.set_font(DEFAULT_FONT, style="B", size=14)
        if title:
            pdf.cell(0, 12, title, ln=True, align='C')
            pdf.ln(4)
        pdf.image(tmpfile.name, x=15, w=180)
    import os
    os.unlink(tmpfile.name)

# ====== MAIN METRICS ======
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

# ====== GENERATE EXCEL ======
def generate_excel(variance_df, npv_column, filtered_wells_df, begin_df, final_df, nri_df):
    excel_buffer = io.BytesIO()
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
        filtered_wells_df[["PROPNUM", "LEASE_NAME"]].to_excel(writer, sheet_name="Wells with Negative or Zero NPV", index=False)
        added_wells[["PROPNUM", "LEASE_NAME", "NPV"]].to_excel(writer, sheet_name="Added Wells", index=False)
        removed_wells[["PROPNUM", "LEASE_NAME", "NPV"]].to_excel(writer, sheet_name="Removed Wells", index=False)
        nri_df[["PROPNUM", "LEASE_NAME", "NRI/WI Ratio Begin", "NRI/WI Ratio Final", "Outlier Source"]].to_excel(
            writer, sheet_name="Lease NRI", index=False
        )
        # Add top contributors for each metric, truncating name to 31 chars
        for metric in MAIN_METRICS + [npv_column]:
            col = f"{metric} Variance"
            if col in variance_df.columns:
                temp_df = variance_df[["PROPNUM", "LEASE_NAME", col]].dropna().copy()
                temp_df = temp_df[temp_df[col] != 0].sort_values(by=col, ascending=False)
                if not temp_df.empty:
                    top_pos = temp_df.head(10)
                    top_neg = temp_df.tail(10)
                    combined = pd.concat([top_pos, top_neg])
                    sheet_name = f"Top {metric} Contributors"
                    if len(sheet_name) > 31:
                        sheet_name = sheet_name[:31]
                    combined.to_excel(writer, sheet_name=sheet_name, index=False)
    excel_buffer.seek(0)
    return excel_buffer

# ====== GENERATE PDF ======
def generate_pdf(variance_df, npv_column, explanation_df, nri_df):
    pdf_buffer = io.BytesIO()
    pdf = PDFWithPageNumbers()
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
        except Exception:
            variance_df[col] = 'Unknown'

    categories = pd.unique(
        pd.concat([variance_df['SE_RSV_CAT_begin'], variance_df['SE_RSV_CAT_final']])
    ).tolist()
    categories = [cat for cat in categories if pd.notna(cat)]
    if not categories:
        categories = ['Summary']

    # Define column widths
    prop_lease_width = 60
    cat_width = 22
    value_width = 26
    line_height = 5
    bottom_margin = 15

    def check_page_break(pdf, needed_height):
        if pdf.get_y() + needed_height > pdf.h - bottom_margin:
            pdf.add_page()
            pdf.set_font(DEFAULT_FONT, style="B", size=11)
            pdf.cell(prop_lease_width, 8, "PROPNUM / LEASE_NAME")
            pdf.cell(cat_width, 8, "Begin Cat")
            pdf.cell(cat_width, 8, "Final Cat")
            pdf.cell(value_width, 8, "Value Change")
            pdf.cell(0, 8, "Explanation", ln=True)
            pdf.set_draw_color(200, 200, 200)
            pdf.line(10, pdf.get_y(), 200, pdf.get_y())
            pdf.set_font(DEFAULT_FONT, size=10)

    def check_outlier_page_break(pdf, needed_height):
        begin_width = 30
        final_width = 30
        outlier_width = 30
        if pdf.get_y() + needed_height > pdf.h - bottom_margin:
            pdf.add_page()
            pdf.set_font(DEFAULT_FONT, style="B", size=11)
            pdf.cell(prop_lease_width, 8, "PROPNUM / LEASE_NAME")
            pdf.cell(begin_width, 8, "Begin Ratio")
            pdf.cell(final_width, 8, "Final Ratio")
            pdf.cell(outlier_width, 8, "Outlier In", ln=True)
            pdf.set_draw_color(200, 200, 200)
            pdf.line(10, pdf.get_y(), 200, pdf.get_y())
            pdf.set_font(DEFAULT_FONT, size=10)

    # --- 1. Collect Plots ---
    metric_figs = []  # [(metric, fig), ...]
    metric_order = [m for m in MAIN_METRICS if m != npv_column]
    npv_fig = plot_top_contributors(variance_df, npv_column)
    for metric in metric_order:
        fig = plot_top_contributors(variance_df, metric)
        if fig:
            metric_figs.append((metric, fig))

    # --- 2. Write all summaries first ---
    for category in categories:
        pdf.add_page()
        pdf.set_font(DEFAULT_FONT, style="B", size=14)
        pdf.cell(0, 10, f"Variance Summary for {category}", ln=True, align='C')
        pdf.ln(2)

        pdf.set_font(DEFAULT_FONT, size=11)
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

        # Header row for main table
        pdf.set_font(DEFAULT_FONT, style="B", size=11)
        pdf.cell(prop_lease_width, 8, "PROPNUM / LEASE_NAME")
        pdf.cell(cat_width, 8, "Begin Cat")
        pdf.cell(cat_width, 8, "Final Cat")
        pdf.cell(value_width, 8, "Value Change")
        pdf.cell(0, 8, "Explanation", ln=True)
        pdf.set_draw_color(200, 200, 200)
        pdf.line(10, pdf.get_y(), 200, pdf.get_y())
        pdf.set_font(DEFAULT_FONT, size=10)

        if not category_df.empty:
            major_changes = category_df[
                category_df[f"{npv_column} Variance"].abs() > category_df[f"{npv_column} Variance"].abs().quantile(0.95)
            ]
            for _, row in major_changes.iterrows():
                value = row["Variance Value"]
                metric = row["Key Metric"]
                explanation = row["Explanation"]
                # Remove decimals for value changes
                if "Revenue" in metric or "$" in metric or "Capex" in metric or metric == npv_column:
                    val_str = f"${value:,.0f}"
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

    # --- 3. Add NPV plot as second page ---
    if npv_fig:
        add_chart_to_pdf(pdf, npv_fig, title=f"Top Contributors to {npv_column} Change")

    # --- 4. Add remaining metric plots in order ---
    for metric, fig in metric_figs:
        add_chart_to_pdf(pdf, fig, title=f"Top Contributors to {metric} Change")

    # --- 5. Add the rest: transitions and NRI outlier pages for all categories ---
    for category in categories:
        pdf.add_page()
        pdf.set_font(DEFAULT_FONT, style="B", size=14)
        pdf.cell(0, 10, f"Transitions and Outliers for {category}", ln=True, align='C')
        pdf.ln(2)

        pdf.set_font(DEFAULT_FONT, size=12)
        pdf.cell(0, 10, "Wells that Changed Reserve Category:", ln=True)
        pdf.set_font(DEFAULT_FONT, size=10)
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
            pdf.set_font(DEFAULT_FONT, style="B", size=12)
            pdf.cell(0, 10, f"NRI/WI Ratio Outliers for {category}", ln=True, align='C')
            pdf.ln(2)
            pdf.set_font(DEFAULT_FONT, style="B", size=11)
            begin_width = 30
            final_width = 30
            outlier_width = 30

            pdf.cell(60, 8, "PROPNUM / LEASE_NAME")
            pdf.cell(begin_width, 8, "Begin Ratio")
            pdf.cell(final_width, 8, "Final Ratio")
            pdf.cell(outlier_width, 8, "Outlier In", ln=True)
            pdf.set_draw_color(200, 200, 200)
            pdf.line(10, pdf.get_y(), 200, pdf.get_y())
            pdf.set_font(DEFAULT_FONT, size=10)

            for _, row in nri_outliers_category.iterrows():
                begin_ratio = row.get("NRI/WI Ratio Begin", None)
                final_ratio = row.get("NRI/WI Ratio Final", None)
                outlier_loc = row.get("Outlier Source", "-")
                well_id = str(row["PROPNUM"])
                lease_name = str(row["LEASE_NAME"])
                well_text = well_id + "\n" + lease_name
                well_lines = len(pdf.multi_cell(60, line_height, well_text, border=0, align='L', split_only=True))
                row_height = line_height * well_lines

                check_outlier_page_break(pdf, row_height)

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

    pdf.output(pdf_buffer)
    pdf_buffer.seek(0)
    return pdf_buffer

# ====== STREAMLIT APP ======

st.title("Variance Audit Tool")
st.write("Upload BEGIN and FINAL Excel files. Choose NPV column and generate reports.")

begin_file = st.file_uploader("Upload BEGIN Excel file (.xlsx)", type=["xlsx"])
final_file = st.file_uploader("Upload FINAL Excel file (.xlsx)", type=["xlsx"])

npv_column = None
excel_buffer = None
pdf_buffer = None

if begin_file and final_file:
    begin_df = load_oneline(begin_file)
    final_df = load_oneline(final_file)

    # Find available NPV columns
    npv_candidates = [col for col in begin_df.columns if col.startswith("NPV at")]
    selected_npv = st.selectbox("Select NPV column", npv_candidates)
    npv_column = selected_npv

    # Calculate everything only once for download buttons
    @st.cache_data(show_spinner=False)
    def compute_reports(begin_df, final_df, npv_column):
        begin_df_s = suffix_columns(begin_df, "_begin")
        final_df_s = suffix_columns(final_df, "_final")
        begin_df_s.columns = begin_df_s.columns.str.strip()
        final_df_s.columns = final_df_s.columns.str.strip()

        variance_df = begin_df_s.merge(final_df_s, on=["PROPNUM", "LEASE_NAME"], how="outer")

        # Fill missing category columns if needed
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

        excel_buffer = generate_excel(
            variance_df, npv_column, negative_npv_wells, begin_df_s, final_df_s, nri_df
        )
        pdf_buffer = generate_pdf(
            variance_df, npv_column, explanation_df, nri_df
        )
        return excel_buffer, pdf_buffer

    if st.button("Generate and Download Reports"):
        excel_buffer, pdf_buffer = compute_reports(begin_df, final_df, npv_column)
        st.download_button(
            label="Download Excel Report",
            data=excel_buffer,
            file_name=f"variance_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="excel_dl"
        )
        st.download_button(
            label="Download PDF Report",
            data=pdf_buffer,
            file_name=f"variance_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
            mime="application/pdf",
            key="pdf_dl"
        )
