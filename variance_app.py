import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import io
import zipfile
import os
import tempfile
from fpdf import FPDF
import matplotlib.ticker as mtick
from matplotlib.ticker import FuncFormatter


# ==== CONFIGURATION ====
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

# right below MAIN_METRICS
PLOT_METRICS = [m for m in MAIN_METRICS 
                if m not in ("BFIT IRR (%)","BFIT Payout (years)")]


# Path to your logo file in the repo root:
LOGO_PATH = "logo-schaperintl-1.png"

# ==== UTILS ====
def load_oneline(file):
    xl = pd.ExcelFile(file)
    if "Oneline" not in xl.sheet_names:
        raise ValueError(f"'Oneline' sheet not found. Available: {xl.sheet_names}")
    df = xl.parse("Oneline")
    df.columns = df.columns.str.strip()
    for req in ("PROPNUM", "LEASE_NAME"):
        if req not in df.columns:
            raise ValueError(f"Required column '{req}' missing in Oneline sheet.")
    if "SE_RSV_CAT" not in df.columns:
        df["SE_RSV_CAT"] = "Unknown"
    return df

def suffix_columns(df, suffix, ignore=["PROPNUM","LEASE_NAME"]):
    return df.rename(columns={c: f"{c}{suffix}" for c in df.columns if c not in ignore})

def identify_negative_npv_wells(var_df, npv_col):
    return var_df[
        (var_df.get(f"{npv_col}_begin", 0) > 0) &
        (var_df.get(f"{npv_col}_final", 0) <= 0)
    ]

def calculate_nri_wi_ratio(begin_df, final_df):
    def _ratio(df, wi, nri, tag):
        df = df[df[wi] != 0]
        df[f"NRI/WI Ratio {tag}"] = df[nri] / df[wi]
        return df[["PROPNUM", "LEASE_NAME", f"NRI/WI Ratio {tag}"]]
    b = _ratio(begin_df, "Inital Approx WI_begin", "Initial Approx NRI_begin", "Begin")
    f = _ratio(final_df, "Inital Approx WI_final", "Initial Approx NRI_final", "Final")
    merged = b.merge(f, on=["PROPNUM", "LEASE_NAME"], how="outer")
    def oob(x): return pd.notna(x) and (x < 0.7 or x > 0.85)
    merged["Outlier Source"] = merged.apply(
        lambda r: "Both" if oob(r["NRI/WI Ratio Begin"]) and oob(r["NRI/WI Ratio Final"])
        else "Begin" if oob(r["NRI/WI Ratio Begin"])
        else "Final" if oob(r["NRI/WI Ratio Final"])
        else None,
        axis=1
    )
    return merged[merged["Outlier Source"].notna()]

# ==== EXPLANATIONS ====
def generate_explanations(var_df, npv_col):
    base_metrics = [
        "Net Res Oil (Mbbl)",
        "Net Res Gas (MMcf)",
        "Net Res NGL (Mbbl)",
        "Net Total Revenue ($)",
        "Net Operating Expense ($)",
        "Net Capex ($)"
    ]
    METRICS = base_metrics + [npv_col]
    rows = []
    for _, r in var_df.iterrows():
        drivers = []
        for m in METRICS:
            b_col, f_col = f"{m}_begin", f"{m}_final"
            if b_col in r and f_col in r:
                vb, vf = r[b_col], r[f_col]
                if pd.notna(vb) and vb != 0 and pd.notna(vf):
                    delta = vf - vb
                    pct   = delta / abs(vb) * 100
                    drivers.append((m, delta, pct))
        if not drivers:
            rows.append({
                "PROPNUM": r["PROPNUM"],
                "LEASE_NAME": r["LEASE_NAME"],
                "Key Metric": "",
                "Variance Value": 0,
                "Explanation": ""
            })
            continue
        top3 = sorted(drivers, key=lambda x: abs(x[2]), reverse=True)[:3]
        km, kv, _ = top3[0]
        parts = []
        for m, d, p in top3:
            sign = "increased" if d > 0 else "decreased"
            if m.endswith("$") or "NPV" in m:
                parts.append(f"{m} {sign} by ${abs(d):,.0f} ({p:.0f}%)")
            else:
                parts.append(f"{m} {sign} by {abs(d):,.0f} ({p:.0f}%)")
        explanation = "; ".join(parts) + "."
        rows.append({
            "PROPNUM": r["PROPNUM"],
            "LEASE_NAME": r["LEASE_NAME"],
            "Key Metric": km,
            "Variance Value": kv,
            "Explanation": explanation
        })
    return pd.DataFrame(rows)

# ==== PLOTTING ====

def plot_top_contributors(var_df, metric, top_n=10):
    # … your existing extraction + scaling logic …

    # decide on label
    if metric in ("Net Total Revenue ($)",
                  "Net Operating Expense ($)",
                  "Net Capex ($)"):
        # remove the " ($)" suffix before appending " ($M)"
        clean = metric.replace(" ($)", "")
        xlabel = f"Change in {clean} ($M)"
    elif "NPV" in metric:
        xlabel = f"Change in {metric} ($M)"
    else:
        xlabel = f"Change in {metric}"

    # plotting…
    fig, ax = plt.subplots(...)
    ax.barh(labels, values, color=colors)

    # apply decimal formatter
    fmt = FuncFormatter(lambda x, _: f"{x:,.0f}")
    ax.xaxis.set_major_formatter(fmt)

    ax.set_xlabel(xlabel)
    ax.set_ylabel("Well (PROPNUM / LEASE_NAME)")
    ax.set_title(f"Top Contributors to {metric} Change")
    ax.invert_yaxis()
    plt.tight_layout()
    return fig





def add_chart_to_pdf(pdf, fig, title=""):
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
        fig.savefig(tmp.name, bbox_inches="tight", dpi=150)
        plt.close(fig)
        pdf.add_page()
        pdf.set_font("Times", "B", 14)
        if title:
            pdf.cell(0, 10, title, ln=True, align="C")
            pdf.ln(4)
        pdf.image(tmp.name, x=15, w=180)
        os.unlink(tmp.name)

# ==== EXCEL EXPORT ====
def generate_excel(var_df, buf, npv_col, neg_df, b_df, f_df, nri_df):
    cols = [
        "Net Total Revenue ($) Variance",
        "Net Operating Expense ($) Variance",
        "Net Capex ($) Variance",
        "Net Res Oil (Mbbl) Variance",
        "Net Res Gas (MMcf) Variance",
        "Net Res NGL (Mbbl) Variance",
        f"{npv_col} Variance",
        "Reserve Category Begin","Reserve Category Final"
    ]
    summary = var_df[["PROPNUM","LEASE_NAME"] + cols]\
              .sort_values(f"{npv_col} Variance", ascending=False)
    bset, fset = set(b_df["PROPNUM"]), set(f_df["PROPNUM"])
    added   = f_df[f_df["PROPNUM"].isin(fset - bset)].copy()
    removed = b_df[b_df["PROPNUM"].isin(bset - fset)].copy()
    added["NPV"]   = added.get(f"{npv_col}_final", np.nan)
    removed["NPV"] = removed.get(f"{npv_col}_begin",  np.nan)
    with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
        summary.to_excel(w, sheet_name="VarianceSummary", index=False)
        neg_df[["PROPNUM","LEASE_NAME"]].to_excel(w, sheet_name="Neg_or_Zero_NPV", index=False)
        added[["PROPNUM","LEASE_NAME","NPV"]].to_excel(w, sheet_name="Added_Wells", index=False)
        removed[["PROPNUM","LEASE_NAME","NPV"]].to_excel(w, sheet_name="Removed_Wells", index=False)
        nri_df[["PROPNUM","LEASE_NAME","NRI/WI Ratio Begin","NRI/WI Ratio Final","Outlier Source"]]\
              .to_excel(w, sheet_name="Lease_NRI", index=False)
        for m in MAIN_METRICS + [npv_col]:
            col = f"{m} Variance"
            if col in var_df:
                tmp = var_df[["PROPNUM","LEASE_NAME",col]].dropna()\
                         .query(f"`{col}`!=0")\
                         .sort_values(col, ascending=False)
                if not tmp.empty:
                    pos = tmp.head(10); neg = tmp.tail(10)
                    combo = pd.concat([pos, neg])
                    name = f"Top_{m}_Ctrb"[:31]
                    combo.to_excel(w, sheet_name=name, index=False)
    buf.seek(0)

# ==== PDF EXPORT ====
class PDF(FPDF):
    def __init__(self, logo_path=None):
        super().__init__()
        self.logo_path = logo_path

    def footer(self):
        self.set_y(-15)
        self.set_font("Times", "I", 8)
        self.cell(0, 10, f"Page {self.page_no()}", 0, 0, "C")
        if self.logo_path and os.path.exists(self.logo_path):
            logo_w = 15
            x = self.w - self.r_margin - logo_w
            y = self.h - logo_w - 5
            self.image(self.logo_path, x=x, y=y, w=logo_w)

def generate_pdf(var_df, buf, npv_col, expl_df, nri_df):
    pdf = PDF(logo_path=LOGO_PATH)
    df = var_df.merge(expl_df, on=["PROPNUM","LEASE_NAME"], how="left")
    for col in ["SE_RSV_CAT_begin","SE_RSV_CAT_final"]:
        if col not in df: df[col] = "Unknown"
    cats = pd.unique(pd.concat([df["SE_RSV_CAT_begin"], df["SE_RSV_CAT_final"]]))
    cats = [c for c in cats if pd.notna(c)]
    if not cats: cats = ["Summary"]

    pw, cw, vw, lh, bm = 60, 22, 26, 5, 15
    def page_break(h):
        if pdf.get_y() + h > pdf.h - bm:
            pdf.add_page()
            pdf.set_font("Times", "B", 11)
            pdf.cell(pw, 8, "PROPNUM / LEASE_NAME")
            pdf.cell(cw, 8, "Begin Cat")
            pdf.cell(cw, 8, "Final Cat")
            pdf.cell(vw, 8, "Value Change")
            pdf.cell(0, 8, "Explanation", ln=True)
            pdf.set_draw_color(200, 200, 200)
            pdf.line(10, pdf.get_y(), 200, pdf.get_y())
            pdf.set_font("Times", "", 10)

    # 1) Variance Summaries
    for cat in cats:
        subgroup = df[(df["SE_RSV_CAT_begin"]==cat)|(df["SE_RSV_CAT_final"]==cat)]
        pdf.add_page()
        pdf.set_font("Times", "B", 14)
        pdf.cell(0,10,f"Variance Summary for {cat}", ln=True, align="C")
        pdf.ln(2)
        pdf.set_font("Times","",11)
        lines = [
            f"Net Oil Change: {subgroup['Net Res Oil (Mbbl) Variance'].sum():,.2f} Mbbl",
            f"Net Gas Change: {subgroup['Net Res Gas (MMcf) Variance'].sum():,.2f} MMcf",
            f"{npv_col} Change: ${subgroup[f'{npv_col} Variance'].sum():,.0f}"
        ]
        for L in lines: pdf.cell(0,8,L,ln=True)
        pdf.ln(4)
        pdf.set_font("Times","B",11)
        pdf.cell(pw,8,"PROPNUM / LEASE_NAME")
        pdf.cell(cw,8,"Begin Cat")
        pdf.cell(cw,8,"Final Cat")
        pdf.cell(vw,8,"Value Change")
        pdf.cell(0,8,"Explanation",ln=True)
        pdf.set_draw_color(200,200,200)
        pdf.line(10,pdf.get_y(),200,pdf.get_y())
        pdf.set_font("Times","",10)
        if not subgroup.empty:
            top5 = subgroup.loc[subgroup[f"{npv_col} Variance"].abs().nlargest(5).index]
            for _, r in top5.iterrows():
                npv_var = r[f"{npv_col} Variance"]
                val_str = f"${int(round(npv_var)):,.0f}"
                well_text = f"{r['PROPNUM']}\n{r['LEASE_NAME']}"
                expl = r["Explanation"]
                hl = len(pdf.multi_cell(pw,lh,well_text,split_only=True))
                el = len(pdf.multi_cell(0,lh,expl,split_only=True))
                rh = max(hl*lh, el*lh)
                page_break(rh)
                x0, y0 = pdf.get_x(), pdf.get_y()
                pdf.set_xy(x0,y0); pdf.multi_cell(pw,lh,well_text)
                pdf.set_xy(x0+pw,y0)
                pdf.cell(cw,rh,str(r["SE_RSV_CAT_begin"]))
                pdf.cell(cw,rh,str(r["SE_RSV_CAT_final"]))
                pdf.cell(vw,rh,val_str)
                pdf.set_xy(x0+pw+2*cw+vw,y0)
                pdf.multi_cell(0,lh,expl)
                pdf.set_draw_color(220,220,220)
                pdf.line(10,pdf.get_y(),200,pdf.get_y())
                pdf.ln(1)

    # 2) Plots
    # inside generate_pdf, after the Variance summaries...
    # 2) Plots
    npv_fig = plot_top_contributors(df, npv_col)
    if npv_fig:
        add_chart_to_pdf(pdf, npv_fig, f"Top Contributors to {npv_col} Change")

    for m in PLOT_METRICS:
        fig = plot_top_contributors(df, m)
        if fig:
            add_chart_to_pdf(pdf, fig, f"Top Contributors to {m} Change")

    # 3) Transitions & Outliers
    for cat in cats:
        subgroup = df[(df["SE_RSV_CAT_begin"]==cat)|(df["SE_RSV_CAT_final"]==cat)]
        pdf.add_page()
        pdf.set_font("Times","B",14)
        pdf.cell(0,10,f"Transitions & Outliers for {cat}",ln=True,align="C")
        pdf.ln(2)
        pdf.set_font("Times","",12)
        pdf.cell(0,10,"Wells that Changed Reserve Category:",ln=True)
        pdf.set_font("Times","",10)
        for _, r in subgroup[subgroup["SE_RSV_CAT_begin"]!=subgroup["SE_RSV_CAT_final"]].iterrows():
            pdf.cell(0,6,f"{r['PROPNUM']} ({r['LEASE_NAME']})",ln=True)
            pdf.cell(0,6,f"  From {r['SE_RSV_CAT_begin']} to {r['SE_RSV_CAT_final']}",ln=True)
            pdf.set_draw_color(220,220,220)
            pdf.line(10,pdf.get_y(),200,pdf.get_y())
            pdf.ln(1)
        out = nri_df[nri_df["PROPNUM"].isin(subgroup["PROPNUM"])]
        if not out.empty:
            pdf.set_font("Times","B",12)
            pdf.cell(0,10,f"NRI/WI Ratio Outliers for {cat}",ln=True,align="C")
            pdf.ln(2)
            pdf.set_font("Times","B",11)
            pdf.cell(60,8,"PROPNUM / LEASE_NAME")
            pdf.cell(30,8,"Begin Ratio")
            pdf.cell(30,8,"Final Ratio")
            pdf.cell(30,8,"Outlier In",ln=True)
            pdf.set_draw_color(200,200,200)
            pdf.line(10,pdf.get_y(),200,pdf.get_y())
            pdf.set_font("Times","",10)
            for _, r in out.iterrows():
                well_text = f"{r['PROPNUM']}\n{r['LEASE_NAME']}"
                hl = len(pdf.multi_cell(60,lh,well_text,split_only=True))
                rh = hl*lh
                x0, y0 = pdf.get_x(), pdf.get_y()
                pdf.multi_cell(60,lh,well_text)
                pdf.set_xy(x0+60,y0)
                pdf.cell(30,rh,f"{r['NRI/WI Ratio Begin']:.3f}")
                pdf.cell(30,rh,f"{r['NRI/WI Ratio Final']:.3f}")
                pdf.cell(30,rh,r["Outlier Source"])
                pdf.ln(rh)
                pdf.set_draw_color(220,220,220)
                pdf.line(10,pdf.get_y(),200,pdf.get_y())

    buf.write(pdf.output(dest="S").encode("latin1"))
    buf.seek(0)

# ==== STREAMLIT UI ====
st.title("Variance Audit Tool")
st.write("Upload BEGIN and FINAL ‘Oneline’ sheets, select NPV column, then generate.")

begin_u = st.file_uploader("BEGIN XLSX", type=["xlsx"], key="b")
final_u = st.file_uploader("FINAL XLSX", type=["xlsx"], key="f")
npv_col = st.selectbox("NPV column", ["NPV at 9%", "NPV at 10%"])
if "zip" not in st.session_state:
    st.session_state.zip = None

if st.button("Generate Reports"):
    try:
        bdf = load_oneline(begin_u)
        fdf = load_oneline(final_u)
    except Exception as e:
        st.error(str(e))
        st.stop()

    bdf_s = suffix_columns(bdf, "_begin")
    fdf_s = suffix_columns(fdf, "_final")
    var_df = bdf_s.merge(fdf_s, on=["PROPNUM", "LEASE_NAME"], how="outer")
    for c in ["SE_RSV_CAT_begin", "SE_RSV_CAT_final"]:
        if c not in var_df:
            var_df[c] = "Unknown"
    var_df["Reserve Category Begin"] = var_df["SE_RSV_CAT_begin"]
    var_df["Reserve Category Final"] = var_df["SE_RSV_CAT_final"]

    for m in MAIN_METRICS + [npv_col]:
        cb, cf = f"{m}_begin", f"{m}_final"
        if cb in var_df and cf in var_df:
            var_df[f"{m} Variance"] = var_df[cf] - var_df[cb]

    expl_df = generate_explanations(var_df, npv_col)
    neg_df  = identify_negative_npv_wells(var_df, npv_col)
    nri_df  = calculate_nri_wi_ratio(bdf_s, fdf_s)

    xbuf = io.BytesIO()
    generate_excel(var_df, xbuf, npv_col, neg_df, bdf_s, fdf_s, nri_df)
    pbuf = io.BytesIO()
    generate_pdf(var_df, pbuf, npv_col, expl_df, nri_df)

    z = io.BytesIO()
    with zipfile.ZipFile(z, "w") as zipf:
        zipf.writestr("variance_report.xlsx", xbuf.getvalue())
        zipf.writestr("variance_report.pdf", pbuf.getvalue())
    z.seek(0)
    st.session_state.zip = z.getvalue()
    st.success("Done! Scroll down to download.")

if st.session_state.zip:
    st.download_button(
        "Download ZIP",
        st.session_state.zip,
        file_name="variance_reports.zip",
        mime="application/zip"
    )
