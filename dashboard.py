import os
import io
from datetime import date

import pandas as pd
import streamlit as st
import altair as alt

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm

import gspread
from google.oauth2.service_account import Credentials

DATA_FILE = "responses.csv"

MSC_YELLOW = "#F8DE8D"
MSC_GREEN = "#00685E"
MSC_RED = "#A6192E"
MSC_GRAY = "#E6E6E6"
MSC_WARM_GREY = "#8B8178"
TEXT_COLOR = "#8C7F72"
BLACK = "#000000"
MSC_LIGHT_BLUE = "#8E9FBC"
MSC_BLUE = "#135193"
MSC_DARK_BLUE = "#1B365D"

COMPANY_DEPARTMENTS = [
    "Administration",
    "Customer Invoicing",
    "Finance & Accounting",
    "Commercial Reporting & BI",
    "Information Technology",
    "OVA",
    "Documentation, Pricing & Legal",
]

CATEGORIES = {
    "Workload & Recovery": ["q1", "q2", "q3"],
    "Team & Leadership": ["q4", "q5", "q6"],
    "Motivation & Wellbeing": ["q7", "q8", "q9"],
    "Work-Life Balance": ["q10", "q11", "q12"],
    "Growth & Recognition": ["q13", "q14", "q15"],
}

@st.cache_data(ttl=20)
def load_data():
    creds_info = st.secrets["gcp_service_account"]
    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
    client = gspread.authorize(creds)

    sheet_id = st.secrets["sheets"]["spreadsheet_id"]
    ws_name = st.secrets["sheets"]["worksheet_name"]
    ws = client.open_by_key(sheet_id).worksheet(ws_name)

    values = ws.get_all_values()
    if len(values) < 2:
        return pd.DataFrame()

    header = values[0]
    rows = values[1:]
    return pd.DataFrame(rows, columns=header)

st.set_page_config(page_title="MSC Latvia – Wellbeing Survey Dashboard", layout="wide")

st.markdown(
    f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Archivo:wght@400;700;900&display=swap');

html, body, [class*="st-"], .stApp {{
  font-family: 'Archivo', sans-serif !important;
  color: {TEXT_COLOR};
}}

h1, h2, h3 {{
  font-weight: 900 !important;
  text-transform: uppercase;
  letter-spacing: 1px;
  color: {TEXT_COLOR};
}}

div[data-testid="stMetricValue"] {{
  color: {TEXT_COLOR} !important;
  font-weight: 900 !important;
}}

section[data-testid="stSidebar"],
div[data-testid="stSidebar"],
section[data-testid="stSidebar"] > div,
div[data-testid="stSidebar"] > div {{
  background: {MSC_WARM_GREY} !important;
  border-right: 2px solid {MSC_YELLOW} !important;
}}

div[data-testid="stSidebar"] *,
section[data-testid="stSidebar"] * {{
  color: #ffffff !important;
}}

div[data-testid="stSidebar"] a,
div[data-testid="stSidebar"] a:visited,
section[data-testid="stSidebar"] a,
section[data-testid="stSidebar"] a:visited {{
  color: #ffffff !important;
}}

div[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] *,
section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] * {{
  color: #ffffff !important;
}}

div[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] label {{
  color: #ffffff !important;
}}

div[data-testid="stSidebar"] h1,
div[data-testid="stSidebar"] h2,
div[data-testid="stSidebar"] h3,
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 {{
  color: #ffffff !important;
}}

div[data-testid="stSidebar"] div[data-baseweb="select"] > div,
div[data-testid="stSidebar"] div[data-baseweb="input"] > div,
section[data-testid="stSidebar"] div[data-baseweb="select"] > div,
section[data-testid="stSidebar"] div[data-baseweb="input"] > div {{
  background: #ffffff22 !important;
  border-color: #ffffff55 !important;
}}

div[data-testid="stSidebar"] div[data-baseweb="select"] span,
section[data-testid="stSidebar"] div[data-baseweb="select"] span {{
  color: #ffffff !important;
}}

div[data-testid="stSidebar"] div[data-baseweb="select"] svg,
section[data-testid="stSidebar"] div[data-baseweb="select"] svg {{
  color: #ffffff !important;
}}

div[data-testid="stSidebar"] .stDateInput input,
section[data-testid="stSidebar"] .stDateInput input {{
  background: #ffffff22 !important;
  border-color: #ffffff55 !important;
  color: #ffffff !important;
}}

div[data-testid="stSidebar"] input::placeholder,
section[data-testid="stSidebar"] input::placeholder {{
  color: #ffffffcc !important;
}}

div[data-testid="stSidebar"] div[data-baseweb="popover"],
section[data-testid="stSidebar"] div[data-baseweb="popover"] {{
  color: {BLACK} !important;
}}

div[data-testid="stSidebar"] div[role="listbox"] *,
section[data-testid="stSidebar"] div[role="listbox"] * {{
  color: {BLACK} !important;
}}

hr {{
  border: none;
  border-top: 2px solid {MSC_YELLOW};
  margin: 1.2rem 0;
}}

button[kind="primary"] {{
  background: {MSC_YELLOW} !important;
  border-color: {MSC_YELLOW} !important;
  color: {BLACK} !important;
}}

button[kind="secondary"] {{
  border-color: #ffffff55 !important;
  color: #ffffff !important;
}}

div[data-testid="collapsedControl"] button,
div[data-testid="collapsedControl"] {{
  background: transparent !important;
}}

div[data-testid="collapsedControl"] button {{
  background: {MSC_YELLOW} !important;
  border: 2px solid {MSC_YELLOW} !important;
  border-radius: 10px !important;
  width: 44px !important;
  height: 38px !important;
  padding: 0 !important;
  display: inline-flex !important;
  align-items: center !important;
  justify-content: center !important;
}}

div[data-testid="collapsedControl"] button svg {{
  display: none !important;
}}

div[data-testid="collapsedControl"] button::before {{
  content: "✕";
  color: {BLACK} !important;
  font-weight: 900 !important;
  font-size: 22px !important;
  line-height: 1 !important;
}}

a, a:visited {{
  color: {TEXT_COLOR};
}}
</style>
""",
    unsafe_allow_html=True,
)

def _logo_path() -> str | None:
    candidates = [
        os.path.join("assets", "msc_logo_1.png"),
        os.path.join("assets", "msc_logo.png"),
        "msc_logo_1.png",
        "msc_logo.png",
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return None

def _msc_altair_theme():
    return {
        "config": {
            "background": "transparent",
            "title": {"font": "Archivo", "fontWeight": 900, "color": TEXT_COLOR},
            "axis": {
                "labelFont": "Archivo",
                "titleFont": "Archivo",
                "labelColor": TEXT_COLOR,
                "titleColor": TEXT_COLOR,
                "gridColor": "#F2F2F2",
                "tickColor": "#DDDDDD",
            },
            "legend": {
                "labelFont": "Archivo",
                "titleFont": "Archivo",
                "labelColor": TEXT_COLOR,
                "titleColor": TEXT_COLOR,
            },
            "view": {"stroke": "transparent"},
        }
    }

alt.themes.register("msc_theme", _msc_altair_theme)
alt.themes.enable("msc_theme")

top_left, top_right = st.columns([0.78, 0.22], vertical_alignment="center")
with top_left:
    st.title("Wellbeing Survey Dashboard")
    st.caption("MSC Latvia Internal Feedback — interactive overview")
with top_right:
    lp = _logo_path()
    if lp:
        st.image(lp, use_container_width=True)

@st.cache_data(ttl=20)
def load_data() -> pd.DataFrame:
    creds_info = st.secrets["gcp_service_account"]
    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
    client = gspread.authorize(creds)

    sheet_id = st.secrets["sheets"]["spreadsheet_id"]
    ws_name = st.secrets["sheets"]["worksheet_name"]
    ws = client.open_by_key(sheet_id).worksheet(ws_name)

    values = ws.get_all_values()
    if len(values) < 2:
        return pd.DataFrame()

    header = [c.strip() for c in values[0]]
    rows = values[1:]
    df = pd.DataFrame(rows, columns=header)

    if "timestamp" in df.columns:
        df["timestamp"] = pd.to_datetime(df["timestamp"], errors="coerce")

    for i in range(1, 16):
        col = f"q{i}"
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


df = load_data(DATA_FILE)
if df.empty:
    st.error(f"Can't find/read `{DATA_FILE}`. Make sure the file is in the same folder as `dashboard.py`.")
    st.stop()

question_cols = [c for c in df.columns if c.lower().startswith("q") and c[1:].isdigit()]
question_cols = sorted(question_cols, key=lambda x: int(x[1:]))

if not question_cols:
    st.error("No question columns found in the CSV file (`q1`, `q2`, ...).")
    st.stop()

for c in question_cols:
    df[c] = pd.to_numeric(df[c], errors="coerce")

has_department = "department" in df.columns
has_survey_date = "survey_date" in df.columns and df["survey_date"].notna().any()

with st.sidebar:
    st.header("Filters")

    selected_dept = "All"
    if has_department:
        selected_dept = st.selectbox("Select department", options=["All"] + COMPANY_DEPARTMENTS, index=0)

    if has_survey_date:
        min_d = df["survey_date"].dropna().min()
        max_d = df["survey_date"].dropna().max()
        dr = st.date_input("Date range", value=(min_d, max_d))
        if isinstance(dr, tuple) and len(dr) == 2:
            start_d, end_d = dr
        else:
            start_d, end_d = min_d, max_d
        if start_d > end_d:
            start_d, end_d = end_d, start_d

    st.divider()

    selected_q = st.selectbox("Question", options=question_cols, index=0)

    if st.button("Refresh data", type="primary"):
        st.cache_data.clear()
        st.rerun()

fdf = df.copy()

if has_department and selected_dept != "All":
    fdf = fdf[fdf["department"] == selected_dept]

if has_survey_date:
    fdf = fdf[(fdf["survey_date"].notna()) & (fdf["survey_date"] >= start_d) & (fdf["survey_date"] <= end_d)]

def compute_category_scores(df_in: pd.DataFrame) -> pd.DataFrame:
    out = df_in.copy()
    for cat, qs in CATEGORIES.items():
        cols = [q for q in qs if q in out.columns]
        if cols:
            out[cat] = out[cols].mean(axis=1, skipna=True)
    cat_cols = [c for c in CATEGORIES.keys() if c in out.columns]
    if cat_cols:
        out["Overall Index"] = out[cat_cols].mean(axis=1, skipna=True)
    return out

adf = compute_category_scores(fdf)
q_series = adf[selected_q].dropna() if selected_q in adf.columns else pd.Series(dtype=float)

def style_worst_per_department_row(df_in: pd.DataFrame):
    df_style = df_in.copy()
    num_cols = df_style.select_dtypes(include="number").columns.tolist()
    if not num_cols:
        return df_style.style
    row_mins = df_style[num_cols].min(axis=1)
    def _apply_row(row: pd.Series):
        styles = [""] * len(row)
        rmin = row_mins[row.name]
        for i, col in enumerate(row.index):
            if col in num_cols:
                v = row[col]
                if pd.notna(v) and pd.notna(rmin) and v == rmin:
                    styles[i] = f"background-color: {MSC_RED}; color: white; font-weight: 800;"
        return styles
    return df_style.style.apply(_apply_row, axis=1).format({c: "{:.2f}" for c in num_cols})

tab_overview, tab_question, tab_dept, tab_heatmap, tab_export = st.tabs(
    ["Overview", "Question Explorer", "Department Compare", "Heatmap", "Export"]
)

with tab_overview:
    st.subheader("Executive summary")

    n_rows = len(adf)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Responses (after filters)", f"{n_rows:,}")

    if "Overall Index" in adf.columns and adf["Overall Index"].notna().any():
        c2.metric("Overall Index (avg)", f"{adf['Overall Index'].mean():.2f}")
        c3.metric("Overall Index (median)", f"{adf['Overall Index'].median():.2f}")
        c4.metric("Overall Index (min / max)", f"{adf['Overall Index'].min():.2f} / {adf['Overall Index'].max():.2f}")
    else:
        c2.metric("Overall Index (avg)", "—")
        c3.metric("Overall Index (median)", "—")
        c4.metric("Overall Index (min / max)", "—")

    st.markdown("---")

    cat_cols = [c for c in CATEGORIES.keys() if c in adf.columns]
    if cat_cols and n_rows > 0:
        cat_avg = pd.DataFrame({"Category": cat_cols, "Average": [adf[c].mean() for c in cat_cols]}).dropna()

        st.subheader("Category averages (after filters)")
        cat_chart = (
            alt.Chart(cat_avg)
            .mark_bar(color=MSC_YELLOW, cornerRadiusTopLeft=6, cornerRadiusTopRight=6)
            .encode(
                x=alt.X("Category:N", sort="-y", title="Category"),
                y=alt.Y("Average:Q", title="Average (1–10)"),
                tooltip=[alt.Tooltip("Category:N"), alt.Tooltip("Average:Q", format=".2f")],
            )
            .properties(height=360)
        )
        st.altair_chart(cat_chart, use_container_width=True)
    else:
        st.info("Insufficient data to display category summary (check if CSV has q1..q15).")

    if has_survey_date and "Overall Index" in adf.columns:
        tdf = adf[["survey_date", "Overall Index"]].dropna()
        if not tdf.empty:
            st.subheader("Overall Index trend (daily average)")
            tdf = tdf.groupby("survey_date", as_index=False)["Overall Index"].mean().rename(columns={"Overall Index": "avg"})
            line = (
                alt.Chart(tdf)
                .mark_line(point=alt.OverlayMarkDef(filled=True), color=MSC_YELLOW, strokeWidth=3)
                .encode(
                    x=alt.X("survey_date:T", title="Date"),
                    y=alt.Y("avg:Q", title="Average"),
                    tooltip=[alt.Tooltip("survey_date:T", title="Date"), alt.Tooltip("avg:Q", title="Avg", format=".2f")],
                )
                .properties(height=300)
            )
            st.altair_chart(line, use_container_width=True)

with tab_question:
    st.subheader(f"Question Explorer — {selected_q}")

    colA, colB, colC, colD = st.columns(4)
    colA.metric("N (valid)", f"{len(q_series):,}")
    colB.metric("Average", f"{q_series.mean():.2f}" if not q_series.empty else "—")
    colC.metric("Median", f"{q_series.median():.0f}" if not q_series.empty else "—")
    colD.metric("Min / Max", f"{int(q_series.min())} / {int(q_series.max())}" if not q_series.empty else "—")

    st.markdown("---")

    left, right = st.columns([1.1, 1])

    with left:
        st.markdown("#### Distribution by department")
        if has_department and not adf.empty:
            by_dept = adf.groupby("department", dropna=False)[selected_q].agg(["count", "mean", "median"]).reset_index()
            by_dept = by_dept[by_dept["count"] > 0]

            if by_dept.empty:
                st.info("No data for the selected filters.")
            else:
                by_dept_disp = by_dept.rename(columns={"department": "Department"})
                bar = (
                    alt.Chart(by_dept_disp)
                    .mark_bar(color=MSC_YELLOW, cornerRadiusTopLeft=6, cornerRadiusTopRight=6)
                    .encode(
                        x=alt.X("Department:N", sort="-y", title="Department"),
                        y=alt.Y("mean:Q", title="Average"),
                        tooltip=[
                            alt.Tooltip("Department:N", title="Department"),
                            alt.Tooltip("count:Q", title="N"),
                            alt.Tooltip("mean:Q", title="Avg", format=".2f"),
                            alt.Tooltip("median:Q", title="Median", format=".0f"),
                        ],
                    )
                    .properties(height=360)
                )
                st.altair_chart(bar, use_container_width=True)
        else:
            st.info("No `department` column or no data after filters.")

    with right:
        st.markdown("#### Response distribution (histogram)")
        if q_series.empty:
            st.info("No data for the selected filters.")
        else:
            hist_df = adf[[selected_q]].dropna().rename(columns={selected_q: "Score"})
            hist = (
                alt.Chart(hist_df)
                .mark_bar(color=MSC_YELLOW, cornerRadiusTopLeft=6, cornerRadiusTopRight=6)
                .encode(
                    x=alt.X("Score:O", title="Score (1–10)"),
                    y=alt.Y("count():Q", title="Count"),
                    tooltip=[alt.Tooltip("Score:O", title="Score"), alt.Tooltip("count():Q", title="Count")],
                )
                .properties(height=360)
            )
            st.altair_chart(hist, use_container_width=True)

    if has_survey_date and selected_q in adf.columns:
        tdf = adf[["survey_date", selected_q]].dropna()
        if not tdf.empty:
            st.markdown("#### Trend over time (daily average)")
            tdf = tdf.groupby("survey_date", as_index=False)[selected_q].mean().rename(columns={selected_q: "avg"})
            line = (
                alt.Chart(tdf)
                .mark_line(point=alt.OverlayMarkDef(filled=True), color=MSC_YELLOW, strokeWidth=3)
                .encode(
                    x=alt.X("survey_date:T", title="Date"),
                    y=alt.Y("avg:Q", title="Average"),
                    tooltip=[alt.Tooltip("survey_date:T", title="Date"), alt.Tooltip("avg:Q", title="Avg", format=".2f")],
                )
                .properties(height=280)
            )
            st.altair_chart(line, use_container_width=True)

with tab_dept:
    st.subheader("Department comparison")

    if not has_department:
        st.info("The CSV file does not contain a `department` column.")
    elif adf.empty:
        st.info("No data after filters.")
    else:
        metric_mode = st.radio("Compare", ["Overall Index", "Category scores", "All questions (avg)"], horizontal=True)

        if metric_mode == "Overall Index":
            if "Overall Index" not in adf.columns:
                st.info("No `Overall Index` found (check if categories contain q1..q15).")
            else:
                comp = adf.groupby("department", as_index=False)["Overall Index"].mean().dropna()
                comp = comp.sort_values("Overall Index", ascending=False)
                comp_disp = comp.rename(columns={"department": "Department"})

                chart = (
                    alt.Chart(comp_disp)
                    .mark_bar(color=MSC_YELLOW, cornerRadiusTopLeft=6, cornerRadiusTopRight=6)
                    .encode(
                        x=alt.X("Department:N", sort="-y", title="Department"),
                        y=alt.Y("Overall Index:Q", title="Average Overall Index"),
                        tooltip=[alt.Tooltip("Department:N"), alt.Tooltip("Overall Index:Q", format=".2f")],
                    )
                    .properties(height=420)
                )
                st.altair_chart(chart, use_container_width=True)
                st.dataframe(comp_disp, use_container_width=True)

        elif metric_mode == "Category scores":
            cat_cols = [c for c in CATEGORIES.keys() if c in adf.columns]
            if not cat_cols:
                st.info("No category columns found (check q1..q15).")
            else:
                dept_cat = adf.groupby("department", as_index=False)[cat_cols].mean()
                dept_cat_disp = dept_cat.rename(columns={"department": "Department"})
                dept_cat_melt = dept_cat_disp.melt("Department", var_name="Category", value_name="Average").dropna()

                chart = (
                    alt.Chart(dept_cat_melt)
                    .mark_bar(color=MSC_YELLOW, cornerRadiusTopLeft=6, cornerRadiusTopRight=6)
                    .encode(
                        x=alt.X("Department:N", title="Department"),
                        y=alt.Y("Average:Q", title="Average (1–10)"),
                        tooltip=[alt.Tooltip("Department:N"), alt.Tooltip("Category:N"), alt.Tooltip("Average:Q", format=".2f")],
                    )
                    .properties(height=460)
                )
                st.altair_chart(chart, use_container_width=True)
                st.dataframe(dept_cat_disp, use_container_width=True)

        else:
            dept_q = adf.groupby("department", as_index=False)[question_cols].mean()
            dept_q_disp = dept_q.rename(columns={"department": "Department"})
            st.dataframe(style_worst_per_department_row(dept_q_disp), use_container_width=True)

with tab_heatmap:
    st.subheader("Heatmap — questions × departments (average)")

    if not has_department:
        st.info("The CSV file does not contain a `department` column.")
    elif adf.empty:
        st.info("No data available for the applied filters.")
    else:
        heat = (
            adf.melt(id_vars=["department"], value_vars=question_cols, var_name="Question", value_name="Score")
            .dropna(subset=["Score"])
            .groupby(["department", "Question"], as_index=False)["Score"]
            .mean()
        )

        if heat.empty:
            st.info("Insufficient data to generate the heatmap with the applied filters.")
        else:
            heat["department"] = pd.Categorical(heat["department"], categories=COMPANY_DEPARTMENTS, ordered=True)
            heat = heat.sort_values(["department", "Question"]).dropna(subset=["department"])
            heat_disp = heat.rename(columns={"department": "Department"})

            hm = (
                alt.Chart(heat_disp)
                .mark_rect(stroke="#FFFFFF", strokeWidth=0.8, cornerRadius=6)
                .encode(
                    x=alt.X("Question:N", title="Question", sort=question_cols, axis=alt.Axis(labelAngle=0, labelPadding=10)),
                    y=alt.Y("Department:N", title="Department", sort=COMPANY_DEPARTMENTS, axis=alt.Axis(labelPadding=10)),
                    color=alt.Color(
                        "Score:Q",
                        title="Average",
                        scale=alt.Scale(domain=[1, 5.5, 10], range=[MSC_LIGHT_BLUE, MSC_BLUE, MSC_DARK_BLUE]),
                        legend=alt.Legend(gradientLength=200),
                    ),
                    tooltip=[
                        alt.Tooltip("Department:N", title="Department"),
                        alt.Tooltip("Question:N", title="Question"),
                        alt.Tooltip("Score:Q", title="Average", format=".2f"),
                    ],
                )
                .properties(height=min(540, 44 * max(5, heat_disp["Department"].nunique())))
            )

            labels = (
                alt.Chart(heat_disp)
                .mark_text(font="Archivo", fontSize=11, fontWeight=700)
                .encode(
                    x=alt.X("Question:N", sort=question_cols),
                    y=alt.Y("Department:N", sort=COMPANY_DEPARTMENTS),
                    text=alt.Text("Score:Q", format=".1f"),
                    color=alt.condition("datum.Score >= 7.0", alt.value("white"), alt.value(BLACK)),
                )
            )

            st.altair_chart(hm + labels, use_container_width=True)

def build_excel_bytes(df_filtered: pd.DataFrame, df_scored: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    df_filtered_out = df_filtered.rename(columns={"department": "Department"}).copy() if "department" in df_filtered.columns else df_filtered.copy()
    df_scored_out = df_scored.rename(columns={"department": "Department"}).copy() if "department" in df_scored.columns else df_scored.copy()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_filtered_out.to_excel(writer, index=False, sheet_name="Filtered Raw")

        q_avg = df_scored_out[question_cols].mean().to_frame("Average").reset_index().rename(columns={"index": "Question"})
        q_avg.to_excel(writer, index=False, sheet_name="Question Avg")

        if has_department:
            cols = [selected_q] + (["Overall Index"] if "Overall Index" in df_scored_out.columns else [])
            cols = [c for c in cols if c in df_scored_out.columns]
            if cols:
                dept_avg = df_scored.groupby("department", as_index=False)[cols].mean()
                dept_avg = dept_avg.rename(columns={"department": "Department"})
                dept_avg.to_excel(writer, index=False, sheet_name="Dept Avg")

        cat_cols = [c for c in CATEGORIES.keys() if c in df_scored_out.columns]
        if cat_cols:
            cat_avg = df_scored_out[cat_cols].mean().to_frame("Average").reset_index().rename(columns={"index": "Category"})
            cat_avg.to_excel(writer, index=False, sheet_name="Category Avg")

    return output.getvalue()

def build_pdf_bytes(df_scored: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4

    font_bold = "Helvetica-Bold"
    font_regular = "Helvetica"

    c.setFont(font_bold, 16)
    c.drawString(2 * cm, h - 2.2 * cm, "MSC Latvia – Wellbeing Survey (Summary)")

    c.setFont(font_regular, 10)
    c.drawString(2 * cm, h - 2.9 * cm, f"Generated: {date.today().isoformat()}")
    c.drawString(2 * cm, h - 3.5 * cm, f"Rows (after filters): {len(df_scored)}")

    y = h - 5.2 * cm

    if "Overall Index" in df_scored.columns and df_scored["Overall Index"].notna().any():
        c.setFont(font_bold, 12)
        c.drawString(2 * cm, y, "Overall Index")
        y -= 0.6 * cm
        c.setFont(font_regular, 10)
        c.drawString(2 * cm, y, f"Average: {df_scored['Overall Index'].mean():.2f}")
        y -= 0.5 * cm
        c.drawString(2 * cm, y, f"Median: {df_scored['Overall Index'].median():.2f}")
        y -= 0.9 * cm

    cat_cols = [cc for cc in CATEGORIES.keys() if cc in df_scored.columns]
    if cat_cols:
        c.setFont(font_bold, 12)
        c.drawString(2 * cm, y, "Category averages (1–10)")
        y -= 0.6 * cm
        c.setFont(font_regular, 10)

        for cc in cat_cols:
            avg = df_scored[cc].mean()
            if pd.notna(avg):
                c.drawString(2 * cm, y, f"- {cc}: {avg:.2f}")
                y -= 0.45 * cm
                if y < 2.5 * cm:
                    c.showPage()
                    y = h - 2.5 * cm
                    c.setFont(font_regular, 10)

        y -= 0.4 * cm

    if has_department and "Overall Index" in df_scored.columns:
        dept_avg = df_scored.groupby("department", as_index=False)["Overall Index"].mean().dropna()
        dept_avg = dept_avg.sort_values("Overall Index", ascending=False)

        if not dept_avg.empty:
            if y < 6 * cm:
                c.showPage()
                y = h - 2.5 * cm

            c.setFont(font_bold, 12)
            c.drawString(2 * cm, y, "Department averages (Overall Index)")
            y -= 0.7 * cm
            c.setFont(font_regular, 10)

            for _, row in dept_avg.iterrows():
                c.drawString(2 * cm, y, f"- {row['department']}: {row['Overall Index']:.2f}")
                y -= 0.45 * cm
                if y < 2.5 * cm:
                    c.showPage()
                    y = h - 2.5 * cm
                    c.setFont(font_regular, 10)

    c.showPage()
    c.save()
    return buf.getvalue()

with tab_export:
    st.subheader("Export")

    if adf.empty:
        st.info("No data available for the applied filters.")
    else:
        excel_bytes = build_excel_bytes(fdf, adf)
        st.download_button(
            label="Download Excel (filtered + summaries)",
            data=excel_bytes,
            file_name="msc_wellbeing_dashboard_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

        pdf_bytes = build_pdf_bytes(adf)
        st.download_button(
            label="Download PDF (summary)",
            data=pdf_bytes,
            file_name="msc_wellbeing_summary.pdf",
            mime="application/pdf",
            type="primary",
        )
