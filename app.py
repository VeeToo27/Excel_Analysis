"""
Student Result Analysis System
Professional Light Theme – Streamlit App
"""

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import plotly.express as px
import plotly.graph_objects as go
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────
# PAGE CONFIG  (must be first Streamlit call)
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Student Result Analysis",
    layout="wide",
    page_icon="🎓",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# GLOBAL CSS  – force light theme + custom look
# ─────────────────────────────────────────────
st.markdown("""
<style>
/* ── Force Light Theme ───────────────────────────────────── */
[data-testid="stAppViewContainer"],
[data-testid="stApp"],
.main, .block-container {
    background-color: #F7F8FC !important;
    color: #1A1D2E !important;
}
[data-testid="stSidebar"] {
    background-color: #FFFFFF !important;
    border-right: 1px solid #E2E6F0;
}
[data-testid="stSidebar"] * { color: #1A1D2E !important; }

/* ── Google Font Import ──────────────────────────────────── */
@import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@600;700&family=DM+Sans:wght@300;400;500;600&display=swap');

/* ── Base Typography ─────────────────────────────────────── */
html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif !important;
    color: #1A1D2E !important;
}

/* ── App Header Banner ───────────────────────────────────── */
.app-header {
    background: linear-gradient(135deg, #1A1D2E 0%, #2D3561 60%, #3D5A80 100%);
    border-radius: 16px;
    padding: 36px 44px;
    margin-bottom: 28px;
    display: flex;
    align-items: center;
    gap: 20px;
    box-shadow: 0 8px 32px rgba(26,29,46,0.18);
    position: relative;
    overflow: hidden;
}
.app-header::before {
    content: '';
    position: absolute;
    top: -40px; right: -40px;
    width: 200px; height: 200px;
    background: rgba(255,255,255,0.04);
    border-radius: 50%;
}
.app-header::after {
    content: '';
    position: absolute;
    bottom: -60px; left: 30%;
    width: 280px; height: 280px;
    background: rgba(255,255,255,0.03);
    border-radius: 50%;
}
.app-header h1 {
    font-family: 'Playfair Display', serif !important;
    font-size: 2.2rem !important;
    font-weight: 700 !important;
    color: #FFFFFF !important;
    margin: 0 !important;
    line-height: 1.2;
}
.app-header p {
    color: rgba(255,255,255,0.7) !important;
    margin: 6px 0 0 0 !important;
    font-size: 0.95rem;
    font-weight: 300;
    letter-spacing: 0.3px;
}
.header-icon {
    font-size: 3rem;
    line-height: 1;
    z-index: 1;
}

/* ── Section Title ───────────────────────────────────────── */
.section-title {
    font-family: 'Playfair Display', serif;
    font-size: 1.35rem;
    font-weight: 600;
    color: #1A1D2E;
    border-left: 4px solid #3D5A80;
    padding-left: 14px;
    margin: 28px 0 16px 0;
    letter-spacing: -0.2px;
}

/* ── Info / Alert Banners ────────────────────────────────── */
.info-banner {
    background: #EEF2FF;
    border: 1px solid #C7D2FE;
    border-radius: 10px;
    padding: 14px 18px;
    color: #3730A3;
    font-size: 0.88rem;
    margin-bottom: 16px;
}
.warn-banner {
    background: #FFFBEB;
    border: 1px solid #FDE68A;
    border-radius: 10px;
    padding: 14px 18px;
    color: #92400E;
    font-size: 0.88rem;
    margin-bottom: 16px;
}
.success-banner {
    background: #ECFDF5;
    border: 1px solid #6EE7B7;
    border-radius: 10px;
    padding: 14px 18px;
    color: #065F46;
    font-size: 0.88rem;
    margin-bottom: 16px;
}
.danger-banner {
    background: #FEF2F2;
    border: 1px solid #FECACA;
    border-radius: 10px;
    padding: 14px 18px;
    color: #991B1B;
    font-size: 0.88rem;
    margin-bottom: 16px;
}

/* ── Upload Zone ─────────────────────────────────────────── */
[data-testid="stFileUploader"] {
    background: #FFFFFF !important;
    border: 2px dashed #CBD5E1 !important;
    border-radius: 14px !important;
    padding: 24px !important;
    transition: border-color 0.2s ease;
}
[data-testid="stFileUploader"]:hover {
    border-color: #3D5A80 !important;
}

/* ── Tabs ────────────────────────────────────────────────── */
[data-testid="stTabs"] > div:first-child {
    background: #FFFFFF;
    border-radius: 12px 12px 0 0;
    border-bottom: 2px solid #E2E6F0;
    padding: 0 8px;
}
button[data-baseweb="tab"] {
    font-family: 'DM Sans', sans-serif !important;
    font-weight: 500 !important;
    font-size: 0.88rem !important;
    color: #6B7280 !important;
    padding: 12px 20px !important;
    border-radius: 0 !important;
    border-bottom: 2px solid transparent !important;
}
button[data-baseweb="tab"][aria-selected="true"] {
    color: #1A1D2E !important;
    border-bottom: 2px solid #3D5A80 !important;
    background: transparent !important;
}
[data-testid="stTabsContent"] {
    background: #FFFFFF;
    border-radius: 0 0 12px 12px;
    border: 1px solid #E2E6F0;
    border-top: none;
    padding: 28px 24px !important;
}

/* ── DataFrames ──────────────────────────────────────────── */
[data-testid="stDataFrame"] {
    border: 1px solid #E2E6F0 !important;
    border-radius: 12px !important;
    overflow: hidden !important;
}
[data-testid="stDataFrame"] th {
    background: #F1F5F9 !important;
    color: #374151 !important;
    font-weight: 600 !important;
    font-size: 0.82rem !important;
    text-transform: uppercase !important;
    letter-spacing: 0.5px !important;
    border-bottom: 2px solid #E2E6F0 !important;
}
[data-testid="stDataFrame"] td {
    font-size: 0.875rem !important;
    color: #374151 !important;
}
[data-testid="stDataFrame"] tr:hover td {
    background: #F8FAFC !important;
}

/* ── Topper cards ────────────────────────────────────────── */
.topper-card {
    background: linear-gradient(135deg, #FFFFFF, #F7F8FC);
    border: 1px solid #E2E6F0;
    border-radius: 14px;
    padding: 18px 20px;
    display: flex;
    align-items: center;
    gap: 16px;
    margin-bottom: 10px;
    box-shadow: 0 2px 8px rgba(26,29,46,0.05);
    transition: box-shadow 0.2s ease;
}
.topper-card:hover { box-shadow: 0 6px 20px rgba(26,29,46,0.1); }
.topper-rank {
    font-family: 'Playfair Display', serif;
    font-size: 1.6rem;
    font-weight: 700;
    color: #3D5A80;
    min-width: 40px;
}
.topper-name { font-weight: 600; font-size: 0.95rem; color: #1A1D2E; }
.topper-roll { font-size: 0.78rem; color: #9CA3AF; margin-top: 2px; }
.topper-score {
    margin-left: auto;
    font-family: 'Playfair Display', serif;
    font-size: 1.4rem;
    font-weight: 700;
    color: #059669;
}

/* ── Student profile box ─────────────────────────────────── */
.student-profile {
    background: #FFFFFF;
    border: 1px solid #E2E6F0;
    border-radius: 14px;
    padding: 24px;
    margin-bottom: 20px;
    box-shadow: 0 2px 12px rgba(26,29,46,0.07);
}
.student-profile h3 {
    font-family: 'Playfair Display', serif;
    font-size: 1.2rem;
    margin: 0 0 4px 0;
    color: #1A1D2E;
}

/* ── Metric Cards (Streamlit native) ─────────────────────── */
[data-testid="stMetric"] {
    background: #FFFFFF;
    border: 1px solid #E2E6F0;
    border-radius: 12px;
    padding: 18px 20px !important;
    box-shadow: 0 2px 8px rgba(26,29,46,0.05);
}
[data-testid="stMetricLabel"] {
    font-size: 0.78rem !important;
    color: #6B7280 !important;
    text-transform: uppercase;
    letter-spacing: 0.6px;
}
[data-testid="stMetricValue"] {
    font-family: 'Playfair Display', serif !important;
    font-size: 1.8rem !important;
    color: #1A1D2E !important;
    font-weight: 700 !important;
}

/* ── Download Button ─────────────────────────────────────── */
[data-testid="stDownloadButton"] > button {
    background: linear-gradient(135deg, #1A1D2E, #3D5A80) !important;
    color: #FFFFFF !important;
    border: none !important;
    border-radius: 10px !important;
    font-family: 'DM Sans', sans-serif !important;
    font-weight: 600 !important;
    font-size: 0.9rem !important;
    padding: 14px 28px !important;
    letter-spacing: 0.3px;
    box-shadow: 0 4px 16px rgba(26,29,46,0.25) !important;
}
[data-testid="stDownloadButton"] > button:hover {
    opacity: 0.9 !important;
}

/* ── Selectbox + Input ───────────────────────────────────── */
[data-baseweb="select"] > div,
[data-baseweb="input"] > div {
    background: #FFFFFF !important;
    border-color: #D1D5DB !important;
    border-radius: 8px !important;
    color: #1A1D2E !important;
}

/* ── Divider ─────────────────────────────────────────────── */
hr {
    border: none !important;
    border-top: 1px solid #E2E6F0 !important;
    margin: 28px 0 !important;
}

/* ── Footer ──────────────────────────────────────────────── */
.app-footer {
    text-align: center;
    padding: 28px;
    color: #9CA3AF;
    font-size: 0.78rem;
    letter-spacing: 0.5px;
    border-top: 1px solid #E2E6F0;
    margin-top: 48px;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# PLOTLY THEME
# ─────────────────────────────────────────────
CHART_THEME = dict(
    template="plotly_white",
    font_family="DM Sans",
    font_color="#374151",
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    margin=dict(l=20, r=20, t=50, b=20),
)

# ─────────────────────────────────────────────
# APP HEADER
# ─────────────────────────────────────────────
st.markdown("""
<div class="app-header">
    <div class="header-icon">🎓</div>
    <div>
        <h1>Student Result Analysis System</h1>
        <p>Academic Performance Intelligence &nbsp;·&nbsp; Upload · Analyse · Report · Download</p>
    </div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# CACHING HELPERS
# ─────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_data(file_bytes):
    try:
        r = pd.read_excel(BytesIO(file_bytes), sheet_name=0)
        f = pd.read_excel(BytesIO(file_bytes), sheet_name=1)
        return r, f, None
    except Exception as e:
        return None, None, str(e)


@st.cache_data(show_spinner=False)
def process_results(raw_json):
    df = pd.read_json(raw_json)
    df = df.iloc[1:].reset_index(drop=True)
    col_lower = {str(c).strip().lower(): c for c in df.columns}
    rename_map = {}
    for key, targets in {
        'RollNo':     ['roll no.', 'roll no', 'rollno', 'roll number'],
        'Name':       ['candidate name', 'name', 'student name'],
        'Status':     ['status', 'result'],
        'BackPaper':  ['back paper', 'backpaper', 'back papers'],
        'Division':   ['per.1', 'division', 'percentage'],
        'GrandTotal': ['total\ni to vi', 'grand total', 'total marks', 'total'],
    }.items():
        for t in targets:
            if t in col_lower:
                rename_map[col_lower[t]] = key
                break
    df.rename(columns=rename_map, inplace=True)
    return df


@st.cache_data(show_spinner=False)
def detect_subject_cols(cols_json):
    cols = pd.read_json(cols_json, typ='series').tolist()
    mapping = {}
    for i, col in enumerate(cols):
        s = str(col).strip()
        if s.replace('P', '').isdigit() and i + 2 < len(cols):
            mapping[s] = cols[i + 2]
    return mapping


@st.cache_data(show_spinner=False)
def compute_subject_analysis(df_json, stc_json):
    df  = pd.read_json(df_json)
    stc = pd.read_json(stc_json, typ='series').to_dict()
    rows = []
    for code, col in stc.items():
        if col not in df.columns:
            continue
        m = pd.to_numeric(df[col], errors='coerce')
        rows.append({
            "Subject Code":    code,
            "Total Appeared":  int(m.count()),
            "Highest Marks":   float(m.max()),
            "Lowest Marks":    float(m.min()),
            "Average Marks":   round(float(m.mean()), 2),
            "Passed":          int((m >= 33).sum()),
            "Failed Students": int((m < 33).sum()),
            "Pass %":          round((m >= 33).mean() * 100, 2),
        })
    return pd.DataFrame(rows)


@st.cache_data(show_spinner=False)
def compute_faculty_analysis(df_json, faculty_json, stc_json):
    df  = pd.read_json(df_json)
    fac = pd.read_json(faculty_json)
    stc = pd.read_json(stc_json, typ='series').to_dict()
    rows = []
    for _, row in fac.iterrows():
        code    = str(row.get('Subject Code', '')).strip()
        faculty = row.get('Faculty Name', 'Unknown')
        if code not in stc or stc[code] not in df.columns:
            continue
        m = pd.to_numeric(df[stc[code]], errors='coerce')
        rows.append({
            "Faculty Name":   faculty,
            "Subject Code":   code,
            "Total Students": int(m.count()),
            "Passed":         int((m >= 33).sum()),
            "Failed":         int((m < 33).sum()),
            "Pass %":         round((m >= 33).mean() * 100, 2),
            "Highest Marks":  float(m.max()),
            "Lowest Marks":   float(m.min()),
            "Average Marks":  round(float(m.mean()), 2),
        })
    return pd.DataFrame(rows)

# ─────────────────────────────────────────────
# FILE UPLOAD
# ─────────────────────────────────────────────
col_up, col_guide = st.columns([2, 1])
with col_up:
    uploaded = st.file_uploader(
        "Upload Result Excel File",
        type=["xlsx"],
        help="Sheet 1: Student Results  |  Sheet 2: Faculty Mapping"
    )
with col_guide:
    st.markdown("""
    <div class="info-banner" style="margin-top:8px">
        <strong>📋 Required Format</strong><br>
        <b>Sheet 1</b> — Student results: Roll No, Name, Status, Subject marks, Grand Total<br>
        <b>Sheet 2</b> — Faculty mapping: Subject Code, Faculty Name
    </div>
    """, unsafe_allow_html=True)

if not uploaded:
    st.markdown("""
    <div class="success-banner">
        ☝️  Awaiting file upload. Once uploaded, the full analysis will appear instantly.
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# ─────────────────────────────────────────────
# DATA LOADING
# ─────────────────────────────────────────────
file_bytes = uploaded.read()
with st.spinner("Loading and processing data…"):
    results_raw, faculty_raw, err = load_data(file_bytes)

if err:
    st.markdown(f'<div class="danger-banner">❌ Failed to read file: {err}</div>', unsafe_allow_html=True)
    st.stop()

if results_raw is None or faculty_raw is None:
    st.markdown('<div class="danger-banner">❌ Could not find both sheets. Ensure Sheet 1 = Results, Sheet 2 = Faculty.</div>', unsafe_allow_html=True)
    st.stop()

try:
    results_df = process_results(results_raw.to_json())
except Exception as e:
    st.markdown(f'<div class="danger-banner">❌ Processing error: {e}</div>', unsafe_allow_html=True)
    st.stop()

required = {'RollNo', 'Name', 'Status', 'GrandTotal'}
missing  = required - set(results_df.columns)
if missing:
    st.markdown(f'<div class="danger-banner">❌ Missing columns: {missing}. Check your Excel headers.</div>', unsafe_allow_html=True)
    st.stop()

total_students = len(results_df)
if total_students == 0:
    st.markdown('<div class="danger-banner">❌ No student records found in the file.</div>', unsafe_allow_html=True)
    st.stop()

# ─────────────────────────────────────────────
# DERIVED DATA
# ─────────────────────────────────────────────
results_df['GrandTotal'] = pd.to_numeric(results_df['GrandTotal'], errors='coerce')

passed_df   = results_df[results_df['Status'].isin(['FIRST', 'SECOND', 'PASS'])]
fail_count  = int((results_df['Status'] == 'FAIL').sum())
later_count = int((results_df['Status'] == 'LATER').sum())
ufm_count   = int((results_df['Status'] == 'UFM').sum())
first_div   = int((results_df['Status'] == 'FIRST').sum())
second_div  = int((results_df['Status'] == 'SECOND').sum())
pass_count  = int((results_df['Status'] == 'PASS').sum())
bp_count    = int(results_df['BackPaper'].notna().sum()) if 'BackPaper' in results_df.columns else 0
pass_pct    = round(len(passed_df) / total_students * 100, 2)

def grade_bucket(s):
    if pd.isna(s): return 'N/A'
    if s >= 90: return 'A+'
    if s >= 75: return 'A'
    if s >= 60: return 'B'
    if s >= 45: return 'C'
    if s >= 33: return 'D'
    return 'F'

results_df['Grade'] = results_df['GrandTotal'].apply(grade_bucket)

reason_map = {'FAIL': 'Low Marks', 'LATER': 'Back Paper / Absent', 'UFM': 'Unfair Means'}
failed_df  = results_df[results_df['Status'].isin(['FAIL', 'LATER', 'UFM'])].copy()
failed_df['Failure Reason'] = failed_df['Status'].map(reason_map).fillna('Unknown')

top5 = results_df.nlargest(5, 'GrandTotal')[['Name', 'RollNo', 'GrandTotal', 'Grade', 'Status']]

cols_ser = pd.Series(list(results_df.columns))
stc      = detect_subject_cols(cols_ser.to_json())
stc_ser  = pd.Series(stc)
subj_df  = compute_subject_analysis(results_df.to_json(), stc_ser.to_json())
fac_df   = compute_faculty_analysis(results_df.to_json(), faculty_raw.to_json(), stc_ser.to_json())

weak_subjects = subj_df[
    (subj_df['Pass %'] < 60) | (subj_df['Failed Students'] > subj_df['Failed Students'].mean())
] if not subj_df.empty else pd.DataFrame()

low_margin_list = []
for code, col in stc.items():
    if col not in results_df.columns:
        continue
    m    = pd.to_numeric(results_df[col], errors='coerce')
    mask = (m >= 30) & (m < 33)
    temp = results_df[mask][['Name', 'RollNo']].copy()
    temp['Subject Code'] = code
    temp['Marks'] = m[mask].values
    low_margin_list.append(temp)
low_margin_df = pd.concat(low_margin_list, ignore_index=True) if low_margin_list else pd.DataFrame()

# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🔍 Filters & Tools")

    st.markdown("**Filter by Status**")
    status_opts = ['All'] + sorted(results_df['Status'].dropna().unique().tolist())
    sel_status  = st.selectbox("Status", status_opts, label_visibility="collapsed")

    st.markdown("---")
    st.markdown("**Student Lookup**")
    search_roll = st.text_input("Roll Number", placeholder="e.g. 2301001", label_visibility="collapsed")

    st.markdown("---")
    st.markdown("**Pass Mark Threshold**")
    pass_mark = st.slider("Pass Mark", 25, 50, 33, label_visibility="collapsed")

    st.markdown("---")
    st.markdown(f"""
    <div style="font-size:0.78rem; color:#6B7280; line-height:1.8">
        <b>File:</b> {uploaded.name}<br>
        <b>Students:</b> {total_students}<br>
        <b>Subjects:</b> {len(stc)}<br>
        <b>Faculty:</b> {len(fac_df)}<br>
        <b>Pass Rate:</b> {pass_pct}%
    </div>
    """, unsafe_allow_html=True)

filtered_df = results_df if sel_status == 'All' else results_df[results_df['Status'] == sel_status]

# ─────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📊  Overview",
    "📘  Subject Analysis",
    "🧑‍🏫  Faculty Performance",
    "👥  Student Details",
    "📈  Charts",
    "📥  Export",
])

# ════════════════════════════════════════════════
# TAB 1 – OVERVIEW
# ════════════════════════════════════════════════
with tab1:
    c1, c2, c3, c4, c5, c6, c7, c8 = st.columns(8)
    metrics = [
        (c1, "Total Students", total_students, f"Enrolled"),
        (c2, "Passed",         len(passed_df), f"{pass_pct}% pass rate"),
        (c3, "First Division", first_div,      "≥ 60%"),
        (c4, "Second Division",second_div,     "45–59%"),
        (c5, "Failed",         fail_count,     "Below pass mark"),
        (c6, "Later / Back",   later_count,    "Incomplete"),
        (c7, "Back Papers",    bp_count,       "Supplementary"),
        (c8, "UFM Cases",      ufm_count,      "Unfair means"),
    ]
    for col, label, val, sub in metrics:
        with col:
            st.metric(label=label, value=val, delta=sub)

    st.markdown("<hr>", unsafe_allow_html=True)

    # Toppers
    st.markdown('<div class="section-title">🏆 Top Performers</div>', unsafe_allow_html=True)
    medals = ["🥇", "🥈", "🥉", "4️⃣", "5️⃣"]
    for i, (_, row) in enumerate(top5.iterrows()):
        gt = row['GrandTotal']
        st.markdown(f"""
        <div class="topper-card">
            <div class="topper-rank">{medals[i]}</div>
            <div>
                <div class="topper-name">{row['Name']}</div>
                <div class="topper-roll">Roll No. {row['RollNo']} &nbsp;·&nbsp; {row.get('Status','')}</div>
            </div>
            <div class="topper-score">{gt:.0f if not pd.isna(gt) else '–'}</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<hr>", unsafe_allow_html=True)

    col_l, col_r = st.columns(2)
    with col_l:
        st.markdown('<div class="section-title">Result Breakdown</div>', unsafe_allow_html=True)
        pie_data = pd.DataFrame({
            'Category': ['First Division', 'Second Division', 'Pass', 'Later', 'Fail', 'UFM'],
            'Count':    [first_div, second_div, pass_count, later_count, fail_count, ufm_count]
        })
        pie_data = pie_data[pie_data['Count'] > 0]
        fig_pie = px.pie(
            pie_data, names='Category', values='Count', hole=0.5,
            color_discrete_sequence=["#3D5A80","#5B8DB8","#8AB4D6","#D97706","#DC2626","#374151"]
        )
        fig_pie.update_traces(textposition='outside', textinfo='label+percent')
        fig_pie.update_layout(**CHART_THEME, showlegend=False, height=360,
                              title="Result Distribution")
        st.plotly_chart(fig_pie, use_container_width=True)

    with col_r:
        st.markdown('<div class="section-title">Grade Distribution</div>', unsafe_allow_html=True)
        grade_order  = ['A+', 'A', 'B', 'C', 'D', 'F', 'N/A']
        grade_colors = ["#059669","#34D399","#60A5FA","#FBBF24","#F97316","#EF4444","#9CA3AF"]
        gc = results_df['Grade'].value_counts().reindex(grade_order, fill_value=0).reset_index()
        gc.columns = ['Grade', 'Count']
        fig_grade = px.bar(gc, x='Grade', y='Count', text='Count',
                           color='Grade', color_discrete_sequence=grade_colors)
        fig_grade.update_traces(textposition='outside')
        fig_grade.update_layout(**CHART_THEME, showlegend=False, height=360,
                                title="Grade Distribution")
        st.plotly_chart(fig_grade, use_container_width=True)

# ════════════════════════════════════════════════
# TAB 2 – SUBJECT ANALYSIS
# ════════════════════════════════════════════════
with tab2:
    st.markdown('<div class="section-title">📘 Subject-wise Performance</div>', unsafe_allow_html=True)

    if subj_df.empty:
        st.markdown('<div class="warn-banner">⚠️ No subject columns detected. Check your Excel structure.</div>', unsafe_allow_html=True)
    else:
        best_sub  = subj_df.loc[subj_df['Pass %'].idxmax()]
        worst_sub = subj_df.loc[subj_df['Pass %'].idxmin()]
        c1, c2, c3 = st.columns(3)
        c1.metric("Best Subject",    best_sub['Subject Code'],  f"{best_sub['Pass %']}% pass rate")
        c2.metric("Weakest Subject", worst_sub['Subject Code'], f"{worst_sub['Pass %']}% pass rate")
        c3.metric("Avg Pass Rate",   f"{subj_df['Pass %'].mean():.1f}%")

        st.markdown("<hr>", unsafe_allow_html=True)
        st.dataframe(
            subj_df.style
                .background_gradient(subset=['Pass %'], cmap='RdYlGn', vmin=0, vmax=100)
                .background_gradient(subset=['Failed Students'], cmap='Reds')
                .format({"Pass %": "{:.1f}%", "Average Marks": "{:.2f}"}),
            use_container_width=True, height=340
        )

        st.markdown("<hr>", unsafe_allow_html=True)
        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown('<div class="section-title">Pass % by Subject</div>', unsafe_allow_html=True)
            fig_sp = px.bar(
                subj_df.sort_values('Pass %'), x='Pass %', y='Subject Code',
                orientation='h', text='Pass %',
                color='Pass %', color_continuous_scale='RdYlGn', range_color=[0, 100]
            )
            fig_sp.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
            fig_sp.add_vline(x=60, line_dash='dot', line_color='#DC2626',
                             annotation_text='60% threshold')
            fig_sp.update_layout(**CHART_THEME, showlegend=False, height=360,
                                 coloraxis_showscale=False)
            st.plotly_chart(fig_sp, use_container_width=True)

        with col_b:
            st.markdown('<div class="section-title">Average Marks by Subject</div>', unsafe_allow_html=True)
            fig_avg = px.bar(
                subj_df.sort_values('Average Marks'), x='Average Marks', y='Subject Code',
                orientation='h', text='Average Marks',
                color='Average Marks', color_continuous_scale='Blues'
            )
            fig_avg.update_traces(texttemplate='%{text:.1f}', textposition='outside')
            fig_avg.add_vline(x=pass_mark, line_dash='dot', line_color='#DC2626',
                              annotation_text=f'Pass ({pass_mark})')
            fig_avg.update_layout(**CHART_THEME, showlegend=False, height=360,
                                  coloraxis_showscale=False)
            st.plotly_chart(fig_avg, use_container_width=True)

        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown('<div class="section-title">📉 Subjects Needing Intervention</div>', unsafe_allow_html=True)
        if weak_subjects.empty:
            st.markdown('<div class="success-banner">✅ All subjects performing well — no intervention required.</div>', unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="warn-banner">⚠️ {len(weak_subjects)} subject(s) flagged (pass rate &lt;60% or above-average failures).</div>', unsafe_allow_html=True)
            st.dataframe(weak_subjects, use_container_width=True)

        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown('<div class="section-title">⚠️ Low-Margin Failures (30–32 Marks)</div>', unsafe_allow_html=True)
        if low_margin_df.empty:
            st.markdown('<div class="success-banner">✅ No students in the 30–32 marks band.</div>', unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="warn-banner">⚠️ {len(low_margin_df)} student–subject pair(s) within 3 marks of passing. May benefit from re-checking / grace marks policy.</div>', unsafe_allow_html=True)
            st.dataframe(low_margin_df, use_container_width=True)

# ════════════════════════════════════════════════
# TAB 3 – FACULTY PERFORMANCE
# ════════════════════════════════════════════════
with tab3:
    st.markdown('<div class="section-title">🧑‍🏫 Faculty-wise Performance</div>', unsafe_allow_html=True)

    if fac_df.empty:
        st.markdown('<div class="warn-banner">⚠️ No faculty data matched subject codes.</div>', unsafe_allow_html=True)
    else:
        best_f  = fac_df.loc[fac_df['Pass %'].idxmax()]
        worst_f = fac_df.loc[fac_df['Pass %'].idxmin()]
        c1, c2, c3 = st.columns(3)
        c1.metric("Highest Pass %",     best_f['Faculty Name'],  f"{best_f['Pass %']}%")
        c2.metric("Lowest Pass %",      worst_f['Faculty Name'], f"{worst_f['Pass %']}%")
        c3.metric("Overall Avg Pass %", f"{fac_df['Pass %'].mean():.1f}%")

        st.markdown("<hr>", unsafe_allow_html=True)
        st.dataframe(
            fac_df.sort_values('Pass %', ascending=False)
                  .style
                  .background_gradient(subset=['Pass %'], cmap='RdYlGn', vmin=0, vmax=100)
                  .background_gradient(subset=['Failed'], cmap='Reds')
                  .format({"Pass %": "{:.1f}%", "Average Marks": "{:.2f}"}),
            use_container_width=True
        )

        st.markdown("<hr>", unsafe_allow_html=True)
        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown('<div class="section-title">Pass % by Faculty</div>', unsafe_allow_html=True)
            fig_fp = px.bar(
                fac_df.sort_values('Pass %'), x='Pass %', y='Faculty Name',
                orientation='h', text='Pass %',
                color='Pass %', color_continuous_scale='RdYlGn', range_color=[0, 100]
            )
            fig_fp.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
            fig_fp.add_vline(x=60, line_dash='dot', line_color='#DC2626',
                             annotation_text='60% threshold')
            fig_fp.update_layout(**CHART_THEME, showlegend=False, height=400,
                                 coloraxis_showscale=False)
            st.plotly_chart(fig_fp, use_container_width=True)

        with col_b:
            st.markdown('<div class="section-title">Avg Marks vs Failures (Quadrant)</div>', unsafe_allow_html=True)
            fig_q = px.scatter(
                fac_df, x='Average Marks', y='Failed',
                size='Total Students', color='Pass %',
                color_continuous_scale='RdYlGn', range_color=[0, 100],
                hover_name='Faculty Name',
                hover_data=['Subject Code', 'Passed']
            )
            fig_q.add_hline(y=fac_df['Failed'].mean(), line_dash='dash',
                            line_color='#9CA3AF', annotation_text="Avg Failures")
            fig_q.add_vline(x=fac_df['Average Marks'].mean(), line_dash='dash',
                            line_color='#9CA3AF', annotation_text="Avg Marks")
            fig_q.update_layout(**CHART_THEME, height=400, coloraxis_showscale=False)
            st.plotly_chart(fig_q, use_container_width=True)

        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown('<div class="section-title">Passed vs Failed per Faculty</div>', unsafe_allow_html=True)
        fac_melt = fac_df.melt(
            id_vars='Faculty Name', value_vars=['Passed', 'Failed'],
            var_name='Outcome', value_name='Count'
        )
        fig_stack = px.bar(
            fac_melt, x='Faculty Name', y='Count', color='Outcome',
            barmode='stack', text='Count',
            color_discrete_map={'Passed': '#059669', 'Failed': '#DC2626'}
        )
        fig_stack.update_traces(textposition='inside')
        fig_stack.update_layout(**CHART_THEME, xaxis_tickangle=-25, height=380)
        st.plotly_chart(fig_stack, use_container_width=True)

# ════════════════════════════════════════════════
# TAB 4 – STUDENT DETAILS
# ════════════════════════════════════════════════
with tab4:
    if search_roll:
        st.markdown('<div class="section-title">🔎 Student Profile</div>', unsafe_allow_html=True)
        student = results_df[results_df['RollNo'].astype(str).str.strip() == search_roll.strip()]
        if student.empty:
            st.markdown(f'<div class="danger-banner">❌ No student found with Roll No: <strong>{search_roll}</strong></div>', unsafe_allow_html=True)
        else:
            s = student.iloc[0]
            grade = s.get('Grade', 'N/A')
            grade_color = {"A+":"#059669","A":"#34D399","B":"#60A5FA","C":"#FBBF24","D":"#F97316","F":"#EF4444"}.get(grade, "#9CA3AF")
            gt = s.get('GrandTotal', '')
            gt_str = f"{gt:.0f}" if not pd.isna(gt) else "–"
            st.markdown(f"""
            <div class="student-profile">
                <div style="display:flex;justify-content:space-between;align-items:flex-start">
                    <div>
                        <h3>{s.get('Name','–')}</h3>
                        <div style="color:#6B7280;font-size:0.85rem">
                            Roll No. {s.get('RollNo','–')} &nbsp;·&nbsp;
                            Status: <strong>{s.get('Status','–')}</strong>
                        </div>
                    </div>
                    <div style="text-align:center">
                        <div style="font-family:'Playfair Display',serif;font-size:2.5rem;font-weight:700;color:{grade_color};line-height:1">{grade}</div>
                        <div style="font-size:0.75rem;color:#9CA3AF">Grade</div>
                    </div>
                </div>
                <div style="margin-top:16px">
                    <div style="font-size:0.75rem;color:#9CA3AF;text-transform:uppercase;letter-spacing:0.5px">Grand Total</div>
                    <div style="font-family:'Playfair Display',serif;font-size:1.8rem;font-weight:700;color:#1A1D2E">{gt_str}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            sub_marks = []
            for code, col in stc.items():
                if col in student.columns:
                    val = pd.to_numeric(student[col].values[0], errors='coerce')
                    sub_marks.append({"Subject": code, "Marks": val,
                                      "Status": "Pass" if val >= pass_mark else "Fail"})
            if sub_marks:
                smdf = pd.DataFrame(sub_marks)
                fig_sm = px.bar(smdf, x='Subject', y='Marks', color='Status',
                                color_discrete_map={'Pass':'#059669','Fail':'#DC2626'},
                                text='Marks', title=f"Subject Marks — {s.get('Name','')}")
                fig_sm.add_hline(y=pass_mark, line_dash='dot', line_color='#374151',
                                 annotation_text=f'Pass ({pass_mark})')
                fig_sm.update_traces(texttemplate='%{text:.0f}', textposition='outside')
                fig_sm.update_layout(**CHART_THEME, height=320, showlegend=True)
                st.plotly_chart(fig_sm, use_container_width=True)

        st.markdown("<hr>", unsafe_allow_html=True)

    st.markdown(f'<div class="section-title">📋 Students — {sel_status}</div>', unsafe_allow_html=True)
    disp_cols = [c for c in ['RollNo','Name','Status','GrandTotal','Grade','Division'] if c in filtered_df.columns]
    st.dataframe(filtered_df[disp_cols], use_container_width=True, height=320)

    st.markdown("<hr>", unsafe_allow_html=True)
    col_a, col_b = st.columns(2)

    with col_a:
        st.markdown('<div class="section-title">❌ Failed / Later Students</div>', unsafe_allow_html=True)
        show_cols = [c for c in ['Name','RollNo','Status','Failure Reason'] if c in failed_df.columns]
        st.dataframe(failed_df[show_cols], use_container_width=True, height=300)

    with col_b:
        st.markdown('<div class="section-title">🚨 Multiple Back Papers</div>', unsafe_allow_html=True)
        if 'BackPaper' in results_df.columns:
            bp_students = results_df[results_df['BackPaper'].notna()]
            mbp = bp_students.groupby('RollNo').size().reset_index(name='BP Count')
            mbp = mbp[mbp['BP Count'] > 1].merge(
                results_df[['RollNo','Name']].drop_duplicates(), on='RollNo', how='left'
            )
            if mbp.empty:
                st.markdown('<div class="success-banner">✅ No students with multiple back papers.</div>', unsafe_allow_html=True)
            else:
                st.dataframe(mbp[['Name','RollNo','BP Count']], use_container_width=True, height=300)
        else:
            st.markdown('<div class="info-banner">ℹ️ Back paper column not found in this file.</div>', unsafe_allow_html=True)

# ════════════════════════════════════════════════
# TAB 5 – CHARTS
# ════════════════════════════════════════════════
with tab5:
    st.markdown('<div class="section-title">📈 Analytics Dashboard</div>', unsafe_allow_html=True)

    if not subj_df.empty:
        col_a, col_b = st.columns(2)
        with col_a:
            # Radar chart
            fig_r = go.Figure()
            codes = subj_df['Subject Code'].tolist()
            vals  = subj_df['Pass %'].tolist()
            fig_r.add_trace(go.Scatterpolar(
                r=vals + [vals[0]],
                theta=codes + [codes[0]],
                fill='toself',
                fillcolor='rgba(61,90,128,0.12)',
                line=dict(color='#3D5A80', width=2),
                name='Pass %'
            ))
            fig_r.update_layout(
                **CHART_THEME,
                polar=dict(
                    bgcolor='rgba(0,0,0,0)',
                    radialaxis=dict(visible=True, range=[0,100], gridcolor='#E2E6F0', tickfont_size=10),
                    angularaxis=dict(gridcolor='#E2E6F0')
                ),
                title="Subject Pass % — Radar View", height=400
            )
            st.plotly_chart(fig_r, use_container_width=True)

        with col_b:
            # Bubble chart
            fig_b = px.scatter(
                subj_df, x='Average Marks', y='Pass %',
                size='Total Appeared', color='Failed Students',
                hover_name='Subject Code', color_continuous_scale='Reds',
                size_max=50, title="Avg Marks vs Pass % (bubble = class size)"
            )
            fig_b.add_hline(y=60, line_dash='dash', line_color='#9CA3AF',
                            annotation_text='60% threshold')
            fig_b.update_layout(**CHART_THEME, height=400, coloraxis_showscale=False)
            st.plotly_chart(fig_b, use_container_width=True)

        st.markdown("<hr>", unsafe_allow_html=True)

        st.markdown('<div class="section-title">Grand Total Distribution</div>', unsafe_allow_html=True)
        fig_hist = px.histogram(
            results_df.dropna(subset=['GrandTotal']), x='GrandTotal', nbins=25,
            color_discrete_sequence=['#3D5A80'],
            labels={'GrandTotal': 'Grand Total Marks'},
            title="Marks Distribution — All Students"
        )
        mean_val = results_df['GrandTotal'].mean()
        fig_hist.add_vline(x=mean_val, line_dash='dash', line_color='#DC2626',
                           annotation_text=f"Mean: {mean_val:.1f}")
        fig_hist.update_layout(**CHART_THEME, height=340, bargap=0.06)
        st.plotly_chart(fig_hist, use_container_width=True)

    if not fac_df.empty and len(fac_df) > 1:
        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown('<div class="section-title">Faculty × Subject Pass % Heatmap</div>', unsafe_allow_html=True)
        try:
            heat_pivot = fac_df.pivot_table(
                index='Faculty Name', columns='Subject Code', values='Pass %', aggfunc='mean'
            )
            fig_heat = px.imshow(
                heat_pivot, color_continuous_scale='RdYlGn',
                aspect='auto', range_color=[0, 100],
                text_auto='.1f', title="Faculty × Subject Pass % Heatmap"
            )
            fig_heat.update_layout(**CHART_THEME, height=400)
            st.plotly_chart(fig_heat, use_container_width=True)
        except Exception:
            st.markdown('<div class="info-banner">ℹ️ Heatmap requires multiple faculty/subject combinations.</div>', unsafe_allow_html=True)

# ════════════════════════════════════════════════
# TAB 6 – EXPORT
# ════════════════════════════════════════════════
with tab6:
    st.markdown('<div class="section-title">📥 Export Full Analysis Report</div>', unsafe_allow_html=True)

    def style_ws(ws, accent="1A1D2E"):
        h_fill = PatternFill("solid", fgColor=accent)
        h_font = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
        border = Border(
            left=Side(style='thin', color='D1D5DB'),
            right=Side(style='thin', color='D1D5DB'),
            top=Side(style='thin', color='D1D5DB'),
            bottom=Side(style='thin', color='D1D5DB'),
        )
        alt_fill = PatternFill("solid", fgColor="F7F8FC")
        for cell in ws[1]:
            cell.fill   = h_fill
            cell.font   = h_font
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = border
        ws.row_dimensions[1].height = 22
        for ri, row in enumerate(ws.iter_rows(min_row=2), start=2):
            for cell in row:
                cell.border    = border
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.font      = Font(name="Calibri", size=9, color="1A1D2E")
                if ri % 2 == 0:
                    cell.fill = alt_fill
        for col in ws.columns:
            max_len = max((len(str(c.value)) for c in col if c.value is not None), default=8)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 32)

    def apply_pass_color(ws):
        headers = [c.value for c in ws[1]]
        if "Pass %" not in headers:
            return
        ci = headers.index("Pass %") + 1
        for row in ws.iter_rows(min_row=2, min_col=ci, max_col=ci):
            for cell in row:
                v = cell.value
                if isinstance(v, (int, float)):
                    if v >= 80:   cell.fill = PatternFill("solid", fgColor="C6EFCE")
                    elif v >= 60: cell.fill = PatternFill("solid", fgColor="FFEB9C")
                    else:         cell.fill = PatternFill("solid", fgColor="FFC7CE")

    out = BytesIO()
    export_sheets = {
        "Overview": pd.DataFrame({
            "Metric": ["Total Students","Passed","Failed","Later/Back","First Division",
                       "Second Division","Back Papers","UFM Cases","Pass %"],
            "Value":  [total_students, len(passed_df), fail_count, later_count,
                       first_div, second_div, bp_count, ufm_count, f"{pass_pct}%"],
        }),
        "Subject Analysis":    subj_df,
        "Faculty Analysis":    fac_df,
        "Weak Subjects":       weak_subjects if not weak_subjects.empty else pd.DataFrame({"Note": ["No weak subjects detected"]}),
        "Low Margin Failures": low_margin_df if not low_margin_df.empty else pd.DataFrame({"Note": ["No low-margin failures"]}),
        "Top 5 Toppers":       top5,
        "Failed Students":     failed_df[[c for c in ['Name','RollNo','Status','Failure Reason'] if c in failed_df.columns]],
        "All Students":        results_df[[c for c in ['RollNo','Name','Status','GrandTotal','Grade'] if c in results_df.columns]],
    }
    accents = {
        "Overview":"1A1D2E", "Subject Analysis":"1E3A5F", "Faculty Analysis":"14532D",
        "Weak Subjects":"7F1D1D", "Low Margin Failures":"78350F",
        "Top 5 Toppers":"4C1D95", "Failed Students":"7F1D1D", "All Students":"1E3A5F"
    }

    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        for name, df in export_sheets.items():
            df.to_excel(writer, sheet_name=name[:31], index=False)
        wb = writer.book
        for name in export_sheets:
            ws = wb[name[:31]]
            style_ws(ws, accents.get(name, "1A1D2E"))
            if name in ("Subject Analysis", "Faculty Analysis"):
                apply_pass_color(ws)

    col_dl, col_info2 = st.columns([1, 1])
    with col_dl:
        st.download_button(
            label="📥  Download Complete Report (.xlsx)",
            data=out.getvalue(),
            file_name=f"Result_Analysis_{uploaded.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.markdown("""
        <div class="info-banner" style="margin-top:14px">
            <strong>Report includes 8 sheets:</strong><br>
            Overview · Subject Analysis · Faculty Analysis ·
            Weak Subjects · Low-Margin Failures · Top 5 Toppers ·
            Failed Students · All Students<br><br>
            Formatted with alternating row colours, colour-coded Pass % cells
            (green/amber/red), professional headers, and auto-sized columns.
        </div>
        """, unsafe_allow_html=True)

    with col_info2:
        st.markdown('<div class="section-title">Export Summary</div>', unsafe_allow_html=True)
        st.metric("Total Records",      total_students)
        st.metric("Subjects Analysed",  len(subj_df))
        st.metric("Faculty Members",    len(fac_df))
        if not subj_df.empty:
            st.metric("Overall Pass Rate", f"{subj_df['Pass %'].mean():.1f}%")

# ─────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────
st.markdown("""
<div class="app-footer">
    Student Result Analysis System &nbsp;·&nbsp; Academic Performance Intelligence<br>
    All data is processed locally in your browser session and is never stored or transmitted.
</div>
""", unsafe_allow_html=True)


