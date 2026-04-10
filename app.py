import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import warnings
warnings.filterwarnings('ignore')

st.set_page_config(
    page_title="Semester Result Analysis",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Styles ───────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
html,body,[class*="css"]{font-family:'Inter',sans-serif}
.main-header{background:linear-gradient(135deg,#1a237e,#3949ab);color:#fff;padding:1.8rem 2.5rem;
  border-radius:16px;margin-bottom:1.5rem;text-align:center;box-shadow:0 8px 32px rgba(26,35,126,.3)}
.main-header h1{font-size:1.9rem;font-weight:700;margin:0 0 .4rem}
.main-header p{font-size:.95rem;opacity:.85;margin:0}
.kpi{background:#fff;border-radius:12px;padding:1.1rem 1.4rem;box-shadow:0 2px 12px rgba(0,0,0,.08);
  border-left:5px solid;margin-bottom:.5rem}
.kpi-val{font-size:1.9rem;font-weight:700}
.kpi-lbl{font-size:.75rem;color:#666;text-transform:uppercase;letter-spacing:.5px}
.sec{font-size:1.1rem;font-weight:600;color:#1a237e;border-bottom:2px solid #e8eaf6;
  padding-bottom:.4rem;margin:1.2rem 0 .8rem}
.info-box{background:#e8eaf6;border-radius:8px;padding:.8rem 1.2rem;
  border-left:4px solid #3949ab;margin:.8rem 0;font-size:.9rem}
</style>
""", unsafe_allow_html=True)

# ── Constants ─────────────────────────────────────────────────────────────────
SUBJECT_COLS = {
    '501': '501_Ttl',
    '502': '502_Ttl',
    '503': '503_Ttl',
    '504': '504_Ttl',
}
SUBJECT_LABELS = {
    '501': 'DBMS',
    '502': 'Java',
    '503': 'Computer Network',
    '504': 'Numerical Methods',
}
COLORS = ['#1a237e','#283593','#3949ab','#5c6bc0','#7986cb','#9fa8da']

# ── Parsing ───────────────────────────────────────────────────────────────────

def flatten_multiindex(df):
    flat = []
    seen = {}
    for top, sub in df.columns:
        t = str(top).strip().replace('\n',' ')
        s = str(sub).strip().replace('\n',' ')
        if 'Unnamed' in t: t = ''
        if 'Unnamed' in s: s = ''
        key = f"{t}_{s}".strip('_') or '_'
        seen[key] = seen.get(key, 0) + 1
        flat.append(key if seen[key] == 1 else f"{key}_{seen[key]}")
    df.columns = flat
    return df


def parse_sheet1(uploaded_file):
    uploaded_file.seek(0)
    df = pd.read_excel(uploaded_file, sheet_name=0, header=[0, 1])
    df = flatten_multiindex(df)
    df = df.dropna(subset=['S. No.']).reset_index(drop=True)
    for col in list(SUBJECT_COLS.values()) + ['Totel (V)_600','Grand Total_3000','PER_%','PER_%.1']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
    return df


def parse_sheet2(uploaded_file):
    uploaded_file.seek(0)
    raw = pd.read_excel(uploaded_file, sheet_name=1, header=None)
    df = raw.iloc[1:].copy()
    df.columns = ['Subject Code','Subject Name','Faculty Name','Sections']
    df = df.dropna(subset=['Subject Code'])
    df['Subject Code'] = df['Subject Code'].astype(str).str.strip()
    return df.reset_index(drop=True)

# ── Analysis ──────────────────────────────────────────────────────────────────

def result_summary(df):
    s = df['Status'].astype(str).str.strip().str.upper()
    total   = len(df)
    passed  = int((s == 'PASS').sum())
    later   = int((s == 'LATER').sum())
    bp      = int((s == 'BP').sum())
    fail    = int((s == 'FAIL').sum())
    ufm     = int((s == 'UFM').sum())
    pct_v   = round(passed/total*100, 2) if total else 0
    pct_all = round((passed+later)/total*100, 2) if total else 0
    return {
        'Total Registered': total,
        'Total Appeared': total,
        'Pass (V Sem)': passed,
        'Later / Overall Pass': later,
        'Back Paper': bp,
        'Fail': fail,
        'UFM': ufm,
        '% Pass in V Sem': pct_v,
        '% Pass (I–V Sem)': pct_all,
    }


def subject_stats(df, sf_df):
    records = []
    for code, col in SUBJECT_COLS.items():
        if col not in df.columns:
            continue
        vals     = df[col].dropna()
        appeared = len(vals)
        if not appeared:
            continue
        back     = int((vals < 40).sum())
        passed   = appeared - back
        pct_pass = round(passed/appeared*100, 2)
        avg      = round(float(vals.mean()), 2)
        above60  = round((vals>=60).sum()/appeared*100, 2)
        highest  = int(vals.max())
        fac_rows = sf_df[sf_df['Subject Code'].astype(str).str.contains(code, regex=False)]
        subj_name= fac_rows['Subject Name'].dropna().iloc[0] if not fac_rows.empty else SUBJECT_LABELS.get(code, code)
        faculties= ', '.join(fac_rows['Faculty Name'].dropna().tolist()) if not fac_rows.empty else 'N/A'
        records.append({
            'Code': code, 'Subject Name': str(subj_name).strip(),
            'Faculty': faculties, 'Appeared': appeared,
            'Back Paper': back, 'Pass': passed,
            'Pass %': pct_pass, 'Avg Score': avg,
            '% > 60': above60, 'Highest Marks': highest,
        })
    return pd.DataFrame(records)


def faculty_stats(df, sf_df):
    records = []
    grp_col = 'GR.' if 'GR.' in df.columns else None
    for _, row in sf_df.iterrows():
        code     = str(row['Subject Code']).strip()
        faculty  = str(row['Faculty Name']).strip()
        sections = str(row['Sections']).strip()
        subj     = str(row['Subject Name']).strip()
        col      = SUBJECT_COLS.get(code)
        if not col or col not in df.columns:
            continue
        if grp_col and sections not in ('nan',''):
            sec_list = [s.strip() for s in sections.split(',')]
            sub_df   = df[df[grp_col].astype(str).str.strip().isin(sec_list)]
        else:
            sub_df = df
        vals     = sub_df[col].dropna()
        appeared = len(vals)
        if not appeared:
            continue
        back  = int((vals<40).sum())
        passed= appeared - back
        pct   = round(passed/appeared*100, 2)
        records.append({
            'S.No': len(records)+1, 'Faculty Name': faculty,
            'Sections': sections, 'Subject': subj, 'Code': code,
            'Appeared': appeared, 'Back Paper': back,
            'Pass': passed, 'Pass %': pct,
        })
    return pd.DataFrame(records)


def grade_distribution(df, sf_df):
    records = []
    for code, col in SUBJECT_COLS.items():
        if col not in df.columns:
            continue
        vals     = df[col].dropna()
        appeared = len(vals)
        if not appeared:
            continue
        fac_rows = sf_df[sf_df['Subject Code'].astype(str).str.contains(code, regex=False)]
        subj     = fac_rows['Subject Name'].dropna().iloc[0] if not fac_rows.empty else code
        lt50 = int((vals<50).sum())
        b5060= int(((vals>=50)&(vals<60)).sum())
        b6075= int(((vals>=60)&(vals<75)).sum())
        ab75 = int((vals>=75).sum())
        records.append({
            'Subject': f"{code} – {str(subj).strip()[:22]}",
            '< 50': lt50, '50–59.9': b5060, '60–74.9': b6075, '≥ 75': ab75,
            '< 50 %':    round(lt50/appeared*100, 1),
            '50–59.9 %': round(b5060/appeared*100, 1),
            '60–74.9 %': round(b6075/appeared*100, 1),
            '≥ 75 %':    round(ab75/appeared*100, 1),
        })
    return pd.DataFrame(records)


def toppers(df, n=10):
    sort_col = 'Grand Total_3000' if 'Grand Total_3000' in df.columns else 'Totel (V)_600'
    cols = [c for c in ['Roll No.','CANDIDATE NAME', sort_col, 'PER_%.1'] if c in df.columns]
    out  = df[cols].copy()
    out[sort_col] = pd.to_numeric(out[sort_col], errors='coerce')
    top  = out.nlargest(n, sort_col).reset_index(drop=True)
    top.index += 1; top.index.name = 'Rank'
    return top.reset_index()

# ── Charts ────────────────────────────────────────────────────────────────────

def _layout(fig, title, h=400):
    fig.update_layout(
        title=title, title_x=.5,
        plot_bgcolor='white', paper_bgcolor='white',
        font=dict(family='Inter', size=12),
        height=h, margin=dict(t=65,b=50,l=50,r=20)
    )
    return fig


def chart_pie(summary):
    cats   = ['Pass (V Sem)','Later / Overall Pass','Back Paper','Fail','UFM']
    vals   = [summary.get(c,0) for c in cats]
    colors = ['#2e7d32','#1565c0','#e65100','#c62828','#6a1b9a']
    fig = go.Figure(go.Pie(labels=cats,values=vals,hole=.4,
                           marker=dict(colors=colors),
                           hovertemplate='%{label}: %{value} (%{percent})<extra></extra>'))
    return _layout(fig,'Result Status Breakdown',360)


def chart_bar_summary(summary):
    cats   = ['Pass (V Sem)','Later / Overall Pass','Back Paper','Fail','UFM']
    vals   = [summary.get(c,0) for c in cats]
    colors = ['#2e7d32','#1565c0','#e65100','#c62828','#6a1b9a']
    fig = go.Figure(go.Bar(x=cats,y=vals,marker_color=colors,
                           text=vals,textposition='outside'))
    fig.update_yaxes(title='No. of Students')
    return _layout(fig,'Overall Result Distribution',360)


def chart_subject_pass(sdf):
    fig = go.Figure()
    fig.add_trace(go.Bar(name='Pass %',x=sdf['Code'],y=sdf['Pass %'],
                         marker_color='#1a237e',text=sdf['Pass %'],textposition='outside'))
    fig.add_trace(go.Bar(name='% > 60',x=sdf['Code'],y=sdf['% > 60'],
                         marker_color='#7986cb',text=sdf['% > 60'],textposition='outside'))
    fig.update_layout(barmode='group',yaxis_range=[0,115],yaxis_title='Percentage')
    return _layout(fig,'Subject Wise: Pass % vs Students Scoring Above 60%',400)


def chart_avg(sdf):
    fig = go.Figure(go.Bar(x=sdf['Code'],y=sdf['Avg Score'],
                           marker_color=COLORS[:len(sdf)],
                           text=sdf['Avg Score'],textposition='outside'))
    fig.update_yaxes(title='Average Score (out of 100)')
    return _layout(fig,'Average Score per Subject',380)


def chart_grade_dist(gdf):
    bands  = ['< 50 %','50–59.9 %','60–74.9 %','≥ 75 %']
    bcolor = ['#c62828','#e65100','#1565c0','#2e7d32']
    fig = go.Figure()
    for band,col in zip(bands,bcolor):
        fig.add_trace(go.Bar(name=band,x=gdf['Subject'],y=gdf[band],marker_color=col))
    fig.update_layout(barmode='stack',yaxis_title='% of Students',
                      legend=dict(orientation='h',y=-0.35),
                      margin=dict(t=65,b=120,l=50,r=20),height=460)
    return _layout(fig,'Score Band Distribution by Subject (%)',460)


def chart_faculty(fdf):
    fig = go.Figure(go.Bar(
        x=fdf['Faculty Name'],y=fdf['Pass %'],
        marker_color=[COLORS[i%len(COLORS)] for i in range(len(fdf))],
        text=fdf['Pass %'],textposition='outside'))
    fig.update_layout(xaxis_tickangle=-30,yaxis_range=[0,115],yaxis_title='Pass %',
                      margin=dict(t=65,b=130,l=50,r=20),height=460)
    return _layout(fig,'Faculty Wise Pass %',460)


def chart_score_hist(df):
    col = 'PER_%'
    if col not in df.columns: return None
    vals = pd.to_numeric(df[col],errors='coerce').dropna()
    fig = go.Figure(go.Histogram(x=vals,nbinsx=20,marker_color='#3949ab'))
    fig.update_xaxes(title='V Sem Percentage')
    fig.update_yaxes(title='No. of Students')
    return _layout(fig,'Distribution of Student Scores (V Sem %)',380)


def chart_toppers(top_df):
    name_col  = [c for c in top_df.columns if 'CANDIDATE' in c or 'NAME' in c]
    score_col = [c for c in top_df.columns if 'Grand' in c or '3000' in c or 'Totel' in c]
    if not name_col or not score_col: return None
    fig = go.Figure(go.Bar(
        x=top_df[name_col[0]].astype(str),
        y=pd.to_numeric(top_df[score_col[0]],errors='coerce'),
        marker_color=COLORS[:len(top_df)],
        text=top_df[score_col[0]],textposition='outside'))
    fig.update_layout(xaxis_tickangle=-25,yaxis_title='Grand Total (out of 3000)',
                      margin=dict(t=65,b=120,l=50,r=20),height=440)
    return _layout(fig,'Top 10 Students by Grand Total (I–V Sem)',440)

# ── Excel export ──────────────────────────────────────────────────────────────

def build_excel(summary, sdf, fdf, gdf, top_df, student_df):
    wb = openpyxl.Workbook()
    hf = Font(bold=True,color='FFFFFF',name='Arial',size=10)
    hfl= PatternFill('solid',start_color='1A237E')
    tf = Font(bold=True,name='Arial',size=13,color='1A237E')
    df2= Font(name='Arial',size=10)
    ctr= Alignment(horizontal='center',vertical='center',wrap_text=True)
    lft= Alignment(horizontal='left',  vertical='center',wrap_text=True)
    th = Side(style='thin',color='CCCCCC')
    bd = Border(left=th,right=th,top=th,bottom=th)

    def hdr(ws,row,cols):
        for j,v in enumerate(cols,1):
            c=ws.cell(row=row,column=j,value=v)
            c.font=hf;c.fill=hfl;c.alignment=ctr;c.border=bd

    def drow(ws,row,vals,alt=False):
        for j,v in enumerate(vals,1):
            c=ws.cell(row=row,column=j,value=v)
            c.font=df2;c.alignment=ctr;c.border=bd
            if alt: c.fill=PatternFill('solid',start_color='F5F5F5')

    def aw(ws):
        for col in ws.columns:
            ml=max((len(str(cell.value or '')) for cell in col),default=0)
            ws.column_dimensions[get_column_letter(col[0].column)].width=min(ml+4,40)

    def titl(ws,text,ncols):
        ws.merge_cells(f'A1:{get_column_letter(max(ncols,1))}1')
        c=ws['A1'];c.value=text;c.font=tf;c.alignment=ctr

    # Sheet 1: Summary
    ws1=wb.active; ws1.title='Overall Summary'
    ws1.merge_cells('A1:C1')
    ws1['A1'].value='BCA V SEMESTER RESULT ANALYSIS'
    ws1['A1'].font=tf; ws1['A1'].alignment=ctr
    ws1.merge_cells('A2:C2')
    ws1['A2'].value='RESULT STATUS EXAM DECEMBER 2025 | SESSION: 2025-26'
    ws1['A2'].font=Font(italic=True,name='Arial',size=10); ws1['A2'].alignment=ctr
    hdr(ws1,4,['Metric','Value'])
    for i,(k,v) in enumerate(summary.items(),5):
        ws1.cell(row=i,column=1,value=k).font=Font(bold=True,name='Arial',size=10)
        ws1.cell(row=i,column=1).border=bd; ws1.cell(row=i,column=1).alignment=lft
        c=ws1.cell(row=i,column=2,value=v)
        c.font=Font(bold=True,name='Arial',size=11,color='1A237E')
        c.alignment=ctr;c.border=bd
    ws1.column_dimensions['A'].width=30; ws1.column_dimensions['B'].width=16

    # Sheet 2: Subject
    ws2=wb.create_sheet('Subject Wise Analysis')
    if not sdf.empty:
        titl(ws2,'SUBJECT WISE RESULT ANALYSIS',len(sdf.columns))
        hdr(ws2,3,list(sdf.columns))
        for i,row in sdf.iterrows(): drow(ws2,i+4,list(row),i%2==0)
    aw(ws2)

    # Sheet 3: Faculty
    ws3=wb.create_sheet('Faculty Wise Analysis')
    if not fdf.empty:
        titl(ws3,'FACULTY WISE RESULT ANALYSIS',len(fdf.columns))
        hdr(ws3,3,list(fdf.columns))
        for i,row in fdf.iterrows(): drow(ws3,i+4,list(row),i%2==0)
    aw(ws3)

    # Sheet 4: Grade dist
    ws4=wb.create_sheet('Grade Distribution')
    if not gdf.empty:
        dc=['Subject','< 50','50–59.9','60–74.9','≥ 75','< 50 %','50–59.9 %','60–74.9 %','≥ 75 %']
        titl(ws4,'SCORE BAND DISTRIBUTION',len(dc))
        hdr(ws4,3,dc)
        for i,row in gdf[dc].iterrows(): drow(ws4,i+4,list(row),i%2==0)
    aw(ws4)

    # Sheet 5: Toppers
    ws5=wb.create_sheet('Toppers')
    if not top_df.empty:
        titl(ws5,'TOP STUDENTS – I TO V SEMESTER',len(top_df.columns))
        hdr(ws5,3,list(top_df.columns))
        for i,row in top_df.iterrows(): drow(ws5,i+4,list(row),i%2==0)
    aw(ws5)

    # Sheet 6: Student data
    ws6=wb.create_sheet('Student Data')
    cols6=list(student_df.columns)
    titl(ws6,'COMPLETE STUDENT RESULT DATA',len(cols6))
    hdr(ws6,3,cols6)
    for i,row in student_df.iterrows():
        for j,val in enumerate(row,1):
            c=ws6.cell(row=i+4,column=j,value=val)
            c.font=df2;c.alignment=ctr;c.border=bd
            if i%2==0: c.fill=PatternFill('solid',start_color='F5F5F5')
    aw(ws6)

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.getvalue()

# ── Print HTML ────────────────────────────────────────────────────────────────

def build_print_html(summary, sdf, fdf, gdf, top_df):
    def tbl(df, title, sel=None):
        if df is None or df.empty:
            return f'<h3 class="stl">{title}</h3><p>No data.</p>'
        d = df[sel] if sel else df
        th=''.join(f'<th>{c}</th>' for c in d.columns)
        rows=''.join('<tr>'+''.join(f'<td>{v}</td>' for v in r)+'</tr>'
                     for _,r in d.iterrows())
        return f'<h3 class="stl">{title}</h3><table><thead><tr>{th}</tr></thead><tbody>{rows}</tbody></table>'

    sum_rows=''.join(f'<tr><td class="lb">{k}</td><td class="vl">{v}</td></tr>'
                     for k,v in summary.items())
    gc=['Subject','< 50 %','50–59.9 %','60–74.9 %','≥ 75 %']

    return f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<title>Result Analysis Report</title>
<style>
@page{{size:A4;margin:14mm}}
body{{font-family:Arial,sans-serif;font-size:9.5pt;color:#000;background:#fff}}
.ph{{text-align:center;border-bottom:3px solid #000;padding-bottom:8px;margin-bottom:14px}}
.ph h1{{font-size:15pt;margin:0 0 3px}}.ph p{{font-size:9pt;margin:0}}
.stl{{font-size:11pt;font-weight:bold;margin:14px 0 5px;
      border-bottom:2px solid #333;padding-bottom:3px}}
table{{width:100%;border-collapse:collapse;margin-bottom:13px;font-size:8.5pt}}
th{{background:#333;color:#fff;padding:4px 6px;border:1px solid #555;text-align:center}}
td{{padding:3px 5px;border:1px solid #bbb;text-align:center}}
tr:nth-child(even) td{{background:#f5f5f5}}
.lb{{font-weight:bold;background:#eee;text-align:left;padding-left:8px}}
.vl{{font-weight:bold}}
.ft{{text-align:center;margin-top:16px;font-size:7.5pt;color:#555;
    border-top:1px solid #ccc;padding-top:5px}}
.no-print{{margin-bottom:10px}}
@media print{{.no-print{{display:none}}}}
</style></head>
<body>
<div class="ph">
  <h1>SEMESTER RESULT ANALYSIS</h1>
  <p>Result Status Exam December 2025 &nbsp;|&nbsp; Batch 2023-2026 &nbsp;|&nbsp;
     Session 2025-26 &nbsp;|&nbsp; 935 IMS</p>
</div>
<div class="no-print">
  <button onclick="window.print()"
    style="padding:7px 18px;font-size:11pt;cursor:pointer;
           background:#1a237e;color:#fff;border:none;border-radius:6px">
    🖨 Print / Save as PDF
  </button>
</div>
<h3 class="stl">Overall Result Summary</h3>
<table style="width:48%"><tbody>{sum_rows}</tbody></table>
{tbl(sdf,'Subject Wise Result Analysis')}
{tbl(fdf,'Faculty Wise Result Analysis')}
{tbl(gdf,'Grade Distribution (%)',gc)}
{tbl(top_df,'Top 10 Students (I–V Semester)')}
<div class="ft">
  Generated by Result Analysis System &nbsp;|&nbsp;
  (Dr. Gagan Varshney) &nbsp;|&nbsp; Head – Dept. of Computer Science
</div>
</body></html>"""

# ═══════════════════════════════════════════════════════════════════════════
#  MAIN APP
# ═══════════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="main-header">
  <h1>🎓 Semester – Result Analysis System</h1>
  <p>Upload the Result Format Excel → Auto-generate complete result analysis with charts, Excel export &amp; PDF report</p>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("## 📂 Upload Result File")
    uploaded = st.file_uploader(
        "Result Format (.xlsx)", type=['xlsx'],
        help="RESULT_FORMAT.xlsx — Sheet 1: student data | Sheet 2: subject-faculty map"
    )
    st.markdown("---")
    st.markdown("""
### 📌 Instructions
1. Upload the **Result Format** Excel  
2. Analysis is generated automatically  
3. View charts across each tab  
4. Download Excel or Print-ready PDF  
---
### 📖 Sheet Format
**Sheet 1** – Student marks + status  
**Sheet 2** – Subject–Faculty mapping
""")

if not uploaded:
    st.markdown("""
    <div style="border:2px dashed #3949ab;border-radius:12px;padding:2.5rem;
                text-align:center;background:#f3f4ff;margin:1rem 0">
      <h3 style="color:#1a237e">📤 Upload Result Format Excel</h3>
      <p>Use the <strong>sidebar</strong> to upload your result file</p>
      <p style="color:#666;font-size:.9rem">Sheet 1: Student Data &nbsp;|&nbsp; Sheet 2: Subject–Faculty Map</p>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

with st.spinner("⏳ Processing your Excel file…"):
    try:
        df      = parse_sheet1(uploaded)
        sf_df   = parse_sheet2(uploaded)
        summary = result_summary(df)
        sdf     = subject_stats(df, sf_df)
        fdf     = faculty_stats(df, sf_df)
        gdf     = grade_distribution(df, sf_df)
        top_df  = toppers(df, 10)
    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
        st.exception(e)
        st.stop()

st.success(f"✅ File processed successfully — **{len(df)} students** loaded.")

# KPI row
kpi_colors = ['#1a237e','#2e7d32','#1565c0','#e65100','#c62828','#6a1b9a']
kpi_metrics = [
    ('Total Students',        summary.get('Total Registered',0)),
    ('Pass (V Sem)',           summary.get('Pass (V Sem)',0)),
    ('Later / Overall Pass',  summary.get('Later / Overall Pass',0)),
    ('Back Paper',            summary.get('Back Paper',0)),
    ('Fail',                  summary.get('Fail',0)),
    ('% Pass in V Sem',       f"{summary.get('% Pass in V Sem',0)}%"),
]
cols = st.columns(6)
for col,(lbl,val),color in zip(cols,kpi_metrics,kpi_colors):
    with col:
        st.markdown(
            f'<div class="kpi" style="border-color:{color}">'
            f'<div class="kpi-val" style="color:{color}">{val}</div>'
            f'<div class="kpi-lbl">{lbl}</div></div>',
            unsafe_allow_html=True
        )

st.markdown("---")

# Tabs
t1,t2,t3,t4,t5,t6 = st.tabs([
    "📊 Overview","📚 Subject Analysis",
    "👩‍🏫 Faculty Analysis","📈 Grade Distribution",
    "🏆 Toppers","📋 Student Data"
])

with t1:
    st.markdown('<div class="sec">Overall Result Summary</div>', unsafe_allow_html=True)
    c1,c2 = st.columns(2)
    with c1:
        sum_df = pd.DataFrame(list(summary.items()),columns=['Metric','Value']).set_index('Metric')
        st.dataframe(sum_df, use_container_width=True, height=330)
    with c2:
        st.plotly_chart(chart_pie(summary), use_container_width=True)
    st.plotly_chart(chart_bar_summary(summary), use_container_width=True)
    h = chart_score_hist(df)
    if h: st.plotly_chart(h, use_container_width=True)

with t2:
    st.markdown('<div class="sec">Subject Wise Result Analysis</div>', unsafe_allow_html=True)
    if not sdf.empty:
        st.dataframe(sdf, use_container_width=True, hide_index=True)
        c1,c2 = st.columns(2)
        with c1: st.plotly_chart(chart_subject_pass(sdf), use_container_width=True)
        with c2: st.plotly_chart(chart_avg(sdf), use_container_width=True)
    else:
        st.info("No subject data found.")

with t3:
    st.markdown('<div class="sec">Faculty Wise Result Analysis</div>', unsafe_allow_html=True)
    if not fdf.empty:
        st.dataframe(fdf, use_container_width=True, hide_index=True)
        st.plotly_chart(chart_faculty(fdf), use_container_width=True)
    else:
        st.info("No faculty data found.")

with t4:
    st.markdown('<div class="sec">Score Band Distribution by Subject</div>', unsafe_allow_html=True)
    if not gdf.empty:
        disp=['Subject','< 50','50–59.9','60–74.9','≥ 75','< 50 %','50–59.9 %','60–74.9 %','≥ 75 %']
        st.dataframe(gdf[disp], use_container_width=True, hide_index=True)
        st.plotly_chart(chart_grade_dist(gdf), use_container_width=True)
    else:
        st.info("No grade data found.")

with t5:
    st.markdown('<div class="sec">🏆 Top 10 Students – I to V Semester</div>', unsafe_allow_html=True)
    if not top_df.empty:
        st.dataframe(top_df, use_container_width=True, hide_index=True)
        ct = chart_toppers(top_df)
        if ct: st.plotly_chart(ct, use_container_width=True)

with t6:
    st.markdown('<div class="sec">Complete Student Result Data</div>', unsafe_allow_html=True)
    search = st.text_input("🔍 Search by name, roll no. or status","")
    disp_df = df.copy()
    if search:
        mask = disp_df.astype(str).apply(
            lambda x: x.str.contains(search, case=False, na=False)
        ).any(axis=1)
        disp_df = disp_df[mask]
    st.dataframe(disp_df, use_container_width=True, hide_index=True, height=480)
    st.caption(f"Showing {len(disp_df)} of {len(df)} students")

# Export
st.markdown("---")
st.markdown("### 📥 Export Options")
d1,d2 = st.columns(2)
with d1:
    excel_data = build_excel(summary,sdf,fdf,gdf,top_df,df)
    st.download_button(
        label="⬇️ Download Excel Analysis (.xlsx)",
        data=excel_data,
        file_name="BCA_V_Result_Analysis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
with d2:
    html_data = build_print_html(summary,sdf,fdf,gdf,top_df)
    st.download_button(
        label="🖨️ Download Print-Ready Report (.html)",
        data=html_data,
        file_name="BCA_V_Result_Analysis_Report.html",
        mime="text/html",
        use_container_width=True,
        help="Open in browser → Ctrl+P → Save as PDF (Black & White)"
    )

st.markdown("""
<div class="info-box">
  <strong>🖨️ To print as PDF:</strong> Download the HTML report → open in any browser (Chrome/Edge/Firefox)
  → press <strong>Ctrl+P</strong> (Windows) or <strong>⌘P</strong> (Mac)
  → select <em>Save as PDF</em> → optionally choose <em>Black &amp; White</em> → Print.
</div>
""", unsafe_allow_html=True)
