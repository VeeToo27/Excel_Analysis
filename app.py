import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import plotly.express as px
import plotly.graph_objects as go
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==========================
# PAGE CONFIG
# ==========================
st.set_page_config(page_title="Student Result Analysis", layout="wide", page_icon="🎓")

st.title("🎓 Student Result Analysis System")
st.caption("Upload Excel → Full Analysis → Graphs → Download")

# ==========================
# CACHING
# ==========================
@st.cache_data
def load_data(file_bytes):
    try:
        results_df = pd.read_excel(file_bytes, sheet_name=0)
        faculty_df = pd.read_excel(file_bytes, sheet_name=1)
        return results_df, faculty_df, None
    except Exception as e:
        return None, None, str(e)

@st.cache_data
def process_results(results_raw):
    df = results_raw.iloc[1:].reset_index(drop=True)
    rename_map = {}
    col_lower = {str(c).strip().lower(): c for c in df.columns}
    
    for key, targets in {
        'RollNo':     ['roll no.', 'roll no', 'rollno'],
        'Name':       ['candidate name', 'name'],
        'Status':     ['status'],
        'BackPaper':  ['back paper'],
        'Division':   ['per.1', 'division'],
        'GrandTotal': ['total\ni to vi', 'grand total', 'total']
    }.items():
        for t in targets:
            if t in col_lower:
                rename_map[col_lower[t]] = key
                break

    df.rename(columns=rename_map, inplace=True)
    return df

# ==========================
# FILE UPLOAD
# ==========================
uploaded_file = st.file_uploader("📂 Upload Result Excel File", type=["xlsx"])

if not uploaded_file:
    st.info("👆 Please upload an Excel file with results on Sheet 1 and faculty mapping on Sheet 2.")
    st.stop()

# Load with caching
file_bytes = uploaded_file.read()
results_raw, faculty_df, load_error = load_data(BytesIO(file_bytes))

if load_error:
    st.error(f"❌ Failed to read the file: {load_error}")
    st.stop()

if results_raw is None or faculty_df is None:
    st.error("❌ Could not find both required sheets. Ensure Sheet 1 = Results, Sheet 2 = Faculty.")
    st.stop()

# ==========================
# PROCESS DATA
# ==========================
try:
    results_df = process_results(results_raw)
except Exception as e:
    st.error(f"❌ Error processing results: {e}")
    st.stop()

required_cols = {'RollNo', 'Name', 'Status', 'GrandTotal'}
missing = required_cols - set(results_df.columns)
if missing:
    st.error(f"❌ Could not find these columns: {missing}. Check your Excel column headers.")
    st.stop()

# ==========================
# OVERALL STATS
# ==========================
total_students = len(results_df)
if total_students == 0:
    st.error("❌ No student records found.")
    st.stop()

passed_df   = results_df[results_df['Status'].isin(['FIRST', 'SECOND', 'PASS'])]
later_count = len(results_df[results_df['Status'] == 'LATER'])
fail_count  = len(results_df[results_df['Status'] == 'FAIL'])
ufm_count   = len(results_df[results_df['Status'] == 'UFM'])
bp_count    = results_df['BackPaper'].notna().sum() if 'BackPaper' in results_df.columns else 0
first_div   = len(results_df[results_df['Status'] == 'FIRST'])
second_div  = len(results_df[results_df['Status'] == 'SECOND'])
pass_pct    = round(len(passed_df) / total_students * 100, 2)

# ==========================
# SUBJECT TOTAL DETECTION
# ==========================
subject_total_cols = {}
columns = list(results_df.columns)
for i, col in enumerate(columns):
    col_str = str(col).strip()
    if col_str.replace('P', '').isdigit():
        if i + 2 < len(columns):
            subject_total_cols[col_str] = columns[i + 2]

# ==========================
# SUBJECT ANALYSIS (cached)
# ==========================
@st.cache_data
def compute_subject_analysis(df_json, subject_total_cols):
    df = pd.read_json(df_json)
    rows = []
    for code, total_col in subject_total_cols.items():
        if total_col not in df.columns:
            continue
        marks = pd.to_numeric(df[total_col], errors='coerce')
        rows.append({
            "Subject Code":   code,
            "Highest Marks":  marks.max(),
            "Average Marks":  round(marks.mean(), 2),
            "Failed Students":(marks < 33).sum(),
            "Pass %":         round((marks >= 33).mean() * 100, 2)
        })
    return pd.DataFrame(rows)

@st.cache_data
def compute_faculty_analysis(df_json, faculty_json, subject_total_cols):
    df         = pd.read_json(df_json)
    faculty_df = pd.read_json(faculty_json)
    rows = []
    for _, row in faculty_df.iterrows():
        code    = str(row.get('Subject Code', ''))
        faculty = row.get('Faculty Name', 'Unknown')
        if code not in subject_total_cols:
            continue
        total_col = subject_total_cols[code]
        if total_col not in df.columns:
            continue
        marks = pd.to_numeric(df[total_col], errors='coerce')
        rows.append({
            "Faculty Name":   faculty,
            "Subject Code":   code,
            "Total Students": int(marks.count()),
            "Passed":         int((marks >= 33).sum()),
            "Failed":         int((marks < 33).sum()),
            "Pass %":         round((marks >= 33).mean() * 100, 2),
            "Highest Marks":  marks.max(),
            "Lowest Marks":   marks.min(),
            "Average Marks":  round(marks.mean(), 2)
        })
    return pd.DataFrame(rows)

df_json      = results_df.to_json()
faculty_json = faculty_df.to_json()

subject_df       = compute_subject_analysis(df_json, subject_total_cols)
faculty_df_final = compute_faculty_analysis(df_json, faculty_json, subject_total_cols)

# Weak subjects
weak_subjects = subject_df[
    (subject_df['Pass %'] < 60) |
    (subject_df['Failed Students'] > subject_df['Failed Students'].mean())
] if not subject_df.empty else pd.DataFrame()

# Top 5
top5 = results_df.sort_values(by="GrandTotal", ascending=False)[
    ['Name', 'RollNo', 'GrandTotal']
].head(5)

# Failed df
failed_df = results_df[results_df['Status'].isin(['FAIL', 'LATER', 'UFM'])].copy()
def failure_reason(row):
    return {'FAIL': 'Low Marks', 'LATER': 'Back Paper / Absent', 'UFM': 'UFM'}.get(row['Status'], 'Unknown')
failed_df['Failure Reason'] = failed_df.apply(failure_reason, axis=1)

# Grade distribution
results_df['GrandTotal'] = pd.to_numeric(results_df['GrandTotal'], errors='coerce')
def grade_bucket(score):
    if pd.isna(score): return 'N/A'
    if score >= 90: return 'A+ (90-100)'
    if score >= 75: return 'A  (75-89)'
    if score >= 60: return 'B  (60-74)'
    if score >= 45: return 'C  (45-59)'
    if score >= 33: return 'D  (33-44)'
    return 'F  (<33)'
results_df['Grade'] = results_df['GrandTotal'].apply(grade_bucket)

# ==========================
# SIDEBAR FILTERS
# ==========================
st.sidebar.header("🔍 Filters")
status_options = ['All'] + sorted(results_df['Status'].dropna().unique().tolist())
selected_status = st.sidebar.selectbox("Filter by Status", status_options)

filtered_df = results_df if selected_status == 'All' else results_df[results_df['Status'] == selected_status]

st.sidebar.markdown("---")
st.sidebar.markdown("### 🔎 Student Lookup")
search_roll = st.sidebar.text_input("Enter Roll Number")

# ==========================
# TABS
# ==========================
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📊 Summary", "📘 Subject Analysis",
    "🧑‍🏫 Faculty Analysis", "👥 Student Details",
    "📈 Charts", "📥 Download"
])

# ── TAB 1: SUMMARY ──────────────────────────────────────────────────────────
with tab1:
    st.subheader("📊 Overall Summary")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Students", total_students)
    c2.metric("Passed",         len(passed_df))
    c3.metric("Pass %",         f"{pass_pct}%")
    c4.metric("Later",          later_count)

    c5, c6, c7, c8 = st.columns(4)
    c5.metric("First Division",  first_div)
    c6.metric("Second Division", second_div)
    c7.metric("Back Papers",     bp_count)
    c8.metric("Failed",          fail_count)

    st.divider()
    st.subheader("🏆 Top 5 Toppers")
    st.dataframe(top5, use_container_width=True)

    st.divider()
    st.subheader("📊 Grade Distribution")
    grade_order = ['A+ (90-100)', 'A  (75-89)', 'B  (60-74)', 'C  (45-59)', 'D  (33-44)', 'F  (<33)', 'N/A']
    grade_counts = results_df['Grade'].value_counts().reindex(grade_order, fill_value=0).reset_index()
    grade_counts.columns = ['Grade', 'Count']
    fig_grade = px.bar(
        grade_counts, x='Grade', y='Count', text='Count',
        color='Grade',
        color_discrete_sequence=px.colors.sequential.Reds[::-1],
        title="Overall Grade Distribution"
    )
    fig_grade.update_layout(showlegend=False, template="plotly_white")
    st.plotly_chart(fig_grade, use_container_width=True)

    st.divider()
    st.subheader("🥧 Pass / Fail Breakdown")
    pie_data = pd.DataFrame({
        'Category': ['First Division', 'Second Division', 'Pass', 'Later', 'Fail', 'UFM'],
        'Count': [
            first_div, second_div,
            len(results_df[results_df['Status'] == 'PASS']),
            later_count, fail_count, ufm_count
        ]
    })
    pie_data = pie_data[pie_data['Count'] > 0]
    fig_pie = px.pie(pie_data, names='Category', values='Count',
                     title="Result Breakdown", hole=0.4)
    st.plotly_chart(fig_pie, use_container_width=True)

# ── TAB 2: SUBJECT ANALYSIS ─────────────────────────────────────────────────
with tab2:
    st.subheader("📘 Subject-wise Analysis")
    if subject_df.empty:
        st.warning("No subject columns detected.")
    else:
        st.dataframe(
            subject_df.style.background_gradient(subset=['Pass %'], cmap='Reds')
                            .background_gradient(subset=['Failed Students'], cmap='Reds'),
            use_container_width=True
        )

        st.divider()
        st.subheader("📉 Weak Subjects (Pass% < 60 or High Failures)")
        if weak_subjects.empty:
            st.success("No weak subjects detected 🎯")
        else:
            st.warning(f"{len(weak_subjects)} subject(s) need academic intervention:")
            st.dataframe(weak_subjects, use_container_width=True)

        st.divider()
        st.subheader("⚠️ Low-Margin Failures (30–32 Marks)")
        low_margin_list = []
        for code, total_col in subject_total_cols.items():
            if total_col not in results_df.columns:
                continue
            marks = pd.to_numeric(results_df[total_col], errors='coerce')
            temp = results_df[(marks >= 30) & (marks < 33)][['Name', 'RollNo']].copy()
            temp['Subject Code'] = code
            temp['Marks'] = marks[(marks >= 30) & (marks < 33)].values
            low_margin_list.append(temp)

        if low_margin_list:
            lm_df = pd.concat(low_margin_list, ignore_index=True)
            if not lm_df.empty:
                st.dataframe(lm_df, use_container_width=True)
            else:
                st.success("No low-margin failures found 🎯")
        else:
            st.success("No low-margin failures found 🎯")

# ── TAB 3: FACULTY ANALYSIS ─────────────────────────────────────────────────
with tab3:
    st.subheader("🧑‍🏫 Faculty-wise Performance")
    if faculty_df_final.empty:
        st.warning("No faculty data matched subject codes.")
    else:
        # Quadrant annotation helper
        avg_pass = faculty_df_final['Pass %'].mean()
        avg_fail = faculty_df_final['Failed'].mean()

        st.dataframe(
            faculty_df_final.sort_values('Pass %', ascending=False)
                            .style.background_gradient(subset=['Pass %'], cmap='Reds'),
            use_container_width=True
        )

        st.divider()
        st.subheader("📈 Faculty Performance: Avg Marks vs Failures")
        fig_scatter = px.scatter(
            faculty_df_final,
            x="Average Marks", y="Failed",
            size="Failed", color="Pass %",
            hover_name="Faculty Name",
            hover_data=["Subject Code", "Passed"],
            title="Faculty: Avg Marks vs Failures (size = failure count)",
            color_continuous_scale="Reds"
        )
        fig_scatter.add_hline(y=avg_fail,  line_dash="dash", line_color="red",
                              annotation_text="Avg Failures")
        fig_scatter.add_vline(x=faculty_df_final['Average Marks'].mean(),
                              line_dash="dash", line_color="blue",
                              annotation_text="Avg Marks")
        fig_scatter.update_layout(template="plotly_white")
        st.plotly_chart(fig_scatter, use_container_width=True)

# ── TAB 4: STUDENT DETAILS ──────────────────────────────────────────────────
with tab4:
    st.subheader("👥 Student Details")

    # Student lookup
    if search_roll:
        student = results_df[results_df['RollNo'].astype(str) == search_roll.strip()]
        if student.empty:
            st.warning(f"No student found with Roll No: {search_roll}")
        else:
            st.success(f"Found: {student['Name'].values[0]}")
            st.dataframe(student.T, use_container_width=True)

            # Subject-wise marks for this student
            sub_marks = []
            for code, total_col in subject_total_cols.items():
                if total_col in student.columns:
                    sub_marks.append({
                        "Subject": code,
                        "Marks":   pd.to_numeric(student[total_col].values[0], errors='coerce')
                    })
            if sub_marks:
                sub_marks_df = pd.DataFrame(sub_marks)
                fig_radar = px.bar(sub_marks_df, x='Subject', y='Marks',
                                   title=f"Subject-wise Marks – {student['Name'].values[0]}",
                                   color='Marks', color_continuous_scale='Blues')
                fig_radar.add_hline(y=33, line_dash='dot', line_color='red',
                                    annotation_text='Pass Mark (33)')
                fig_radar.update_layout(template='plotly_white')
                st.plotly_chart(fig_radar, use_container_width=True)
        st.divider()

    # Filtered table
    st.subheader(f"📋 Students — Status: {selected_status}")
    display_cols = [c for c in ['RollNo', 'Name', 'Status', 'GrandTotal', 'Grade', 'BackPaper'] if c in filtered_df.columns]
    st.dataframe(filtered_df[display_cols], use_container_width=True)

    st.divider()
    st.subheader("❌ Failed / LATER Students")
    fail_cols = [c for c in ['Name', 'RollNo', 'Status', 'Failure Reason'] if c in failed_df.columns]
    st.dataframe(failed_df[fail_cols], use_container_width=True)

    st.divider()
    st.subheader("🚨 Students with Multiple Back Papers")
    if 'BackPaper' in results_df.columns:
        bp_students = results_df[results_df['BackPaper'].notna()]
        multiple_bp = bp_students['RollNo'].value_counts()
        multiple_bp = multiple_bp[multiple_bp > 1]
        if not multiple_bp.empty:
            st.dataframe(
                multiple_bp.reset_index().rename(columns={'RollNo': 'Roll No', 'count': 'BP Count'}),
                use_container_width=True
            )
        else:
            st.success("No students with multiple back papers 🎉")
    else:
        st.info("Back paper column not found.")

# ── TAB 5: CHARTS ───────────────────────────────────────────────────────────
with tab5:
    st.subheader("📊 Interactive Analytics")

    if not subject_df.empty:
        st.markdown("### 📘 Subject-wise Pass Percentage")
        fig_sub = px.bar(subject_df, x="Subject Code", y="Pass %", text="Pass %",
                         color="Pass %", color_continuous_scale="Blues",
                         title="Subject-wise Pass %")
        fig_sub.update_layout(yaxis_range=[0, 100], template="plotly_white")
        st.plotly_chart(fig_sub, use_container_width=True)

        st.markdown("### ❌ Subject-wise Failure Count")
        fig_fail = px.bar(subject_df, x="Subject Code", y="Failed Students",
                          text="Failed Students", color="Failed Students",
                          color_continuous_scale="Reds", title="Subject-wise Failures")
        fig_fail.update_layout(template="plotly_white")
        st.plotly_chart(fig_fail, use_container_width=True)

        st.markdown("### 📊 Average Marks by Subject")
        fig_avg = px.bar(subject_df, x="Subject Code", y="Average Marks",
                         text="Average Marks", color="Average Marks",
                         color_continuous_scale="Viridis", title="Subject Average Marks")
        fig_avg.add_hline(y=33, line_dash='dot', line_color='red',
                          annotation_text='Pass Mark')
        fig_avg.update_layout(template="plotly_white")
        st.plotly_chart(fig_avg, use_container_width=True)

    if not faculty_df_final.empty:
        st.markdown("### 🧑‍🏫 Faculty-wise Pass Percentage")
        fig_fac = px.bar(faculty_df_final, x="Faculty Name", y="Pass %",
                         text="Pass %", color="Pass %",
                         color_continuous_scale="Greens", title="Faculty Pass %")
        fig_fac.update_layout(yaxis_range=[0, 100], xaxis_tickangle=-30,
                              template="plotly_white")
        st.plotly_chart(fig_fac, use_container_width=True)

# ── TAB 6: DOWNLOAD ─────────────────────────────────────────────────────────
with tab6:
    st.subheader("📥 Download Analyzed Excel Report")

    def style_sheet(ws, header_color="1F4E79"):
        """Apply professional styling to a worksheet."""
        header_fill = PatternFill("solid", fgColor=header_color)
        header_font = Font(bold=True, color="FFFFFF")
        thin_border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = thin_border

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='center')

        for col in ws.columns:
            max_len = max((len(str(c.value)) for c in col if c.value), default=10)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 30)

    def pass_color_fill(val):
        if isinstance(val, (int, float)):
            if val >= 80:  return PatternFill("solid", fgColor="C6EFCE")
            if val >= 60:  return PatternFill("solid", fgColor="FFEB9C")
            return         PatternFill("solid", fgColor="FFC7CE")
        return None

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        sheets = {
            "Subject Analysis": subject_df,
            "Faculty Analysis": faculty_df_final,
            "Weak Subjects":    weak_subjects if not weak_subjects.empty else pd.DataFrame({"Message": ["No weak subjects"]}),
            "Top 5 Toppers":    top5,
            "Failed Students":  failed_df[[c for c in ['Name','RollNo','Status','Failure Reason'] if c in failed_df.columns]],
            "All Students":     results_df[[c for c in ['RollNo','Name','Status','GrandTotal','Grade'] if c in results_df.columns]]
        }
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

        wb = writer.book
        colors = {
            "Subject Analysis": "1F4E79",
            "Faculty Analysis": "375623",
            "Weak Subjects":    "843C0C",
            "Top 5 Toppers":    "7030A0",
            "Failed Students":  "C00000",
            "All Students":     "1F4E79"
        }
        for sheet_name in sheets:
            ws = wb[sheet_name]
            style_sheet(ws, colors.get(sheet_name, "1F4E79"))

            # Color-code Pass % column
            if sheet_name in ("Subject Analysis", "Faculty Analysis"):
                headers = [cell.value for cell in ws[1]]
                if "Pass %" in headers:
                    pass_col_idx = headers.index("Pass %") + 1
                    for row in ws.iter_rows(min_row=2, min_col=pass_col_idx, max_col=pass_col_idx):
                        for cell in row:
                            fill = pass_color_fill(cell.value)
                            if fill:
                                cell.fill = fill

    st.download_button(
        label="📥 Download Full Analysis Report (.xlsx)",
        data=output.getvalue(),
        file_name="Complete_Result_Analysis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.markdown("""
    **📋 Report contains:**
    - Subject Analysis (color-coded Pass %)
    - Faculty Analysis (color-coded Pass %)
    - Weak Subjects
    - Top 5 Toppers
    - Failed Students
    - All Students with Grades
    """)


