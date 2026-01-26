import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import matplotlib.pyplot as plt
import plotly.express as px


# ==========================
# PAGE CONFIG
# ==========================
st.set_page_config(page_title="Student Result Analysis", layout="wide")
st.title("🎓 Student Result Analysis System")
st.caption("Upload Excel → Full Analysis → Graphs → Download")

# ==========================
# FILE UPLOAD
# ==========================
uploaded_file = st.file_uploader("Upload Result Excel File", type=["xlsx"])

if uploaded_file:

    # ==========================
    # LOAD SHEETS
    # ==========================
    results_df = pd.read_excel(uploaded_file, sheet_name=0)
    faculty_df = pd.read_excel(uploaded_file, sheet_name=1)

    # ==========================
    # CLEAN RESULT SHEET
    # ==========================
    results_df = results_df.iloc[1:].reset_index(drop=True)

    results_df.rename(columns={
        'Roll No.': 'RollNo',
        'CANDIDATE NAME': 'Name',
        'Status': 'Status',
        'BACK PAPER': 'BackPaper',
        'PER.1': 'Division',
        'Total\nI to VI': 'GrandTotal'
    }, inplace=True)

    # ==========================
    # OVERALL STATS
    # ==========================
    total_students = len(results_df)
    passed_df = results_df[results_df['Status'].isin(['FIRST', 'SECOND', 'PASS'])]

    later_count = len(results_df[results_df['Status'] == 'LATER'])
    fail_count = len(results_df[results_df['Status'] == 'FAIL'])
    ufm_count = len(results_df[results_df['Status'] == 'UFM'])
    bp_count = results_df['BackPaper'].notna().sum()

    first_div = len(results_df[results_df['Status'] == 'FIRST'])
    second_div = len(results_df[results_df['Status'] == 'SECOND'])

    pass_percentage = round((len(passed_df) / total_students) * 100, 2)

    # ==========================
    # DASHBOARD METRICS
    # ==========================
    st.subheader("📊 Overall Summary")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Students", total_students)
    c2.metric("Passed", len(passed_df))
    c3.metric("Later", later_count)
    c4.metric("Pass %", pass_percentage)

    c5, c6, c7, c8 = st.columns(4)
    c5.metric("First Division", first_div)
    c6.metric("Second Division", second_div)
    c7.metric("Back Papers", bp_count)
    c8.metric("Failed", fail_count)

    # ==========================
    # TOPPERS
    # ==========================
    st.subheader("🏆 Top 5 Toppers")

    top5 = results_df.sort_values(
        by="GrandTotal", ascending=False
    )[['Name', 'RollNo', 'GrandTotal']].head(5)

    st.dataframe(top5, use_container_width=True)

    # ==========================
    # FAILED / LATER STUDENTS
    # ==========================
    st.subheader("❌ Failed / LATER Students")

    failed_df = results_df[
        results_df['Status'].isin(['FAIL', 'LATER', 'UFM'])
    ].copy()

    def failure_reason(row):
        if row['Status'] == 'FAIL':
            return 'Low Marks'
        elif row['Status'] == 'LATER':
            return 'Back Paper / Absent'
        elif row['Status'] == 'UFM':
            return 'UFM'
        return 'Unknown'

    failed_df['Failure Reason'] = failed_df.apply(failure_reason, axis=1)

    st.dataframe(
        failed_df[['Name', 'RollNo', 'Status', 'Failure Reason']],
        use_container_width=True
    )

    # ==========================
    # SUBJECT TOTAL DETECTION (NUMERIC + P)
    # ==========================
    subject_total_cols = {}
    columns = list(results_df.columns)

    for i, col in enumerate(columns):
        col_str = str(col).strip()
        if col_str.replace('P', '').isdigit():
            if i + 2 < len(columns):
                subject_total_cols[col_str] = columns[i + 2]

    # ==========================
    # SUBJECT-WISE ANALYSIS
    # ==========================
    st.subheader("📘 Subject-wise Analysis")

    subject_analysis = []

    for code, total_col in subject_total_cols.items():
        marks = pd.to_numeric(results_df[total_col], errors='coerce')

        subject_analysis.append({
            "Subject Code": code,
            "Highest Marks": marks.max(),
            "Average Marks": round(marks.mean(), 2),
            "Failed Students": (marks < 33).sum(),
            "Pass %": round((marks >= 33).mean() * 100, 2)
        })

    subject_df = pd.DataFrame(subject_analysis)
    st.dataframe(subject_df, use_container_width=True)

    # ==========================
    # FACULTY-WISE PERFORMANCE
    # ==========================
    st.subheader("🧑‍🏫 Faculty-wise Performance")

    faculty_analysis = []

    for _, row in faculty_df.iterrows():
        code = str(row['Subject Code'])
        faculty = row['Faculty Name']

        if code not in subject_total_cols:
            continue

        marks = pd.to_numeric(results_df[subject_total_cols[code]], errors='coerce')

        faculty_analysis.append({
            "Faculty Name": faculty,
            "Subject Code": code,
            "Total Students": marks.count(),
            "Passed": (marks >= 33).sum(),
            "Failed": (marks < 33).sum(),
            "Pass %": round((marks >= 33).mean() * 100, 2),
            "Highest Marks": marks.max(),
            "Lowest Marks": marks.min(),
            "Average Marks": round(marks.mean(), 2)
        })

    faculty_df_final = pd.DataFrame(faculty_analysis)
    st.dataframe(faculty_df_final, use_container_width=True)

    # ==========================
    # MULTIPLE BACK PAPERS
    # ==========================
    st.subheader("🚨 Students with Multiple Back Papers")

    bp_students = results_df[results_df['BackPaper'].notna()]
    multiple_bp = bp_students['RollNo'].value_counts()
    multiple_bp = multiple_bp[multiple_bp > 1]

    if not multiple_bp.empty:
        st.dataframe(
            multiple_bp.reset_index().rename(
                columns={'index': 'Roll No', 'RollNo': 'BP Count'}
            ),
            use_container_width=True
        )
    else:
        st.success("No students with multiple back papers 🎉")

    # ==========================
    # LOW MARGIN FAILURES
    # ==========================
    st.subheader("⚠️ Low-Margin Failures (30–32 Marks)")

    low_margin_list = []

    for code, total_col in subject_total_cols.items():
        marks = pd.to_numeric(results_df[total_col], errors='coerce')
        temp = results_df[(marks >= 30) & (marks < 33)][['Name', 'RollNo']]
        temp['Subject Code'] = code
        temp['Marks'] = marks[(marks >= 30) & (marks < 33)]
        low_margin_list.append(temp)

    if low_margin_list:
        st.dataframe(pd.concat(low_margin_list), use_container_width=True)
    else:
        st.success("No low-margin failures found 🎯")

    # ==========================
    # 📉 WEAK SUBJECT DETECTION
    # ==========================
    st.subheader("📉 Weak Subject Detection")

    weak_subjects = subject_df[
        (subject_df['Pass %'] < 60) |
        (subject_df['Failed Students'] > subject_df['Failed Students'].mean())
    ]

    if not weak_subjects.empty:
        st.warning("Subjects needing academic intervention:")
        st.dataframe(weak_subjects, use_container_width=True)
    else:
        st.success("No weak subjects detected 🎯")

    # ==========================
    # 📊 INTERACTIVE PLOTLY CHARTS
    # ==========================
    st.subheader("📊 Interactive Analytics Dashboard")
    
    # --------------------------
    # SUBJECT-WISE PASS %
    # --------------------------
    st.markdown("### 📘 Subject-wise Pass Percentage")
    
    fig_subject = px.bar(
        subject_df,
        x="Subject Code",
        y="Pass %",
        text="Pass %",
        color="Pass %",
        color_continuous_scale="Blues",
        title="Subject-wise Pass Percentage",
    )
    
    fig_subject.update_layout(
        xaxis_title="Subject Code",
        yaxis_title="Pass Percentage",
        yaxis_range=[0, 100],
        template="plotly_white"
    )
    
    st.plotly_chart(fig_subject, use_container_width=True)
    
    # --------------------------
    # SUBJECT FAILURE COUNT
    # --------------------------
    st.markdown("### ❌ Subject-wise Failure Count")
    
    fig_subject_fail = px.bar(
        subject_df,
        x="Subject Code",
        y="Failed Students",
        text="Failed Students",
        color="Failed Students",
        color_continuous_scale="Reds",
        title="Subject-wise Failure Count",
    )
    
    fig_subject_fail.update_layout(
        xaxis_title="Subject Code",
        yaxis_title="Number of Failed Students",
        template="plotly_white"
    )
    
    st.plotly_chart(fig_subject_fail, use_container_width=True)
    
    # --------------------------
    # FACULTY-WISE PASS %
    # --------------------------
    st.markdown("### 🧑‍🏫 Faculty-wise Pass Percentage")
    
    fig_faculty = px.bar(
        faculty_df_final,
        x="Faculty Name",
        y="Pass %",
        text="Pass %",
        color="Pass %",
        color_continuous_scale="Greens",
        title="Faculty-wise Pass Percentage",
    )
    
    fig_faculty.update_layout(
        xaxis_title="Faculty Name",
        yaxis_title="Pass Percentage",
        yaxis_range=[0, 100],
        template="plotly_white",
        xaxis_tickangle=-30
    )
    
    st.plotly_chart(fig_faculty, use_container_width=True)
    
    # --------------------------
    # FACULTY AVG MARKS vs FAIL
    # --------------------------
    st.markdown("### 📈 Faculty Avg Marks vs Failures")
    
    fig_faculty_scatter = px.scatter(
        faculty_df_final,
        x="Average Marks",
        y="Failed",
        size="Failed",
        color="Pass %",
        hover_name="Faculty Name",
        title="Faculty Performance: Avg Marks vs Failures",
        color_continuous_scale="Viridis"
    )
    
    fig_faculty_scatter.update_layout(
        xaxis_title="Average Marks",
        yaxis_title="Failed Students",
        template="plotly_white"
    )
    
    st.plotly_chart(fig_faculty_scatter, use_container_width=True)


    # ==========================
    # 📥 DOWNLOAD ANALYZED EXCEL
    # ==========================
    st.subheader("📥 Download Analyzed Excel")

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        subject_df.to_excel(writer, sheet_name="Subject Analysis", index=False)
        faculty_df_final.to_excel(writer, sheet_name="Faculty Analysis", index=False)
        weak_subjects.to_excel(writer, sheet_name="Weak Subjects", index=False)
        top5.to_excel(writer, sheet_name="Top 5 Toppers", index=False)
        failed_df.to_excel(writer, sheet_name="Failed Students", index=False)

    st.download_button(
        label="📥 Download Full Analysis Report",
        data=output.getvalue(),
        file_name="Complete_Result_Analysis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


