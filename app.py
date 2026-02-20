from flask import Flask, render_template, request, send_file
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import os
import math
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
import matplotlib.pyplot as plt

def generate_pie_chart(subject_marks, enrollment):
    labels = list(subject_marks.keys())
    values = list(subject_marks.values())

    plt.figure(figsize=(10, 8))  # Bigger chart size

    wedges, texts, autotexts = plt.pie(
        values,
        autopct="%1.1f%%",
        startangle=140,
        pctdistance=0.8,
        labeldistance=1.1
    )

    # TITLE
    plt.title("Subject-wise Marks Distribution", fontsize=16)

    # LEGEND WITH COLORS
    plt.legend(
        wedges,
        labels,
        title="Subjects",
        loc="center left",
        bbox_to_anchor=(1, 0, 0.3, 1),
        fontsize=10
    )

    # Make sure layout doesn't cut labels
    plt.tight_layout()

    filepath = f"static/pie_{enrollment}.png"
    plt.savefig(filepath, dpi=150, bbox_inches="tight")
    plt.close()

    return filepath
app = Flask(__name__)

FILE_PATH = "sem5_result.xlsx"

# ---------------------------------------------
# SUBJECT MAPPING (FINAL & CORRECT)
# ---------------------------------------------
SUBJECTS = {
    "Python Programming (Theory)": ("Internal", "External", "Total"),
    "Cloud Computing": ("Internal.1", "External.1", "Total.1"),
    "Information Security": ("Internal.2", "External.2", "Total.2"),
    "Python Programming (Practical)": ("Internal.3", "External.3", "Total.3"),
    "Mobile App Development (Theory)": ("Internal.4", "External.4", "Total.4"),
    "Mobile App Development (Practical)": ("Internal.5", "External.5", "Total.5"),
    "Software Project Management": ("Internal.6", "External.6", "Total.6"),
    "Internship / Project – I": ("Internal.7", "External.7", "Total.7")
}

# ---------------------------------------------
# LOAD DATA CLEANLY
# ---------------------------------------------
def load_data():
    df = pd.read_excel(FILE_PATH, skiprows=6)
    df.columns = df.columns.astype(str).str.strip()

    df.rename(columns={
        "Enrollement No.": "EnrollmentNo",
        "Roll No": "RollNo"
    }, inplace=True)

    df["EnrollmentNo"] = df["EnrollmentNo"].astype(str).str.replace(".0", "", regex=False)
    df["RollNo"] = df["RollNo"].astype(str).str.replace(".0", "", regex=False)

    # Convert numeric fields only where required
    numeric_cols = ["Obtain", "SGPA V", "Max Marks", "T Cr", "T GP", "Total CP"]
    for col in numeric_cols:
        if col in df:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


# ---------------------------------------------
# GRADE CALCULATION
# ---------------------------------------------
def grade_from_marks(m):
    if pd.isna(m):
        return "-"
    if m >= 90: return "A+"
    if m >= 80: return "A"
    if m >= 70: return "B+"
    if m >= 60: return "B"
    if m >= 50: return "C"
    if m >= 40: return "Pass"
    return "Fail"


def pass_fail(total, internal, external):
    if pd.isna(total):
        return "Fail"
    if total >= 40 and internal >= 24 and external >= 16:
        return "Pass"
    return "Fail"


# ---------------------------------------------
# DASHBOARD
# ---------------------------------------------
@app.route("/")
def dashboard():
    df = load_data()

    total_students = len(df)
    avg_marks = round(df["Obtain"].mean(), 2)
    avg_sgpa = round(df["SGPA V"].mean(), 2)

    topper = df.sort_values(by="Obtain", ascending=False).iloc[0].to_dict()

    return render_template(
        "dashboard.html",
        total_students=total_students,
        avg_marks=avg_marks,
        avg_sgpa=avg_sgpa,
        topper=topper
    )


# ---------------------------------------------
# STUDENT LIST (PAGINATION)
# ---------------------------------------------
@app.route("/students")
def students():
    df = load_data()

    page = int(request.args.get("page", 1))
    per_page = 20
    total_pages = math.ceil(len(df) / per_page)
    start = (page - 1) * per_page
    end = start + per_page

    student_rows = df.iloc[start:end]

    return render_template(
        "students.html",
        students=student_rows.to_dict(orient="records"),
        page=page,
        total_pages=total_pages
    )


# ---------------------------------------------
# SEARCH STUDENT
# ---------------------------------------------
@app.route("/search")
def search_page():
    return render_template("search.html")


@app.route("/student")
def student_detail():
    df = load_data()
    enrollment = request.args.get("enrollment")

    student = df[df["EnrollmentNo"] == enrollment]

    if student.empty:
        return render_template("not_found.html")

    student = student.iloc[0]

    # SUBJECT-WISE DETAILS
    subject_data = []
    total_marks_list = []

    for name, cols in SUBJECTS.items():
        internal = student[cols[0]]
        external = student[cols[1]]
        total = student[cols[2]]
        grade = grade_from_marks(total)
        status = pass_fail(total, internal, external)

        subject_data.append({
            "name": name,
            "internal": internal,
            "external": external,
            "total": total,
            "grade": grade,
            "status": status
        })

        total_marks_list.append(total)

    # PIE CHART
    if not os.path.exists("static/charts"):
        os.makedirs("static/charts")

    plt.figure(figsize=(6, 6))
    plt.pie(total_marks_list, labels=[s["name"] for s in subject_data], autopct="%1.1f%%")
    plt.title("Subject-wise Marks Distribution")
    pie_path = f"static/charts/{enrollment}_pie.png"
    plt.savefig(pie_path)
    plt.close()

    return render_template(
        "student_detail.html",
        student=student,
        subject_data=subject_data,
        pie_chart=pie_path
    )


# ---------------------------------------------
# DOWNLOAD PDF
# ---------------------------------------------
@app.route("/download/<enrollment>")
def download_pdf(enrollment):
    df = load_data()
    student = df[df["EnrollmentNo"] == enrollment]

    if student.empty:
        return "Student Not Found"

    student = student.iloc[0]

    # Calculate Rank
    # Calculate Rank
    df_sorted = df.sort_values(by="Obtain", ascending=False).reset_index(drop=True)

    df_sorted["Rank"] = df_sorted.index + 1  # Avoid Series issue
    df_sorted["EnrollmentNo"] = df_sorted["EnrollmentNo"].astype(str)

    student_en = str(enrollment)

    rank = int(df_sorted[df_sorted["EnrollmentNo"] == student_en]["Rank"].iloc[0])

    filename = f"{enrollment}_result.pdf"
    doc = SimpleDocTemplate(filename)
    styles = getSampleStyleSheet()
    elements = []

    # Title
    elements.append(Paragraph("<b><font size=18>Semester 5 Result Report</font></b>", styles["Title"]))
    elements.append(Spacer(1, 0.3 * inch))

    # Student info
    elements.append(Paragraph("<b>Student Details</b>", styles["Heading2"]))
    elements.append(Paragraph(f"Enrollment No: {student.EnrollmentNo}", styles["Normal"]))
    elements.append(Paragraph(f"Roll No: {student.RollNo}", styles["Normal"]))
    elements.append(Paragraph(f"Rank: {rank}", styles["Normal"]))
    elements.append(Spacer(1, 0.2 * inch))

    # Subject Table
    elements.append(Paragraph("<b>Subject Wise Marks</b>", styles["Heading2"]))

    table_data = [["Subject", "Internal", "External", "Total", "Grade", "Status"]]

    for name, cols in SUBJECTS.items():
        internal = student[cols[0]]
        external = student[cols[1]]
        total = student[cols[2]]
        grade = grade_from_marks(total)
        status = pass_fail(total, internal, external)

        table_data.append([name, internal, external, total, grade, status])

    from reportlab.platypus import Table, TableStyle

    table = Table(table_data)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), "#cce5ff"),
        ("TEXTCOLOR", (0, 0), (-1, 0), "black"),
        ("GRID", (0, 0), (-1, -1), 0.5, "black"),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("ALIGN", (1, 1), (-1, -1), "CENTER"),
    ]))

    elements.append(table)
    elements.append(Spacer(1, 0.3 * inch))

    # Summary
    elements.append(Paragraph("<b>Summary</b>", styles["Heading2"]))
    elements.append(Paragraph(f"Total Marks: {student.Obtain} / {student['Max Marks']}", styles["Normal"]))
    elements.append(Paragraph(f"SGPA: {student['SGPA V']}", styles["Normal"]))
    elements.append(Paragraph(f"Total Credits: {student['T Cr']}", styles["Normal"]))
    elements.append(Paragraph(f"Grade Points: {student['T GP']}", styles["Normal"]))
    elements.append(Paragraph(f"Total CP: {student['Total CP']}", styles["Normal"]))

    doc.build(elements)

    return send_file(filename, as_attachment=True)


if __name__ == "__main__":
    app.run()