import docx
from docx.shared import Inches, Pt
from pathlib import Path

# --- Configuration ---
# Updated filename for Experian
path = Path(Path.home().joinpath("Desktop", "CSV", "Jeremiah_Mulwa_Experian_Analyst_CV.docx"))
path.parent.mkdir(parents=True, exist_ok=True)

# Personal details
FULL_NAME = "MULWA KITALA JEREMIAH"
ADDRESS_LINES = [
    "Meru University of Science and Technology",
    "P.O Box 972-60200",
    "MERU"
]
EMAIL = "mulwajeremy24@gmail.com"
PHONE = "0713916894"
GITHUB = "https://github.com/jeremy-dhy98"
LINKEDIN = "https://www.linkedin.com/in/jeremiah-kitala-15aa57288"

# Build document
doc = docx.Document()

# --- Header / Contact ---
header = doc.add_paragraph()
header.alignment = 1 # Center align header
header_run = header.add_run(FULL_NAME + "\n")
header_run.bold = True
header_run.font.size = Pt(18)
header.add_run(f"\nEmail: {EMAIL}\nPhone: {PHONE}\n")
header.add_run(f"\nLinkedIn: {LINKEDIN}\nGitHub: {GITHUB}\n")

# --- Professional Summary (Tailored for Experian & Financial Data) ---
doc.add_heading("Professional Summary", level=1)
p_summary = doc.add_paragraph()
p_summary.add_run("Technically adept ")
p_summary.add_run("Applied Statistics").bold = True
p_summary.add_run(" professional with a focus on ")
p_summary.add_run("Predictive Modeling and Quantitative Analysis").bold = True
p_summary.add_run(". Specialized in leveraging ")
p_summary.add_run("Python and SQL").bold = True
p_summary.add_run(" to derive actionable insights from complex datasets. Expert in ")
p_summary.add_run("Statistical Inference and Time-Series Analysis").bold = True
p_summary.add_run(", with a strong commitment to data integrity and governance. Eager to contribute to Experian’s global mission by applying rigorous mathematical frameworks to enhance credit risk assessment, decision analytics, and consumer insights.")

# --- Education ---
doc.add_heading("Education", level=1)
p_edu = doc.add_paragraph()
p_edu.add_run("Bachelor of Science in Mathematics and Computer Science").bold = True
p_edu.add_run("\nSpecialization: Applied Statistics (2021 - 2025)")
p_edu.add_run("\nMeru University of Science and Technology")

# --- Relevant Coursework (Refocused for Decision Analytics) ---
doc.add_heading("Relevant Coursework", level=2)
courses = [
    "Statistical Inference & Regression Modeling (Credit Risk Focus)",
    "Time Series Analysis & Economic Forecasting",
    "Probability Theory & Mathematical Statistics",
    "Advanced Database Systems (SQL & Schema Optimization)",
    "Design and Analysis of Algorithms"
]
for course in courses:
    doc.add_paragraph(course, style='List Bullet')

# --- Technical Skills (Prioritizing Scalability & Precision) ---
doc.add_heading("Technical Skills", level=1)
skills = {
    "Programming": "Python (Pandas, Scikit-learn, NumPy), C#",
    "Modeling & Analytics": "Predictive Modeling, Regression Analysis, Hypothesis Testing, Outlier Detection",
    "Data Engineering": "SQL (PostgreSQL, SQLite), ETL Pipeline Construction, API/WebSocket Integration",
    "Analytics Tools": "SPSS, STATA, MATLAB, Excel (VBA & Pivot Tables)"
}
for k, v in skills.items():
    p = doc.add_paragraph()
    p.add_run(f"{k}: ").bold = True
    p.add_run(v)

# --- Projects & Experience (Reframed for Financial Data Impact) ---
doc.add_heading("Projects & Experience", level=1)

# Project 1: Focus on Pipeline Integrity (Crucial for Financial Data)
p1_head = doc.add_paragraph()
p1_head.add_run("End-to-End Automated Data Pipelines").bold = True
doc.add_paragraph("Engineered robust ETL pipelines using Python to fetch and validate high-frequency data, ensuring high availability and data quality.", style='List Bullet')
doc.add_paragraph("Implemented comprehensive error-handling and data-logging protocols to maintain data integrity, a core requirement for regulatory financial environments.", style='List Bullet')

# Project 2: Focus on Quantitative Insight
p2_head = doc.add_paragraph()
p2_head.add_run("Statistical Modeling & Predictive Analytics").bold = True
doc.add_paragraph("Performed in-depth Exploratory Data Analysis (EDA) on large-scale datasets to uncover hidden patterns and correlations used for forecasting.", style='List Bullet')
doc.add_paragraph("Developed regression models using Python's scikit-learn to identify key variables impacting outcome probability, supporting data-driven decision making.", style='List Bullet')
doc.add_paragraph("Created automated summary reports and visualizations in Matplotlib/Seaborn to present complex quantitative findings to executive stakeholders.", style='List Bullet')

# Project 3: Automation
p3_head = doc.add_paragraph()
p3_head.add_run("Business Process Automation").bold = True
doc.add_paragraph("Streamlined data extraction and transformation processes using custom Python scripts, reducing manual processing time by approximately 40%.", style='List Bullet')

# --- Achievements & Certifications ---
doc.add_heading("Achievements & Certifications", level=1)
doc.add_paragraph("Professional Certification: Statistical Data Analysis using SPSS & STATA", style='List Bullet')
doc.add_paragraph("Innovation Contributor: Meru University Innovation Club", style='List Bullet')

# --- References ---
doc.add_heading("References", level=1)
doc.add_paragraph("Mrs. Christine Gacheri | H.O.D Mathematics\nEmail: cmutuura@must.ac.ke\nPhone: 0723674987")

# Save
doc.save(path)
print(f"Experian-Targeted CV saved to: {path}")