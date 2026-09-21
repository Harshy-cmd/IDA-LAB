# 📊 Introduction to Data Analytics (IDA) Lab

> **Course:** Introduction to Data Analytics (IDA) Lab  
> **Curriculum:** 3rd Year / 5th Semester — Computer Science & Engineering  
> **Academic Year:** 2026  
> **Repository:** [https://github.com/Harshy-cmd/IDA-LAB.git](https://github.com/Harshy-cmd/IDA-LAB.git)

---

## 🎯 Repository Overview & Lab Roadmap

This repository contains the complete laboratory curriculum, investigative data missions, code implementations, interactive dashboards, and academic lab reports for the **Introduction to Data Analytics (IDA)** laboratory course.

The curriculum is structured across **5 Core Activities**:

| Activity | Name / Topic | Dataset / Focus | Status | Key Deliverable |
| :--- | :--- | :--- | :---: | :--- |
| **Activity 1** | **Data Detective Mission** | UCI Online Retail (541k rows) | ✅ **Completed** | Full EDA Report (`.docx`, `.html`, `.pdf`, `.py`, `.ipynb`) |
| **Activity 2** | *Data Preprocessing & Feature Engineering* | Ingestion, Cleaning & Transformation | ⏳ *Upcoming* | Pipeline Scripts & Transformation Manifesto |
| **Activity 3** | *Supervised Predictive Modeling* | Classification & Regression Pipelines | ⏳ *Upcoming* | ML Model Benchmark & Validation Suite |
| **Activity 4** | *Customer Segmentation & Clustering* | RFM Analysis & Unsupervised Learning | ⏳ *Upcoming* | Cluster Profiles & Segment Insights |
| **Activity 5** | *Advanced Analytics & Capstone Project* | Production Dashboard & Capstone Dossier | ⏳ *Upcoming* | Capstone Report & Interactive UI |

---

## 📁 Repository Directory Structure

```text
IDA-LAB/
├── README.md                                    # Master repository documentation & viewing guide
├── .gitignore                                   # Git exclusions for temporary files & caches
├── Acitivity-1/                                 # Activity 1: Data Detective Mission (UCI Online Retail)
│   ├── IDA_ACTIVITY_1_Data_Detective_Report.docx # ⭐ PRIMARY ACADEMIC SUBMISSION WORD DOCUMENT
│   ├── Data_Detective_EDA_Report.html          # 🌐 Interactive Executive HTML Dashboard (Dark/Light)
│   ├── Data_Detective_EDA_Report.pdf           # 📑 High-Fidelity Print-Ready PDF Report
│   ├── Data_Detective_Mission_Complete.py      # 🐍 End-to-End Python EDA Pipeline
│   ├── Data_Detective_Mission_EDA.ipynb        # 📓 Google Colab / Jupyter Interactive Notebook
│   ├── build_ida_report_docx.py                # ⚙️ Script that compiles the 22-section Word report
│   ├── create_pdf_report.py                    # ⚙️ Headless Chrome HTML-to-PDF renderer
│   ├── eda_summary_results.csv                 # 📊 Extracted Numerical Metrics & Key KPIs
│   ├── Online Retail-2.xlsx                    # 📦 Raw UCI Transactional Dataset (541k records)
│   ├── extracted_images/                       # 🖼️ Formatted charts embedded in Word document
│   │   ├── page_19_X44.png                     # Figure 1: Monthly Revenue Trend
│   │   ├── page_20_X48.png                     # Figure 2: Top 10 Products by Revenue
│   │   ├── page_20_X49.png                     # Figure 3: Top 10 Countries by Revenue
│   │   └── page_21_X52.png                     # Figure 4: Quantity vs Unit Price Scatter
│   └── viz1_revenue_by_country.png ... viz8    # 📈 Ultra-High Resolution Matplotlib/Seaborn Charts
├── Activity-2/                                 # Activity 2 workspace (Upcoming)
│   └── .gitkeep
├── Activity-3/                                 # Activity 3 workspace (Upcoming)
│   └── .gitkeep
├── Activity-4/                                 # Activity 4 workspace (Upcoming)
│   └── .gitkeep
└── Activity-5/                                 # Activity 5 workspace (Upcoming)
    └── .gitkeep
```

---

## 🧭 Document Viewing Guide — Which Document to Watch for What?

### 📌 Activity 1: Data Detective Mission (UCI Online Retail)

When reviewing or grading Activity 1, refer to the following documents according to your evaluation needs:

| Purpose / Use Case | File to Open | Description |
| :--- | :--- | :--- |
| **Official Academic Lab Submission** | **[`Acitivity-1/IDA_ACTIVITY_1_Data_Detective_Report.docx`](./Acitivity-1/IDA_ACTIVITY_1_Data_Detective_Report.docx)** | **The official Word Document format for IDA Activity 1.** Formatted with academic styling, 22 numbered sections, institutional header, clean zebra-striped tables, code blocks in shaded callouts, embedded figures with captions, key findings, and highlighted recommendations. |
| **Interactive Executive Web Review** | **[`Acitivity-1/Data_Detective_EDA_Report.html`](./Acitivity-1/Data_Detective_EDA_Report.html)** | **Modern browser-based interactive dashboard.** Features a sticky top navigation bar with quick jump anchors, Light/Dark theme toggle, animated KPI stat cards, interactive hover tables, and a one-click print-to-PDF button. Double-click to open in Chrome, Edge, or Firefox. |
| **Formal Print / Archival PDF** | **[`Acitivity-1/Data_Detective_EDA_Report.pdf`](./Acitivity-1/Data_Detective_EDA_Report.pdf)** | **Executive print dossier.** 12-page compiled monograph rendered with headless Chrome, containing base64 embedded charts and A4 page breaks. |
| **Automated Code Execution** | **[`Acitivity-1/Data_Detective_Mission_Complete.py`](./Acitivity-1/Data_Detective_Mission_Complete.py)** | **Self-contained Python script.** Run this in any terminal or IDE to execute the entire data cleaning, statistical modeling, and plot generation pipeline in one shot. |
| **Interactive Colab / Jupyter Exploration** | **[`Acitivity-1/Data_Detective_Mission_EDA.ipynb`](./Acitivity-1/Data_Detective_Mission_EDA.ipynb)** | **Jupyter Notebook.** Ideal for inspecting cell-by-cell execution outputs, intermediate dataframes, and visualizations inside Google Colab or VS Code. |
| **Quantitative Summary Metrics** | **[`Acitivity-1/eda_summary_results.csv`](./Acitivity-1/eda_summary_results.csv)** | Raw tabular export of numerical findings (Revenue, Transaction Averages, Null counts, Country breakdown). |

---

## 🔎 Summary of Findings for Activity 1

- **Gross Clean Revenue:** **£10,666,684.54** across **524,878 clean transactions** (97.8% clean retention rate).
- **Geographic Dominance:** The **United Kingdom** accounts for **84.6%** (£9.03M) of total revenue, followed by the Netherlands, EIRE, Germany, and France.
- **Seasonality Surge:** Sales peak dramatically in **Q4**, rising from £523k in February to **£1.51M in November** (2.9× monthly trough).
- **Pareto Concentration:** The top **20% of customer accounts** generate **74.6%** of gross sales volume.
- **Data Quality Audit:**
  - **135,080 missing CustomerIDs (24.93%)** identified as guest checkouts.
  - **10,624 negative quantity entries** and **9,288 'C'-prefixed cancellation records** filtered from standard sales analysis.
  - **5,268 duplicate transactions** pruned.

---

## 🛠️ Environment Setup & Installation

To run the analysis scripts or regenerate the reports locally:

### 1. Clone the Repository
```bash
git clone https://github.com/Harshy-cmd/IDA-LAB.git
cd IDA-LAB
```

### 2. Install Required Dependencies
```bash
pip install pandas numpy matplotlib seaborn openpyxl python-docx pypdf
```

### 3. Generate Reports
- **Generate Word Document (`.docx`):**
  ```bash
  cd Acitivity-1
  python build_ida_report_docx.py
  ```
- **Recompile HTML to PDF:**
  ```bash
  cd Acitivity-1
  python create_pdf_report.py
  ```
- **Execute Full Analysis Pipeline:**
  ```bash
  cd Acitivity-1
  python Data_Detective_Mission_Complete.py
  ```

---

## 📅 Roadmap for Future Activities (2 to 5)

As upcoming lab sessions are conducted, each activity folder will be populated following the same uniform standard:
1. `Activity-X/` dedicated workspace.
2. Executable analysis scripts (`.py` and `.ipynb`).
3. Formatted academic lab reports (`.docx` and `.pdf`).
4. Output figures and metrics summaries.

---
*Maintained for IDE LAB (5th Semester / 3rd Year) • Computer Science & Engineering*
