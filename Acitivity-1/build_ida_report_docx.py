"""
Full generator script for the IDA Activity 1 Word Document.
Creates IDA_ACTIVITY_1_Data_Detective_Report.docx with high-fidelity formatting.
"""

import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DOCX = os.path.join(CURRENT_DIR, "IDA_ACTIVITY_1_Data_Detective_Report.docx")
IMG_DIR = os.path.join(CURRENT_DIR, "extracted_images")
WORKSPACE_DIR = CURRENT_DIR

COLOR_PRIMARY = RGBColor(15, 41, 74)       # Deep Academic Navy #0F294A
COLOR_SECONDARY = RGBColor(30, 64, 175)    # Royal Blue #1E40AF
COLOR_TEXT = RGBColor(31, 41, 55)          # Charcoal Slate #1F2937
COLOR_MUTED = RGBColor(100, 116, 139)      # Slate Gray #64748B

def create_report():
    doc = docx.Document()
    
    # Page setup - Standard Letter / A4 with 1 inch margins
    sections = doc.sections
    for s in sections:
        s.top_margin = Inches(1.0)
        s.bottom_margin = Inches(1.0)
        s.left_margin = Inches(1.0)
        s.right_margin = Inches(1.0)

    # Helper functions
    def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = parse_xml(
            f'<w:tcMar {nsdecls("w")}>'
            f'<w:top w:w="{top}" w:type="dxa"/>'
            f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
            f'<w:left w:w="{left}" w:type="dxa"/>'
            f'<w:right w:w="{right}" w:type="dxa"/>'
            f'</w:tcMar>'
        )
        tcPr.append(tcMar)

    def set_cell_shading(cell, color_hex):
        tcPr = cell._tc.get_or_add_tcPr()
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
        tcPr.append(shd)

    def set_table_borders(table, color="CBD5E1", sz="4"):
        tblPr = table._tbl.tblPr
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'<w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:insideH w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:insideV w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:left w:val="none"/>'
            f'<w:right w:val="none"/>'
            f'</w:tblBorders>'
        )
        tblPr.append(borders)

    def add_title_header():
        # Institutional / Lab Header
        p_pre = doc.add_paragraph()
        p_pre.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_pre.paragraph_format.space_before = Pt(0)
        p_pre.paragraph_format.space_after = Pt(2)
        r_pre = p_pre.add_run("IDE LAB — 3RD YEAR / 5TH SEMESTER  |  IDA ACTIVITY 1")
        r_pre.font.name = 'Calibri'
        r_pre.font.size = Pt(9.5)
        r_pre.font.bold = True
        r_pre.font.color.rgb = COLOR_MUTED
        
        # Main Title
        p_title = doc.add_paragraph()
        p_title.paragraph_format.space_before = Pt(4)
        p_title.paragraph_format.space_after = Pt(2)
        r_title = p_title.add_run("DATA DETECTIVE MISSION")
        r_title.font.name = 'Calibri'
        r_title.font.size = Pt(22)
        r_title.font.bold = True
        r_title.font.color.rgb = COLOR_PRIMARY
        
        # Subtitle
        p_sub = doc.add_paragraph()
        p_sub.paragraph_format.space_before = Pt(0)
        p_sub.paragraph_format.space_after = Pt(8)
        r_sub = p_sub.add_run("Exploratory Data Analysis Report — UCI Online Retail Dataset")
        r_sub.font.name = 'Calibri'
        r_sub.font.size = Pt(13)
        r_sub.font.color.rgb = COLOR_SECONDARY
        
        # Divider Line
        add_divider()

    def add_divider():
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(12)
        pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="8" w:space="1" w:color="3B82F6"/></w:pBdr>')
        p._p.get_or_add_pPr().append(pBdr)

    def add_h1(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(16)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = 'Calibri'
        r.font.size = Pt(13.5)
        r.font.bold = True
        r.font.color.rgb = COLOR_PRIMARY
        return p

    def add_h2(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(11)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = 'Calibri'
        r.font.size = Pt(11.5)
        r.font.bold = True
        r.font.color.rgb = COLOR_SECONDARY
        return p

    def add_h3(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(7)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = 'Calibri'
        r.font.size = Pt(10.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(51, 65, 85)
        return p

    def add_p(text, bold_prefix=""):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(5)
        p.paragraph_format.line_spacing = 1.15
        if bold_prefix:
            rb = p.add_run(bold_prefix)
            rb.font.name = 'Calibri'
            rb.font.size = Pt(10.5)
            rb.font.bold = True
            rb.font.color.rgb = COLOR_PRIMARY
        r = p.add_run(text)
        r.font.name = 'Calibri'
        r.font.size = Pt(10.5)
        r.font.color.rgb = COLOR_TEXT
        return p

    def add_bullet(text, bold_prefix=""):
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2.5)
        p.paragraph_format.line_spacing = 1.15
        if bold_prefix:
            rb = p.add_run(bold_prefix)
            rb.font.name = 'Calibri'
            rb.font.size = Pt(10.5)
            rb.font.bold = True
            rb.font.color.rgb = COLOR_PRIMARY
        r = p.add_run(text)
        r.font.name = 'Calibri'
        r.font.size = Pt(10.5)
        r.font.color.rgb = COLOR_TEXT
        return p

    def add_numbered(text, bold_prefix=""):
        p = doc.add_paragraph(style='List Number')
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2.5)
        p.paragraph_format.line_spacing = 1.15
        if bold_prefix:
            rb = p.add_run(bold_prefix)
            rb.font.name = 'Calibri'
            rb.font.size = Pt(10.5)
            rb.font.bold = True
            rb.font.color.rgb = COLOR_PRIMARY
        r = p.add_run(text)
        r.font.name = 'Calibri'
        r.font.size = Pt(10.5)
        r.font.color.rgb = COLOR_TEXT
        return p

    def add_callout(title, text):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.cell(0, 0)
        cell.width = Inches(6.5)
        
        tcPr = cell._tc.get_or_add_tcPr()
        borders = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'<w:left w:val="single" w:sz="32" w:space="0" w:color="2563EB"/>'
            f'<w:top w:val="none"/>'
            f'<w:bottom w:val="none"/>'
            f'<w:right w:val="none"/>'
            f'</w:tcBorders>'
        )
        tcPr.append(borders)
        set_cell_shading(cell, "EFF6FF")
        set_cell_margins(cell, top=140, bottom=140, left=180, right=180)
        
        p1 = cell.paragraphs[0]
        p1.paragraph_format.space_before = Pt(2)
        p1.paragraph_format.space_after = Pt(4)
        r1 = p1.add_run(title)
        r1.font.name = 'Calibri'
        r1.font.size = Pt(11)
        r1.font.bold = True
        r1.font.color.rgb = COLOR_PRIMARY
        
        p2 = cell.add_paragraph()
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(2)
        p2.paragraph_format.line_spacing = 1.15
        r2 = p2.add_run(text)
        r2.font.name = 'Calibri'
        r2.font.size = Pt(10.5)
        r2.font.color.rgb = COLOR_TEXT
        
        doc.add_paragraph().paragraph_format.space_after = Pt(4)

    def add_code(code_str):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.cell(0, 0)
        cell.width = Inches(6.5)
        
        tcPr = cell._tc.get_or_add_tcPr()
        borders = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'<w:left w:val="single" w:sz="20" w:space="0" w:color="3B82F6"/>'
            f'<w:top w:val="single" w:sz="4" w:space="0" w:color="E2E8F0"/>'
            f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="E2E8F0"/>'
            f'<w:right w:val="single" w:sz="4" w:space="0" w:color="E2E8F0"/>'
            f'</w:tcBorders>'
        )
        tcPr.append(borders)
        set_cell_shading(cell, "F8FAFC")
        set_cell_margins(cell, top=100, bottom=100, left=140, right=140)
        
        lines = code_str.strip().split('\n')
        for i, line in enumerate(lines):
            p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.1
            r = p.add_run(line)
            r.font.name = 'Consolas'
            r.font.size = Pt(9.5)
            r.font.color.rgb = RGBColor(15, 23, 42)
            
        doc.add_paragraph().paragraph_format.space_after = Pt(4)

    def add_table(headers, data, col_widths=None, alignments=None, header_bg="1E3A8A"):
        rows_cnt = len(data) + 1
        cols_cnt = len(headers)
        tbl = doc.add_table(rows=rows_cnt, cols=cols_cnt)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(tbl, color="CBD5E1", sz="4")
        
        # Repeat header
        header_tr = tbl.rows[0]._tr.get_or_add_trPr()
        header_tr.append(OxmlElement('w:tblHeader'))
        
        # Format Header
        for j, h in enumerate(headers):
            cell = tbl.cell(0, j)
            set_cell_shading(cell, header_bg)
            set_cell_margins(cell, top=100, bottom=100, left=120, right=120)
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            if alignments and j < len(alignments):
                p.alignment = alignments[j]
            r = p.add_run(h)
            r.font.name = 'Calibri'
            r.font.size = Pt(9.5)
            r.font.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)
            
        # Format Rows
        for i, row_data in enumerate(data):
            row_idx = i + 1
            trPr = tbl.rows[row_idx]._tr.get_or_add_trPr()
            trPr.append(OxmlElement('w:cantSplit'))
            
            bg = "F8FAFC" if i % 2 == 1 else "FFFFFF"
            for j, val in enumerate(row_data):
                cell = tbl.cell(row_idx, j)
                if bg != "FFFFFF":
                    set_cell_shading(cell, bg)
                set_cell_margins(cell, top=70, bottom=70, left=100, right=100)
                p = cell.paragraphs[0]
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = 1.15
                if alignments and j < len(alignments):
                    p.alignment = alignments[j]
                r = p.add_run(str(val))
                r.font.name = 'Calibri'
                r.font.size = Pt(9)
                r.font.color.rgb = COLOR_TEXT
                
        if col_widths:
            for row in tbl.rows:
                for j, w in enumerate(col_widths):
                    if j < len(row.cells):
                        row.cells[j].width = w
                        
        doc.add_paragraph().paragraph_format.space_after = Pt(6)
        return tbl

    def add_image(img_name, caption, width_in=5.8):
        img_path = os.path.join(IMG_DIR, img_name)
        if not os.path.exists(img_path):
            img_path = os.path.join(WORKSPACE_DIR, img_name)
        if os.path.exists(img_path):
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(8)
            p_img.paragraph_format.space_after = Pt(3)
            r_img = p_img.add_run()
            r_img.add_picture(img_path, width=Inches(width_in))
            
            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cap.paragraph_format.space_before = Pt(0)
            p_cap.paragraph_format.space_after = Pt(8)
            r_cap = p_cap.add_run(caption)
            r_cap.font.name = 'Calibri'
            r_cap.font.size = Pt(9)
            r_cap.font.italic = True
            r_cap.font.color.rgb = COLOR_MUTED

    # =========================================================================
    # BUILD DOCUMENT CONTENT (SECTIONS 1 TO 22)
    # =========================================================================

    # Title & Header
    add_title_header()

    # 1. Activity Information
    add_h1("1. Activity Information")
    info_headers = ["Item", "Details"]
    info_data = [
        ["Activity Type", "Team-based Investigation / EDA Lab"],
        ["Team Size", "3–4 students"],
        ["Duration", "2 Hours"],
        ["Tools Used", "Python, Google Colab, Pandas, NumPy, Matplotlib, Seaborn"],
        ["Dataset", "UCI Online Retail"],
        ["Dataset Source", "UCI Machine Learning Repository"]
    ]
    add_table(info_headers, info_data, col_widths=[Inches(2.2), Inches(4.3)])

    # 2. Introduction
    add_h1("2. Introduction")
    add_p(
        "The Data Detective Mission was an Exploratory Data Analysis (EDA) activity based on the "
        "UCI Online Retail dataset. The dataset contains real transaction records from an online retail "
        "business based in the UK."
    )
    add_p(
        "The main aim of this activity was to understand the data before making any conclusions from it. "
        "We looked at the structure of the dataset, checked for missing and duplicate values, identified "
        "unusual transactions, and then used different visualizations to find useful patterns."
    )
    add_p(
        "As junior data analysts, our focus was not just on finding numbers but also on understanding "
        "what those numbers could tell us about the business."
    )

    # 3. Real-World Scenario
    add_h1("3. Real-World Scenario")
    add_p(
        "For this activity, we were given historical sales data from an online retailer. The company wanted "
        "us to investigate the data and answer questions such as:"
    )
    add_bullet("Which products generate the most revenue?")
    add_bullet("How does revenue change from month to month?")
    add_bullet("Which countries contribute the most to sales?")
    add_bullet("Are there missing or inconsistent records?")
    add_bullet("Are there unusual transactions or returns?")
    add_bullet("What useful business decisions can be made from the data?")
    add_p(
        "Our task was to inspect, clean, analyze, and visualize the data and then use our findings to "
        "give a practical business recommendation."
    )

    # 4. Learning Objectives
    add_h1("4. Learning Objectives")
    add_p("Through this activity, we aimed to:")
    add_numbered("Understand the structure and important variables in a real-world dataset.")
    add_numbered("Find missing values, duplicate records, and unusual observations.")
    add_numbered("Use descriptive statistics to understand the data better.")
    add_numbered("Create suitable visualizations to identify patterns.")
    add_numbered("Convert the results of our analysis into meaningful business insights.")

    # 5. Dataset Description
    add_h1("5. Dataset Description")
    add_p(
        "The UCI Online Retail dataset contains 541,909 transaction records. "
        "It includes information about invoices, products, quantities, prices, customers, dates, and countries."
    )
    ds_headers = ["Variable", "Description"]
    ds_data = [
        ["InvoiceNo", "Unique invoice or transaction number (6 digits; prefixed with 'C' if cancelled)"],
        ["StockCode", "Unique product identification code"],
        ["Description", "Name or description of the product"],
        ["Quantity", "Number of items purchased per transaction"],
        ["InvoiceDate", "Date and time of the transaction"],
        ["UnitPrice", "Price of one item in GBP (£)"],
        ["CustomerID", "Unique customer identification number"],
        ["Country", "Country of the customer"]
    ]
    add_table(ds_headers, ds_data, col_widths=[Inches(1.8), Inches(4.7)])
    add_p("During preprocessing, we also created a new variable called Revenue:")
    add_callout("Feature Engineering Formula", "Revenue = Quantity × UnitPrice\n\nThis helped us understand how much revenue was generated by individual transactions and products.")

    # 6. Data Loading and Initial Inspection
    add_h1("6. Data Loading and Initial Inspection")
    add_p("We loaded the dataset into Google Colab using Python and Pandas.")
    add_code(
"""import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

df = pd.read_excel('/content/Online Retail.xlsx')
print("Dataset Shape:", df.shape)
display(df.head())"""
    )
    add_h2("Initial Observations")
    add_p("After loading the dataset, we found that it contained:")
    add_bullet("541,909 rows")
    add_bullet("8 original columns")
    add_bullet("Transaction-level information")
    add_bullet("Numerical, categorical, and date/time variables")
    add_p("We also looked at the first few rows to get a basic idea of how the transactions were recorded.")

    # 7. Data Dictionary
    add_h1("7. Data Dictionary")
    add_p(
        "Before starting the analysis, we prepared a simple data dictionary so that we could clearly "
        "understand what each column represented."
    )
    dict_headers = ["Variable", "Data Type", "Purpose"]
    dict_data = [
        ["InvoiceNo", "Categorical", "Identifies an invoice"],
        ["StockCode", "Categorical", "Identifies a product"],
        ["Description", "Categorical", "Gives the product name"],
        ["Quantity", "Numerical", "Shows the number of products purchased"],
        ["InvoiceDate", "DateTime", "Shows when the transaction took place"],
        ["UnitPrice", "Numerical", "Shows the price of one product"],
        ["CustomerID", "Categorical", "Identifies the customer"],
        ["Country", "Categorical", "Shows the customer's country"],
        ["Revenue", "Numerical", "Calculated transaction value (Quantity × UnitPrice)"]
    ]
    add_table(dict_headers, dict_data, col_widths=[Inches(1.6), Inches(1.5), Inches(3.4)])

    # 8. Data Quality Investigation
    add_h1("8. Data Quality Investigation")
    add_p("Before analyzing the sales patterns, we first checked whether the dataset itself had any problems. We specifically looked for:")
    add_bullet("Missing values")
    add_bullet("Duplicate records")
    add_bullet("Incorrect data types")
    add_bullet("Negative quantities")
    add_bullet("Invalid prices")
    add_bullet("Cancelled transactions")
    add_bullet("Extremely large or unusual values")
    add_p("We used the following commands:")
    add_code(
"""print("DATA TYPES")
display(df.dtypes)

print("\\nMISSING VALUES")
display(df.isnull().sum())

print("\\nDUPLICATE ROWS:", df.duplicated().sum())

print("\\nSTATISTICAL SUMMARY")
display(df.describe(include="all").T)"""
    )
    add_p("This step was important because analyzing unclean data could lead to incorrect conclusions.")
    
    add_h2("8.1 Missing Values")
    add_p("We found missing values mainly in:")
    add_bullet("CustomerID (135,080 missing values, accounting for 24.93% of the dataset)")
    add_bullet("Description (1,454 missing values, accounting for 0.27% of the dataset)")
    add_p(
        "The missing CustomerID values mean that some transactions cannot be linked to a particular customer."
    )
    add_h3("Impact")
    add_p(
        "Because of this, customer-level analysis needs to be done carefully. We cannot assume that every "
        "transaction belongs to a known customer."
    )

    # 9. Duplicate Records
    add_h1("9. Duplicate Records")
    add_p("We checked for duplicate rows using:")
    add_code('df.duplicated().sum()')
    add_p("Duplicate records can affect the results by making it appear that:")
    add_bullet("More products were sold than actually were")
    add_bullet("More revenue was generated")
    add_bullet("More transactions took place")
    add_p(
        "They can also affect statistical calculations. Therefore, 5,268 duplicate records were removed during the cleaning process."
    )

    # 10. Unusual Transactions
    add_h1("10. Unusual Transactions")
    add_p(
        "While exploring the dataset, we found transactions with negative quantities. In a retail dataset, "
        "a negative quantity usually indicates that products were returned or that a transaction was "
        "cancelled rather than representing a normal purchase."
    )
    add_p(
        "We also noticed that invoice numbers beginning with C represented cancellation transactions. "
        "Therefore, these transactions were treated separately from normal sales during preprocessing."
    )

    # 11. Data Preprocessing
    add_h1("11. Data Preprocessing")
    add_p("Before performing the main analysis, we cleaned the dataset. The following steps were carried out:")
    add_numbered("Converted InvoiceDate into the correct DateTime format.")
    add_numbered("Removed duplicate records.")
    add_numbered("Identified cancelled invoices.")
    add_numbered("Removed transactions with non-positive quantities from the normal-sales analysis.")
    add_numbered("Removed transactions with non-positive unit prices.")
    add_numbered("Created the Revenue column.")
    add_numbered("Created additional time-related columns such as month and day.")
    
    add_h2("Preprocessing Code")
    add_code(
"""clean = df.copy()
clean["InvoiceDate"] = pd.to_datetime(clean["InvoiceDate"])
clean = clean.drop_duplicates()
clean = clean[~clean["InvoiceNo"].astype(str).str.startswith("C")]
clean = clean[(clean["Quantity"] > 0) & (clean["UnitPrice"] > 0)]
clean["Revenue"] = clean["Quantity"] * clean["UnitPrice"]
clean["Month"] = clean["InvoiceDate"].dt.to_period("M").astype(str)
clean["Day"] = clean["InvoiceDate"].dt.day_name()"""
    )
    add_p("After this step, the dataset was in a much better condition for analyzing normal sales (524,878 clean records remaining).")

    # 12. Descriptive Statistics
    add_h1("12. Descriptive Statistics")
    add_p("We used descriptive statistics to get a better understanding of the numerical data. The main measures we looked at were:")
    add_bullet("Count")
    add_bullet("Mean")
    add_bullet("Standard deviation")
    add_bullet("Minimum")
    add_bullet("Maximum")
    add_bullet("Quartiles (25%, 50% Median, 75%)")
    add_p(
        "The results showed that quantities and prices were not evenly distributed. Some transactions had much "
        "larger quantities or values than most other transactions."
    )
    add_p(
        "These unusual observations can be considered outliers, so they need to be kept in mind when interpreting averages."
    )

    # 13. Visualization 1 — Monthly Revenue Trend
    add_h1("13. Visualization 1 — Monthly Revenue Trend")
    add_h3("Objective")
    add_p("The purpose of this visualization was to see how the company's revenue changed over time.")
    add_code(
"""monthly_sales = clean.groupby("Month")["Revenue"].sum()

plt.figure(figsize=(14,5))
monthly_sales.plot(marker="o")
plt.title("Monthly Revenue Trend")
plt.xlabel("Month")
plt.ylabel("Revenue (£)")
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()"""
    )
    add_image("page_19_X44.png", "Figure 1: Monthly Revenue Trend across December 2010 to December 2011", width_in=5.8)
    add_h3("Finding")
    add_p("The graph shows that revenue was not the same every month. Some months had noticeably higher sales activity than others.")
    add_h3("Business Insight")
    add_p("The company can use these historical patterns to plan:")
    add_bullet("Inventory")
    add_bullet("Promotions")
    add_bullet("Staffing")
    add_bullet("Product availability")
    add_p("This can help the business prepare better for periods of higher demand.")

    # 14. Visualization 2 — Top 10 Products by Revenue
    add_h1("14. Visualization 2 — Top 10 Products by Revenue")
    add_h3("Objective")
    add_p("The purpose of this analysis was to find the products that contributed the most revenue.")
    add_code(
"""top_products = (
    clean.groupby("Description")["Revenue"]
    .sum()
    .sort_values(ascending=False)
    .head(10)
    .sort_values()
)

plt.figure(figsize=(10,6))
top_products.plot(kind="barh")
plt.title("Top 10 Products by Revenue")
plt.xlabel("Revenue (£)")
plt.ylabel("Product")
plt.tight_layout()
plt.show()"""
    )
    add_image("page_20_X48.png", "Figure 2: Top 10 Products by Revenue Contribution", width_in=5.6)
    add_h3("Finding")
    add_p("The analysis showed that revenue was concentrated among a smaller group of products.")
    add_h3("Business Insight")
    add_p("These high-revenue products deserve more attention when it comes to:")
    add_bullet("Maintaining stock")
    add_bullet("Reordering products")
    add_bullet("Promotions")
    add_bullet("Forecasting future demand")
    add_p("Keeping popular products available can help prevent missed sales opportunities.")

    # 15. Visualization 3 — Top 10 Countries by Revenue
    add_h1("15. Visualization 3 — Top 10 Countries by Revenue")
    add_h3("Objective")
    add_p("We used this visualization to understand which countries contributed the most to the company's revenue.")
    add_code(
"""country_sales = (
    clean.groupby("Country")["Revenue"]
    .sum()
    .sort_values(ascending=False)
    .head(10)
    .sort_values()
)

plt.figure(figsize=(10,6))
country_sales.plot(kind="barh")
plt.title("Top 10 Countries by Revenue")
plt.xlabel("Revenue (£)")
plt.ylabel("Country")
plt.tight_layout()
plt.show()"""
    )
    add_image("page_20_X49.png", "Figure 3: Top 10 Countries by Total Revenue Generation", width_in=5.6)
    add_h3("Finding")
    add_p("The United Kingdom is the dominant market in the dataset, while several other countries also contribute to international sales.")
    add_h3("Business Insight")
    add_p(
        "Looking at sales country-wise allows the company to compare its domestic and international markets "
        "and identify areas that may need more attention."
    )

    # 16. Visualization 4 — Quantity vs Unit Price
    add_h1("16. Visualization 4 — Quantity vs Unit Price")
    add_h3("Objective")
    add_p("The purpose of this visualization was to see how quantity and unit price are distributed and to identify unusual transactions.")
    add_code(
"""sample = clean.sample(
    min(10000, len(clean)),
    random_state=42
)

plt.figure(figsize=(9,6))
sns.scatterplot(
    data=sample,
    x="UnitPrice",
    y="Quantity",
    alpha=0.4
)
plt.title("Quantity vs Unit Price")
plt.xlabel("Unit Price (£)")
plt.ylabel("Quantity")
plt.tight_layout()
plt.show()"""
    )
    add_image("page_21_X52.png", "Figure 4: Scatter Plot of Quantity vs Unit Price", width_in=5.4)
    add_h3("Finding")
    add_p(
        "Most transactions are grouped within relatively smaller quantity and price ranges. However, "
        "there are some transactions with unusually high quantities."
    )
    add_h3("Business Insight")
    add_p("These unusual transactions could be worth investigating because they might represent:")
    add_bullet("Bulk purchases")
    add_bullet("Special orders")
    add_bullet("Data-entry errors")
    add_bullet("Exceptionally high customer demand")

    # 17. Pattern Identification
    add_h1("17. Pattern Identification")
    add_p(
        "After completing the main visualizations, we carried out some additional analysis to understand "
        "the sales patterns in more detail."
    )
    
    add_h2("Top Products by Quantity")
    add_code(
"""top_quantity = (
    clean.groupby("Description")["Quantity"]
    .sum()
    .sort_values(ascending=False)
    .head(10)
)
display(top_quantity)"""
    )
    add_p("This helped us identify the products that were sold in the highest quantities.")

    add_h2("Highest-Value Customers")
    add_code(
"""top_customers = (
    clean.groupby("CustomerID")["Revenue"]
    .sum()
    .sort_values(ascending=False)
    .head(10)
)
display(top_customers)"""
    )
    add_p(
        "This analysis helped identify customers who contributed significantly to the company's revenue. "
        "However, because some CustomerID values are missing, the results should be interpreted with that limitation in mind."
    )

    add_h2("Average Transaction Value")
    add_code(
"""invoice_value = clean.groupby("InvoiceNo")["Revenue"].sum()
print("Average transaction value:", round(invoice_value.mean(), 2))
print("Median transaction value:", round(invoice_value.median(), 2))"""
    )
    add_p(
        "We compared the average and median transaction values to understand whether unusually large "
        "transactions were having a noticeable effect on the average."
    )

    # 18. Key Data-Quality Findings
    add_h1("18. Key Data-Quality Findings")
    add_p("The audit of the raw dataset revealed several critical findings summarized below:")
    dq_headers = ["Issue", "What We Found", "Effect on Analysis"]
    dq_data = [
        ["Missing CustomerID", "Many transactions do not have customer information (135,080 records / 24.93%)", "Makes customer-level analysis difficult"],
        ["Missing Description", "Some products have no description (1,454 records / 0.27%)", "Makes product analysis less clear"],
        ["Duplicate records", "Duplicate rows are present (5,268 records)", "Can increase calculated sales and transaction counts"],
        ["Negative Quantity", "Returns and cancellations are present (10,624 records)", "Should not be treated as normal sales"],
        ["Cancelled invoices", "Some invoice numbers begin with C (9,288 records)", "Need to be handled separately"],
        ["Extreme values", "Some quantities and prices are unusually large", "Can affect averages and statistics"]
    ]
    add_table(dq_headers, dq_data, col_widths=[Inches(1.6), Inches(2.5), Inches(2.4)])

    # 19. Key Business Findings
    add_h1("19. Key Business Findings")
    add_p("After completing the analysis, we identified several important patterns:")
    add_numbered("Sales change over time: Sales activity is not the same throughout the year. Some months show stronger activity than others, especially leading into the Q4 holiday surge.")
    add_numbered("Some products contribute much more revenue: A smaller group of products contributes a significant portion of the overall revenue (classic 80/20 Pareto distribution).")
    add_numbered("The UK is the main market: The dataset shows that the UK accounts for a large share (>85%) of total revenue.")
    add_numbered("Returns and cancellations matter: Negative quantities and cancelled invoices need to be separated from normal sales to avoid misleading results.")
    add_numbered("Customer information is incomplete: Since many transactions do not have a CustomerID, complete customer segmentation is not possible using this dataset alone.")
    add_numbered("Outliers need attention: Some transactions have unusually large quantities or values. These can influence statistical measures, particularly averages.")

    # 20. Business Recommendation
    add_h1("20. Business Recommendation")
    add_p(
        "Based on what we found during the analysis, the retailer could use a data-driven approach to "
        "inventory and sales management. The company should:"
    )
    add_bullet("Keep a close watch on high-revenue products.")
    add_bullet("Use monthly sales patterns when planning inventory.")
    add_bullet("Track returns and cancellations separately.")
    add_bullet("Investigate unusually large transactions.")
    add_bullet("Improve the collection of customer information.")
    add_bullet("Analyze international markets separately to understand their potential.")
    
    add_callout(
        "Main Recommendation",
        "The retailer should focus inventory planning on high-revenue products and periods with stronger sales, "
        "while keeping a separate track of returns, cancellations, and unusual transactions.\n\n"
        "This recommendation connects our data-quality findings with the sales patterns we identified during the analysis."
    )

    # 21. Outputs
    add_h1("21. Outputs")
    add_p("The complete execution outputs, console logs, and generated tables from the EDA process are recorded below:")
    
    add_h2("Dataset Ingestion & Shape")
    add_code(
"""Online Retail.xlsx (application/vnd.openxmlformats-officedocument.spreadsheetml.sheet) - 23715344 bytes, last modified: 9/21/2026 - 100% done
Saving Online Retail.xlsx to Online Retail.xlsx
Dataset Shape: (541909, 8)"""
    )
    
    add_h2("Initial Records (df.head())")
    head_headers = ["InvoiceNo", "StockCode", "Description", "Qty", "InvoiceDate", "UnitPrice", "CustomerID", "Country"]
    head_data = [
        ["536365", "85123A", "WHITE HANGING HEART T-LIGHT HOLDER", "6", "2010-12-01 08:26:00", "£ 2.55", "17850.0", "United Kingdom"],
        ["536365", "71053", "WHITE METAL LANTERN", "6", "2010-12-01 08:26:00", "£ 3.39", "17850.0", "United Kingdom"],
        ["536365", "84406B", "CREAM CUPID HEARTS COAT HANGER", "8", "2010-12-01 08:26:00", "£ 2.75", "17850.0", "United Kingdom"],
        ["536365", "84029G", "KNITTED UNION FLAG HOT WATER BOTTLE", "6", "2010-12-01 08:26:00", "£ 3.39", "17850.0", "United Kingdom"],
        ["536365", "84029E", "RED WOOLLY HOTTIE WHITE HEART.", "6", "2010-12-01 08:26:00", "£ 3.39", "17850.0", "United Kingdom"]
    ]
    aligns_head = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.LEFT]
    add_table(head_headers, head_data, col_widths=[Inches(0.8), Inches(0.7), Inches(2.2), Inches(0.4), Inches(1.3), Inches(0.7), Inches(0.8), Inches(0.9)], alignments=aligns_head)

    add_h2("Data Types & Missing Values")
    dtypes_headers = ["Column Name", "Data Type", "Missing Count", "Missing Percentage (%)"]
    dtypes_data = [
        ["InvoiceNo", "object", "0", "0.00%"],
        ["StockCode", "object", "0", "0.00%"],
        ["Description", "object", "1,454", "0.27%"],
        ["Quantity", "int64", "0", "0.00%"],
        ["InvoiceDate", "datetime64[ns]", "0", "0.00%"],
        ["UnitPrice", "float64", "0", "0.00%"],
        ["CustomerID", "float64", "135,080", "24.93%"],
        ["Country", "object", "0", "0.00%"]
    ]
    aligns_dt = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT]
    add_table(dtypes_headers, dtypes_data, col_widths=[Inches(1.8), Inches(1.5), Inches(1.4), Inches(1.8)], alignments=aligns_dt)

    add_h2("Data Quality Anomalies & Cleaning Summary")
    add_code(
"""DUPLICATE ROWS: 5,268
NEGATIVE QUANTITY RECORDS: 10,624
INVALID / ZERO UNIT PRICE RECORDS: 2,517
CANCELLATION INVOICES ('C'): 9,288

CLEAN DATASET SHAPE: (524,878 rows, 11 columns)"""
    )

    add_h2("Statistical Summary (df.describe(include='all').T)")
    desc_headers = ["Column", "Count", "Unique", "Top / Mean", "Freq / Std", "Min", "50% (Median)", "Max"]
    desc_data = [
        ["InvoiceNo", "541,909", "25,900", "573585", "1,114", "—", "—", "—"],
        ["StockCode", "541,909", "4,070", "85123A", "2,313", "—", "—", "—"],
        ["Description", "540,455", "4,223", "WHITE HANGING...", "2,369", "—", "—", "—"],
        ["Quantity", "541,909", "—", "9.55", "218.08", "-80,995", "3.00", "80,995"],
        ["InvoiceDate", "541,909", "—", "2011-07-04...", "—", "2010-12-01", "2011-07-19", "2011-12-09"],
        ["UnitPrice", "541,909", "—", "£ 4.61", "£ 96.76", "-£ 11,062", "£ 2.08", "£ 38,970"],
        ["CustomerID", "406,829", "—", "15287.69", "1713.60", "12346.0", "15152.0", "18287.0"],
        ["Country", "541,909", "38", "United Kingdom", "495,478", "—", "—", "—"]
    ]
    aligns_desc = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT]
    add_table(desc_headers, desc_data, col_widths=[Inches(1.1), Inches(0.8), Inches(0.7), Inches(1.3), Inches(0.8), Inches(0.8), Inches(0.8), Inches(0.8)], alignments=aligns_desc)

    add_h2("Clean Dataset Sample (clean.head())")
    clean_headers = ["InvoiceNo", "StockCode", "Description", "Qty", "UnitPrice", "CustomerID", "Revenue", "Month", "Day"]
    clean_data = [
        ["536365", "85123A", "WHITE HANGING HEART T-LIGHT HOLDER", "6", "£ 2.55", "17850.0", "£ 15.30", "2010-12", "Wednesday"],
        ["536365", "71053", "WHITE METAL LANTERN", "6", "£ 3.39", "17850.0", "£ 20.34", "2010-12", "Wednesday"],
        ["536365", "84406B", "CREAM CUPID HEARTS COAT HANGER", "8", "£ 2.75", "17850.0", "£ 22.00", "2010-12", "Wednesday"],
        ["536365", "84029G", "KNITTED UNION FLAG HOT WATER BOTTLE", "6", "£ 3.39", "17850.0", "£ 20.34", "2010-12", "Wednesday"],
        ["536365", "84029E", "RED WOOLLY HOTTIE WHITE HEART.", "6", "£ 3.39", "17850.0", "£ 20.34", "2010-12", "Wednesday"]
    ]
    aligns_clean = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT]
    add_table(clean_headers, clean_data, col_widths=[Inches(0.75), Inches(0.65), Inches(2.0), Inches(0.4), Inches(0.65), Inches(0.75), Inches(0.65), Inches(0.65), Inches(0.8)], alignments=aligns_clean)

    add_h2("Output Visualizations")
    add_p("The four core analytical visualizations produced during exploratory inspection:")
    add_image("page_19_X44.png", "Monthly Revenue Trend (£ GBP) across the operating timeline", width_in=5.8)
    add_image("page_20_X48.png", "Top 10 Products by Total Generated Sales (£ GBP)", width_in=5.6)
    add_image("page_20_X49.png", "Top 10 Country Markets by Revenue (£ GBP)", width_in=5.6)
    add_image("page_21_X52.png", "Quantity vs Unit Price Scatter Distribution", width_in=5.4)

    add_h2("Top Products by Quantity (Output)")
    qty_headers = ["Description", "Total Quantity Sold"]
    qty_data = [
        ["PAPER CRAFT , LITTLE BIRDIE", "80,995"],
        ["MEDIUM CERAMIC TOP STORAGE JAR", "78,033"],
        ["WORLD WAR 2 GLIDERS ASSTD DESIGNS", "54,951"],
        ["JUMBO BAG RED RETROSPOT", "48,371"],
        ["WHITE HANGING HEART T-LIGHT HOLDER", "37,872"],
        ["POPCORN HOLDER", "36,749"],
        ["PACK OF 72 RETROSPOT CAKE CASES", "36,396"],
        ["ASSORTED COLOUR BIRD ORNAMENT", "36,362"],
        ["RABBIT NIGHT LIGHT", "30,739"],
        ["MINI PAINT SET VINTAGE", "26,633"]
    ]
    aligns_qty = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT]
    add_table(qty_headers, qty_data, col_widths=[Inches(4.5), Inches(2.0)], alignments=aligns_qty)

    add_h2("Highest-Value Customers (Output)")
    cust_headers = ["CustomerID", "Total Revenue Generated (£)"]
    cust_data = [
        ["14646.0", "£ 280,206.02"],
        ["18102.0", "£ 259,657.30"],
        ["17450.0", "£ 194,390.79"],
        ["16446.0", "£ 168,472.50"],
        ["14911.0", "£ 143,711.17"],
        ["12415.0", "£ 124,914.53"],
        ["14156.0", "£ 117,210.08"],
        ["17511.0", "£ 91,062.38"],
        ["16029.0", "£ 80,850.84"],
        ["12346.0", "£ 77,183.60"]
    ]
    aligns_cust = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT]
    add_table(cust_headers, cust_data, col_widths=[Inches(3.2), Inches(3.3)], alignments=aligns_cust)

    add_h2("Transaction Value Metrics (Output)")
    add_code(
"""Average transaction value: £ 533.17
Median transaction value: £ 303.30"""
    )
    add_p(
        "Note: The substantial divergence between the average (£533.17) and median (£303.30) values illustrates "
        "a pronounced right-skewed distribution, driven by exceptionally high-value wholesale transactions."
    )

    # 22. Conclusion
    add_h1("22. Conclusion")
    add_p(
        "The Data Detective Mission gave us a practical experience of how raw business data can be explored "
        "and turned into useful information."
    )
    add_p(
        "We started by understanding the dataset and checking its quality. We found issues such as missing "
        "customer information, duplicate records, cancelled transactions, and unusual values. After cleaning "
        "the data, we were able to analyze the sales information more reliably."
    )
    add_p(
        "We then used statistics and visualizations to understand monthly sales, product performance, "
        "country-wise revenue, and unusual transaction patterns. These helped us see where the business "
        "was performing strongly and which areas needed closer attention."
    )
    add_p(
        "Overall, this activity showed us that data cleaning is just as important as data analysis. "
        "Once the data was properly prepared, it became much easier to identify meaningful patterns and "
        "connect them to real business decisions, such as better inventory planning, monitoring returns, "
        "and focusing on high-performing products."
    )

    # Save document
    doc.save(OUTPUT_DOCX)
    print(f"[SUCCESS] Successfully generated report at: {OUTPUT_DOCX}")
    
    # Also save a copy in Downloads for convenient access
    downloads_copy = os.path.join(os.path.expanduser("~"), "Downloads", "IDA_ACTIVITY_1_Data_Detective_Report.docx")
    try:
        doc.save(downloads_copy)
        print(f"[SUCCESS] Also saved a copy in Downloads: {downloads_copy}")
    except Exception as e:
        print(f"Could not copy to Downloads: {e}")

if __name__ == "__main__":
    create_report()
