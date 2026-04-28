"""
Generates: Contoso_COO_Agent_Setup_Guide.docx
"""
import shutil
from pathlib import Path
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

OUTPUT_PATH = Path(r"C:\Users\peiyiyap\AppData\Local\Temp\ihh-copilot-immersion\samples\COO_Fabric_Agent\Contoso_COO_Agent_Setup_Guide.docx")
COPY_PATH = Path(r"C:\Users\peiyiyap\OneDrive - Microsoft\Documents\Customers\IHH\IHH Demo and Immersion Experience\samples\COO_Fabric_Agent\Contoso_COO_Agent_Setup_Guide.docx")

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def set_cell_bg(cell, hex_color: str):
    """Set table cell background colour."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def add_header_footer(doc, header_text: str, footer_text: str):
    section = doc.sections[0]
    # Header
    header = section.header
    hp = header.paragraphs[0]
    hp.clear()
    run = hp.add_run(header_text)
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
    run.font.italic = True
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    # Footer
    footer = section.footer
    fp = footer.paragraphs[0]
    fp.clear()
    run = fp.add_run(footer_text)
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER


def heading(doc, text: str, level: int):
    p = doc.add_heading(text, level=level)
    return p


def body(doc, text: str):
    return doc.add_paragraph(text, style="Normal")


def bullet(doc, text: str, level: int = 0):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.left_indent = Inches(0.25 * (level + 1))
    p.add_run(text)
    return p


def numbered(doc, text: str, level: int = 0):
    p = doc.add_paragraph(style="List Number")
    p.paragraph_format.left_indent = Inches(0.25 * (level + 1))
    p.add_run(text)
    return p


def tip(doc, text: str):
    p = doc.add_paragraph()
    r = p.add_run("TIP:  ")
    r.bold = True
    r.italic = True
    r.font.color.rgb = RGBColor(0x00, 0x70, 0xC0)
    r2 = p.add_run(text)
    r2.italic = True
    r2.font.color.rgb = RGBColor(0x00, 0x70, 0xC0)
    return p


def warn(doc, text: str):
    p = doc.add_paragraph()
    r = p.add_run("NOTE:  ")
    r.bold = True
    r.font.color.rgb = RGBColor(0xC0, 0x50, 0x00)
    r2 = p.add_run(text)
    r2.font.color.rgb = RGBColor(0xC0, 0x50, 0x00)
    return p


def code_block(doc, code: str):
    """Simulate a code block: Courier New 9pt in a 1-cell table with grey fill."""
    table = doc.add_table(rows=1, cols=1)
    table.style = "Table Grid"
    cell = table.rows[0].cells[0]
    set_cell_bg(cell, "F2F2F2")
    cell.paragraphs[0].clear()
    for line in code.strip().splitlines():
        p = cell.add_paragraph()
        run = p.add_run(line if line else " ")
        run.font.name = "Courier New"
        run.font.size = Pt(9)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)
    # Remove the first empty paragraph that add_table inserts
    first_p = cell.paragraphs[0]
    if not first_p.text:
        fp = first_p._element
        fp.getparent().remove(fp)
    doc.add_paragraph()  # spacing after


def spacer(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)


# ---------------------------------------------------------------------------
# Build document
# ---------------------------------------------------------------------------

def build():
    doc = Document()

    # ---- Page margins ----
    section = doc.sections[0]
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)

    add_header_footer(
        doc,
        "CONFIDENTIAL — Contoso Healthcare COO Agent Setup Guide",
        "Microsoft Internal  |  Prepared by Contoso Healthcare Transformation Team  |  2026",
    )

    # =====================================================================
    # TITLE
    # =====================================================================
    title_p = doc.add_heading(
        "Building a COO Intelligence Agent:\nMicrosoft Fabric Data Agent + Copilot Studio\nStep-by-Step Guide",
        level=0,
    )
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()

    # Sub-title / audience note
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Prepared for: Peiyi  |  Closing Demo — IHH Healthcare Executive Leadership")
    run.bold = True
    run.font.size = Pt(11)

    doc.add_paragraph()

    # =====================================================================
    # OVERVIEW
    # =====================================================================
    heading(doc, "OVERVIEW", 1)

    heading(doc, "What We Are Building", 2)
    body(doc,
         "A COO-facing Copilot Studio agent that enables Contoso Healthcare's Chief Operating Officer "
         "to ask natural-language questions about hospital operations, clinical quality, financials, and "
         "workforce data — all powered by a Microsoft Fabric Data Agent connected directly to a Fabric Lakehouse.")

    spacer(doc)
    body(doc, "Example questions the COO can ask:")
    for q in [
        "Which hospital had the highest readmission rate last quarter?",
        "Show me ED wait times across all hospitals in Q1 2026.",
        "What's our total procurement spend by category this year?",
        "Which hospitals are at risk based on the KPI scorecard?",
        "How does our bed occupancy compare across ASEAN vs EMEA?",
    ]:
        bullet(doc, q)

    spacer(doc)
    heading(doc, "Solution Architecture", 2)
    body(doc,
         "The diagram below describes the end-to-end data flow:")
    for step in [
        "Excel / CSV source files  →",
        "Microsoft Fabric Lakehouse (ContosoHealthcareLH)  →",
        "SQL Analytics Endpoint (auto-generated)  →",
        "Fabric Data Agent (natural language → SQL)  →",
        "Copilot Studio Agent (conversation layer)  →",
        "COO / Executive (Teams, SharePoint, Web)",
    ]:
        bullet(doc, step)

    tip(doc, "All components run within your Microsoft 365 / Azure tenant — no data leaves your environment.")

    # =====================================================================
    # PART 1
    # =====================================================================
    heading(doc, "PART 1: PREPARE DATA IN MICROSOFT FABRIC", 1)

    # Step 1
    heading(doc, "Step 1: Upload Data to Fabric Lakehouse", 2)
    warn(doc, "Pre-requisite: Microsoft Fabric capacity (F2 or higher), or an active Fabric Trial. "
              "You need a Fabric-enabled workspace to proceed.")

    body(doc, "Follow these steps to set up your Fabric workspace and upload the source data files:")
    numbered(doc, "Navigate to app.fabric.microsoft.com and sign in with your Microsoft 365 account.")
    numbered(doc, "Click 'Workspaces' in the left navigation → 'New workspace'.")
    numbered(doc, "Name the workspace: Contoso-Healthcare-COO → click 'Apply'.")
    numbered(doc, "Inside the workspace, click 'New' → 'Lakehouse'.")
    numbered(doc, "Name the Lakehouse: ContosoHealthcareLH → click 'Create'.")
    numbered(doc, "In the Lakehouse explorer, click 'Files' in the left pane.")
    numbered(doc, "Click 'Upload' → 'Upload files' → select all 5 Excel files from the COO_Fabric_Agent folder:")

    for f in [
        "Contoso_Patient_Operations.xlsx",
        "Contoso_Workforce_HR.xlsx",
        "Contoso_Supply_Chain.xlsx",
        "Contoso_Clinical_Quality.xlsx",
        "Contoso_Executive_KPIs.xlsx",
    ]:
        bullet(doc, f, level=1)

    numbered(doc, "Wait for all uploads to complete. Verify files appear under 'Files' in the Lakehouse explorer.")
    tip(doc, "Keep all 5 files in the root of 'Files' (not in sub-folders) to match the notebook paths in Step 2.")

    # Step 2
    heading(doc, "Step 2: Create Dataflow Gen2 or Notebook to Load Tables", 2)
    body(doc,
         "You need to convert the Excel files into Fabric Lakehouse Delta tables so the SQL Analytics "
         "Endpoint — and ultimately the Data Agent — can query them. Two options are available:")

    heading(doc, "Option A — No-Code: Dataflow Gen2", 3)
    numbered(doc, "In your Fabric workspace, click 'New' → 'Dataflow Gen2'.")
    numbered(doc, "Click 'Get data' → search for 'Excel' → select the file from your Lakehouse Files.")
    numbered(doc, "Select the sheet you want to load (e.g., Admissions_Bed_Occupancy).")
    numbered(doc, "In the query settings, set the output destination to 'Lakehouse' → select ContosoHealthcareLH.")
    numbered(doc, "Set the table name (e.g., patient_operations_admissions) → click 'Save and publish'.")
    numbered(doc, "Repeat for all 15 sheets across the 5 Excel files.")
    tip(doc, "Dataflow Gen2 is recommended for non-technical users. It requires no coding and handles data type inference automatically.")

    heading(doc, "Option B — Code: PySpark Notebook (Recommended for Speed)", 3)
    body(doc, "This single notebook loads all 15 tables in one run:")
    numbered(doc, "In the workspace, click 'New' → 'Notebook'.")
    numbered(doc, "Ensure the notebook is attached to your Lakehouse: click 'Add Lakehouse' → select ContosoHealthcareLH.")
    numbered(doc, "Paste the following code into the notebook cell and click 'Run All':")

    code_block(doc, """\
import pandas as pd
from pyspark.sql import SparkSession

spark = SparkSession.builder.getOrCreate()

files = {
    "patient_operations_admissions": ("Contoso_Patient_Operations.xlsx", "Admissions_Bed_Occupancy"),
    "patient_operations_dept":       ("Contoso_Patient_Operations.xlsx", "Dept_Utilisation"),
    "patient_operations_ed":         ("Contoso_Patient_Operations.xlsx", "Emergency_Dept_KPIs"),
    "patient_operations_theatre":    ("Contoso_Patient_Operations.xlsx", "Surgical_Theatre"),
    "workforce_headcount":           ("Contoso_Workforce_HR.xlsx",       "Headcount_Monthly"),
    "workforce_staff_directory":     ("Contoso_Workforce_HR.xlsx",       "Staff_Directory"),
    "workforce_training":            ("Contoso_Workforce_HR.xlsx",       "Training_Development"),
    "supply_procurement":            ("Contoso_Supply_Chain.xlsx",       "Procurement_Orders"),
    "supply_inventory":              ("Contoso_Supply_Chain.xlsx",       "Inventory_Levels"),
    "clinical_incidents":            ("Contoso_Clinical_Quality.xlsx",   "Patient_Safety_Incidents"),
    "clinical_infection":            ("Contoso_Clinical_Quality.xlsx",   "Infection_Control"),
    "clinical_satisfaction":         ("Contoso_Clinical_Quality.xlsx",   "Patient_Satisfaction"),
    "clinical_outcomes":             ("Contoso_Clinical_Quality.xlsx",   "Clinical_Outcomes"),
    "kpi_scorecard":                 ("Contoso_Executive_KPIs.xlsx",     "Monthly_KPI_Scorecard"),
    "kpi_dashboard":                 ("Contoso_Executive_KPIs.xlsx",     "COO_Dashboard_Summary"),
}

for table_name, (file_name, sheet_name) in files.items():
    pdf = pd.read_excel(f"/lakehouse/default/Files/{file_name}", sheet_name=sheet_name)
    df = spark.createDataFrame(pdf)
    df.write.mode("overwrite").format("delta").saveAsTable(table_name)
    print(f"Loaded {table_name}: {df.count()} rows")

print("All 15 tables loaded successfully.")""")

    tip(doc, "The notebook typically completes in 2–4 minutes. Check the output cell for row counts to confirm successful loads.")

    # Step 3
    heading(doc, "Step 3: Verify Tables in SQL Analytics Endpoint", 2)
    numbered(doc, "In the Lakehouse view, click the toggle at the top-right and switch to 'SQL analytics endpoint'.")
    numbered(doc, "In the left pane under 'Tables', verify all 15 tables are listed.")
    numbered(doc, "Click 'New query' and run the following test query to confirm data integrity:")

    code_block(doc, """\
SELECT
    Hospital,
    Country,
    AVG(Bed_Occupancy_Rate_Pct) AS Avg_Occupancy
FROM patient_operations_admissions
WHERE Year = 2025
GROUP BY Hospital, Country
ORDER BY Avg_Occupancy DESC""")

    numbered(doc, "Confirm that results are returned with hospital names, countries, and occupancy rates.")
    warn(doc, "If a table is missing, re-run the notebook or Dataflow for that specific table. "
              "Column names in the SQL endpoint must match exactly what is used in the Data Agent semantic layer.")
    tip(doc, "Save any useful test queries — you can reuse them to validate Data Agent responses during testing in Step 7.")

    # =====================================================================
    # PART 2
    # =====================================================================
    heading(doc, "PART 2: CREATE THE FABRIC DATA AGENT", 1)

    # Step 4
    heading(doc, "Step 4: Create the Fabric Data Agent", 2)
    warn(doc, "Fabric Data Agent is currently in Preview. Ensure your Fabric tenant admin has enabled preview features.")
    numbered(doc, "In your Fabric workspace, click 'New' → scroll to the 'AI' section.")
    numbered(doc, "Select 'Data Agent (preview)'.")
    numbered(doc, "In the creation dialog, enter the name: Contoso Healthcare COO Agent.")
    numbered(doc, "Under 'Data source', select 'Lakehouse' → choose ContosoHealthcareLH.")
    numbered(doc, "In the table selection screen, tick all 15 tables.")
    numbered(doc, "Click 'Create'. The Data Agent editor will open.")
    tip(doc, "Fabric Data Agent uses the SQL Analytics Endpoint under the hood. All queries it generates run against your Lakehouse Delta tables — no separate database is needed.")

    # Step 5
    heading(doc, "Step 5: Configure Agent Instructions (System Prompt)", 2)
    body(doc,
         "In the Data Agent editor, locate the 'Instructions' panel on the right-hand side. "
         "Replace the default text with the following system prompt:")

    code_block(doc, """\
You are a healthcare analytics assistant for Contoso Healthcare Group. You help the COO
and senior executives understand operational performance, clinical quality, financial
results, and workforce metrics across all 24 hospitals in 11 countries.

When answering questions:
- Always specify which hospital(s) and time period you are referring to
- Provide comparisons where helpful (e.g., vs benchmark, vs prior period, vs budget)
- Highlight any hospitals that are underperforming or at risk
- Use RAG status (Red/Amber/Green) language where available
- Format numbers clearly (e.g., USD millions, percentages with 1 decimal place)
- If asked about trends, compare the most recent 3-6 months""")

    tip(doc, "Keep the system prompt concise and domain-specific. Overly long prompts can reduce response quality.")

    # Step 6
    heading(doc, "Step 6: Add Table Descriptions (Semantic Layer)", 2)
    body(doc,
         "Click each table name in the Data Agent editor and add a description. "
         "These descriptions form a semantic layer that helps the AI understand the business meaning of each table:")

    table_descs = [
        ("patient_operations_admissions",
         "Monthly hospital-level data on bed occupancy, admissions, length of stay, ICU occupancy, and surgical volumes for all 24 Contoso hospitals from Jan 2024 to Dec 2025."),
        ("patient_operations_dept",
         "Monthly department-level utilisation data including outpatient and inpatient activity by specialty and hospital."),
        ("patient_operations_ed",
         "Emergency Department KPIs including door-to-doctor time, triage category volumes, ED attendance, and re-attendance rates per hospital per month."),
        ("patient_operations_theatre",
         "Surgical theatre utilisation metrics: cases performed, theatre occupancy rate, cancellation rate, and average turnaround time."),
        ("workforce_headcount",
         "Monthly headcount, vacancy rate, attrition rate, and FTE by hospital and job family."),
        ("workforce_staff_directory",
         "Employee master data including role, department, hospital, region, and employment type."),
        ("workforce_training",
         "Training completion and compliance rates by hospital, department, and training type."),
        ("supply_procurement",
         "Procurement orders including supplier, category (Pharmaceuticals, Medical Equipment, Consumables, etc.), order value, and delivery status."),
        ("supply_inventory",
         "Monthly inventory levels by hospital and item category, including stockout incidents and days-on-hand."),
        ("clinical_incidents",
         "Patient safety incident reports by hospital, incident type, severity, and outcome."),
        ("clinical_infection",
         "Healthcare-associated infection rates including CLABSI, CAUTI, SSI, and C. diff per hospital per month."),
        ("clinical_satisfaction",
         "Patient satisfaction survey scores (overall and by dimension) per hospital, benchmarked against national targets."),
        ("clinical_outcomes",
         "Clinical outcome metrics: 30-day readmission rate, in-hospital mortality, complication rate, and average length of stay by specialty."),
        ("kpi_scorecard",
         "Monthly KPI scorecard with Financial, Operations, Quality, People and ESG scores for each hospital. "
         "Overall_Score determines Status: On Track (>=75), At Risk (62-74), Off Track (<62)."),
        ("kpi_dashboard",
         "High-level COO dashboard summary with group-wide aggregated metrics, regional averages, and month-on-month trends."),
    ]

    for tbl, desc in table_descs:
        p = doc.add_paragraph(style="List Bullet")
        r1 = p.add_run(tbl + ":  ")
        r1.bold = True
        r1.font.name = "Courier New"
        r1.font.size = Pt(9)
        p.add_run(desc)

    tip(doc, "The better your table descriptions, the more accurately the Data Agent will choose the right tables and columns when generating SQL. Invest time here — it directly improves demo quality.")

    # Step 7
    heading(doc, "Step 7: Test the Fabric Data Agent", 2)
    body(doc, "Use the Data Agent playground (right-hand chat panel) to validate the agent before connecting it to Copilot Studio.")
    body(doc, "Test with the following prompts:")

    test_prompts_7 = [
        "Which 3 hospitals had the lowest bed occupancy in Q4 2025?",
        "Show me all hospitals currently Off Track on their KPI scorecard.",
        "What is the total procurement spend on Pharmaceuticals year-to-date?",
        "Which hospital has the highest ED door-to-doctor wait time on average?",
        "Compare staff attrition rates between ASEAN and EMEA regions.",
    ]
    for i, p_text in enumerate(test_prompts_7, 1):
        numbered(doc, f'"{p_text}"')

    body(doc, "For each prompt, verify:")
    bullet(doc, "A SQL query is generated and shown in the 'SQL' tab of the response.")
    bullet(doc, "The result contains correct hospital names, values, and time periods.")
    bullet(doc, "No 'table not found' or 'column not found' errors appear.")

    warn(doc, "If you see SQL errors, check the column names in your table descriptions match the actual column headers in your Excel files. "
              "Common issues: spaces in column names, special characters, or sheet name mismatches.")

    # =====================================================================
    # PART 3
    # =====================================================================
    heading(doc, "PART 3: CONNECT FABRIC DATA AGENT TO COPILOT STUDIO", 1)

    # Step 8
    heading(doc, "Step 8: Get the Fabric Data Agent Connection Details", 2)
    numbered(doc, "In your Fabric Data Agent, click the 'Settings' icon (gear icon, top-right).")
    numbered(doc, "Copy the Agent ID or the connection endpoint URL — you will need this when adding the agent as a knowledge source in Copilot Studio.")
    numbered(doc, "Note the Fabric workspace name and region — these must match your Copilot Studio environment region.")
    tip(doc, "Fabric Data Agent exposes an API endpoint that is natively compatible with Copilot Studio's 'Knowledge' connector. No custom API or Azure Function is required.")

    # Step 9
    heading(doc, "Step 9: Create New Agent in Copilot Studio", 2)
    warn(doc, "Pre-requisite: You need a Copilot Studio licence (included in Microsoft 365 Copilot or available as a standalone add-on) and access to copilotstudio.microsoft.com.")
    numbered(doc, "Navigate to copilotstudio.microsoft.com and sign in.")
    numbered(doc, "Click 'Create' → 'New agent'.")
    numbered(doc, "Enter the agent name:  COO Intelligence Agent – Contoso Healthcare.")
    numbered(doc, "Enter the description:")
    bullet(doc, "Executive AI assistant for COO providing real-time insights on hospital operations, clinical quality, financials and workforce across all 24 Contoso hospitals.", level=1)
    numbered(doc, "In the 'Instructions' field, enter the following system prompt:")

    code_block(doc, """\
You are the COO Intelligence Agent for Contoso Healthcare Group. Your role is to provide
senior executives with clear, concise, data-driven insights about hospital performance.

You have access to live operational data across 24 hospitals in Malaysia, Singapore,
Turkey, India, China, Hong Kong, Brunei, Bulgaria, Netherlands, and Serbia.

Always:
- Be concise and executive-focused (lead with the key insight)
- Quantify answers with specific numbers
- Highlight top 3 best and worst performers when comparing hospitals
- Flag anything that needs immediate attention (Off Track / Red status)
- Offer to drill down into specifics when giving summaries

You can answer questions about:
- Patient Operations: bed occupancy, admissions, ED wait times, surgical theatre utilisation
- Clinical Quality: patient safety incidents, infection rates, satisfaction scores, outcomes
- Workforce: headcount, attrition, vacancy rates, training compliance
- Supply Chain: procurement spend, inventory levels, stockouts
- Executive KPIs: overall scorecard, RAG status, trend analysis""")

    numbered(doc, "Click 'Create agent'.")
    tip(doc, "The system prompt in Copilot Studio acts as the agent's personality and scope. Keep it executive-focused and action-oriented so the COO gets concise, useful answers.")

    # Step 10
    heading(doc, "Step 10: Add Fabric Data Agent as a Knowledge Source", 2)
    numbered(doc, "In the Copilot Studio agent editor, click 'Knowledge' in the top navigation.")
    numbered(doc, "Click 'Add knowledge'.")
    numbered(doc, "Select 'Dataverse and Fabric (preview)' or 'Microsoft Fabric' from the connector list.")
    numbered(doc, "Authenticate with your Microsoft 365 account if prompted.")
    numbered(doc, "Select your Fabric workspace: Contoso-Healthcare-COO.")
    numbered(doc, "Select the Data Agent: Contoso Healthcare COO Agent.")
    numbered(doc, "Click 'Add'. The knowledge source will appear in the Knowledge panel.")
    warn(doc, "The Fabric knowledge connector is in preview. If you cannot find it, ensure your tenant admin has enabled 'Generative AI features' in the Copilot Studio admin settings.")

    # Step 11
    heading(doc, "Step 11: (Optional) Add Pre-Built Conversation Topics", 2)
    body(doc,
         "For a polished demo experience, add 3 pre-built topics that map common COO questions to "
         "specific Data Agent queries. This ensures consistent, fast responses for the most likely demo questions.")

    heading(doc, "Topic 1: Hospital Performance Overview", 3)
    bullet(doc, "Trigger phrases:  'performance overview', 'how are we doing', 'hospital status', 'KPI summary'")
    bullet(doc, "Action: Call Fabric Data Agent with the following prompt:")
    code_block(doc, "Give me a summary of the top 5 and bottom 5 hospitals by overall KPI score for the most recent month.")

    heading(doc, "Topic 2: Operational Alerts", 3)
    bullet(doc, "Trigger phrases:  'any alerts', 'what needs attention', 'red status', 'off track hospitals'")
    bullet(doc, "Action: Call Fabric Data Agent with the following prompt:")
    code_block(doc, "List all hospitals with Off Track KPI status. For each, show the specific metrics causing the red flag.")

    heading(doc, "Topic 3: Regional Comparison", 3)
    bullet(doc, "Trigger phrases:  'regional breakdown', 'compare regions', 'ASEAN vs EMEA'")
    bullet(doc, "Action: Call Fabric Data Agent with the following prompt:")
    code_block(doc, """\
Compare key operational metrics (bed occupancy, patient satisfaction, EBITDA margin,
staff attrition) across regions: ASEAN, EMEA, APAC, South Asia for the latest quarter.""")

    tip(doc, "Test each topic trigger phrase before the demo. Topics take priority over the generative AI fallback, so they guarantee a fast, well-formatted response for your most important demo questions.")

    # Step 12
    heading(doc, "Step 12: Test End-to-End in Copilot Studio", 2)
    body(doc, "Use the 'Test' panel (right-hand side in Copilot Studio) to run end-to-end validation:")

    test_prompts_12 = [
        ("'How are we doing overall?'", "should trigger Topic 1 — Hospital Performance Overview"),
        ("'Are there any hospitals I should be worried about?'", "should trigger Topic 2 — Operational Alerts"),
        ("'How does ASEAN compare to EMEA?'", "should trigger Topic 3 — Regional Comparison"),
        ("'What is the average length of stay for cardiology patients in Singapore?'", "Data Agent answers directly from patient_operations_admissions / clinical_outcomes"),
        ("'Which hospitals have stockout incidents above 2 in the last 6 months?'", "tests supply chain data from supply_inventory"),
    ]
    for prompt, expected in test_prompts_12:
        p = doc.add_paragraph(style="List Number")
        r1 = p.add_run(prompt + " ")
        r1.bold = True
        p.add_run(f"→ {expected}")

    body(doc, "For each test, verify:")
    bullet(doc, "Response is factually accurate and references correct hospitals/metrics.")
    bullet(doc, "Response time is acceptable (typically 5–15 seconds for Data Agent queries).")
    bullet(doc, "No error messages or empty responses.")

    warn(doc, "If a topic is not triggering, check that the trigger phrases in Step 11 closely match the test input. "
              "Copilot Studio uses semantic matching — exact wording is not required, but the intent must be clear.")

    # Step 13
    heading(doc, "Step 13: Publish and Deploy", 2)
    numbered(doc, "Click 'Publish' in the Copilot Studio top bar.")
    numbered(doc, "Click 'Publish' again on the confirmation dialog.")
    numbered(doc, "Wait for the publish to complete (typically 1–3 minutes).")
    numbered(doc, "Click 'Channels' to select your deployment target:")

    body(doc, "Recommended deployment options:")
    bullet(doc, "Microsoft Teams:  Click 'Microsoft Teams' → 'Turn on Teams' → 'Open in Teams' to add as a Teams app. Share the app link with the COO and CFO channels.")
    bullet(doc, "SharePoint:  Use the Web Chat channel. Copy the embed code and paste into a SharePoint page as an 'Embed' web part on your executive dashboard.")
    bullet(doc, "Web Chat:  Embed on your internal intranet portal for broader executive access.")

    tip(doc, "For the closing demo, use the Teams deployment. It is the most familiar channel for healthcare executives and demonstrates seamless integration with their existing workflow.")

    # =====================================================================
    # PART 4 — DEMO SCRIPT
    # =====================================================================
    heading(doc, "PART 4: DEMO SCRIPT (For Closing Demo)", 1)

    heading(doc, "Suggested Demo Flow — 10 to 15 Minutes", 2)
    body(doc,
         "Open the Copilot Studio Test panel or the Teams deployment. "
         "Walk through the following script, narrating the key messages at each step:")

    demo_steps = [
        (
            "Open the agent and introduce",
            "\"Let me show you how Contoso Healthcare's COO can get instant answers from 24 hospitals, "
            "across 11 countries, in plain English.\"",
            None,
        ),
        (
            "Type: 'Give me a hospital performance overview for this month'",
            "Agent returns KPI scorecard summary with top and bottom performers.",
            "\"Instead of waiting for a weekly PowerPoint, the COO can see the full picture in seconds.\"",
        ),
        (
            "Type: 'Which hospitals are off track and why?'",
            "Agent lists hospitals with specific metrics flagging red — e.g., high readmissions, low satisfaction.",
            "\"The agent doesn't just show a red dot on a dashboard — it explains what is causing the problem.\"",
        ),
        (
            "Type: 'Drill into [specific hospital] — what is causing the low quality score?'",
            "Agent queries clinical_incidents, clinical_infection, clinical_satisfaction and returns a summary.",
            "\"From group overview to individual hospital root cause — all in one conversation.\"",
        ),
        (
            "Type: 'What is our total pharmaceutical spend this quarter vs last quarter?'",
            "Agent queries supply_procurement and returns a comparison with percentage change.",
            "\"Finance questions answered in real time — no need to wait for the CFO's monthly report.\"",
        ),
        (
            "Type: 'If we improve bed occupancy in our ASEAN hospitals to 88%, what is the revenue impact?'",
            "Agent combines operational data with reasoning to provide an analytical estimate.",
            "\"This is where it gets powerful — not just reporting, but strategic what-if analysis.\"",
        ),
    ]

    for i, (action, result, message) in enumerate(demo_steps, 1):
        heading(doc, f"Demo Step {i}: {action}", 3)
        p = doc.add_paragraph()
        p.add_run("Expected response: ").bold = True
        p.add_run(result)
        if message:
            p2 = doc.add_paragraph()
            r = p2.add_run("Suggested narration: ")
            r.bold = True
            r.italic = True
            p2.add_run(message).italic = True

    heading(doc, "Key Messages for IHH Executives", 2)
    key_messages = [
        ("This is not a dashboard — this is a conversation with your data.",
         "Dashboards show what you configured. This agent answers any question you have not thought of yet."),
        ("The COO can ask any question in plain English, on any device, in real time.",
         "No waiting for analysts. No SQL knowledge required. Available 24/7 on mobile, desktop, and Teams."),
        ("No SQL knowledge needed — the AI translates business questions into data queries.",
         "The Fabric Data Agent automatically writes and executes the SQL — executives just ask."),
        ("Built on Microsoft Fabric + Copilot Studio — enterprise-grade, secure, compliant.",
         "All data stays within your Microsoft tenant. Full audit trail. HIPAA/PDPA-aligned architecture."),
        ("Time to value: from data upload to working agent in under 2 hours.",
         "This is not a 6-month implementation. With your existing data, you can have this running today."),
    ]

    for headline, detail in key_messages:
        p = doc.add_paragraph(style="List Bullet")
        r1 = p.add_run(f'"{headline}"  ')
        r1.bold = True
        p.add_run(detail)

    # =====================================================================
    # APPENDIX
    # =====================================================================
    heading(doc, "APPENDIX: Table Reference", 1)

    headers = ["Table Name", "Source File", "Sheet Name", "Key Columns (examples)"]
    appendix_data = [
        ("patient_operations_admissions", "Contoso_Patient_Operations.xlsx", "Admissions_Bed_Occupancy",
         "Hospital, Country, Region, Year, Month, Bed_Occupancy_Rate_Pct, Total_Admissions, Avg_LOS_Days"),
        ("patient_operations_dept",       "Contoso_Patient_Operations.xlsx", "Dept_Utilisation",
         "Hospital, Department, Specialty, Outpatient_Visits, Inpatient_Admissions"),
        ("patient_operations_ed",         "Contoso_Patient_Operations.xlsx", "Emergency_Dept_KPIs",
         "Hospital, Month, Door_to_Doctor_Min, ED_Attendance, Triage_1_Count"),
        ("patient_operations_theatre",    "Contoso_Patient_Operations.xlsx", "Surgical_Theatre",
         "Hospital, Month, Theatre_Utilisation_Pct, Cases_Performed, Cancellation_Rate_Pct"),
        ("workforce_headcount",           "Contoso_Workforce_HR.xlsx",       "Headcount_Monthly",
         "Hospital, Month, Total_FTE, Vacancy_Rate_Pct, Attrition_Rate_Pct"),
        ("workforce_staff_directory",     "Contoso_Workforce_HR.xlsx",       "Staff_Directory",
         "Employee_ID, Name, Hospital, Department, Role, Region, Employment_Type"),
        ("workforce_training",            "Contoso_Workforce_HR.xlsx",       "Training_Development",
         "Hospital, Training_Type, Compliance_Rate_Pct, Completion_Count"),
        ("supply_procurement",            "Contoso_Supply_Chain.xlsx",       "Procurement_Orders",
         "Hospital, Supplier, Category, Order_Value_USD, Order_Date, Delivery_Status"),
        ("supply_inventory",              "Contoso_Supply_Chain.xlsx",       "Inventory_Levels",
         "Hospital, Item_Category, Stock_Level, Days_on_Hand, Stockout_Incidents"),
        ("clinical_incidents",            "Contoso_Clinical_Quality.xlsx",   "Patient_Safety_Incidents",
         "Hospital, Incident_Type, Severity, Date, Outcome"),
        ("clinical_infection",            "Contoso_Clinical_Quality.xlsx",   "Infection_Control",
         "Hospital, Month, CLABSI_Rate, CAUTI_Rate, SSI_Rate, C_diff_Rate"),
        ("clinical_satisfaction",         "Contoso_Clinical_Quality.xlsx",   "Patient_Satisfaction",
         "Hospital, Month, Overall_Score, Communication_Score, Cleanliness_Score"),
        ("clinical_outcomes",             "Contoso_Clinical_Quality.xlsx",   "Clinical_Outcomes",
         "Hospital, Specialty, Month, Readmission_Rate_Pct, Mortality_Rate_Pct, Avg_LOS"),
        ("kpi_scorecard",                 "Contoso_Executive_KPIs.xlsx",     "Monthly_KPI_Scorecard",
         "Hospital, Month, Financial_Score, Operations_Score, Quality_Score, People_Score, ESG_Score, Overall_Score, Status"),
        ("kpi_dashboard",                 "Contoso_Executive_KPIs.xlsx",     "COO_Dashboard_Summary",
         "Region, Month, Avg_Overall_Score, Hospitals_On_Track, Hospitals_At_Risk, Hospitals_Off_Track"),
    ]

    tbl = doc.add_table(rows=1, cols=4)
    tbl.style = "Table Grid"
    hdr_cells = tbl.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = h
        hdr_cells[i].paragraphs[0].runs[0].bold = True
        set_cell_bg(hdr_cells[i], "1F5C99")
        hdr_cells[i].paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    for row_data in appendix_data:
        row_cells = tbl.add_row().cells
        for i, val in enumerate(row_data):
            row_cells[i].text = val
            if i == 0:
                row_cells[i].paragraphs[0].runs[0].font.name = "Courier New"
                row_cells[i].paragraphs[0].runs[0].font.size = Pt(8)

    doc.add_paragraph()
    tip(doc, "Bookmark this appendix — it is your reference when fixing SQL errors or adding new table descriptions to the Fabric Data Agent.")

    # ---- Save ----
    doc.save(str(OUTPUT_PATH))
    print(f"Saved: {OUTPUT_PATH}")
    print(f"File size: {OUTPUT_PATH.stat().st_size:,} bytes")

    # ---- Copy ----
    COPY_PATH.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(str(OUTPUT_PATH), str(COPY_PATH))
    print(f"Copied to: {COPY_PATH}")
    print(f"Copy size: {COPY_PATH.stat().st_size:,} bytes")


if __name__ == "__main__":
    build()
