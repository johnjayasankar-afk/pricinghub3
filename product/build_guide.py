"""Generates the buyer's Quick-Start Guide PDF (v2, 23 sheets)."""
import os
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Quick-Start-Guide.pdf")
ss = getSampleStyleSheet()
H1 = ParagraphStyle("H1", parent=ss["Heading1"], textColor=colors.HexColor("#1F3A5F"), spaceAfter=8)
H2 = ParagraphStyle("H2", parent=ss["Heading2"], textColor=colors.HexColor("#2A7F8E"), spaceBefore=12, spaceAfter=5)
P = ParagraphStyle("P", parent=ss["BodyText"], leading=14, spaceAfter=5)
SM = ParagraphStyle("SM", parent=P, fontSize=8.5, leading=11, textColor=colors.HexColor("#555555"))
TS = TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F3A5F")), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                 ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#BFBFBF")), ("FONTSIZE", (0, 0), (-1, -1), 8.5), ("VALIGN", (0, 0), (-1, -1), "TOP")])
def tbl(rows, widths):
    rows = [[Paragraph(str(c), SM if i else ParagraphStyle("h", parent=SM, textColor=colors.white)) for c in r] for i, r in enumerate(rows)]
    t = Table(rows, colWidths=widths); t.setStyle(TS); return t

doc = SimpleDocTemplate(OUT, pagesize=letter, leftMargin=0.85*inch, rightMargin=0.85*inch, topMargin=0.75*inch, bottomMargin=0.75*inch,
                        title="Contractor Job Costing & Estimate Spreadsheet — Quick-Start Guide")
s = [Paragraph("Contractor Job Costing &amp; Estimate Spreadsheet", H1), Paragraph("Quick-Start Guide · Excel and Google Sheets · 24 sheets · Version 1.0 (2026-09)", ss["Heading3"]), Spacer(1, 6),
     Paragraph("Thank you for your purchase. This workbook runs a job from first quote to final payment: estimate, proposal, schedule, crew hours, costs, "
               "mileage, daily log, change orders, invoices, payments, and profit reporting. Setup takes about ten minutes.", P)]
s.append(Paragraph("Files in your download", H2))
s.append(tbl([["File", "Use it for"], ["Contractor-Job-Costing-SAMPLE.xlsx", "Three worked jobs (kitchen, bathroom, deck) with crew, vendors, schedule, invoices. Open first."],
              ["Contractor-Job-Costing-BLANK.xlsx", "Your working copy. Same formulas, no sample rows."], ["Quick-Start-Guide.pdf", "This guide."]], [2.5*inch, 4.2*inch]))
s.append(Paragraph("Open it", H2))
s.append(Paragraph("<b>Excel:</b> double-click the .xlsx (Excel 2016 or newer, Microsoft 365, Mac, iPad/web). If Excel opens it in Protected View click <b>Enable Editing</b>. No macros, no add-ins.<br/>"
                   "<b>Google Sheets:</b> upload the .xlsx to Google Drive, open it, then <b>File → Save as Google Sheets</b>. Formulas, dropdowns, conditional formatting and the Gantt chart carry over; the three bar charts may need a one-time reformat (right-click → Edit chart).", P))
s.append(Paragraph("Color key", H2))
s.append(Paragraph("<b>Light-blue cells are inputs.</b> Every other cell is a formula: type only in light-blue cells (Undo with Ctrl+Z / Cmd+Z if you overwrite one). "
                   "Red = over budget, overdue, late, expired, or a data problem. Yellow = watch. Green = on track, paid, done.", P))
s.append(Paragraph("Ten-minute setup", H2))
steps = [("Settings", "Company name, address, phone, license. Default markup % per category (Labor, Material, Subcontractor, Equipment, Permit/Fee, Other). Contingency, sales tax on materials, deposit and progress %, overdue grace days, reporting year, mileage rate, proposal validity, overhead %. Edit the 24 cost codes and their categories to match your trade. Document text for proposals, invoices and change orders."),
         ("Vendors &amp; Crew", "Crew table: each worker with a Cost Rate (what you pay) and Bill Rate (what you charge). Labor burden % in Settings (payroll taxes, insurance) is added to cost on the Timesheet. Vendor table: subs and suppliers with insurance expiry (status turns red when expired), W-9 and 1099-eligible flags, paid year-to-date and the unpaid balance you still owe each one."),
         ("Cost Library", "Optional price book. Items with cost code, unit and unit cost. On the Estimate, pick a Library Item and the unit and unit cost fill in unless you type your own. 'Used in estimates' shows how often each item is picked."),
         ("Clients", "One row per client with contact and billing details. Jobs pick clients from this list; lifetime value and balance due calculate."),
         ("Jobs", "One row per job: Job ID (for example J-1001), client, description, status, start and target end dates. Estimate Price fills from the Estimate. When a contract is signed type the amount in <b>Contract Value</b>; blank means the estimate price is used. Approved change orders are added automatically.")]
for n, t in steps: s.append(Paragraph(f"<b>{n}.</b> {t}", P))
s.append(Paragraph("Running a job", H2))
run = [("Estimate", "One line per item: Job ID, cost code, optional library item, description, qty, unit, unit cost. Markup comes from the category; type a Markup Override to change one line. The Check column flags unknown jobs or codes."),
       ("Proposal", "Pick a Job ID in cell I2, to the right of the printable area. You get a client-ready document with a proposal number: your details, the client's details, scope summary by category, contingency, tax, total, payment schedule (deposit / progress / final), up to 40 detailed lines, terms, exclusions and assumptions (from Settings), and signature lines. Columns H–J are internal (cost and margin) and are outside the print area. Print or File → Save as PDF."),
       ("Schedule", "One row per task: Job ID, task, assigned worker, start date, duration in workdays and % complete. End date skips weekends. Status shows Not started / In progress / Late / Done. The 12-week Gantt chart starts at the date in B2 (defaults to this Monday); type any date to scroll. Dark bars = completed portion, blue = planned, yellow column = today."),
       ("Timesheet", "Crew hours by day, worker, job and cost code. Rate Used comes from the crew table unless you type an override (overtime, for example). Labor cost flows into the job's actual cost; billable value uses the bill rate. The weekly payroll summary on the right totals hours and cost per worker for the week ending in Q4. Log crew labor here <b>or</b> in Job Costs, never both."),
       ("Job Costs", "Materials, subcontractors, equipment, permits, dumpsters and other actual costs. Qty × Unit Cost = Amount. Category comes from the cost code. Mark Paid? No for bills you still owe; the dashboard totals them."),
       ("Mileage", "Job-related driving. Amount = miles × the rate in Settings and counts toward the job under the cost code you pick."),
       ("Daily Log", "Date, job, weather, crew count, work performed, issues and delays, inspections. Hours pull from the Timesheet for that date and job. This is your record if a delay or change is ever disputed."),
       ("Change Orders", "Log added work with its cost; Price to Client applies the markup (the Other category by default, or type a Markup %). Only <b>Approved</b> change orders raise the revised contract and the budget. Record the schedule impact and the client-approved date, then print the <b>CO Form</b> for signature; it shows the revised target completion date."),
       ("Invoices and Payments", "Record invoices with due dates and a description (the sheet suggests the next invoice number); record payments with the Invoice #. Balance, status (Open / Partial / Paid / Overdue), days overdue and aging bucket calculate. <b>Invoice Print</b> builds a printable invoice with a contract summary and a tear-off remittance slip for any Invoice #."),
       ("Client Statement", "Pick a client: every invoice with amount, paid, balance and status, total balance due, and an aging strip. Print or save as PDF and send it monthly to anyone who owes you money.")]
for n, t in run: s.append(Paragraph(f"<b>{n}.</b> {t}", P))
s.append(PageBreak())
s.append(Paragraph("Reading the reports", H2))
rep = [("Dashboard", "Sixteen KPIs: contract value, estimated and actual cost, projected profit and margin, margin after overhead, invoiced, collected, outstanding, overdue, backlog (active jobs not yet invoiced), unpaid bills, active jobs, jobs over budget, labor hours this month and win rate. An alerts row (overdue invoices, late tasks, insurance, data issues), a This-Month row (invoiced, collected, costs, cash profit), three charts, a per-job table and a monthly P&amp;L with invoiced, collected, costs and hours."),
       ("Job Profit Report", "Pick a Job ID: contract, invoiced, collected, profit and margin, overhead allocation and net margin, estimated vs actual labor hours; budget vs actual by category and by cost code (estimate + approved change orders vs Job Costs + Timesheet + Mileage) with OVER / Watch / On track flags; unpaid bills, pending change orders, late tasks, daily-log count."),
       ("Reports", "Receivables aging (Current, 1–30, 31–60, 61–90, 90+), a six-month cash forecast from open invoice due dates, pipeline by status with win rate, and cost-code performance across all jobs so you can see which codes you consistently under-estimate."),
       ("Tax Summary", "For the reporting year: cash collected, invoiced, job costs by category from all three sources, gross profit, business miles and mileage amount, a quarter-by-quarter table for estimated-tax planning, and a 1099 review list (vendors marked eligible and paid $600 or more). Informational only; confirm with your tax professional."),
       ("Health Check", "Counts data problems from every Check column (unknown Job IDs, duplicate invoice numbers, missing dates, workers not in the crew table and so on), attention items (overdue invoices, late tasks, expired insurance) and formula errors per sheet. Keep it at ALL CLEAR.")]
for n, t in rep: s.append(Paragraph(f"<b>{n}.</b> {t}", P))
s.append(Paragraph("How profit is calculated", H2))
s.append(Paragraph("Revised Contract = Contract Value (or Estimate Price if blank) + approved change-order prices.<br/>Estimated Cost = estimate line costs + approved change-order costs.<br/>"
                   "Actual Cost = Job Costs + Timesheet labor cost (hours × rate × (1 + labor burden)) + Mileage.<br/>Projected Gross Profit = Revised Contract − the larger of Estimated Cost and Actual Cost. Once actual costs exceed the budget, profit falls immediately.<br/>"
                   "Projected Margin = Projected Gross Profit ÷ Revised Contract. Margin after overhead subtracts the overhead % from Settings.", P))
s.append(Paragraph("Frequently asked questions", H2))
faq = [("Can I stop myself overwriting formulas?", "Yes. Input cells are already unlocked, so on any sheet choose Review → Protect Sheet (no password) and only the light-blue cells stay editable. Unprotect before inserting rows."),
       ("Sales tax on labor too?", "Settings → 'Sales tax applies to': Material only (default) or All lines."),
       ("Can I add more rows?", "Yes. Insert rows inside a table (above the TOTAL row), then copy the formula cells from the row above into the new rows."),
       ("Can I add more cost codes or crew?", "Settings holds 24 codes and the crew table holds 30 workers; the reports mirror them. Rename unused ones rather than adding rows."),
       ("A dropdown is empty.", "Add the item to its source first: jobs on Jobs, clients on Clients, workers and vendors on Vendors &amp; Crew, invoice numbers on Invoices."),
       ("Timesheet hours show a rate of 0.", "The worker name does not match the crew table exactly. Pick it from the dropdown."),
       ("The Gantt chart is empty.", "Change the Chart start date in Schedule B2 to a date near your tasks."),
       ("I see #NAME? or #REF!.", "A formula was overwritten or a column deleted. Undo, or copy the same cell from the SAMPLE file. Health Check shows which sheet."),
       ("Dates look wrong in Google Sheets.", "Set File → Settings → Locale to your region."),
       ("Multiple companies or years?", "One copy per company. For a new year, copy the file, change the reporting year in Settings and clear closed jobs."),
       ("Other trades?", "Any job-based business works: landscaping, painting, roofing, electrical, plumbing, cleaning, events. Rename the cost codes."),
       ("Named ranges?", "JobIDs, ClientNames, CostCodes, CrewNames, VendorNames, InvoiceNumbers, LibraryItems and MarkupTable are defined for your own formulas (Formulas → Name Manager)."),
       ("Another currency?", "Amounts are formatted in US dollars. Select the cells and change the number format (Format Cells) to your currency symbol."),
       ("Is my data private?", "Completely. The file lives on your computer or your Google Drive. Nothing is sent anywhere.")]
for q, a in faq: s.append(Paragraph(f"<b>{q}</b> {a}", P))
s.append(Paragraph("Support", H2))
s.append(Paragraph("Message the shop where you bought this with the sheet and cell you are looking at. Typical reply within one business day.", P))
s.append(Paragraph("License and disclaimer", H2))
s.append(Paragraph("Licensed for use within your own business on any number of your devices. You may not resell, share or redistribute the template or a modified version of it. "
                   "This workbook is an organizational and estimating tool, not accounting, tax, legal or financial advice; verify figures before quoting, invoicing or filing. "
                   "Sample company, client, worker and vendor names are fictional.", SM))
doc.build(s); print("saved", OUT)
