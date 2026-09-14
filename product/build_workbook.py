"""Builds the Contractor Job Costing & Estimate workbook v2.

Usage: python3 build_workbook.py [--blank] OUTPUT.xlsx
Set MINI=1 to build a small-capacity variant (for compatibility tests).
"""
import sys, json, os
from openpyxl import Workbook
from layout import *
import sheets_setup, sheets_jobs, sheets_costs, sheets_reports
import sample_data as SD

BLANK = "--blank" in sys.argv
OUT = [a for a in sys.argv[1:] if not a.startswith("--")][0]

wb = Workbook(); wb.active.title = "Start Here"
sheets_setup.build_settings(wb, BLANK)
sheets_setup.build_clients(wb)
sheets_setup.build_vendors(wb)
sheets_setup.build_library(wb)
sheets_jobs.build_jobs(wb)
sheets_jobs.build_estimate(wb)
sheets_jobs.build_proposal(wb)
sheets_costs.build_schedule(wb)
sheets_costs.build_timesheet(wb)
sheets_costs.build_job_costs(wb)
sheets_costs.build_mileage(wb)
sheets_costs.build_daily_log(wb)
sheets_costs.build_change_orders(wb)
sheets_costs.build_co_form(wb)
sheets_costs.build_invoices(wb)
sheets_costs.build_invoice_print(wb)
sheets_costs.build_payments(wb)
_, stmt = sheets_costs.build_client_statement(wb)
_, jpr = sheets_reports.build_job_profit_report(wb)
_, rep = sheets_reports.build_reports(wb)
sheets_reports.build_tax_summary(wb)
_, hc = sheets_reports.build_health_check(wb)
_, dash = sheets_reports.build_dashboard(wb, hc)
sheets_reports.build_start_here(wb)

def fill(ws_name, rows, headers):
    ws = wb[ws_name]; cm = C[ws_name]
    for i, row in enumerate(rows):
        for h, v in zip(headers, row):
            if v is not None and v != "":
                ws[f"{cm[h]}{R1 + i}"] = v

if not BLANK:
    fill("Clients", SD.CLIENTS, ["Client", "Contact person", "Phone", "Email", "Billing address", "Notes"])
    vc = wb["Vendors & Crew"]
    for i, row in enumerate(SD.CREW):
        for h, v in zip(["Worker", "Role", "Cost Rate ($/hr)", "Bill Rate ($/hr)", "Phone"], row): vc[f"{CC[h]}{CREW_R1 + i}"] = v
    for i, row in enumerate(SD.VENDORS):
        for h, v in zip(["Vendor / Sub", "Type", "Trade", "Phone", "Email", "License #", "Insurance Expiry", "W-9 on file", "1099-eligible"], row):
            if v is not None and v != "": vc[f"{VC[h]}{VEND_R1 + i}"] = v
    fill("Cost Library", SD.LIBRARY, ["Item", "Cost Code", "Unit", "Unit Cost"])
    fill("Jobs", SD.JOBS, ["Job ID", "Client", "Job address / description", "Status", "Start date", "Target end", "% Complete (manual)", "Contract Value (signed)"])
    fill("Estimate", SD.ESTIMATE, ["Job ID", "Cost Code", "Library Item", "Description", "Qty", "Unit", "Unit Cost"])
    fill("Schedule", SD.SCHEDULE, ["Job ID", "Task", "Assigned to", "Start", "Duration (workdays)", "% Complete"])
    fill("Timesheet", SD.TIMESHEET, ["Date", "Worker", "Job ID", "Cost Code", "Hours", "Cost Rate Override"])
    fill("Job Costs", SD.JOB_COSTS, ["Date", "Job ID", "Cost Code", "Vendor", "Description", "Qty or Hours", "Unit Cost / Rate", "Paid?", "Payment Method", "Receipt / Invoice #"])
    fill("Mileage", SD.MILEAGE, ["Date", "Job ID", "Cost Code", "From", "To", "Purpose", "Miles"])
    fill("Daily Log", SD.DAILY_LOG, ["Date", "Job ID", "Weather", "Crew on site", "Work performed", "Issues / delays", "Inspections / visitors"])
    fill("Change Orders", SD.CHANGE_ORDERS, ["CO #", "Job ID", "Date", "Description", "Status", "Markup %", "Added Cost (budget)", "Schedule impact (days)", "Client approved date"])
    fill("Invoices", SD.INVOICES, ["Invoice #", "Job ID", "Invoice Date", "Due Date", "Amount", "Description"])
    fill("Payments", SD.PAYMENTS, ["Date", "Job ID", "Invoice #", "Amount", "Method", "Reference / Check #"])
    for ref, v in SD.SELECTORS.items():
        s, a = ref.split("!"); wb[s][a] = v

from openpyxl.workbook.defined_name import DefinedName
for nm, ref in [("JobIDs", JOB_IDS), ("ClientNames", CLIENT_NAMES), ("CostCodes", CODE_RANGE), ("CrewNames", CREW_NAMES), ("VendorNames", VEND_NAMES),
                ("InvoiceNumbers", INV_NUMS), ("LibraryItems", LIB_NAMES), ("MarkupTable", MARKUP_TABLE)]:
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)
wb.properties.title = "Contractor Job Costing & Estimate Spreadsheet"
wb.properties.creator = "Contractor Job Costing template"
wb.properties.subject = "Job costing, estimating, scheduling, invoicing and profit reporting for small contractors"
wb.properties.keywords = "contractor, job costing, estimate, proposal, gantt, timesheet, change order, invoice, budget vs actual"
wb.properties.description = "Version 1.0 (2026-09). 24 sheets. Light-blue cells are inputs."
wb.calculation.fullCalcOnLoad = True
if os.environ.get("MINI"):
    wb.move_sheet("Health Check", offset=-(len(wb.sheetnames) - 1))
# First-open experience: cursor on the selector or first input cell of every sheet; wide tables open at 90 % zoom.
CURSOR = {"Start Here": "B1", "Settings": "B5", "Clients": "A5", "Vendors & Crew": "A5", "Cost Library": "A5", "Jobs": "A5", "Estimate": "A5",
          "Proposal": "I2", "Schedule": "A5", "Timesheet": "A5", "Job Costs": "A5", "Mileage": "A5", "Daily Log": "A5", "Change Orders": "A5",
          "CO Form": "I2", "Invoices": "A5", "Invoice Print": "I2", "Payments": "A5", "Client Statement": "L2", "Job Profit Report": "C4",
          "Reports": "B1", "Tax Summary": "B1", "Dashboard": "B1", "Health Check": "B1"}
ZOOM = {"Estimate": 90, "Timesheet": 90, "Job Costs": 90, "Invoices": 90, "Schedule": 90, "Vendors & Crew": 90, "Change Orders": 90, "Settings": 90, "Dashboard": 90}
from openpyxl.worksheet.views import Selection
for name, cell in CURSOR.items():
    ws = wb[name]; sv = ws.sheet_view
    if sv.selection: sv.selection[-1].activeCell = cell; sv.selection[-1].sqref = cell
    else: sv.selection.append(Selection(activeCell=cell, sqref=cell))
    if name in ZOOM: sv.zoomScale = ZOOM[name]
    sv.tabSelected = (name == "Start Here")
wb.active = 0
wb.save(OUT)
meta = {"C": C, "N": N, "R1": R1, "CC": CC, "VC": VC, "CREW_R1": CREW_R1, "VEND_R1": VEND_R1, "VEND_R2": VEND_R2,
        "jobs_total_row": jobs_total_row(), "jpr": jpr, "rep": rep, "hc": hc, "dash": dash, "stmt": stmt, "MARKUP_R1": MARKUP_R1, "CODE_R1": CODE_R1,
        "N_PROP": N_PROP, "GANTT_C1": GANTT_C1}
with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "layout.json"), "w") as f: json.dump(meta, f, indent=1)
print("saved", OUT, "sheets:", len(wb.sheetnames))
