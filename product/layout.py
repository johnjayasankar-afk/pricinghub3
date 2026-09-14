"""Single source of truth for sheet names, headers, column letters and cross-sheet ranges (v3)."""
from wb_common import *

R1 = 5  # first data row on every table sheet

SHEETS = ["Start Here", "Dashboard", "Settings", "Clients", "Vendors & Crew", "Cost Library", "Jobs", "Estimate", "Proposal",
          "Schedule", "Timesheet", "Job Costs", "Mileage", "Daily Log", "Change Orders", "CO Form", "Invoices", "Invoice Print",
          "Payments", "Client Statement", "Job Profit Report", "Reports", "Tax Summary", "Health Check"]

H = {}; C = {}; N = {}
def define(name, headers, n):
    H[name] = headers; C[name] = colmap(headers); N[name] = n

define("Clients", ["Client", "Contact person", "Phone", "Email", "Billing address", "Notes", "Jobs", "Contract value (all jobs)", "Collected", "Balance due", "Check"], N_CLI)
define("Jobs", ["Job ID", "Client", "Job address / description", "Status", "Start date", "Target end", "% Complete (manual)",
                "Estimate Price", "Contract Value (signed)", "Approved Change Orders", "Revised Contract", "Estimated Cost",
                "Actual Cost to Date", "Cost Variance", "% of Budget Used", "Labor Hours", "Invoiced", "Collected", "Balance Due",
                "Projected Gross Profit", "Projected Margin %", "Estimated Margin %", "Days to Target", "Check", "Notes"], N_JOBS)
define("Estimate", ["Job ID", "Cost Code", "Library Item", "Description", "Qty", "Unit", "Unit Cost", "Category", "Unit Cost Used", "Unit Used",
                    "Est. Cost", "Markup Override", "Price to Client", "Markup Used", "Proposal Line #", "Check", "Notes"], N_EST)
define("Cost Library", ["Item", "Cost Code", "Unit", "Unit Cost", "Category", "Used in estimates", "Notes"], N_LIB)
define("Schedule", ["Job ID", "Task", "Assigned to", "Start", "Duration (workdays)", "End", "% Complete", "Status", "Check", "Notes"], N_SCH)
define("Timesheet", ["Date", "Worker", "Job ID", "Cost Code", "Hours", "Cost Rate Override", "Rate Used", "Labor Cost (incl. burden)", "Bill Rate",
                     "Billable Value", "Week Ending", "Category", "Check", "Notes"], N_TS)
define("Job Costs", ["Date", "Job ID", "Cost Code", "Category", "Vendor", "Description", "Qty or Hours", "Unit Cost / Rate", "Amount",
                     "Paid?", "Payment Method", "Receipt / Invoice #", "Check", "Notes"], N_COST)
define("Mileage", ["Date", "Job ID", "Cost Code", "From", "To", "Purpose", "Miles", "Rate", "Amount", "Category", "Check", "Notes"], N_MIL)
define("Daily Log", ["Date", "Job ID", "Weather", "Crew on site", "Hours logged (Timesheet)", "Work performed", "Issues / delays",
                     "Inspections / visitors", "Check"], N_LOG)
define("Change Orders", ["CO #", "Job ID", "Date", "Description", "Status", "Markup %", "Added Cost (budget)", "Price to Client",
                         "Added Profit", "Schedule impact (days)", "Client approved date", "Check", "Notes"], N_CO)
define("Invoices", ["Invoice #", "Job ID", "Invoice Date", "Due Date", "Amount", "Description", "Paid to Date", "Balance", "Status",
                    "Days Overdue", "Aging Bucket", "Check", "Client", "Stmt Line #", "Notes"], N_INV)
define("Payments", ["Date", "Job ID", "Invoice #", "Amount", "Method", "Reference / Check #", "Check", "Notes"], N_PAY)

CREW_H = ["Worker", "Role", "Cost Rate ($/hr)", "Bill Rate ($/hr)", "Phone", "Hours YTD", "Labor Cost YTD", "Check", "Notes"]
VEND_H = ["Vendor / Sub", "Type", "Trade", "Phone", "Email", "License #", "Insurance Expiry", "Insurance Status", "W-9 on file",
          "1099-eligible", "Paid YTD", "Unpaid Balance", "1099 Flag", "Check", "Notes"]
CREW_R1, CREW_R2 = 5, 4 + N_CREW
VEND_HDR = CREW_R2 + 6
VEND_R1, VEND_R2 = VEND_HDR + 1, VEND_HDR + N_VEND
CC = colmap(CREW_H); VC = colmap(VEND_H)

def last(name): return R1 + N[name] - 1
def col(name, hdr): return C[name][hdr]
def rr(name, hdr): return rng(name, C[name][hdr], R1, last(name))
def rc(name, hdr, r): return f"{C[name][hdr]}{r}"

# Settings cells (company block rows 5-19, markup rows 22-27, document text rows 30-34)
S = {"company": "Settings!$B$5", "address": "Settings!$B$6", "phone": "Settings!$B$7", "license": "Settings!$B$8",
     "contingency": "Settings!$B$9", "tax": "Settings!$B$10", "deposit": "Settings!$B$11", "progress": "Settings!$B$12",
     "grace": "Settings!$B$13", "year": "Settings!$B$14", "mileage": "Settings!$B$15", "valid_days": "Settings!$B$16",
     "overhead": "Settings!$B$17", "burden": "Settings!$B$18", "tax_scope": "Settings!$B$19",
     "proposal_terms": "Settings!$B$30", "invoice_instructions": "Settings!$B$31", "invoice_footer": "Settings!$B$32", "co_terms": "Settings!$B$33",
     "proposal_exclusions": "Settings!$B$34"}
DOC_HDR, DOC_R1 = 29, 30
MARKUP_R1, MARKUP_R2 = 22, 27
MARKUP_TABLE = f"Settings!$A${MARKUP_R1}:$B${MARKUP_R2}"
CAT_RANGE = f"Settings!$A${MARKUP_R1}:$A${MARKUP_R2}"
CODE_R1, CODE_R2 = 5, 4 + N_CODES
CODE_TABLE = f"Settings!$E${CODE_R1}:$G${CODE_R2}"
CODE_RANGE = f"Settings!$E${CODE_R1}:$E${CODE_R2}"
STATUS_LIST = ["Lead", "Quoted", "Approved", "In Progress", "Complete", "Closed", "Lost"]
UNIT_LIST = ["ea", "hr", "day", "sq ft", "lin ft", "sq yd", "cu yd", "gal", "box", "lot"]
PAYM_LIST = ["Check", "ACH / Bank", "Card", "Cash", "Zelle", "Other"]
VTYPE_LIST = ["Subcontractor", "Supplier", "Equipment rental", "Other"]
TAX_SCOPE_LIST = ["Material only", "All lines"]
STATUS_RANGE = f"Settings!$I$5:$I${4+len(STATUS_LIST)}"
UNIT_RANGE = f"Settings!$J$5:$J${4+len(UNIT_LIST)}"
PAYM_RANGE = f"Settings!$K$5:$K${4+len(PAYM_LIST)}"
VTYPE_RANGE = f"Settings!$L$5:$L${4+len(VTYPE_LIST)}"
CATS = [("Labor", 0.35), ("Material", 0.20), ("Subcontractor", 0.15), ("Equipment", 0.15), ("Permit/Fee", 0.0), ("Other", 0.10)]

JOB_IDS = rr("Jobs", "Job ID")
CLIENT_NAMES = rr("Clients", "Client")
LIB_TABLE = f"'Cost Library'!$A${R1}:$D${last('Cost Library')}"
LIB_NAMES = rr("Cost Library", "Item")
CREW_NAMES = rng("Vendors & Crew", CC["Worker"], CREW_R1, CREW_R2)
CREW_TABLE = f"'Vendors & Crew'!$A${CREW_R1}:$D${CREW_R2}"
VEND_NAMES = rng("Vendors & Crew", VC["Vendor / Sub"], VEND_R1, VEND_R2)
INV_NUMS = rr("Invoices", "Invoice #")
CO_NUMS = rr("Change Orders", "CO #")

PROPOSAL_SEL = "Proposal!$I$2"
JPR_SEL = "'Job Profit Report'!$C$4"
COFORM_SEL = "'CO Form'!$I$2"
INVPRINT_SEL = "'Invoice Print'!$I$2"
STATEMENT_SEL = "'Client Statement'!$L$2"
GANTT_C1 = 11
SCHED_START = "Schedule!$B$2"
TS_WEEK = "Timesheet!$Q$3"
N_STMT = 30

def jobs_total_row(): return last("Jobs") + 1
