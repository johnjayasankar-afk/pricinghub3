"""Schedule (Gantt), Timesheet, Job Costs, Mileage, Daily Log, Change Orders, CO Form, Invoices, Invoice Print, Payments."""
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from layout import *

def ci(name, h): return ord(C[name][h]) - 64
def chk_fill(ws, n):
    c = C[n]; ws.conditional_formatting.add(f"{c['Check']}{R1}:{c['Check']}{last(n)}", FormulaRule(formula=[f'{c["Check"]}{R1}<>""'], fill=RED_FILL))
def unknown_job(g): return f'COUNTIF({JOB_IDS},{g})=0'

def build_schedule(wb):
    ws = wb.create_sheet("Schedule"); n = "Schedule"; c = C[n]
    title(ws, "Schedule", "One row per task. End date skips weekends. The 12-week chart starts at the date in B2 (defaults to this Monday); type any date to scroll.", TAB["input"])
    ws["A2"] = "Chart start"; ws["A2"].font = BOLD
    inp(ws, 2, 2, "=TODAY()-WEEKDAY(TODAY(),2)+1", DATEF)
    ws["C2"] = "Legend:"; ws["C2"].font = NOTE
    ws["D2"] = "planned"; ws["D2"].fill = BAR_FILL; ws["D2"].font = Font(color="FFFFFF", size=9)
    ws["E2"] = "done"; ws["E2"].fill = BAR_DONE_FILL; ws["E2"].font = Font(color="FFFFFF", size=9)
    ws["F2"] = "today"; ws["F2"].fill = TODAY_FILL; ws["F2"].font = Font(size=9)
    ws["G2"] = "weekend"; ws["G2"].fill = GRAY_FILL; ws["G2"].font = Font(size=9)
    header(ws, 4, H[n], [11, 30, 14, 11, 10, 11, 9, 12, 20, 22])
    ws.freeze_panes = f"{get_column_letter(GANTT_C1)}5"
    g1, g2 = GANTT_C1, GANTT_C1 + GANTT_DAYS - 1
    for k in range(g1, g2 + 1):
        L = get_column_letter(k); ws.column_dimensions[L].width = 2.6
        prev = get_column_letter(k - 1)
        ws.cell(row=3, column=k, value=f"=$B$2" if k == g1 else f"={prev}3+1").number_format = "d"
        ws.cell(row=2, column=k, value=f'=IF(OR(DAY({L}3)=1,{k}={g1}),TEXT({L}3,"mmm"),"")').font = Font(bold=True, size=8, color=NAVY)
        x = ws.cell(row=4, column=k, value=f'=LEFT(TEXT({L}3,"ddd"),1)'); x.fill = HEAD_FILL; x.font = Font(color="FFFFFF", size=8); x.alignment = Alignment(horizontal="center")
        ws.cell(row=3, column=k).font = Font(size=8); ws.cell(row=3, column=k).alignment = Alignment(horizontal="center")
    for r in range(R1, last(n) + 1):
        for h in ["Job ID", "Task", "Assigned to", "Start", "Duration (workdays)", "% Complete", "Notes"]:
            inp(ws, r, ci(n, h), fmt={"Start": DATEF, "% Complete": PCT}.get(h))
        g = lambda h: f"{c[h]}{r}"
        fx(ws, r, ci(n, "End"), f'=IF(OR({g("Start")}="",{g("Duration (workdays)")}=""),"",WORKDAY({g("Start")},MAX({g("Duration (workdays)")},1)-1))', DATEF)
        fx(ws, r, ci(n, "Status"), f'=IF({g("Start")}="","",IF({g("% Complete")}>=1,"Done",IF(TODAY()>{g("End")},"Late",IF(TODAY()>={g("Start")},"In progress","Not started"))))')
        fx(ws, r, ci(n, "Check"), f'=IF({g("Task")}="","",IF(AND({g("Job ID")}<>"",{unknown_job(g("Job ID"))}),"Unknown Job ID",IF({g("Start")}="","Missing start",IF({g("Duration (workdays)")}="","Missing duration",""))))')
        for k in range(g1, g2 + 1): ws.cell(row=r, column=k).border = Border(left=Side(style="hair", color="DDDDDD"), right=Side(style="hair", color="DDDDDD"))
    grid = f"{get_column_letter(g1)}{R1}:{get_column_letter(g2)}{last(n)}"
    K = get_column_letter(g1); D, E, G_ = c["Start"], c["End"], c["% Complete"]
    ws.conditional_formatting.add(grid, FormulaRule(formula=[f'AND(${D}{R1}<>"",{K}$3>=${D}{R1},{K}$3<=${D}{R1}+(${E}{R1}-${D}{R1}+1)*${G_}{R1}-1)'], fill=BAR_DONE_FILL, stopIfTrue=True))
    ws.conditional_formatting.add(grid, FormulaRule(formula=[f'AND(${D}{R1}<>"",{K}$3>=${D}{R1},{K}$3<=${E}{R1})'], fill=BAR_FILL, stopIfTrue=True))
    ws.conditional_formatting.add(grid, FormulaRule(formula=[f'{K}$3=TODAY()'], fill=TODAY_FILL, stopIfTrue=True))
    ws.conditional_formatting.add(grid, FormulaRule(formula=[f'WEEKDAY({K}$3,2)>5'], fill=GRAY_FILL))
    ws.conditional_formatting.add(f"{K}3:{get_column_letter(g2)}3", FormulaRule(formula=[f'{K}$3=TODAY()'], fill=TODAY_FILL))
    s = f"{c['Status']}{R1}:{c['Status']}{last(n)}"
    ws.conditional_formatting.add(s, CellIsRule(operator="equal", formula=['"Late"'], fill=RED_FILL))
    ws.conditional_formatting.add(s, CellIsRule(operator="equal", formula=['"In progress"'], fill=YEL_FILL))
    ws.conditional_formatting.add(s, CellIsRule(operator="equal", formula=['"Done"'], fill=GRN_FILL))
    chk_fill(ws, n)
    dv_list(ws, f"={JOB_IDS}", f"A{R1}:A{last(n)}"); dv_list(ws, f"={CREW_NAMES}", f"{c['Assigned to']}{R1}:{c['Assigned to']}{last(n)}")
    print_setup(ws, title_rows="3:4")
    return ws

def build_timesheet(wb):
    ws = wb.create_sheet("Timesheet"); n = "Timesheet"; c = C[n]
    title(ws, "Timesheet", "Crew hours by day and job. Rate Used comes from Vendors & Crew unless you type an override. Log crew labor HERE (or in Job Costs), never both.", TAB["input"])
    header(ws, 4, H[n], [12, 20, 11, 10, 8, 12, 11, 13, 11, 13, 12, 13, 22, 26])
    ws.freeze_panes = "C5"
    for r in range(R1, last(n) + 1):
        for h in ["Date", "Worker", "Job ID", "Cost Code", "Hours", "Cost Rate Override", "Notes"]:
            inp(ws, r, ci(n, h), fmt={"Date": DATEF, "Cost Rate Override": MONEY, "Hours": HOURS}.get(h))
        g = lambda h: f"{c[h]}{r}"; W = g("Worker")
        fx(ws, r, ci(n, "Rate Used"), f'=IF({W}="","",IF({g("Cost Rate Override")}<>"",{g("Cost Rate Override")},IFERROR(VLOOKUP({W},{CREW_TABLE},3,FALSE),0)))', MONEY)
        fx(ws, r, ci(n, "Labor Cost (incl. burden)"), f'=IF({W}="","",{g("Hours")}*{g("Rate Used")}*(1+{S["burden"]}))', MONEY)
        fx(ws, r, ci(n, "Bill Rate"), f'=IF({W}="","",IFERROR(VLOOKUP({W},{CREW_TABLE},4,FALSE),0))', MONEY)
        fx(ws, r, ci(n, "Billable Value"), f'=IF({W}="","",{g("Hours")}*{g("Bill Rate")})', MONEY)
        fx(ws, r, ci(n, "Week Ending"), f'=IF({g("Date")}="","",{g("Date")}-WEEKDAY({g("Date")},2)+7)', DATEF)
        fx(ws, r, ci(n, "Category"), f'=IF({W}="","",IF({g("Cost Code")}="","Labor",IFERROR(VLOOKUP({g("Cost Code")},{CODE_TABLE},3,FALSE),"Labor")))')
        fx(ws, r, ci(n, "Check"), f'=IF({W}="","",IF(COUNTIF({CREW_NAMES},{W})=0,"Worker not in Crew table",IF({g("Job ID")}="","Missing Job ID",IF({unknown_job(g("Job ID"))},"Unknown Job ID",IF({g("Hours")}="","Missing hours",IF({g("Date")}="","Missing date",""))))))')
    tr = last(n) + 1
    ws.cell(row=tr, column=1, value="TOTAL")
    for h in ["Hours", "Labor Cost (incl. burden)", "Billable Value"]:
        L = c[h]; fx(ws, tr, ci(n, h), f"=SUM({L}{R1}:{L}{last(n)})", HOURS if h == "Hours" else MONEY, True)
    style_range(ws, tr, tr, 1, len(H[n]), fill=TOTAL_FILL, font=BOLD)
    dv_list(ws, f"={CREW_NAMES}", f"{c['Worker']}{R1}:{c['Worker']}{last(n)}", "Workers come from the Crew table on Vendors & Crew.")
    dv_list(ws, f"={JOB_IDS}", f"{c['Job ID']}{R1}:{c['Job ID']}{last(n)}", "Job IDs come from the Jobs sheet.")
    dv_list(ws, f"={CODE_RANGE}", f"{c['Cost Code']}{R1}:{c['Cost Code']}{last(n)}")
    chk_fill(ws, n)
    # Weekly summary (columns P..S); header shares row 4 so data rows keep their height
    ws.column_dimensions["O"].width = 3
    for col_, w in zip("PQRSTU", [22, 12, 14, 14, 12, 3]): ws.column_dimensions[col_].width = w
    ws["P2"] = "WEEKLY PAYROLL SUMMARY"; ws["P2"].font = H2
    ws["P3"] = "Week ending (Sunday)"; ws["P3"].font = BOLD
    inp(ws, 3, 17, "=TODAY()-WEEKDAY(TODAY(),2)+7", DATEF)
    ws["R3"] = "← type any Sunday"; ws["R3"].font = NOTE
    header(ws, 4, ["Worker", "Hours", "Labor cost", "Billable", "OT hours (>40)"], start_col=16)
    for i in range(N_CREW):
        r = 5 + i; cr = CREW_R1 + i
        fx(ws, r, 16, f'=IF({cell("Vendors & Crew", CC["Worker"], cr)}="","",{cell("Vendors & Crew", CC["Worker"], cr)})')
        fx(ws, r, 17, f'=IF(P{r}="","",SUMIFS({rr(n,"Hours")},{rr(n,"Worker")},P{r},{rr(n,"Week Ending")},$Q$3))', HOURS)
        fx(ws, r, 18, f'=IF(P{r}="","",SUMIFS({rr(n,"Labor Cost (incl. burden)")},{rr(n,"Worker")},P{r},{rr(n,"Week Ending")},$Q$3))', MONEY)
        fx(ws, r, 19, f'=IF(P{r}="","",SUMIFS({rr(n,"Billable Value")},{rr(n,"Worker")},P{r},{rr(n,"Week Ending")},$Q$3))', MONEY)
        fx(ws, r, 20, f'=IF(P{r}="","",MAX(0,Q{r}-40))', HOURS)
    t = 5 + N_CREW
    ws.cell(row=t, column=16, value="TOTAL")
    fx(ws, t, 17, f"=SUM(Q5:Q{t-1})", HOURS, True); fx(ws, t, 18, f"=SUM(R5:R{t-1})", MONEY, True); fx(ws, t, 19, f"=SUM(S5:S{t-1})", MONEY, True); fx(ws, t, 20, f"=SUM(T5:T{t-1})", HOURS, True)
    style_range(ws, t, t, 16, 20, fill=TOTAL_FILL, font=BOLD)
    ws.conditional_formatting.add(f"T5:T{t-1}", FormulaRule(formula=['AND(T5<>"",T5>0)'], fill=YEL_FILL))
    print_setup(ws, title_rows="4:4")
    return ws

def build_job_costs(wb):
    ws = wb.create_sheet("Job Costs"); n = "Job Costs"; c = C[n]
    title(ws, "Job Costs (actuals)", "Materials, subcontractors, equipment, permits and other real costs. Qty × Unit Cost = Amount. Category comes from the cost code. Crew hours belong on the Timesheet.", TAB["input"])
    header(ws, 4, H[n], [12, 12, 10, 14, 22, 34, 11, 14, 14, 8, 14, 16, 22, 26])
    ws.freeze_panes = "C5"
    for r in range(R1, last(n) + 1):
        for h in ["Date", "Job ID", "Cost Code", "Vendor", "Description", "Qty or Hours", "Unit Cost / Rate", "Paid?", "Payment Method", "Receipt / Invoice #", "Notes"]:
            inp(ws, r, ci(n, h), fmt={"Date": DATEF, "Unit Cost / Rate": MONEY}.get(h))
        g = lambda h: f"{c[h]}{r}"; J = g("Job ID")
        fx(ws, r, ci(n, "Category"), f'=IF({g("Cost Code")}="","",IFERROR(VLOOKUP({g("Cost Code")},{CODE_TABLE},3,FALSE),"Other"))')
        fx(ws, r, ci(n, "Amount"), f'=IF({J}="","",{g("Qty or Hours")}*{g("Unit Cost / Rate")})', MONEY)
        fx(ws, r, ci(n, "Check"), f'=IF({J}="","",IF({unknown_job(J)},"Unknown Job ID",IF({g("Cost Code")}="","Missing cost code",IF(COUNTIF({CODE_RANGE},{g("Cost Code")})=0,"Unknown cost code",IF({g("Date")}="","Missing date",IF(OR({g("Qty or Hours")}="",{g("Unit Cost / Rate")}=""),"Missing qty or cost",""))))))')
    tr = last(n) + 1
    ws.cell(row=tr, column=ci(n, "Description"), value="TOTAL")
    fx(ws, tr, ci(n, "Amount"), f"=SUM({c['Amount']}{R1}:{c['Amount']}{last(n)})", MONEY, True)
    style_range(ws, tr, tr, 1, len(H[n]), fill=TOTAL_FILL, font=BOLD)
    dv_list(ws, f"={JOB_IDS}", f"{c['Job ID']}{R1}:{c['Job ID']}{last(n)}", "Job IDs come from the Jobs sheet.")
    dv_list(ws, f"={CODE_RANGE}", f"{c['Cost Code']}{R1}:{c['Cost Code']}{last(n)}", "Cost code sets the category for Budget vs Actual.")
    dv_list(ws, f"={VEND_NAMES}", f"{c['Vendor']}{R1}:{c['Vendor']}{last(n)}", "Vendors come from Vendors & Crew (drives Paid YTD and 1099 flags).")
    dv_list(ws, '"Yes,No"', f"{c['Paid?']}{R1}:{c['Paid?']}{last(n)}")
    dv_list(ws, f"={PAYM_RANGE}", f"{c['Payment Method']}{R1}:{c['Payment Method']}{last(n)}")
    ws.conditional_formatting.add(f"{c['Paid?']}{R1}:{c['Paid?']}{last(n)}", CellIsRule(operator="equal", formula=['"No"'], fill=YEL_FILL))
    chk_fill(ws, n); print_setup(ws, title_rows="4:4")
    return ws

def build_mileage(wb):
    ws = wb.create_sheet("Mileage"); n = "Mileage"; c = C[n]
    title(ws, "Mileage log", "Job-related driving. Rate comes from Settings. Amount counts toward the job's actual cost under the cost code you pick (default 22 Fuel & mileage).", TAB["input"])
    header(ws, 4, H[n], [12, 12, 10, 24, 30, 26, 8, 10, 12, 12, 20, 24])
    ws.freeze_panes = "C5"
    for r in range(R1, last(n) + 1):
        for h in ["Date", "Job ID", "Cost Code", "From", "To", "Purpose", "Miles", "Notes"]:
            inp(ws, r, ci(n, h), fmt={"Date": DATEF, "Miles": '#,##0.0'}.get(h))
        g = lambda h: f"{c[h]}{r}"; J = g("Job ID")
        fx(ws, r, ci(n, "Rate"), f'=IF({J}="","",{S["mileage"]})', MONEY)
        fx(ws, r, ci(n, "Amount"), f'=IF({J}="","",{g("Miles")}*{g("Rate")})', MONEY)
        fx(ws, r, ci(n, "Category"), f'=IF({J}="","",IF({g("Cost Code")}="","Other",IFERROR(VLOOKUP({g("Cost Code")},{CODE_TABLE},3,FALSE),"Other")))')
        fx(ws, r, ci(n, "Check"), f'=IF({J}="","",IF({unknown_job(J)},"Unknown Job ID",IF({g("Miles")}="","Missing miles",IF({g("Date")}="","Missing date",""))))')
    tr = last(n) + 1
    ws.cell(row=tr, column=1, value="TOTAL")
    fx(ws, tr, ci(n, "Miles"), f"=SUM({c['Miles']}{R1}:{c['Miles']}{last(n)})", '#,##0.0', True)
    fx(ws, tr, ci(n, "Amount"), f"=SUM({c['Amount']}{R1}:{c['Amount']}{last(n)})", MONEY, True)
    style_range(ws, tr, tr, 1, len(H[n]), fill=TOTAL_FILL, font=BOLD)
    dv_list(ws, f"={JOB_IDS}", f"{c['Job ID']}{R1}:{c['Job ID']}{last(n)}")
    dv_list(ws, f"={CODE_RANGE}", f"{c['Cost Code']}{R1}:{c['Cost Code']}{last(n)}")
    chk_fill(ws, n); print_setup(ws, title_rows="4:4")
    return ws

def build_daily_log(wb):
    ws = wb.create_sheet("Daily Log"); n = "Daily Log"; c = C[n]
    title(ws, "Daily Log", "A dated record per job: weather, crew, work done, problems, inspections. Hours pull from the Timesheet for that date and job. Useful for change-order and delay disputes.", TAB["input"])
    header(ws, 4, H[n], [12, 12, 14, 10, 12, 44, 34, 26, 20])
    ws.freeze_panes = "C5"
    for r in range(R1, last(n) + 1):
        for h in ["Date", "Job ID", "Weather", "Crew on site", "Work performed", "Issues / delays", "Inspections / visitors"]:
            x = inp(ws, r, ci(n, h), fmt=DATEF if h == "Date" else None); x.alignment = Alignment(wrap_text=True, vertical="top")
        g = lambda h: f"{c[h]}{r}"; J = g("Job ID")
        fx(ws, r, ci(n, "Hours logged (Timesheet)"), f'=IF(OR({J}="",{g("Date")}=""),"",SUMIFS({rr("Timesheet","Hours")},{rr("Timesheet","Job ID")},{J},{rr("Timesheet","Date")},{g("Date")}))', HOURS).alignment = Alignment(vertical="top")
        fx(ws, r, ci(n, "Check"), f'=IF({J}="","",IF({unknown_job(J)},"Unknown Job ID",IF({g("Date")}="","Missing date","")))').alignment = Alignment(vertical="top")
    dv_list(ws, f"={JOB_IDS}", f"{c['Job ID']}{R1}:{c['Job ID']}{last(n)}")
    dv_list(ws, '"Sunny,Cloudy,Rain,Snow,Wind,Extreme heat,Extreme cold"', f"{c['Weather']}{R1}:{c['Weather']}{last(n)}")
    chk_fill(ws, n); print_setup(ws, title_rows="4:4")
    return ws

def build_change_orders(wb):
    ws = wb.create_sheet("Change Orders"); n = "Change Orders"; c = C[n]
    title(ws, "Change Orders", "Approved change orders add to the Revised Contract and the job budget. Pending and Rejected are ignored. Print a signature-ready form on the CO Form sheet.", TAB["input"])
    header(ws, 4, H[n], [8, 12, 12, 40, 12, 10, 16, 16, 14, 12, 14, 20, 26])
    ws.freeze_panes = "C5"
    ws["A3"] = f'="Suggested next CO #: "&(MAX(A{R1}:A{last(n)})+1)'; ws["A3"].font = NOTE
    for r in range(R1, last(n) + 1):
        for h in ["CO #", "Job ID", "Date", "Description", "Status", "Markup %", "Added Cost (budget)", "Schedule impact (days)", "Client approved date", "Notes"]:
            inp(ws, r, ci(n, h), fmt={"Date": DATEF, "Markup %": PCT, "Added Cost (budget)": MONEY, "Client approved date": DATEF}.get(h))
        g = lambda h: f"{c[h]}{r}"; J = g("Job ID")
        fx(ws, r, ci(n, "Price to Client"), f'=IF({J}="","",{g("Added Cost (budget)")}*(1+IF({g("Markup %")}="",{cell("Settings","B",MARKUP_R2)},{g("Markup %")})))', MONEY)
        fx(ws, r, ci(n, "Added Profit"), f'=IF({J}="","",{g("Price to Client")}-{g("Added Cost (budget)")})', MONEY)
        fx(ws, r, ci(n, "Check"), f'=IF({J}="","",IF({unknown_job(J)},"Unknown Job ID",IF(COUNTIF({CO_NUMS},{g("CO #")})>1,"Duplicate CO #",IF({g("Status")}="","Missing status",IF(AND({g("Status")}="Approved",{g("Client approved date")}=""),"Approved without client date","")))))')
    dv_list(ws, f"={JOB_IDS}", f"{c['Job ID']}{R1}:{c['Job ID']}{last(n)}")
    dv_list(ws, '"Pending,Approved,Rejected"', f"{c['Status']}{R1}:{c['Status']}{last(n)}")
    s = f"{c['Status']}{R1}:{c['Status']}{last(n)}"
    ws.conditional_formatting.add(s, CellIsRule(operator="equal", formula=['"Approved"'], fill=GRN_FILL))
    ws.conditional_formatting.add(s, CellIsRule(operator="equal", formula=['"Pending"'], fill=YEL_FILL))
    chk_fill(ws, n); print_setup(ws, title_rows="4:4")
    return ws

def build_co_form(wb):
    ws = wb.create_sheet("CO Form"); ws.sheet_properties.tabColor = TAB["report"]; ws.sheet_view.showGridLines = False
    for col_, w in zip("ABCDEFGHI", [9, 40, 14, 14, 18, 14, 3, 12, 20]): ws.column_dimensions[col_].width = w
    ws["A1"] = "CHANGE ORDER"; ws["A1"].font = Font(bold=True, size=22, color=NAVY)
    ws["H2"] = "CO #"; ws["H2"].font = BOLD; inp(ws, 2, 9).alignment = Alignment(horizontal="left"); dv_list(ws, f"={CO_NUMS}", "I2", "Pick a change order number; the form fills in.")
    ws["H3"] = "← pick a change order; the form at left fills in. Columns H–I never print."; ws["H3"].font = NOTE
    M = f"MATCH($I$2,{CO_NUMS},0)"
    def cof(h): return f'IFERROR(INDEX({rr("Change Orders", h)},{M}),"")'
    JOB = cof("Job ID"); JR = f"MATCH({JOB},{JOB_IDS},0)"
    def jobf(h): return f'IFERROR(INDEX({rr("Jobs", h)},{JR}),"")'
    CLIENT = jobf("Client")
    def clif(h): return f'IFERROR(INDEX({rr("Clients", h)},MATCH({CLIENT},{CLIENT_NAMES},0)),"")'
    ws["B4"] = "CONTRACTOR"; ws["E4"] = "CLIENT"
    for a in ("B4", "E4"): ws[a].font = Font(bold=True, color=TEAL)
    ws["B5"] = f"={S['company']}&\"\""; ws["B5"].font = BOLD; ws["B6"] = f"={S['address']}&\"\""; ws["B7"] = f"={S['phone']}&\"\""
    ws["E5"] = f"={CLIENT}"; ws["E5"].font = BOLD; ws["E6"] = f"={clif('Contact person')}"; ws["E7"] = f"={clif('Billing address')}"
    rows = [("Job", f'=IF($I$2="","",{JOB}&"  —  "&{jobf("Job address / description")})'), ("Change order date", f"={cof('Date')}"),
            ("Description of change", f"={cof('Description')}"), ("Schedule impact (days)", f"={cof('Schedule impact (days)')}"),
            ("Revised target completion", f'=IF(OR($I$2="",{jobf("Target end")}=""),"",{jobf("Target end")}+IF({cof("Schedule impact (days)")}="",0,{cof("Schedule impact (days)")}))'),
            ("Price of this change", f"={cof('Price to Client')}"), ("Original contract value", f'=IF({jobf("Contract Value (signed)")}="",{jobf("Estimate Price")},{jobf("Contract Value (signed)")})'),
            ("Previously approved change orders", f'=SUMIFS({rr("Change Orders","Price to Client")},{rr("Change Orders","Job ID")},{JOB},{rr("Change Orders","Status")},"Approved")-IF({cof("Status")}="Approved",{cof("Price to Client")},0)'),
            ("Revised contract total after this change", '=IF($I$2="","",E15+E16+E14)'), ("Status", f"={cof('Status')}")]
    for i, (k, v) in enumerate(rows, start=9):
        ws.cell(row=i, column=2, value=k).font = BOLD
        x = ws.cell(row=i, column=5, value=v); x.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=i, start_column=5, end_row=i, end_column=6)
        if k in ("Price of this change", "Original contract value", "Previously approved change orders", "Revised contract total after this change"): x.number_format = MONEY
        if k in ("Change order date", "Revised target completion"): x.number_format = DATEF
        if k == "Description of change": ws.row_dimensions[i].height = 60
        if k == "Job": ws.row_dimensions[i].height = 30
        style_range(ws, i, i, 2, 6); ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=4)
        ws.cell(row=i, column=2).alignment = Alignment(vertical="top")
    style_range(ws, 17, 17, 2, 6, fill=TOTAL_FILL, font=BOLD)
    ws["B20"] = "Terms"; ws["B20"].font = Font(bold=True, color=TEAL)
    x = ws.cell(row=21, column=2, value=f"={S['co_terms']}"); x.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("B21:F21"); ws.row_dimensions[21].height = 50
    ws["B23"] = "Client signature: ______________________________    Date: ____________"
    ws["B24"] = "Contractor signature: __________________________    Date: ____________"
    print_setup(ws, landscape=False, area="A1:F24")
    return ws

def build_invoices(wb):
    ws = wb.create_sheet("Invoices"); n = "Invoices"; c = C[n]
    title(ws, "Invoices", "One row per invoice sent. Paid to Date pulls from Payments by Invoice #. Status and aging use today's date and the grace days in Settings.", TAB["input"])
    header(ws, 4, H[n], [14, 12, 12, 12, 14, 30, 14, 14, 11, 10, 11, 22, 20, 9, 26])
    ws.freeze_panes = "C5"
    ws["A3"] = f'="Suggested next: INV-"&{S["year"]}&"-"&TEXT(COUNTA(A{R1}:A{last(n)})+1,"000")'; ws["A3"].font = NOTE
    for r in range(R1, last(n) + 1):
        for h in ["Invoice #", "Job ID", "Invoice Date", "Due Date", "Amount", "Description", "Notes"]:
            inp(ws, r, ci(n, h), fmt={"Invoice Date": DATEF, "Due Date": DATEF, "Amount": MONEY}.get(h))
        g = lambda h: f"{c[h]}{r}"; A = g("Invoice #")
        fx(ws, r, ci(n, "Paid to Date"), f'=IF({A}="","",SUMIFS({rr("Payments","Amount")},{rr("Payments","Invoice #")},{A}))', MONEY)
        fx(ws, r, ci(n, "Balance"), f'=IF({A}="","",{g("Amount")}-{g("Paid to Date")})', MONEY)
        fx(ws, r, ci(n, "Status"), f'=IF({A}="","",IF({g("Balance")}<=0,"Paid",IF(AND({g("Due Date")}<>"",TODAY()>{g("Due Date")}+{S["grace"]}),"Overdue",IF({g("Paid to Date")}>0,"Partial","Open"))))')
        fx(ws, r, ci(n, "Days Overdue"), f'=IF({A}="","",IF({g("Status")}="Overdue",TODAY()-{g("Due Date")},0))')
        fx(ws, r, ci(n, "Aging Bucket"), f'=IF({A}="","",IF({g("Balance")}<=0,"Paid",IF(OR({g("Due Date")}="",TODAY()<={g("Due Date")}),"Current",IF(TODAY()-{g("Due Date")}<=30,"1-30",IF(TODAY()-{g("Due Date")}<=60,"31-60",IF(TODAY()-{g("Due Date")}<=90,"61-90","90+"))))))')
        fx(ws, r, ci(n, "Check"), f'=IF({A}="","",IF(COUNTIF({INV_NUMS},{A})>1,"Duplicate invoice #",IF({g("Job ID")}="","Missing Job ID",IF({unknown_job(g("Job ID"))},"Unknown Job ID",IF({g("Due Date")}="","Missing due date",IF({g("Amount")}="","Missing amount",""))))))')
        fx(ws, r, ci(n, "Client"), f'=IF({A}="","",IFERROR(INDEX({rr("Jobs","Client")},MATCH({g("Job ID")},{JOB_IDS},0)),""))')
        fx(ws, r, ci(n, "Stmt Line #"), f'=IF(AND({A}<>"",{g("Client")}<>"",{g("Client")}={STATEMENT_SEL}),COUNTIF({c["Client"]}${R1}:{c["Client"]}{r},{STATEMENT_SEL}),"")')
    tr = last(n) + 1
    ws.cell(row=tr, column=1, value="TOTAL")
    for h in ["Amount", "Paid to Date", "Balance"]:
        L = c[h]; fx(ws, tr, ci(n, h), f"=SUM({L}{R1}:{L}{last(n)})", MONEY, True)
    style_range(ws, tr, tr, 1, len(H[n]), fill=TOTAL_FILL, font=BOLD)
    dv_list(ws, f"={JOB_IDS}", f"{c['Job ID']}{R1}:{c['Job ID']}{last(n)}", "Job IDs come from the Jobs sheet.")
    s = f"{c['Status']}{R1}:{c['Status']}{last(n)}"
    ws.conditional_formatting.add(s, CellIsRule(operator="equal", formula=['"Overdue"'], fill=RED_FILL))
    ws.conditional_formatting.add(s, CellIsRule(operator="equal", formula=['"Paid"'], fill=GRN_FILL))
    ws.conditional_formatting.add(s, CellIsRule(operator="equal", formula=['"Partial"'], fill=YEL_FILL))
    chk_fill(ws, n); print_setup(ws, title_rows="4:4")
    return ws

def build_invoice_print(wb):
    ws = wb.create_sheet("Invoice Print"); ws.sheet_properties.tabColor = TAB["report"]; ws.sheet_view.showGridLines = False
    for col_, w in zip("ABCDEFGHI", [10, 40, 14, 14, 18, 14, 3, 12, 20]): ws.column_dimensions[col_].width = w
    ws["A1"] = "INVOICE"; ws["A1"].font = Font(bold=True, size=22, color=NAVY)
    ws["H2"] = "Invoice #"; ws["H2"].font = BOLD; inp(ws, 2, 9); dv_list(ws, f"={INV_NUMS}", "I2", "Pick an invoice number; the invoice fills in.")
    ws["H3"] = "← pick an invoice; print or save as PDF. Columns H–I never print."; ws["H3"].font = NOTE
    M = f"MATCH($I$2,{INV_NUMS},0)"
    def invf(h): return f'IFERROR(INDEX({rr("Invoices", h)},{M}),"")'
    JOB = invf("Job ID"); JR = f"MATCH({JOB},{JOB_IDS},0)"
    def jobf(h): return f'IFERROR(INDEX({rr("Jobs", h)},{JR}),"")'
    CLIENT = jobf("Client")
    def clif(h): return f'IFERROR(INDEX({rr("Clients", h)},MATCH({CLIENT},{CLIENT_NAMES},0)),"")'
    ws["B4"] = "FROM"; ws["E4"] = "BILL TO"
    for a in ("B4", "E4"): ws[a].font = Font(bold=True, color=TEAL)
    ws["B5"] = f"={S['company']}&\"\""; ws["B5"].font = BOLD; ws["B6"] = f"={S['address']}&\"\""; ws["B7"] = f"={S['phone']}&\"\""; ws["B8"] = f'=IF({S["license"]}="","","License # "&{S["license"]})'
    ws["E5"] = f"={CLIENT}"; ws["E5"].font = BOLD; ws["E6"] = f"={clif('Contact person')}"; ws["E7"] = f"={clif('Billing address')}"; ws["E8"] = f"={clif('Email')}"
    ws["B10"] = "Invoice number"; ws["E10"] = "=$I$2&\"\""
    ws["B11"] = "Invoice date"; ws["E11"] = f"={invf('Invoice Date')}"; ws["E11"].number_format = DATEF
    ws["B12"] = "Due date"; ws["E12"] = f"={invf('Due Date')}"; ws["E12"].number_format = DATEF
    ws["B13"] = "Job"; ws["E13"] = f'=IF($I$2="","",{JOB}&"  —  "&{jobf("Job address / description")})'
    for r in range(10, 14):
        ws.cell(row=r, column=2).font = BOLD; ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=6); style_range(ws, r, r, 2, 6); ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
    ws["E13"].alignment = Alignment(wrap_text=True, vertical="top"); ws["B13"].alignment = Alignment(vertical="top"); ws.row_dimensions[13].height = 30
    header(ws, 15, [None, "Description", "", "", "Amount", ""])
    ws.merge_cells("B15:D15"); ws.merge_cells("E15:F15")
    x = ws.cell(row=19, column=2, value='=IF(AND($I$2<>"",E16<>"",E16>0,E18=0),"PAID IN FULL — thank you","")'); x.font = Font(bold=True, color="1E7B34")
    ws["B16"] = f'=IF({invf("Description")}="","Services per contract",{invf("Description")})'; ws["E16"] = f"={invf('Amount')}"
    ws["B17"] = "Less: payments received on this invoice"; ws["E17"] = f'=IF($I$2="","",-{invf("Paid to Date")})'
    ws["B18"] = "BALANCE DUE"; ws["E18"] = f"={invf('Balance')}"
    for r in (16, 17, 18): ws.cell(row=r, column=5).number_format = MONEY; style_range(ws, r, r, 2, 6); ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4); ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=6)
    style_range(ws, 18, 18, 2, 6, fill=TOTAL_FILL, font=BOLD)
    ws["B20"] = "Contract summary"; ws["B20"].font = Font(bold=True, color=TEAL)
    summ = [("Revised contract value", f"={jobf('Revised Contract')}"), ("Invoiced to date (all invoices)", f"={jobf('Invoiced')}"),
            ("Paid to date (all invoices)", f"={jobf('Collected')}"), ("Remaining to invoice", f'=IF($I$2="","",{jobf("Revised Contract")}-{jobf("Invoiced")})')]
    for i, (k, v) in enumerate(summ, start=21):
        ws.cell(row=i, column=2, value=k); ws.cell(row=i, column=5, value=v).number_format = MONEY; style_range(ws, i, i, 2, 6); ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=4); ws.merge_cells(start_row=i, start_column=5, end_row=i, end_column=6)
    ws["B26"] = "Payment instructions"; ws["B26"].font = Font(bold=True, color=TEAL)
    x = ws.cell(row=27, column=2, value=f"={S['invoice_instructions']}"); x.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("B27:F27"); ws.row_dimensions[27].height = 40
    ws["B29"] = f"={S['invoice_footer']}"; ws["B29"].font = NOTE
    ws["B31"] = "REMITTANCE — detach and return with payment"; ws["B31"].font = Font(bold=True, color=TEAL)
    for i, (k, v, f) in enumerate([("Payable to", f"={S['company']}&\"\"", None), ("Invoice number", "=$I$2&\"\"", None), ("Balance due", f"={invf('Balance')}", MONEY), ("Amount enclosed", "$ ______________", None)], start=32):
        ws.cell(row=i, column=2, value=k).font = BOLD; x = ws.cell(row=i, column=5, value=v)
        if f: x.number_format = f
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=4); ws.merge_cells(start_row=i, start_column=5, end_row=i, end_column=6); style_range(ws, i, i, 2, 6)
    print_setup(ws, landscape=False, area="A1:F35")
    return ws

def build_payments(wb):
    ws = wb.create_sheet("Payments"); n = "Payments"; c = C[n]
    title(ws, "Payments received", "Every client payment. Invoice # links it to an invoice; Job ID drives the job's Collected total and the cash P&L.", TAB["input"])
    header(ws, 4, H[n], [12, 12, 14, 14, 14, 18, 24, 30])
    ws.freeze_panes = "C5"
    for r in range(R1, last(n) + 1):
        for h in ["Date", "Job ID", "Invoice #", "Amount", "Method", "Reference / Check #", "Notes"]:
            inp(ws, r, ci(n, h), fmt={"Date": DATEF, "Amount": MONEY}.get(h))
        g = lambda h: f"{c[h]}{r}"; J = g("Job ID")
        fx(ws, r, ci(n, "Check"), f'=IF(AND({J}="",{g("Amount")}=""),"",IF({J}="","Missing Job ID",IF({unknown_job(J)},"Unknown Job ID",IF(AND({g("Invoice #")}<>"",COUNTIF({INV_NUMS},{g("Invoice #")})=0),"Invoice # not in Invoices",IF({g("Date")}="","Missing date","")))))')
    tr = last(n) + 1
    ws.cell(row=tr, column=1, value="TOTAL"); fx(ws, tr, ci(n, "Amount"), f"=SUM({c['Amount']}{R1}:{c['Amount']}{last(n)})", MONEY, True)
    style_range(ws, tr, tr, 1, len(H[n]), fill=TOTAL_FILL, font=BOLD)
    dv_list(ws, f"={JOB_IDS}", f"{c['Job ID']}{R1}:{c['Job ID']}{last(n)}", "Job IDs come from the Jobs sheet.")
    dv_list(ws, f"={INV_NUMS}", f"{c['Invoice #']}{R1}:{c['Invoice #']}{last(n)}", "Link the payment to an invoice so its balance updates.")
    dv_list(ws, f"={PAYM_RANGE}", f"{c['Method']}{R1}:{c['Method']}{last(n)}")
    chk_fill(ws, n); print_setup(ws, title_rows="4:4")
    return ws


def build_client_statement(wb):
    ws = wb.create_sheet("Client Statement"); ws.sheet_properties.tabColor = TAB["report"]; ws.sheet_view.showGridLines = False
    for col_, w in zip("ABCDEFGHIJKL", [9, 15, 42, 12, 12, 14, 14, 14, 11, 3, 12, 24]): ws.column_dimensions[col_].width = w
    ws["A1"] = "STATEMENT"; ws["A1"].font = Font(bold=True, size=22, color=NAVY)
    ws["K2"] = "Client"; ws["K2"].font = BOLD; inp(ws, 2, 12); dv_list(ws, f"={CLIENT_NAMES}", "L2", "Pick a client; all their invoices list below.")
    ws["K3"] = "← pick a client; print or save as PDF. Columns K–L never print."; ws["K3"].font = NOTE
    def clif(h): return f'IFERROR(INDEX({rr("Clients", h)},MATCH($L$2,{CLIENT_NAMES},0)),"")'
    ws["B4"] = "FROM"; ws["F4"] = "TO"
    for a in ("B4", "F4"): ws[a].font = Font(bold=True, color=TEAL)
    ws["B5"] = f"={S['company']}&\"\""; ws["B5"].font = BOLD; ws["B6"] = f"={S['address']}&\"\""; ws["B7"] = f"={S['phone']}&\"\""; ws["B8"] = f'=IF({S["license"]}="","","License # "&{S["license"]})'
    ws["F5"] = "=$L$2&\"\""; ws["F5"].font = BOLD; ws["F6"] = f"={clif('Contact person')}"; ws["F7"] = f"={clif('Billing address')}"; ws["F8"] = f"={clif('Email')}"
    for r_ in range(5, 9): ws.merge_cells(start_row=r_, start_column=2, end_row=r_, end_column=4); ws.merge_cells(start_row=r_, start_column=6, end_row=r_, end_column=9)
    ws["B10"] = "Statement date"; ws["B10"].font = BOLD; ws["D10"] = "=TODAY()"; ws["D10"].number_format = DATEF
    ws["F10"] = "TOTAL BALANCE DUE"; ws["F10"].font = BOLD
    fx(ws, 10, 8, f'=SUMIFS({rr("Invoices","Balance")},{rr("Invoices","Client")},$L$2)', MONEY, True); ws.merge_cells("H10:I10")
    style_range(ws, 10, 10, 6, 9, fill=TOTAL_FILL, font=BOLD)
    header(ws, 12, [None, "Invoice #", "Job", "Invoice date", "Due date", "Amount", "Paid", "Balance", "Status"])
    ws.cell(row=12, column=1).fill = PatternFill(None)
    SL = rr("Invoices", "Stmt Line #")
    for i in range(1, N_STMT + 1):
        r = 12 + i; m = f"MATCH({i},{SL},0)"
        fx(ws, r, 2, f'=IFERROR(INDEX({rr("Invoices","Invoice #")},{m}),"")')
        fx(ws, r, 3, f'=IFERROR(INDEX({rr("Invoices","Job ID")},{m})&"  "&INDEX({rr("Invoices","Description")},{m}),"")')
        fx(ws, r, 4, f'=IFERROR(INDEX({rr("Invoices","Invoice Date")},{m}),"")', DATEF)
        fx(ws, r, 5, f'=IFERROR(INDEX({rr("Invoices","Due Date")},{m}),"")', DATEF)
        fx(ws, r, 6, f'=IFERROR(INDEX({rr("Invoices","Amount")},{m}),"")', MONEY)
        fx(ws, r, 7, f'=IFERROR(INDEX({rr("Invoices","Paid to Date")},{m}),"")', MONEY)
        fx(ws, r, 8, f'=IFERROR(INDEX({rr("Invoices","Balance")},{m}),"")', MONEY)
        fx(ws, r, 9, f'=IFERROR(INDEX({rr("Invoices","Status")},{m}),"")')
    t = 13 + N_STMT
    ws.cell(row=t, column=2, value="Totals"); ws.cell(row=t, column=3, value=f'=IF(COUNTIF({SL},">0")>{N_STMT},"+ more invoices not shown","")').font = NOTE
    for cc, L in ((6, "F"), (7, "G"), (8, "H")): fx(ws, t, cc, f"=SUM({L}13:{L}{t-1})", MONEY, True)
    style_range(ws, t, t, 2, 9, fill=TOTAL_FILL, font=BOLD)
    ws.conditional_formatting.add(f"I13:I{t-1}", CellIsRule(operator="equal", formula=['"Overdue"'], fill=RED_FILL))
    ws.conditional_formatting.add(f"I13:I{t-1}", CellIsRule(operator="equal", formula=['"Paid"'], fill=GRN_FILL))
    a = t + 2
    ws.cell(row=a, column=2, value="Aging of open balance").font = Font(bold=True, color=TEAL)
    header(ws, a + 1, [None, "Current", "1-30 days", "31-60 days", "61-90 days", "90+ days"])
    ws.cell(row=a + 1, column=1).fill = PatternFill(None)
    for j, b in enumerate(["Current", "1-30", "31-60", "61-90", "90+"], start=2):
        fx(ws, a + 2, j, f'=SUMIFS({rr("Invoices","Balance")},{rr("Invoices","Client")},$L$2,{rr("Invoices","Aging Bucket")},"{b}")', MONEY)
    ws.cell(row=a + 4, column=2, value="Payment instructions").font = Font(bold=True, color=TEAL)
    x = ws.cell(row=a + 5, column=2, value=f"={S['invoice_instructions']}"); x.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=a + 5, start_column=2, end_row=a + 5, end_column=9); ws.row_dimensions[a + 5].height = 40
    ws.cell(row=a + 7, column=2, value=f"={S['invoice_footer']}").font = NOTE
    print_setup(ws, landscape=False, area=f"A1:I{a+7}")
    return ws, {"line_r1": 13, "total_row": t, "aging_row": a + 2, "balance": "H10"}
