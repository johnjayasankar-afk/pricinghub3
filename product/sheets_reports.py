"""Job Profit Report, Reports, Tax Summary, Health Check, Dashboard, Start Here (v3)."""
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.chart import BarChart, Reference

def _axes(ch, fmt):
    """Recent Excel builds hide chart axes unless delete is explicitly False; Sheets ignores the flag."""
    from openpyxl.chart.text import RichText
    from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties, RichTextProperties
    def _lbl(rot=0):
        return RichText(bodyPr=RichTextProperties(rot=rot, vert="horz"), p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(sz=800)), endParaRPr=CharacterProperties(sz=800))])
    ch.x_axis.delete = False; ch.y_axis.delete = False
    ch.y_axis.number_format = fmt; ch.y_axis.txPr = _lbl(); ch.x_axis.txPr = _lbl(-2700000)
    ch.x_axis.tickLblPos = "low"; ch.gapWidth = 60
from layout import *

LC = "Labor Cost (incl. burden)"

def costs_by(job_expr, key_col, key_val):
    return (f'SUMIFS({rr("Job Costs","Amount")},{rr("Job Costs","Job ID")},{job_expr},{rr("Job Costs",key_col)},{key_val})'
            f'+SUMIFS({rr("Timesheet",LC)},{rr("Timesheet","Job ID")},{job_expr},{rr("Timesheet",key_col)},{key_val})'
            f'+SUMIFS({rr("Mileage","Amount")},{rr("Mileage","Job ID")},{job_expr},{rr("Mileage",key_col)},{key_val})')

def costs_all_by(key_col, key_val):
    return (f'SUMIFS({rr("Job Costs","Amount")},{rr("Job Costs",key_col)},{key_val})'
            f'+SUMIFS({rr("Timesheet",LC)},{rr("Timesheet",key_col)},{key_val})'
            f'+SUMIFS({rr("Mileage","Amount")},{rr("Mileage",key_col)},{key_val})')

def costs_between(d1, d2):
    return (f'SUMIFS({rr("Job Costs","Amount")},{rr("Job Costs","Date")},">="&{d1},{rr("Job Costs","Date")},"<="&{d2})'
            f'+SUMIFS({rr("Timesheet",LC)},{rr("Timesheet","Date")},">="&{d1},{rr("Timesheet","Date")},"<="&{d2})'
            f'+SUMIFS({rr("Mileage","Amount")},{rr("Mileage","Date")},">="&{d1},{rr("Mileage","Date")},"<="&{d2})')

def build_job_profit_report(wb):
    ws = wb.create_sheet("Job Profit Report"); ws.sheet_properties.tabColor = TAB["report"]; ws.sheet_view.showGridLines = False
    for col_, w in zip("ABCDEFGHI", [4, 20, 30, 16, 30, 16, 14, 12, 12]): ws.column_dimensions[col_].width = w
    ws["A1"] = "Job Profit Report — budget vs actual"; ws["A1"].font = TITLE
    ws["A2"] = "Pick a Job ID. Budget = estimate lines + approved change-order cost. Actual = Job Costs + Timesheet labor (incl. burden) + Mileage."; ws["A2"].font = NOTE
    ws["B4"] = "Job ID"; ws["B4"].font = BOLD; inp(ws, 4, 3); dv_list(ws, f"={JOB_IDS}", "C4", "Pick the job to report on.")
    JR = f"MATCH($C$4,{JOB_IDS},0)"
    def jobf(h): return f'IFERROR(INDEX({rr("Jobs", h)},{JR}),"")'
    info = [("Client", jobf("Client"), None), ("Status", jobf("Status"), None),
            ("Start → target end", f'IF({jobf("Start date")}="","",TEXT({jobf("Start date")},"yyyy-mm-dd"))&"  →  "&IF({jobf("Target end")}="","",TEXT({jobf("Target end")},"yyyy-mm-dd"))', None),
            ("% complete (manual)", jobf("% Complete (manual)"), PCT), ("Days to target", jobf("Days to Target"), "0")]
    for i, (k, v, f) in enumerate(info, start=5):
        ws.cell(row=i, column=2, value=k).font = BOLD; x = ws.cell(row=i, column=3, value=f"={v}")
        if f: x.number_format = f
    right = [("Revised contract", f"={jobf('Revised Contract')}", None, None),
             ("Invoiced / Collected", f"={jobf('Invoiced')}", f"={jobf('Collected')}", MONEY),
             ("Balance due", f"={jobf('Balance Due')}", None, None),
             ("Projected gross profit / margin", f"={jobf('Projected Gross Profit')}", f"={jobf('Projected Margin %')}", PCT),
             ("Overhead allocation / net margin", f'=IF($C$4="","",{jobf("Revised Contract")}*{S["overhead"]})', f'=IF(OR($C$4="",{jobf("Revised Contract")}=0),"",({jobf("Projected Gross Profit")}-{jobf("Revised Contract")}*{S["overhead"]})/{jobf("Revised Contract")})', PCT),
             ("Labor hours: estimated / actual", f'=IF($C$4="","",SUMIFS({rr("Estimate","Qty")},{rr("Estimate","Job ID")},$C$4,{rr("Estimate","Category")},"Labor",{rr("Estimate","Unit Used")},"hr"))', f"={jobf('Labor Hours')}", HOURS)]
    for i, (k, v1, v2, f2) in enumerate(right, start=4):
        ws.cell(row=i, column=5, value=k).font = BOLD
        x = ws.cell(row=i, column=6, value=v1); x.number_format = HOURS if "hours" in k else MONEY
        if v2 is not None: y = ws.cell(row=i, column=7, value=v2); y.number_format = f2
    header(ws, 12, [None, "By category", "", "Estimated Cost", "Actual Cost", "Variance", "% Used", "Price Quoted", "Actual vs Price"])
    ws.cell(row=12, column=1).fill = PatternFill(None); ws.merge_cells("B12:C12")
    c1 = 13
    for i, (cat, _) in enumerate(CATS, start=c1):
        ws.cell(row=i, column=2, value=cat); ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=3)
        fx(ws, i, 4, f'=SUMIFS({rr("Estimate","Est. Cost")},{rr("Estimate","Job ID")},$C$4,{rr("Estimate","Category")},B{i})', MONEY)
        fx(ws, i, 5, f'={costs_by("$C$4", "Category", f"B{i}")}', MONEY)
        fx(ws, i, 6, f"=D{i}-E{i}", MONEY); fx(ws, i, 7, f"=IF(D{i}=0,IF(E{i}>0,1,0),E{i}/D{i})", PCT)
        fx(ws, i, 8, f'=SUMIFS({rr("Estimate","Price to Client")},{rr("Estimate","Job ID")},$C$4,{rr("Estimate","Category")},B{i})', MONEY)
        fx(ws, i, 9, f"=IF(H{i}=0,0,E{i}/H{i})", PCT)
        style_range(ws, i, i, 2, 3)
    ct = c1 + len(CATS)
    ws.cell(row=ct, column=2, value="Total by category"); ws.merge_cells(start_row=ct, start_column=2, end_row=ct, end_column=3)
    for cc, L in ((4, "D"), (5, "E"), (6, "F"), (8, "H")): fx(ws, ct, cc, f"=SUM({L}{c1}:{L}{ct-1})", MONEY, True)
    fx(ws, ct, 7, f"=IF(D{ct}=0,0,E{ct}/D{ct})", PCT, True); fx(ws, ct, 9, f"=IF(H{ct}=0,0,E{ct}/H{ct})", PCT, True)
    style_range(ws, ct, ct, 2, 9, fill=TOTAL_FILL, font=BOLD)
    hh = ct + 2
    header(ws, hh, [None, "Code", "Cost code", "Estimated Cost", "Actual Cost", "Variance ($)", "Variance (%)", "% Used", "Status"])
    ws.cell(row=hh, column=1).fill = PatternFill(None)
    B1 = hh + 1; B2 = B1 + N_CODES - 1
    for i in range(N_CODES):
        r = B1 + i; s = CODE_R1 + i
        fx(ws, r, 2, f"=Settings!E{s}"); fx(ws, r, 3, f"=Settings!F{s}")
        fx(ws, r, 4, f'=SUMIFS({rr("Estimate","Est. Cost")},{rr("Estimate","Job ID")},$C$4,{rr("Estimate","Cost Code")},B{r})', MONEY)
        fx(ws, r, 5, f'={costs_by("$C$4", "Cost Code", f"B{r}")}', MONEY)
        fx(ws, r, 6, f"=D{r}-E{r}", MONEY); fx(ws, r, 7, f"=IF(D{r}=0,IF(E{r}>0,-1,0),F{r}/D{r})", PCT)
        fx(ws, r, 8, f"=IF(D{r}=0,IF(E{r}>0,1,0),E{r}/D{r})", PCT)
        fx(ws, r, 9, f'=IF(AND(D{r}=0,E{r}=0),"",IF(E{r}>D{r},"OVER",IF(H{r}>=0.9,"Watch","On track")))')
    r = B2 + 1
    ws.cell(row=r, column=3, value="Approved change orders (added cost)")
    fx(ws, r, 4, f'=SUMIFS({rr("Change Orders","Added Cost (budget)")},{rr("Change Orders","Job ID")},$C$4,{rr("Change Orders","Status")},"Approved")', MONEY)
    fx(ws, r, 5, f'=IF($C$4="",0,{jobf("Actual Cost to Date")}-SUM(E{B1}:E{B2}))', MONEY)
    ws.cell(row=r + 1, column=3, value="(Actual on this line = costs logged without a cost code)").font = NOTE
    t2 = B2 + 3
    ws.cell(row=t2, column=3, value="TOTAL JOB BUDGET vs ACTUAL")
    fx(ws, t2, 4, f"=SUM(D{B1}:D{B2+1})", MONEY, True); fx(ws, t2, 5, f"=SUM(E{B1}:E{B2+1})", MONEY, True)
    fx(ws, t2, 6, f"=D{t2}-E{t2}", MONEY, True); fx(ws, t2, 7, f"=IF(D{t2}=0,0,F{t2}/D{t2})", PCT, True); fx(ws, t2, 8, f"=IF(D{t2}=0,0,E{t2}/D{t2})", PCT, True)
    style_range(ws, t2, t2, 2, 9, fill=TOTAL_FILL, font=BOLD)
    extras = [("Unpaid bills on this job (Job Costs, Paid? = No)", f'=SUMIFS({rr("Job Costs","Amount")},{rr("Job Costs","Job ID")},$C$4,{rr("Job Costs","Paid?")},"No")', MONEY),
              ("Pending change orders (price, not yet approved)", f'=SUMIFS({rr("Change Orders","Price to Client")},{rr("Change Orders","Job ID")},$C$4,{rr("Change Orders","Status")},"Pending")', MONEY),
              ("Late schedule tasks", f'=COUNTIFS({rr("Schedule","Job ID")},$C$4,{rr("Schedule","Status")},"Late")', None),
              ("Daily log entries", f'=COUNTIF({rr("Daily Log","Job ID")},$C$4)', None)]
    for i, (k, f, fmt) in enumerate(extras, start=t2 + 2):
        ws.cell(row=i, column=2, value=k).font = BOLD; ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=3); fx(ws, i, 4, f, fmt)
    s_ = f"I{B1}:I{B2}"
    ws.conditional_formatting.add(s_, CellIsRule(operator="equal", formula=['"OVER"'], fill=RED_FILL))
    ws.conditional_formatting.add(s_, CellIsRule(operator="equal", formula=['"Watch"'], fill=YEL_FILL))
    ws.conditional_formatting.add(s_, CellIsRule(operator="equal", formula=['"On track"'], fill=GRN_FILL))
    ws.conditional_formatting.add(f"G{c1}:G{ct-1}", CellIsRule(operator="greaterThan", formula=["1"], fill=RED_FILL))
    print_setup(ws, landscape=False, area=f"A1:I{t2+5}")
    return ws, {"cat_r1": c1, "cat_total": ct, "code_r1": B1, "total_row": t2}

def build_reports(wb):
    ws = wb.create_sheet("Reports"); ws.sheet_properties.tabColor = TAB["report"]; ws.sheet_view.showGridLines = False
    for col_, w in zip("ABCDEFGH", [4, 30, 16, 16, 16, 14, 14, 12]): ws.column_dimensions[col_].width = w
    ws["A1"] = "Company reports"; ws["A1"].font = TITLE
    ws["A2"] = "Receivables aging, cash forecast, pipeline, win rate and cost-code performance across all jobs. All calculated."; ws["A2"].font = NOTE
    ws["B4"] = "RECEIVABLES AGING"; ws["B4"].font = H2
    header(ws, 5, [None, "Bucket", "Open balance", "Invoices"]); ws.cell(row=5, column=1).fill = PatternFill(None)
    for i, b in enumerate(["Current", "1-30", "31-60", "61-90", "90+"], start=6):
        ws.cell(row=i, column=2, value=b).border = BORDER
        fx(ws, i, 3, f'=SUMIFS({rr("Invoices","Balance")},{rr("Invoices","Aging Bucket")},B{i})', MONEY)
        fx(ws, i, 4, f'=COUNTIF({rr("Invoices","Aging Bucket")},B{i})')
    ws.cell(row=11, column=2, value="Total open"); fx(ws, 11, 3, "=SUM(C6:C10)", MONEY, True); fx(ws, 11, 4, "=SUM(D6:D10)", None, True)
    style_range(ws, 11, 11, 2, 4, fill=TOTAL_FILL, font=BOLD)
    ws["B13"] = "CASH FORECAST — open invoices by due month"; ws["B13"].font = H2
    header(ws, 14, [None, "Month", "Due (open balance)", "Invoices"]); ws.cell(row=14, column=1).fill = PatternFill(None)
    ws.cell(row=15, column=2, value="Past due").border = BORDER
    fx(ws, 15, 3, f'=SUMIFS({rr("Invoices","Balance")},{rr("Invoices","Due Date")},"<"&TODAY(),{rr("Invoices","Balance")},">0")', MONEY)
    fx(ws, 15, 4, f'=COUNTIFS({rr("Invoices","Due Date")},"<"&TODAY(),{rr("Invoices","Balance")},">0")')
    for m in range(6):
        r = 16 + m
        fx(ws, r, 2, f"=DATE(YEAR(TODAY()),MONTH(TODAY())+{m},1)", "mmm yyyy")
        fx(ws, r, 3, f'=SUMIFS({rr("Invoices","Balance")},{rr("Invoices","Due Date")},">="&MAX(B{r},TODAY()),{rr("Invoices","Due Date")},"<="&EOMONTH(B{r},0),{rr("Invoices","Balance")},">0")', MONEY)
        fx(ws, r, 4, f'=COUNTIFS({rr("Invoices","Due Date")},">="&MAX(B{r},TODAY()),{rr("Invoices","Due Date")},"<="&EOMONTH(B{r},0),{rr("Invoices","Balance")},">0")')
    ws.cell(row=22, column=2, value="Total"); fx(ws, 22, 3, "=SUM(C15:C21)", MONEY, True); fx(ws, 22, 4, "=SUM(D15:D21)", None, True)
    style_range(ws, 22, 22, 2, 4, fill=TOTAL_FILL, font=BOLD)
    ws["B24"] = "PIPELINE BY STATUS"; ws["B24"].font = H2
    header(ws, 25, [None, "Status", "Jobs", "Revised contract", "Balance due", "Projected profit"]); ws.cell(row=25, column=1).fill = PatternFill(None)
    for i, s in enumerate(STATUS_LIST, start=26):
        ws.cell(row=i, column=2, value=s).border = BORDER
        fx(ws, i, 3, f'=COUNTIF({rr("Jobs","Status")},B{i})')
        fx(ws, i, 4, f'=SUMIFS({rr("Jobs","Revised Contract")},{rr("Jobs","Status")},B{i})', MONEY)
        fx(ws, i, 5, f'=SUMIFS({rr("Jobs","Balance Due")},{rr("Jobs","Status")},B{i})', MONEY)
        fx(ws, i, 6, f'=SUMIFS({rr("Jobs","Projected Gross Profit")},{rr("Jobs","Status")},B{i})', MONEY)
    t = 26 + len(STATUS_LIST)
    ws.cell(row=t, column=2, value="Total"); fx(ws, t, 3, f"=SUM(C26:C{t-1})", None, True)
    for cc, L in ((4, "D"), (5, "E"), (6, "F")): fx(ws, t, cc, f"=SUM({L}26:{L}{t-1})", MONEY, True)
    style_range(ws, t, t, 2, 6, fill=TOTAL_FILL, font=BOLD)
    ws.cell(row=t + 1, column=2, value="Win rate (won ÷ won + lost)").font = BOLD
    fx(ws, t + 1, 3, f'=IF(C28+C29+C30+C31+C32=0,0,(C28+C29+C30+C31)/(C28+C29+C30+C31+C32))', PCT)
    cc_r = t + 4
    ws.cell(row=cc_r - 1, column=2, value="COST CODE PERFORMANCE — all jobs").font = H2
    header(ws, cc_r, [None, "Cost code", "Estimated", "Actual", "Variance", "Variance %", "Flag"]); ws.cell(row=cc_r, column=1).fill = PatternFill(None)
    for i in range(N_CODES):
        r = cc_r + 1 + i; s = CODE_R1 + i
        fx(ws, r, 2, f'=Settings!E{s}&"  "&Settings!F{s}')
        fx(ws, r, 3, f'=SUMIFS({rr("Estimate","Est. Cost")},{rr("Estimate","Cost Code")},Settings!E{s})', MONEY)
        fx(ws, r, 4, f'={costs_all_by("Cost Code", f"Settings!E{s}")}', MONEY)
        fx(ws, r, 5, f"=C{r}-D{r}", MONEY); fx(ws, r, 6, f"=IF(C{r}=0,IF(D{r}>0,-1,0),E{r}/C{r})", PCT)
        fx(ws, r, 7, f'=IF(AND(C{r}=0,D{r}=0),"",IF(D{r}>C{r},"OVER",IF(D{r}>=0.9*C{r},"Watch","OK")))')
    e = cc_r + N_CODES + 1
    ws.cell(row=e, column=2, value="Total"); fx(ws, e, 3, f"=SUM(C{cc_r+1}:C{e-1})", MONEY, True); fx(ws, e, 4, f"=SUM(D{cc_r+1}:D{e-1})", MONEY, True)
    fx(ws, e, 5, f"=C{e}-D{e}", MONEY, True); fx(ws, e, 6, f"=IF(C{e}=0,0,E{e}/C{e})", PCT, True)
    style_range(ws, e, e, 2, 7, fill=TOTAL_FILL, font=BOLD)
    fl = f"G{cc_r+1}:G{e-1}"
    ws.conditional_formatting.add(fl, CellIsRule(operator="equal", formula=['"OVER"'], fill=RED_FILL))
    ws.conditional_formatting.add(fl, CellIsRule(operator="equal", formula=['"Watch"'], fill=YEL_FILL))
    ws.conditional_formatting.add(fl, CellIsRule(operator="equal", formula=['"OK"'], fill=GRN_FILL))
    print_setup(ws, landscape=False)
    return ws, {"aging_r1": 6, "forecast_r1": 15, "pipeline_r1": 26, "cc_r1": cc_r + 1, "winrate": f"C{t+1}"}

def build_tax_summary(wb):
    ws = wb.create_sheet("Tax Summary"); ws.sheet_properties.tabColor = TAB["report"]; ws.sheet_view.showGridLines = False
    for col_, w in zip("ABCDEF", [4, 40, 18, 18, 18, 22]): ws.column_dimensions[col_].width = w
    ws["A1"] = "Year-end summary (cash basis, informational)"; ws["A1"].font = TITLE
    ws["A2"] = "Totals for the reporting year set in Settings. Hand this to your accountant with your bank statements. Not tax advice."; ws["A2"].font = NOTE
    ws["B4"] = "Reporting year"; ws["B4"].font = BOLD; ws["C4"] = f"={S['year']}"
    y = "$C$4"; d1 = f"DATE({y},1,1)"; d2 = f"DATE({y},12,31)"
    ws["B6"] = "INCOME"; ws["B6"].font = H2
    ws["B7"] = "Payments received (cash collected)"; fx(ws, 7, 3, f'=SUMIFS({rr("Payments","Amount")},{rr("Payments","Date")},">="&{d1},{rr("Payments","Date")},"<="&{d2})', MONEY)
    ws["B8"] = "Invoiced (accrual reference)"; fx(ws, 8, 3, f'=SUMIFS({rr("Invoices","Amount")},{rr("Invoices","Invoice Date")},">="&{d1},{rr("Invoices","Invoice Date")},"<="&{d2})', MONEY)
    ws["B10"] = "JOB COSTS BY CATEGORY (as logged)"; ws["B10"].font = H2
    header(ws, 11, [None, "Category", "Job Costs", "Timesheet labor", "Mileage", "Total"]); ws.cell(row=11, column=1).fill = PatternFill(None)
    for i, (cat, _) in enumerate(CATS, start=12):
        ws.cell(row=i, column=2, value=cat).border = BORDER
        fx(ws, i, 3, f'=SUMIFS({rr("Job Costs","Amount")},{rr("Job Costs","Category")},B{i},{rr("Job Costs","Date")},">="&{d1},{rr("Job Costs","Date")},"<="&{d2})', MONEY)
        fx(ws, i, 4, f'=SUMIFS({rr("Timesheet",LC)},{rr("Timesheet","Category")},B{i},{rr("Timesheet","Date")},">="&{d1},{rr("Timesheet","Date")},"<="&{d2})', MONEY)
        fx(ws, i, 5, f'=SUMIFS({rr("Mileage","Amount")},{rr("Mileage","Category")},B{i},{rr("Mileage","Date")},">="&{d1},{rr("Mileage","Date")},"<="&{d2})', MONEY)
        fx(ws, i, 6, f"=C{i}+D{i}+E{i}", MONEY)
    ws.cell(row=18, column=2, value="Total job costs")
    for cc, L in ((3, "C"), (4, "D"), (5, "E"), (6, "F")): fx(ws, 18, cc, f"=SUM({L}12:{L}17)", MONEY, True)
    style_range(ws, 18, 18, 2, 6, fill=TOTAL_FILL, font=BOLD)
    ws["B19"] = "Gross profit (collected − job costs)"; fx(ws, 19, 6, "=C7-F18", MONEY, True)
    ws["B21"] = "MILEAGE"; ws["B21"].font = H2
    ws["B22"] = "Business miles logged"; fx(ws, 22, 3, f'=SUMIFS({rr("Mileage","Miles")},{rr("Mileage","Date")},">="&{d1},{rr("Mileage","Date")},"<="&{d2})', '#,##0.0')
    ws["B23"] = "Mileage amount at Settings rate"; fx(ws, 23, 3, f'=SUMIFS({rr("Mileage","Amount")},{rr("Mileage","Date")},">="&{d1},{rr("Mileage","Date")},"<="&{d2})', MONEY)
    ws["B25"] = "1099 REVIEW (US)"; ws["B25"].font = H2
    ws["B26"] = "Vendors flagged for 1099 review"; fx(ws, 26, 3, f'=COUNTIF({rng("Vendors & Crew", VC["1099 Flag"], VEND_R1, VEND_R2)},"Review for 1099")')
    ws["B27"] = "Paid YTD uses the reporting year. Flag = marked 1099-eligible and paid $600 or more. Confirm with your tax professional."; ws["B27"].font = NOTE
    ws["B29"] = "BY QUARTER"; ws["B29"].font = H2
    header(ws, 30, [None, "Quarter", "Invoiced", "Collected", "Job costs", "Gross profit (cash)"]); ws.cell(row=30, column=1).fill = PatternFill(None)
    for qn in range(4):
        r = 31 + qn; q1 = f"DATE({y},{3*qn+1},1)"; q2 = f"EOMONTH(DATE({y},{3*qn+3},1),0)"
        ws.cell(row=r, column=2, value=f"Q{qn+1}").border = BORDER
        fx(ws, r, 3, f'=SUMIFS({rr("Invoices","Amount")},{rr("Invoices","Invoice Date")},">="&{q1},{rr("Invoices","Invoice Date")},"<="&{q2})', MONEY)
        fx(ws, r, 4, f'=SUMIFS({rr("Payments","Amount")},{rr("Payments","Date")},">="&{q1},{rr("Payments","Date")},"<="&{q2})', MONEY)
        fx(ws, r, 5, f'={costs_between(q1, q2)}', MONEY)
        fx(ws, r, 6, f"=D{r}-E{r}", MONEY)
    ws.cell(row=35, column=2, value="Year"); 
    for cc, L in ((3, "C"), (4, "D"), (5, "E"), (6, "F")): fx(ws, 35, cc, f"=SUM({L}31:{L}34)", MONEY, True)
    style_range(ws, 35, 35, 2, 6, fill=TOTAL_FILL, font=BOLD)
    ws["B37"] = "1099 REVIEW LIST"; ws["B37"].font = H2
    header(ws, 38, [None, "Vendor", "Type", "Paid YTD", "W-9 on file", "Flag"]); ws.cell(row=38, column=1).fill = PatternFill(None)
    for i in range(N_VEND):
        r = 39 + i; vr = VEND_R1 + i
        v = cell("Vendors & Crew", VC["Vendor / Sub"], vr)
        fx(ws, r, 2, f'=IF({v}="","",{v})'); fx(ws, r, 3, f'=IF({v}="","",{cell("Vendors & Crew", VC["Type"], vr)})')
        fx(ws, r, 4, f'=IF({v}="","",{cell("Vendors & Crew", VC["Paid YTD"], vr)})', MONEY)
        fx(ws, r, 5, f'=IF({v}="","",{cell("Vendors & Crew", VC["W-9 on file"], vr)})'); fx(ws, r, 6, f'=IF({v}="","",{cell("Vendors & Crew", VC["1099 Flag"], vr)})')
    ws.conditional_formatting.add(f"F39:F{38+N_VEND}", CellIsRule(operator="equal", formula=['"Review for 1099"'], fill=YEL_FILL))
    print_setup(ws, landscape=False)
    return ws

def build_health_check(wb):
    ws = wb.create_sheet("Health Check"); ws.sheet_properties.tabColor = TAB["setup"]; ws.sheet_view.showGridLines = False
    for col_, w in zip("ABCDE", [4, 50, 12, 12, 40]): ws.column_dimensions[col_].width = w
    ws["A1"] = "Health Check"; ws["A1"].font = TITLE
    ws["A2"] = "Data problems that would make reports wrong. Each input sheet has a Check column; this page counts them. Fix anything marked Review."; ws["A2"].font = NOTE
    header(ws, 4, [None, "Check", "Issues", "Status", "Where to look"]); ws.cell(row=4, column=1).fill = PatternFill(None)
    def cnt(sheet, col_letter, r1, r2, crit='"?*"'): return f'COUNTIF({rng(sheet, col_letter, r1, r2)},{crit})'
    checks = [("Jobs: duplicate IDs, unknown client, missing status", cnt("Jobs", C["Jobs"]["Check"], R1, last("Jobs")), "Jobs → Check column"),
              ("Clients: duplicate names", cnt("Clients", C["Clients"]["Check"], R1, last("Clients")), "Clients → Check"),
              ("Crew: duplicates, missing cost rate", cnt("Vendors & Crew", CC["Check"], CREW_R1, CREW_R2), "Vendors & Crew → Crew table"),
              ("Vendors: duplicates", cnt("Vendors & Crew", VC["Check"], VEND_R1, VEND_R2), "Vendors & Crew → Vendor table"),
              ("Estimate lines: unknown job/code, missing qty or cost", cnt("Estimate", C["Estimate"]["Check"], R1, last("Estimate")), "Estimate → Check"),
              ("Schedule tasks: unknown job, missing start/duration", cnt("Schedule", C["Schedule"]["Check"], R1, last("Schedule")), "Schedule → Check"),
              ("Timesheet: unknown worker/job, missing hours or date", cnt("Timesheet", C["Timesheet"]["Check"], R1, last("Timesheet")), "Timesheet → Check"),
              ("Job Costs: unknown job/code, missing qty, cost or date", cnt("Job Costs", C["Job Costs"]["Check"], R1, last("Job Costs")), "Job Costs → Check"),
              ("Mileage: unknown job, missing miles or date", cnt("Mileage", C["Mileage"]["Check"], R1, last("Mileage")), "Mileage → Check"),
              ("Daily Log: unknown job, missing date", cnt("Daily Log", C["Daily Log"]["Check"], R1, last("Daily Log")), "Daily Log → Check"),
              ("Change Orders: unknown job, duplicate #, missing status/date", cnt("Change Orders", C["Change Orders"]["Check"], R1, last("Change Orders")), "Change Orders → Check"),
              ("Invoices: duplicate #, unknown job, missing due date/amount", cnt("Invoices", C["Invoices"]["Check"], R1, last("Invoices")), "Invoices → Check"),
              ("Payments: unknown job or invoice #, missing date", cnt("Payments", C["Payments"]["Check"], R1, last("Payments")), "Payments → Check")]
    info = [("Jobs over budget", f'=COUNTIF({rr("Jobs","% of Budget Used")},">1")', "Job Profit Report"),
            ("Invoices overdue", f'=COUNTIF({rr("Invoices","Status")},"Overdue")', "Invoices / Client Statement"),
            ("Schedule tasks late", f'=COUNTIF({rr("Schedule","Status")},"Late")', "Schedule"),
            ("Vendors with expired insurance", f'=COUNTIF({rng("Vendors & Crew", VC["Insurance Status"], VEND_R1, VEND_R2)},"EXPIRED")', "Vendors & Crew"),
            ("Vendors with insurance expiring within 30 days", f'=COUNTIF({rng("Vendors & Crew", VC["Insurance Status"], VEND_R1, VEND_R2)},"Expiring soon")', "Vendors & Crew"),
            ("Approved change orders without client-approved date", f'=COUNTIF({rr("Change Orders","Check")},"Approved without client date")', "Change Orders"),
            ("Unpaid bills logged (count)", f'=COUNTIF({rr("Job Costs","Paid?")},"No")', "Job Costs → Paid?"),
            ("Invoices overpaid (balance below zero)", f'=COUNTIF({rr("Invoices","Balance")},"<0")', "Invoices / Payments")]
    r = 5
    ws.cell(row=r, column=2, value="DATA INTEGRITY").font = H2; r += 1
    first_int = r
    for name, formula, where in checks:
        ws.cell(row=r, column=2, value=name).border = BORDER
        fx(ws, r, 3, f"={formula}"); fx(ws, r, 4, f'=IF(C{r}=0,"OK","Review")'); ws.cell(row=r, column=5, value=where).border = BORDER; r += 1
    last_int = r - 1
    ws.cell(row=r, column=2, value="Total data issues").font = BOLD; fx(ws, r, 3, f"=SUM(C{first_int}:C{last_int})", None, True)
    fx(ws, r, 4, f'=IF(C{r}=0,"ALL CLEAR","Review")', None, True); style_range(ws, r, r, 2, 5, fill=TOTAL_FILL, font=BOLD)
    total_row = r; r += 2
    ws.cell(row=r, column=2, value="ATTENTION ITEMS (not errors)").font = H2; r += 1
    first_info = r
    for name, formula, where in info:
        ws.cell(row=r, column=2, value=name).border = BORDER
        fx(ws, r, 3, formula); fx(ws, r, 4, f'=IF(C{r}=0,"—","Look")'); ws.cell(row=r, column=5, value=where).border = BORDER; r += 1
    r += 1
    ws.cell(row=r, column=2, value="FORMULA DIAGNOSTICS (should all be 0)").font = H2; r += 1
    diag_first = r
    def A(name, extra=1): return f"A{R1}:{get_column_letter(len(H[name]))}{last(name)+extra}"
    diag = {"Jobs": A("Jobs"), "Estimate": A("Estimate"), "Timesheet": f"A{R1}:T{last('Timesheet')+1}", "Job Costs": A("Job Costs"), "Mileage": A("Mileage"),
            "Daily Log": A("Daily Log", 0), "Change Orders": A("Change Orders", 0), "Invoices": A("Invoices"), "Payments": A("Payments"),
            "Schedule": f"A{R1}:J{last('Schedule')}", "Clients": A("Clients", 0), "Vendors & Crew": f"A{CREW_R1}:O{VEND_R2}", "Cost Library": f"A{R1}:G{last('Cost Library')}",
            "Proposal": "A1:J80", "CO Form": "A1:F23", "Invoice Print": "A1:F29", "Client Statement": "A1:I55", "Job Profit Report": "A1:I60",
            "Reports": "A1:G70", "Tax Summary": "A1:F145", "Dashboard": "A1:L100"}
    for sheet, area in diag.items():
        ws.cell(row=r, column=2, value=f"Formula errors on {sheet}").border = BORDER
        a1, a2 = area.split(":")
        fx(ws, r, 3, f"=SUMPRODUCT(--ISERROR({q(sheet)}!{a1}:{a2}))"); fx(ws, r, 4, f'=IF(C{r}=0,"OK","ERROR")'); ws.cell(row=r, column=5, value="").border = BORDER; r += 1
    diag_last = r - 1
    ws.cell(row=r, column=2, value="Total formula errors").font = BOLD; fx(ws, r, 3, f"=SUM(C{diag_first}:C{diag_last})", None, True); fx(ws, r, 4, f'=IF(C{r}=0,"OK","ERROR")', None, True)
    style_range(ws, r, r, 2, 5, fill=TOTAL_FILL, font=BOLD)
    for v, f in (("Review", RED_FILL), ("ERROR", RED_FILL), ("Look", YEL_FILL), ("OK", GRN_FILL), ("ALL CLEAR", GRN_FILL)):
        ws.conditional_formatting.add(f"D5:D{r}", CellIsRule(operator="equal", formula=[f'"{v}"'], fill=f))
    print_setup(ws, landscape=False)
    return ws, {"total_issues": f"C{total_row}", "total_issues_row": total_row, "diag_first": diag_first, "diag_last": diag_last, "diag_total": f"C{r}", "info_first": first_info}

def build_dashboard(wb, hc):
    ws = wb.create_sheet("Dashboard", 1); ws.sheet_properties.tabColor = TAB["report"]; ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    for col_ in "BCDEFGHIJKL": ws.column_dimensions[col_].width = 16
    ws.column_dimensions["M"].width = 3
    ws["A1"] = "Dashboard"; ws["A1"].font = TITLE
    ws["A2"] = "Company-wide view. Everything here is calculated. Change the reporting year in Settings to move the monthly P&L."; ws["A2"].font = NOTE
    ws["K2"] = "As of"; ws["K2"].font = NOTE; ws["K2"].alignment = Alignment(horizontal="right")
    ws["L2"] = "=TODAY()"; ws["L2"].number_format = DATEF; ws["L2"].font = BOLD
    JT = jobs_total_row(); j = C["Jobs"]
    def jt(h): return f"Jobs!{j[h]}{JT}"
    won = "+".join(f'COUNTIF({rr("Jobs","Status")},"{s}")' for s in ("Approved", "In Progress", "Complete", "Closed"))
    lost = f'COUNTIF({rr("Jobs","Status")},"Lost")'
    active = f'SUMIFS({rr("Jobs","Revised Contract")},{rr("Jobs","Status")},"Approved")+SUMIFS({rr("Jobs","Revised Contract")},{rr("Jobs","Status")},"In Progress")'
    active_inv = f'SUMIFS({rr("Jobs","Invoiced")},{rr("Jobs","Status")},"Approved")+SUMIFS({rr("Jobs","Invoiced")},{rr("Jobs","Status")},"In Progress")'
    m1 = "DATE(YEAR(TODAY()),MONTH(TODAY()),1)"; m2 = "EOMONTH(TODAY(),0)"
    kpis = [("Revised contract value (all jobs)", f"={jt('Revised Contract')}", MONEY0), ("Estimated cost", f"={jt('Estimated Cost')}", MONEY0),
            ("Actual cost to date", f"={jt('Actual Cost to Date')}", MONEY0), ("Projected gross profit", f"={jt('Projected Gross Profit')}", MONEY0),
            ("Projected gross margin", f"={jt('Projected Margin %')}", PCT), ("Margin after overhead (Settings %)", f"=IF({jt('Revised Contract')}=0,0,({jt('Projected Gross Profit')}-{jt('Revised Contract')}*{S['overhead']})/{jt('Revised Contract')})", PCT),
            ("Invoiced", f"={jt('Invoiced')}", MONEY0), ("Collected", f"={jt('Collected')}", MONEY0),
            ("Outstanding balance", f"={jt('Balance Due')}", MONEY0), ("Overdue balance", f'=SUMIFS({rr("Invoices","Balance")},{rr("Invoices","Status")},"Overdue")', MONEY0),
            ("Backlog (active jobs not yet invoiced)", f"={active}-({active_inv})", MONEY0), ("Unpaid bills logged (Paid? = No)", f'=SUMIFS({rr("Job Costs","Amount")},{rr("Job Costs","Paid?")},"No")', MONEY0),
            ("Active jobs (Approved / In Progress)", f'=COUNTIF({rr("Jobs","Status")},"Approved")+COUNTIF({rr("Jobs","Status")},"In Progress")', "0"), ("Jobs over budget", f'=COUNTIF({rr("Jobs","% of Budget Used")},">1")', "0"),
            ("Labor hours this month", f'=SUMIFS({rr("Timesheet","Hours")},{rr("Timesheet","Date")},">="&{m1},{rr("Timesheet","Date")},"<="&{m2})', HOURS), ("Win rate (won ÷ won + lost)", f"=IF({won}+{lost}=0,0,({won})/({won}+{lost}))", PCT)]
    for i, (k, f, fmt) in enumerate(kpis):
        r = 4 + (i // 4) * 3; c = 2 + (i % 4) * 3
        lab = ws.cell(row=r, column=c, value=k); lab.font = NOTE
        v = ws.cell(row=r + 1, column=c, value=f); v.font = Font(bold=True, size=14, color=NAVY); v.number_format = fmt
        for rr_ in (r, r + 1):
            for cc in (c, c + 1): ws.cell(row=rr_, column=cc).fill = KPI_FILL
    alerts = [("Overdue invoices", f'=COUNTIF({rr("Invoices","Status")},"Overdue")'), ("Late schedule tasks", f'=COUNTIF({rr("Schedule","Status")},"Late")'),
              ("Vendor insurance expired / expiring", f'=COUNTIF({rng("Vendors & Crew", VC["Insurance Status"], VEND_R1, VEND_R2)},"EXPIRED")+COUNTIF({rng("Vendors & Crew", VC["Insurance Status"], VEND_R1, VEND_R2)},"Expiring soon")'),
              ("Data issues (Health Check)", f"='Health Check'!{hc['total_issues']}")]
    ar = 16
    ws.cell(row=ar, column=2, value="ALERTS").font = Font(bold=True, size=9, color=TEAL)
    for i, (k, f) in enumerate(alerts):
        c = 2 + i * 3
        lab = ws.cell(row=ar + 1, column=c, value=k); lab.font = NOTE
        v = ws.cell(row=ar + 2, column=c, value=f); v.font = Font(bold=True, size=14, color=NAVY); v.number_format = "0"
        for rr_ in (ar + 1, ar + 2):
            for cc in (c, c + 1): ws.cell(row=rr_, column=cc).fill = KPI_FILL
        L = get_column_letter(c)
        ws.conditional_formatting.add(f"{L}{ar+2}:{get_column_letter(c+1)}{ar+2}", FormulaRule(formula=[f"${L}${ar+2}>0"], fill=YEL_FILL))
    tm = 20
    ws.cell(row=tm, column=2, value="THIS MONTH").font = Font(bold=True, size=9, color=TEAL)
    month = [("Invoiced this month", f'=SUMIFS({rr("Invoices","Amount")},{rr("Invoices","Invoice Date")},">="&{m1},{rr("Invoices","Invoice Date")},"<="&{m2})'),
             ("Collected this month", f'=SUMIFS({rr("Payments","Amount")},{rr("Payments","Date")},">="&{m1},{rr("Payments","Date")},"<="&{m2})'),
             ("Costs this month (all sources)", f'={costs_between(m1, m2)}'), ("Cash profit this month", f"=E{tm+2}-H{tm+2}")]
    for i, (k, f) in enumerate(month):
        c = 2 + i * 3
        lab = ws.cell(row=tm + 1, column=c, value=k); lab.font = NOTE
        v = ws.cell(row=tm + 2, column=c, value=f); v.font = Font(bold=True, size=14, color=NAVY); v.number_format = MONEY0
        for rr_ in (tm + 1, tm + 2):
            for cc in (c, c + 1): ws.cell(row=rr_, column=cc).fill = KPI_FILL
    hr = 40
    cols = ["Job ID", "Client", "Status", "Revised Contract", "Actual Cost", "Projected Profit", "Margin %", "% Budget Used", "Balance Due", "Days to Target", "Hours"]
    ws.cell(row=hr - 1, column=2, value="Jobs").font = H2
    header(ws, hr, [None] + cols); ws.cell(row=hr, column=1).fill = PatternFill(None)
    src = ["Job ID", "Client", "Status", "Revised Contract", "Actual Cost to Date", "Projected Gross Profit", "Projected Margin %", "% of Budget Used", "Balance Due", "Days to Target", "Labor Hours"]
    fmts = [None, None, None, MONEY0, MONEY0, MONEY0, PCT, PCT, MONEY0, "0", HOURS]
    for i in range(N_JOBS):
        r = hr + 1 + i; jr = R1 + i
        for k, (h, f) in enumerate(zip(src, fmts), start=2):
            fx(ws, r, k, f'=IF(Jobs!$A{jr}="","",Jobs!{j[h]}{jr})', f)
    JT1, JT2 = hr + 1, hr + N_JOBS
    ws.conditional_formatting.add(f"I{JT1}:I{JT2}", FormulaRule(formula=[f'AND(I{JT1}<>"",I{JT1}>1)'], fill=RED_FILL))
    ws.conditional_formatting.add(f"H{JT1}:H{JT2}", FormulaRule(formula=[f'AND(H{JT1}<>"",H{JT1}<0.1,D{JT1}<>"Lost",D{JT1}<>"Lead")'], fill=YEL_FILL))
    ws.conditional_formatting.add(f"K{JT1}:K{JT2}", FormulaRule(formula=[f'AND(K{JT1}<>"",K{JT1}<0)'], fill=RED_FILL))
    mr = JT2 + 3
    ws.cell(row=mr - 1, column=2, value="Monthly P&L").font = H2
    ws.cell(row=mr - 1, column=4, value="Year:").font = BOLD; ws.cell(row=mr - 1, column=4).alignment = Alignment(horizontal="right"); ws.cell(row=mr - 1, column=5, value=f"={S['year']}")
    header(ws, mr, [None, "Month", "Invoiced", "Payments Received", "Costs (all sources)", "Cash Profit", "Cumulative", "Labor Hours"]); ws.cell(row=mr, column=1).fill = PatternFill(None)
    for m in range(1, 13):
        r = mr + m
        fx(ws, r, 2, f"=DATE($E${mr-1},{m},1)", "mmm yyyy")
        fx(ws, r, 3, f'=SUMIFS({rr("Invoices","Amount")},{rr("Invoices","Invoice Date")},">="&B{r},{rr("Invoices","Invoice Date")},"<="&EOMONTH(B{r},0))', MONEY0)
        fx(ws, r, 4, f'=SUMIFS({rr("Payments","Amount")},{rr("Payments","Date")},">="&B{r},{rr("Payments","Date")},"<="&EOMONTH(B{r},0))', MONEY0)
        fx(ws, r, 5, f'={costs_between(f"B{r}", f"EOMONTH(B{r},0)")}', MONEY0)
        fx(ws, r, 6, f"=D{r}-E{r}", MONEY0); fx(ws, r, 7, f"=SUM($F${mr+1}:F{r})", MONEY0)
        fx(ws, r, 8, f'=SUMIFS({rr("Timesheet","Hours")},{rr("Timesheet","Date")},">="&B{r},{rr("Timesheet","Date")},"<="&EOMONTH(B{r},0))', HOURS)
    tr = mr + 13
    ws.cell(row=tr, column=2, value="TOTAL")
    for cc, L in ((3, "C"), (4, "D"), (5, "E"), (6, "F")): fx(ws, tr, cc, f"=SUM({L}{mr+1}:{L}{mr+12})", MONEY0, True)
    fx(ws, tr, 8, f"=SUM(H{mr+1}:H{mr+12})", HOURS, True)
    style_range(ws, tr, tr, 2, 8, fill=TOTAL_FILL, font=BOLD)
    ws.cell(row=24, column=2, value="Charts").font = H2
    ch1 = BarChart(); ch1.type = "col"; ch1.title = "Projected profit by job"; ch1.style = 10; ch1.legend = None
    ch1.add_data(Reference(ws, min_col=7, min_row=hr, max_row=hr + 12), titles_from_data=True); ch1.set_categories(Reference(ws, min_col=2, min_row=hr + 1, max_row=hr + 12))
    ch1.height = 6.2; ch1.width = 10.6; _axes(ch1, MONEY0); ws.add_chart(ch1, "B25")
    ch2 = BarChart(); ch2.type = "col"; ch2.title = "Payments vs costs by month"; ch2.style = 10
    ch2.add_data(Reference(ws, min_col=4, max_col=5, min_row=mr, max_row=mr + 12), titles_from_data=True); ch2.set_categories(Reference(ws, min_col=2, min_row=mr + 1, max_row=mr + 12))
    ch2.height = 6.2; ch2.width = 10.6; _axes(ch2, MONEY0); ch2.legend.position = "r"; ws.add_chart(ch2, "F25")
    ch3 = BarChart(); ch3.type = "col"; ch3.title = "Receivables aging"; ch3.style = 10; ch3.legend = None
    rp = wb["Reports"]
    ch3.add_data(Reference(rp, min_col=3, min_row=5, max_row=10), titles_from_data=True); ch3.set_categories(Reference(rp, min_col=2, min_row=6, max_row=10))
    ch3.height = 6.2; ch3.width = 10.6; _axes(ch3, MONEY0); ws.add_chart(ch3, "J25")
    print_setup(ws, area=f"A1:L{tr+1}")
    return ws, {"kpi_rows": (4, 5), "alert_row": ar + 2, "month_row": tm + 2, "job_hr": hr, "month_hr": mr}

def build_start_here(wb):
    ws = wb["Start Here"]; ws.sheet_properties.tabColor = TAB["setup"]; ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    for col_ in "BCDEFGHI": ws.column_dimensions[col_].width = 17
    for c in range(2, 10):
        ws.cell(row=1, column=c).fill = HEAD_FILL; ws.cell(row=2, column=c).fill = HEAD_FILL
    ws.merge_cells("B1:I2")
    x = ws["B1"]; x.value = "Contractor Job Costing & Estimate Spreadsheet"; x.font = Font(bold=True, size=20, color="FFFFFF"); x.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[1].height = 26; ws.row_dimensions[2].height = 22
    ws.merge_cells("B3:I3"); ws["B3"] = "Estimate → schedule → track hours and costs → change orders → invoice → collect → see real profit.  Works in Microsoft Excel (2016+, 365, Mac) and Google Sheets.  Version 1.0 (2026-09)."
    ws["B3"].font = NOTE; ws["B3"].alignment = Alignment(indent=1)
    ws["B5"] = "HOW IT FLOWS"; ws["B5"].font = H2
    flow = ["1  Settings", "2  Clients, Vendors & Crew", "3  Jobs", "4  Estimate → Proposal", "5  Schedule", "6  Timesheet, Costs, Mileage, Log", "7  Change Orders", "8  Invoices → Payments"]
    for i, t in enumerate(flow):
        c = ws.cell(row=6, column=2 + i, value=t); c.fill = HEAD_FILL if i % 2 == 0 else PatternFill("solid", fgColor=TEAL)
        c.font = Font(bold=True, color="FFFFFF", size=9); c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[6].height = 34
    ws.merge_cells("B7:I7"); ws["B7"] = "Dashboard, Job Profit Report, Reports, Tax Summary and Health Check update themselves from these inputs."; ws["B7"].font = NOTE; ws["B7"].alignment = Alignment(horizontal="center")
    groups = [("SETUP", TAB["setup"], [("Settings", "Company details, markups, tax, deposit and progress %, mileage rate, labor burden, cost codes, document text. Start here."),
                                      ("Clients", "Client directory with lifetime value and balance due."),
                                      ("Vendors & Crew", "Crew rates for the Timesheet; vendor and sub directory with insurance alerts and 1099 flags."),
                                      ("Cost Library", "Your price book: items with unit and unit cost that fill into estimates.")]),
              ("INPUTS", TAB["input"], [("Jobs", "One row per job. Everything rolls up here."),
                                       ("Estimate", "Line-by-line quote with automatic category markup and library lookups."),
                                       ("Schedule", "Tasks with workday end dates and a 12-week Gantt chart."),
                                       ("Timesheet", "Crew hours by day and job with a weekly payroll summary."),
                                       ("Job Costs", "Materials, subs, equipment, permits and other actual costs."),
                                       ("Mileage", "Job-related driving at your mileage rate."),
                                       ("Daily Log", "Weather, crew, work performed, issues, inspections — your dispute record."),
                                       ("Change Orders", "Added work; approved COs raise the contract and the budget."),
                                       ("Invoices", "Invoices with balances, status, days overdue and aging bucket."),
                                       ("Payments", "Money received, linked to invoices and jobs.")]),
              ("DOCUMENTS & REPORTS", TAB["report"], [("Proposal", "Client-ready proposal: scope, payment schedule, detailed lines, terms, signatures."),
                                                     ("CO Form", "Signature-ready change order document."),
                                                     ("Invoice Print", "Printable invoice with contract summary."),
                                                     ("Client Statement", "All of a client's invoices, balances and aging on one page."),
                                                     ("Job Profit Report", "Budget vs actual by category and cost code for one job."),
                                                     ("Reports", "Receivables aging, cash forecast, pipeline, win rate, cost-code performance."),
                                                     ("Tax Summary", "Year-end totals by category, mileage, 1099 review list."),
                                                     ("Dashboard", "16 KPIs, alerts, charts, per-job table, monthly P&L."),
                                                     ("Health Check", "Counts data problems and formula errors. Keep it at ALL CLEAR.")])]
    r = 9
    for name, color, items in groups:
        c = ws.cell(row=r, column=2, value=name); c.font = Font(bold=True, color="FFFFFF", size=10); c.fill = PatternFill("solid", fgColor=color)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=9); r += 1
        for s, d in items:
            link(ws, r, 2, s, s); ws.cell(row=r, column=3, value=d); ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=9)
            style_range(ws, r, r, 2, 9); r += 1
        r += 1
    blocks = [("TEN-MINUTE SETUP", ["1. Settings: company details, markups, contingency, tax, deposit %, mileage rate, labor burden. Rename cost codes to match your trade.",
                                    "2. Vendors & Crew: add your workers with cost and bill rates, and your subs and suppliers.",
                                    "3. Clients and Jobs: add a client, then a job row with a Job ID such as J-1001 (the sheet suggests the next one).",
                                    "4. Estimate the job, open Proposal, pick the Job ID, print. Enter the signed Contract Value on Jobs.",
                                    "5. As work happens: Timesheet for crew hours, Job Costs for everything else, Mileage, Daily Log, Change Orders.",
                                    "6. Invoice from the Invoices sheet and print with Invoice Print; record Payments; send Client Statements. Watch the Dashboard."]),
              ("COLOR KEY", ["Light-blue cells are inputs. Every other cell is a formula: do not type over them (Undo with Ctrl+Z / Cmd+Z if you do).",
                             "Red = over budget, overdue, late, expired, or a data problem. Yellow = watch. Green = on track / paid / done."]),
              ("PROTECT AGAINST ACCIDENTS (optional)", ["Input cells are already unlocked. On any sheet, Review → Protect Sheet (leave the password blank) and only the light-blue cells stay editable. Unprotect before inserting rows."]),
              ("GOOGLE SHEETS", ["Upload the .xlsx to Google Drive, open it, then File → Save as Google Sheets. Formulas, dropdowns, conditional formatting and the Gantt chart carry over. The three bar charts may need a one-time reformat."]),
              ("CAPACITY", [f"{N_JOBS} jobs · {N_EST} estimate lines · {N_TS} timesheet rows · {N_COST} cost entries · {N_MIL} mileage rows · {N_LOG} daily-log rows · {N_CO} change orders · {N_INV} invoices · {N_PAY} payments · {N_SCH} schedule tasks · {N_CLI} clients · {N_CREW} crew · {N_VEND} vendors · {N_LIB} library items.",
                            "Need more rows? Insert them inside a table (above the TOTAL row), then copy the formula cells of the row above into the new rows."]),
              ("IMPORTANT", ["This is an organizational and estimating tool, not accounting, tax, legal or financial advice. Verify figures before quoting, invoicing or filing.",
                             "Amounts are formatted in US dollars; change the number format (Format Cells) to use another currency symbol.",
                             "Licensed for use in your own business on any number of your devices. Resale or redistribution of the template is not permitted."])]
    for head, lines in blocks:
        ws.cell(row=r, column=2, value=head).font = H2; r += 1
        for ln in lines:
            x = ws.cell(row=r, column=2, value=ln); x.alignment = Alignment(wrap_text=True, vertical="top")
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=9); ws.row_dimensions[r].height = 30 if len(ln) > 140 else 16; r += 1
        r += 1
    print_setup(ws, landscape=False, area=f"A1:I{r}")
    return ws
