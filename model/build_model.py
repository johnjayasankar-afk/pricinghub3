"""Editable financial model for the Etsy + Payhip digital-download plan. All inputs are on one sheet; scenarios are formulas."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

wb = Workbook(); ws = wb.active; ws.title = "Model"
INP = PatternFill("solid", fgColor="EAF3FF"); HEAD = PatternFill("solid", fgColor="1F3A5F"); TOT = PatternFill("solid", fgColor="F2F2F2")
W = Font(bold=True, color="FFFFFF"); B = Font(bold=True); N = Font(italic=True, color="666666", size=9)
M = '"$"#,##0.00;[Red]-"$"#,##0.00'; P = '0.0%'
ws.column_dimensions["A"].width = 46
for c in "BCDEF": ws.column_dimensions[c].width = 16
ws["A1"] = "Financial model — Contractor Job Costing Spreadsheet (Etsy primary, Payhip secondary)"; ws["A1"].font = Font(bold=True, size=14, color="1F3A5F")
ws["A2"] = "Blue cells are assumptions. Sources: Etsy fee page (help.etsy.com/hc/en-us/articles/115014483627), Etsy US processing fee 3% + $0.25 (third-party summaries; verify at shop setup), Payhip pricing page. Accessed 2026-09-04."; ws["A2"].font = N

def row(r, label, val=None, fmt=None, inp=False, note=None):
    ws.cell(row=r, column=1, value=label)
    if val is not None:
        c = ws.cell(row=r, column=2, value=val)
        if fmt: c.number_format = fmt
        if inp: c.fill = INP; c.font = Font(color="1F4E9A")
    if note: ws.cell(row=r, column=3, value=note).font = N

ws["A4"] = "ASSUMPTIONS (per Etsy sale)"; ws["A4"].font = B
row(5, "List price", 29.00, M, True, "Competitors $9.97–$97 (Gantt+forms tier $47–$97); anchor $29, launch 25% off")
row(6, "Typical discount %", 0.25, P, True)
row(7, "Realized price", "=B5*(1-B6)", M)
row(8, "Etsy transaction fee %", 0.065, P, True, "Official")
row(9, "Etsy payment processing % (US)", 0.03, P, True, "Verify at setup")
row(10, "Etsy payment processing fixed ($)", 0.25, M, True)
row(11, "Etsy auto-renew-on-sale fee ($)", 0.20, M, True, "Official: $0.20 each time a digital listing sells")
row(12, "Offsite Ads share of orders", 0.10, P, True, "Mandatory opt-in under $10k/yr; 15% fee only on ad-attributed orders")
row(13, "Offsite Ads fee %", 0.15, P, True)
row(14, "Refund rate (goodwill refunds on non-returnable item)", 0.03, P, True)
row(15, "Fulfillment / AI / hosting cost per sale", 0.00, M, True, "Etsy hosts and delivers the file")
row(16, "Net contribution per Etsy sale", "=B7*(1-B8-B9-B12*B13)-B10-B11-B15-B14*B7", M)
row(17, "Contribution margin", "=B16/B7", P)
ws["A19"] = "ASSUMPTIONS (per Payhip sale, direct link)"; ws["A19"].font = B
row(20, "Realized price", "=B7", M)
row(21, "Payhip fee %", 0.05, P, True, "Free plan")
row(22, "PayPal fee % / fixed", 0.0349, P, True, "PayPal US commercial rate; verify")
row(23, "PayPal fixed ($)", 0.49, M, True)
row(24, "Net contribution per Payhip sale", "=B20*(1-B21-B22)-B23-B14*B20", M)
ws["A26"] = "FIXED / ONE-TIME COSTS"; ws["A26"].font = B
row(27, "Etsy shop set-up fee (one-time, owner approval)", 29.00, M, True, "Etsy says 'varies by location'; third parties report $15–$29 US")
row(28, "Etsy listing fees (2 listings × $0.20)", 0.40, M, True)
row(29, "Listing renewals per month if unsold (2 × $0.20 / 4 months)", "=0.40/4", M)
row(30, "Payhip monthly plan", 0.00, M, True)
row(31, "Fixed monthly expenses", "=B29+B30", M)
row(32, "Break-even sales per month (fixed / contribution)", "=IF(B16<=0,\"n/a\",B31/B16)", '0.00')
row(33, "Sales to recover one-time setup", "=IF(B16<=0,\"n/a\",(B27+B28)/B16)", '0.0')
ws["A35"] = "SCENARIOS (month 1–3 average)"; ws["A35"].font = B
for j, h in enumerate(["Zero sales", "Conservative", "Base", "Optimistic"], start=2):
    c = ws.cell(row=36, column=j, value=h); c.fill = HEAD; c.font = W; c.alignment = Alignment(horizontal="center")
ws.cell(row=36, column=1, value="Metric").fill = HEAD; ws.cell(row=36, column=1).font = W
labels = ["Etsy sales / month", "Payhip sales / month", "Revenue / month", "Contribution / month", "Fixed monthly", "Net profit / month", "Cumulative net after 3 months (incl. setup)", "Cash actually in bank by day 45"]
vals = {"Etsy sales / month": [0, 2, 6, 15], "Payhip sales / month": [0, 0, 1, 3]}
for i, lab in enumerate(labels, start=37):
    ws.cell(row=i, column=1, value=lab)
for j in range(2, 6):
    L = "BCDE"[j - 2]
    ws.cell(row=37, column=j, value=vals["Etsy sales / month"][j - 2]).fill = INP
    ws.cell(row=38, column=j, value=vals["Payhip sales / month"][j - 2]).fill = INP
    ws.cell(row=39, column=j, value=f"={L}37*$B$7+{L}38*$B$20").number_format = M
    ws.cell(row=40, column=j, value=f"={L}37*$B$16+{L}38*$B$24").number_format = M
    ws.cell(row=41, column=j, value="=$B$31").number_format = M
    ws.cell(row=42, column=j, value=f"={L}40-{L}41").number_format = M
    ws.cell(row=43, column=j, value=f"=3*{L}42-$B$27-$B$28").number_format = M
    ws.cell(row=44, column=j, value=f"=({L}37*$B$16)*0.5+{L}38*$B$24*1.5").number_format = M
ws.cell(row=44, column=6, value="Etsy: only sales older than 14 days deposited; assumes half of month-1 Etsy sales cleared by day 45. Payhip: instant.").font = N
ws["A46"] = "CASH TIMING"; ws["A46"].font = B
row(47, "Etsy: sale → eligible for deposit", "14 days (new seller)", note="help.etsy.com deposit article")
row(48, "Etsy: deposit schedule", "Weekly, Monday (default)")
row(49, "Etsy: bank-detail change hold", "5 days")
row(50, "Payhip: sale → PayPal balance", "Instant")
row(51, "PayPal → bank", "1–3 business days (standard transfer)")
ws["A53"] = "NOTES"; ws["A53"].font = B
ws["A54"] = "Revenue ≠ profit ≠ cash. Personal income tax is not modeled (jurisdiction-specific). Sales tax on digital goods is collected and remitted by Etsy as marketplace facilitator; Payhip sales may require the seller to handle sales tax in some states."; ws["A54"].alignment = Alignment(wrap_text=True); ws.row_dimensions[54].height = 45
ws["A55"] = "A single abusive buyer costs at most one refund (~$18). No usage-based costs exist, so growth cannot erase margin."; ws["A55"].alignment = Alignment(wrap_text=True); ws.row_dimensions[55].height = 30
for r in range(5, 34):
    for c in (1, 2): ws.cell(row=r, column=c).border = Border(bottom=Side(style="thin", color="DDDDDD"))
wb.save("model/financial-model.xlsx"); print("saved")
