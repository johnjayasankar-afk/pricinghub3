"""v2: exports real sheet renders from Excel (workbook PDF -> per-sheet pages -> high-res PNG) and composes 10 Etsy listing images (2000x2000).
Run with --render to re-export from Excel; without it, reuses images/raw."""
import subprocess, os, glob, sys
from PIL import Image, ImageDraw, ImageFont, ImageChops
from pypdf import PdfReader, PdfWriter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XLSX = os.path.join(ROOT, "product", "Contractor-Job-Costing-SAMPLE.xlsx")
OUT = os.path.join(ROOT, "listing", "images"); RAW = os.path.join(OUT, "raw")
os.makedirs(RAW, exist_ok=True)
AREAS = {"Dashboard": "A1:N38", "Estimate": "A1:O27", "Job Profit Report": "A1:I54", "Proposal": "A1:F54", "Job Costs": "A1:N17",
         "Jobs": "A1:M12", "Invoices": "A1:O12", "Client Statement": "A1:I48", "Change Orders": "A1:M10", "Settings": "A1:L36", "Schedule": "A1:BG22",
         "Timesheet": "A1:T16", "Invoice Print": "A1:F35", "CO Form": "A1:F24", "Reports": "A1:G21", "Tax Summary": "A1:F36",
         "Health Check": None, "Vendors & Crew": "A39:N52", "Start Here": "A1:I48", "Daily Log": "A1:I10", "Mileage": "A1:L10", "Cost Library": "A1:G17", "Payments": "A1:H24"}
TITLES = {"Dashboard": "Dashboard", "Estimate": "Estimate", "Job Profit Report": "Job Profit Report", "Proposal": "PROPOSAL", "Job Costs": "Job Costs (actuals)",
          "Jobs": "Jobs", "Invoices": "Invoices", "Change Orders": "Change Orders", "Settings": "Settings", "Schedule": "Schedule", "Timesheet": "Timesheet",
          "Invoice Print": "INVOICE", "Client Statement": "STATEMENT", "CO Form": "CHANGE ORDER", "Reports": "Company reports", "Tax Summary": "Year-end summary", "Health Check": "Health Check",
          "Vendors & Crew": "Vendors & Crew", "Start Here": "HOW IT FLOWS", "Daily Log": "Daily Log", "Mileage": "Mileage log", "Cost Library": "Cost Library (price book)", "Payments": "Payments received"}
def fname(name): return name.replace(" ", "_").replace("&", "and")
# SAMPLE_DOCS=1: export the four client documents with their full print areas (for the landing page) instead of the image crops, then stop.
SAMPLE_DOCS = bool(os.environ.get("SAMPLE_DOCS"))
if SAMPLE_DOCS:
    AREAS.update({"Proposal": "A1:F79", "Invoice Print": "A1:F35", "CO Form": "A1:F24", "Client Statement": "A1:I52"})
if "--render" in sys.argv:
    for f in glob.glob(os.path.join(RAW, "*")): os.remove(f)
    ALLPDF = os.path.join(RAW, "workbook.pdf")
    script = f'tell application "Microsoft Excel"\n set wb to open workbook workbook file name (POSIX file "{XLSX}")\n calculate\n'
    for s_, area in AREAS.items():
        script += f' set ps to page setup object of worksheet "{s_}" of wb\n' + (f' set print area of ps to "{area}"\n' if area else '') + ' set zoom of ps to false\n set fit to pages wide of ps to 1\n set fit to pages tall of ps to 1\n set center footer of ps to ""\n'
    script += f' save workbook as wb filename (POSIX file "{ALLPDF}") file format PDF file format\n close wb saving no\nend tell\n'
    open(os.path.join(RAW, "export.applescript"), "w").write(script)
    r = subprocess.run(["osascript", os.path.join(RAW, "export.applescript")], capture_output=True, text=True); print("applescript:", r.returncode, r.stderr[:300])
    reader = PdfReader(ALLPDF); print("pages", len(reader.pages))
    found = set()
    for i, page in enumerate(reader.pages):
        txt = (page.extract_text() or "").strip(); first = txt.split("\n")[0] if txt else ""
        for name in sorted(TITLES, key=len, reverse=True):
            if first.startswith(TITLES[name]) and name not in found:
                w = PdfWriter(); w.add_page(page)
                with open(os.path.join(RAW, fname(name) + ".pdf"), "wb") as f: w.write(f)
                found.add(name); print("page", i, "->", name); break
    missing = [n for n in TITLES if n not in found]
    if missing:  # second export with no print areas at all (Excel drops some sheets when print areas are set)
        ALL2 = os.path.join(RAW, "workbook2.pdf")
        script = f'tell application "Microsoft Excel"\n set wb to open workbook workbook file name (POSIX file "{XLSX}")\n calculate\n'
        for s_ in TITLES:
            script += f' set ps to page setup object of worksheet "{s_}" of wb\n set zoom of ps to false\n set fit to pages wide of ps to 1\n set fit to pages tall of ps to 1\n set center footer of ps to ""\n'
        script += f' save workbook as wb filename (POSIX file "{ALL2}") file format PDF file format\n close wb saving no\nend tell\n'
        open(os.path.join(RAW, "export2.applescript"), "w").write(script)
        r = subprocess.run(["osascript", os.path.join(RAW, "export2.applescript")], capture_output=True, text=True); print("fallback export", missing, r.returncode, r.stderr[:300])
        reader = PdfReader(ALL2)
        for i, page in enumerate(reader.pages):
            txt = (page.extract_text() or "").strip(); first = txt.split("\n")[0] if txt else ""
            for name in missing:
                if first.startswith(TITLES[name]) and name not in found:
                    w = PdfWriter(); w.add_page(page)
                    with open(os.path.join(RAW, fname(name) + ".pdf"), "wb") as f: w.write(f)
                    found.add(name); print("page2", i, "->", name); break
    for pdf in glob.glob(os.path.join(RAW, "*.pdf")):
        if pdf == ALLPDF: continue
        subprocess.run(["qlmanage", "-t", "-s", "2800", "-o", RAW, pdf], capture_output=True)
        png = pdf + ".png"; dst = pdf[:-4] + ".png"
        if os.path.exists(png):
            im = Image.open(png).convert("RGBA"); bg = Image.new("RGBA", im.size, (255, 255, 255, 255)); bg.alpha_composite(im); bg.convert("RGB").save(dst); os.remove(png)
            print("rendered", os.path.basename(dst), Image.open(dst).size)

if SAMPLE_DOCS:
    dst = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "assets", "samples"); os.makedirs(dst, exist_ok=True)
    for name, fn in (("Proposal", "proposal.pdf"), ("Invoice Print", "invoice-print.pdf"), ("CO Form", "co-form.pdf"), ("Client Statement", "client-statement.pdf")):
        import shutil; shutil.copy(os.path.join(RAW, fname(name) + ".pdf"), os.path.join(dst, fn)); print("sample document", fn)
    sys.exit(0)

NAVY = (31, 58, 95); TEAL = (42, 127, 142); LIGHT = (238, 246, 247); WHITE = (255, 255, 255); DARK = (40, 40, 40)
def font(size, bold=False):
    p = "/System/Library/Fonts/Supplemental/Arial Bold.ttf" if bold else "/System/Library/Fonts/Supplemental/Arial.ttf"
    return ImageFont.truetype(p, size) if os.path.exists(p) else ImageFont.load_default()
def _gridlines(im):
    """Row indices that are table borders: mostly grey/navy pixels across the width."""
    import numpy as np
    a = np.asarray(im.convert("RGB")).astype(int); r, g, b = a[..., 0], a[..., 1], a[..., 2]
    grey = (abs(r - g) < 18) & (abs(g - b) < 18) & (r > 120) & (r < 236)
    navy = (r < 80) & (g < 90) & (b > 80)
    frac = (grey | navy).mean(axis=1)
    return [i for i, f in enumerate(frac) if f > 0.45]
def shot(name, box_w, box_h, crop_top=0.0, crop_h=1.0, crop_w=1.0):
    im = Image.open(os.path.join(RAW, fname(name) + ".png")).convert("RGB")
    bbox = ImageChops.difference(im, Image.new("RGB", im.size, WHITE)).getbbox()
    if bbox: im = im.crop(bbox)
    W, H = im.size
    if crop_top or crop_h < 1.0:  # snap the cut to the nearest table gridline so no row is sliced
        lines = _gridlines(im); top = int(H * crop_top); bot = int(H * min(1.0, crop_top + crop_h))
        if lines:
            below = [l for l in lines if l >= top]; above = [l for l in lines if l <= bot]
            if below and crop_top: top = below[0]
            if above and crop_h < 1.0: bot = above[-1] + 1
        im = im.crop((0, top, int(W * crop_w), bot))
    elif crop_w < 1.0: im = im.crop((0, 0, int(W * crop_w), H))
    W, H = im.size; sc = min(box_w / W, box_h / H)
    return im.resize((int(W * sc), int(H * sc)), Image.LANCZOS)
def place(img, im, x, y, pad=18, radius=24):
    w, h = im.size; card = Image.new("RGB", (w + 2 * pad, h + 2 * pad), WHITE); card.paste(im, (pad, pad))
    mask = Image.new("L", card.size, 0); ImageDraw.Draw(mask).rounded_rectangle([0, 0, card.size[0] - 1, card.size[1] - 1], radius, fill=255)
    img.paste(Image.new("RGB", card.size, (200, 208, 214)), (x + 14, y + 14), mask); img.paste(card, (x, y), mask)
def base(t, sub=None):
    img = Image.new("RGB", (2000, 2000), LIGHT); d = ImageDraw.Draw(img); d.rectangle([0, 0, 2000, 300], fill=NAVY)
    d.text((80, 70), t, font=font(84, True), fill=WHITE)
    if sub: d.text((80, 190), sub, font=font(42), fill=(205, 222, 235))
    return img, d
def footer(d, text="Excel + Google Sheets  •  24 sheets  •  Instant download  •  Sample + blank + PDF guide"):
    d.rectangle([0, 1900, 2000, 2000], fill=TEAL); d.text((80, 1926), text, font=font(38, True), fill=WHITE)
def check(d, x, y, t, size=46):
    d.line([(x, y + 30), (x + 18, y + 50), (x + 52, y + 8)], fill=TEAL, width=9); d.text((x + 75, y), t, font=font(size, True), fill=NAVY)
def place_center(img, im, y): place(img, im, (2000 - (im.size[0] + 36)) // 2, y)
def caption(d, x, y, t): d.text((x, y), t.upper(), font=font(30, True), fill=TEAL)
def save(img, name): img.save(os.path.join(OUT, name))
for f in glob.glob(os.path.join(OUT, "*.png")): os.remove(f)

img, d = base("Contractor Job Costing", "Estimate → schedule → track hours & costs → invoice → see real profit")
# Etsy's search thumbnail is a 4:3 centre crop (y 250..1750): the dashboard card and the checklist both sit inside that safe zone.
place(img, shot("Dashboard", 1804, 880), 80, 330)
d.rounded_rectangle([80, 1290, 1920, 1700], 30, fill=WHITE)
for i, t in enumerate(["Estimates, proposals & price book", "12-week Gantt schedule", "Crew timesheets & payroll summary", "Change orders with signable forms",
                       "Invoices, statements, receivables aging", "Budget vs actual by cost code", "Tax-time summary & 1099 review", "Dashboard: 16 KPIs, alerts, charts"]):
    check(d, 130 + (i % 2) * 920, 1330 + (i // 2) * 92, t, 42)
footer(d); save(img, "01-hero.png")
img, d = base("12-week Gantt schedule", "Workday end dates. Blue = planned, dark = done, yellow = today. Late tasks turn red.")
caption(d, 80, 360, "Schedule — tasks and Gantt (first seven weeks shown)"); place(img, shot("Schedule", 1804, 560), 80, 410)
caption(d, 80, 830, "Daily Log — weather, crew, work done, issues: your dispute record"); place(img, shot("Daily Log", 1804, 340), 80, 880)
caption(d, 80, 1290, "Mileage — miles × your rate, counted into job cost"); place(img, shot("Mileage", 1804, 340), 80, 1340)
footer(d); save(img, "02-schedule.png")
img, d = base("Build the estimate line by line", "Cost code → category → markup, automatically. Pull items from your price book.")
caption(d, 80, 360, "Estimate — one line per item, markup and proposal line number calculated"); place(img, shot("Estimate", 1804, 640), 80, 410)
caption(d, 80, 1120, "Cost Library — your price book, picked from a dropdown"); place(img, shot("Cost Library", 1804, 640), 80, 1170)
footer(d); save(img, "03-estimate.png")
img, d = base("Client-ready proposal in one click", "Pick a Job ID: scope, payment schedule, detailed lines, terms, signatures.")
place_center(img, shot("Proposal", 1300, 1460), 360); footer(d); save(img, "04-proposal.png")
img, d = base("Budget vs actual, per cost code", "Estimate + approved change orders vs Job Costs + Timesheet + Mileage.")
place_center(img, shot("Job Profit Report", 1804, 1460), 360); footer(d); save(img, "05-job-profit-report.png")
img, d = base("Crew timesheets, paid right", "Rates from your crew table, weekly payroll summary, billable value.")
caption(d, 80, 360, "Timesheet — hours by day, job and cost code, with the weekly payroll summary"); place(img, shot("Timesheet", 1804, 560), 80, 410)
caption(d, 80, 790, "Crew — cost rate, bill rate, hours and labor cost year-to-date"); place(img, shot("Vendors & Crew", 1804, 330, 0.0, 0.068, 0.62), 80, 840)
caption(d, 80, 1250, "Job Costs — materials, subs, permits, equipment; paid or still owed"); place(img, shot("Job Costs", 1804, 500), 80, 1300)
footer(d); save(img, "06-timesheet-costs.png")
img, d = base("Change orders that get signed", "Log it, approve it, print the form. Contract and budget update automatically.")
caption(d, 80, 360, "Change Orders — approved ones raise the contract and the budget"); place(img, shot("Change Orders", 1804, 480), 80, 410)
caption(d, 80, 800, "CO Form — signature-ready, with the revised contract total and completion date"); place_center(img, shot("CO Form", 1100, 900), 850)
footer(d); save(img, "07-change-orders.png")
img, d = base("Invoices, statements, aging", "Printable invoices and client statements. Overdue turns red. Cash forecast by month.")
caption(d, 80, 360, "Invoice Print"); place(img, shot("Invoice Print", 864, 760), 80, 410)
caption(d, 1020, 360, "Client Statement"); place(img, shot("Client Statement", 864, 760), 1020, 410)
caption(d, 80, 1220, "Reports — aging and cash forecast"); place(img, shot("Reports", 864, 560), 80, 1270)
caption(d, 1020, 1220, "Payments — linked to invoices"); place(img, shot("Payments", 864, 560), 1020, 1270)
footer(d); save(img, "08-invoices-reports.png")
img, d = base("Tax time, vendors & health check", "Year-end totals by category, 1099 review, insurance alerts, data-problem counter.")
caption(d, 80, 360, "Tax Summary"); place(img, shot("Tax Summary", 864, 760), 80, 410)
caption(d, 1020, 360, "Health Check"); place(img, shot("Health Check", 864, 760, 0, 0.575), 1020, 410)
caption(d, 80, 1220, "Vendors & Subcontractors — insurance alerts, W-9 and 1099 flags, unpaid balances"); place(img, shot("Vendors & Crew", 1804, 520, 0.255, 0.125, 0.82), 80, 1270)
footer(d); save(img, "09-tax-vendors-health.png")
img, d = base("What you get", "Instant download after purchase")
items = [("SAMPLE workbook (.xlsx)", "Three worked jobs, crew, vendors, schedule, invoices: see how everything connects."),
         ("BLANK workbook (.xlsx)", "Ready for your business. 40 jobs · 400 estimate lines · 1,000 timesheet rows · 1,000 cost entries."),
         ("Quick-start guide (PDF)", "Ten-minute setup, sheet-by-sheet walkthrough, FAQ."),
         ("24 connected sheets", "Settings, Clients, Vendors & Crew, Cost Library, Jobs, Estimate, Proposal, Schedule, Timesheet, Job Costs, Mileage, Daily Log, Change Orders, CO Form, Invoices, Invoice Print, Payments, Client Statement, Job Profit Report, Reports, Tax Summary, Health Check, Dashboard."),
         ("Google Sheets ready", "Upload to Drive → File > Save as Google Sheets. Same formulas, dropdowns, Gantt."),
         ("No subscription. No add-ins. No macros.", "Plain formulas you can audit. Windows, Mac, iPad and web.")]
y = 370
for t, sdesc in items:
    lines, line = [], ""
    for w in sdesc.split():
        if d.textlength((line + " " + w).strip(), font=font(34)) > 1700: lines.append(line); line = w
        else: line = (line + " " + w).strip()
    lines.append(line); h = 105 + 42 * len(lines) + 28
    d.rounded_rectangle([80, y, 1920, y + h], 24, fill=WHITE); d.text((130, y + 30), t, font=font(50, True), fill=NAVY)
    for k, ln in enumerate(lines): d.text((130, y + 105 + 42 * k), ln, font=font(34), fill=DARK)
    y += h + 22
footer(d); save(img, "10-included.png")
print("done", sorted(f for f in os.listdir(OUT) if f.endswith(".png")))
