"""Fictional sample dataset (three jobs) used by the SAMPLE workbook and by the tests."""
from datetime import date as D

CLIENTS = [("Alvarez residence", "Maria Alvarez", "(555) 201-7788", "maria.alvarez@example.com", "118 Maple St, Dover, DE 19901", ""),
           ("Kim & Patel", "Daniel Kim", "(555) 318-4402", "dkim@example.com", "42 Orchard Ave, Dover, DE 19904", "Referred by Alvarez"),
           ("Riverside HOA", "L. Chen (board treasurer)", "(555) 555-0190", "treasurer@riversidehoa.example", "1 Riverside Dr, Dover, DE 19901", "Net-30 terms")]
CREW = [("M. Ortiz", "Lead carpenter", 42, 65, "(555) 400-1001"), ("D. Lee", "Carpenter", 38, 55, "(555) 400-1002"),
        ("S. Reyes", "Apprentice", 24, 40, "(555) 400-1003"), ("J. Owner", "Owner / PM", 65, 85, "(555) 010-2200")]
# name, type, trade, phone, email, license, insurance expiry, w9, 1099
VENDORS = [("Bright Plumbing", "Subcontractor", "Plumbing", "(555) 600-2001", "office@brightplumbing.example", "PL-2231", D(2027, 3, 1), "Yes", "Yes"),
           ("Volt Electric", "Subcontractor", "Electrical", "(555) 600-2002", "jobs@voltelectric.example", "EL-8810", D(2026, 9, 20), "Yes", "Yes"),
           ("Smooth Drywall", "Subcontractor", "Drywall", "(555) 600-2003", "", "", D(2026, 6, 30), "Yes", "Yes"),
           ("Stone & Co", "Subcontractor", "Countertops", "(555) 600-2004", "", "", D(2027, 1, 15), "No", "Yes"),
           ("Builders Supply", "Supplier", "Lumber & hardware", "(555) 600-3001", "", "", None, "No", "No"),
           ("Tile Outlet", "Supplier", "Tile", "(555) 600-3002", "", "", None, "No", "No"),
           ("CabinetDirect", "Supplier", "Cabinets", "(555) 600-3003", "", "", None, "No", "No"),
           ("Fixture Depot", "Supplier", "Plumbing fixtures", "(555) 600-3004", "", "", None, "No", "No"),
           ("Paint Pro", "Supplier", "Paint", "(555) 600-3005", "", "", None, "No", "No"),
           ("Metro Dumpster", "Other", "Waste", "(555) 600-4001", "", "", D(2027, 5, 31), "Yes", "No"),
           ("City of Dover", "Other", "Permits", "(555) 736-7000", "", "", None, "No", "No")]
LIBRARY = [("Demolition — kitchen (per lot)", "02", "lot", 1800), ("Demolition — bathroom (per lot)", "02", "lot", 900),
           ("Framing labor", "05", "hr", 55), ("Finish carpentry labor", "14", "hr", 55), ("Paint labor", "14", "hr", 50),
           ("Project supervision", "01", "hr", 65), ("Porcelain tile (material)", "16", "sq ft", 6.5),
           ("Quartz countertop installed", "17", "sq ft", 85), ("Building permit", "20", "ea", 450),
           ("Dumpster pull (20 yd)", "21", "ea", 425), ("Composite decking package", "06", "lot", 7400),
           ("Auger / lift rental (day)", "19", "day", 240)]
# id, client, desc, status, start, target end, pct, contract
JOBS = [("J-1001", "Alvarez residence", "Kitchen remodel — 118 Maple St", "In Progress", D(2026, 7, 6), D(2026, 9, 11), 0.6, 42500),
        ("J-1002", "Kim & Patel", "Bathroom remodel — 42 Orchard Ave", "Complete", D(2026, 6, 1), D(2026, 6, 26), 1.0, None),
        ("J-1003", "Riverside HOA", "Deck rebuild — clubhouse", "Quoted", D(2026, 9, 15), D(2026, 10, 2), 0.0, None),
        ("J-1004", "Kim & Patel", "Basement office build-out", "Lost", D(2026, 8, 3), None, 0.0, None)]
# job, code, library item, description, qty, unit, unit cost (None -> library)
ESTIMATE = [("J-1001", "02", "Demolition — kitchen (per lot)", "", 1, "", None), ("J-1001", "06", "", "Framing lumber & blocking", 1, "lot", 950),
            ("J-1001", "05", "Framing labor", "", 24, "", None), ("J-1001", "10", "", "Plumbing rough & finish (sub)", 1, "lot", 4200),
            ("J-1001", "11", "", "Electrical rough & finish (sub)", 1, "lot", 3900), ("J-1001", "13", "", "Drywall patch & finish (sub)", 1, "lot", 1600),
            ("J-1001", "17", "", "Cabinets", 1, "lot", 9800), ("J-1001", "17", "Quartz countertop installed", "", 48, "", None),
            ("J-1001", "16", "Porcelain tile (material)", "Tile floor material", 220, "", None), ("J-1001", "14", "Finish carpentry labor", "Tile & trim install labor", 40, "", None),
            ("J-1001", "15", "", "Paint materials", 1, "lot", 380), ("J-1001", "14", "Paint labor", "", 16, "", None),
            ("J-1001", "18", "", "Sink, faucet, disposal", 1, "lot", 1150), ("J-1001", "20", "Building permit", "", 1, "", None),
            ("J-1001", "21", "Dumpster pull (20 yd)", "Dumpster (2 pulls)", 2, "", None), ("J-1001", "01", "Project supervision", "", 30, "", None),
            ("J-1002", "02", "Demolition — bathroom (per lot)", "", 1, "", None), ("J-1002", "10", "", "Plumbing (sub)", 1, "lot", 2600),
            ("J-1002", "11", "", "Electrical (sub)", 1, "lot", 1400), ("J-1002", "16", "", "Tile material", 140, "sq ft", 7.25),
            ("J-1002", "14", "", "Tile install labor", 32, "hr", 55), ("J-1002", "18", "", "Vanity, toilet, fixtures", 1, "lot", 2100),
            ("J-1002", "15", "", "Paint materials", 1, "lot", 160), ("J-1002", "14", "", "Finish carpentry & paint labor", 14, "hr", 50),
            ("J-1002", "20", "", "Permit", 1, "ea", 250), ("J-1002", "01", "", "Supervision", 12, "hr", 65),
            ("J-1003", "02", "", "Remove old deck", 1, "lot", 1200), ("J-1003", "06", "Composite decking package", "PT lumber, composite decking, hardware", 1, "", None),
            ("J-1003", "05", "", "Deck framing & decking labor", 90, "hr", 55), ("J-1003", "04", "", "Footings (sub)", 1, "lot", 1800),
            ("J-1003", "20", "", "Permit & inspections", 1, "ea", 320), ("J-1003", "19", "Auger / lift rental (day)", "", 2, "", None)]
# job, task, assigned, start, duration, pct
SCHEDULE = [("J-1001", "Demolition", "M. Ortiz", D(2026, 7, 6), 2, 1.0), ("J-1001", "Framing & blocking", "M. Ortiz", D(2026, 7, 8), 3, 1.0),
            ("J-1001", "Rough plumbing (Bright)", "J. Owner", D(2026, 7, 13), 2, 1.0), ("J-1001", "Rough electrical (Volt)", "J. Owner", D(2026, 7, 15), 2, 1.0),
            ("J-1001", "Rough inspection", "J. Owner", D(2026, 7, 17), 1, 1.0), ("J-1001", "Drywall patch & finish", "J. Owner", D(2026, 7, 27), 4, 1.0),
            ("J-1001", "Cabinet install", "M. Ortiz", D(2026, 8, 4), 3, 1.0), ("J-1001", "Tile floor", "D. Lee", D(2026, 8, 11), 4, 1.0),
            ("J-1001", "Countertops (Stone & Co)", "J. Owner", D(2026, 8, 18), 1, 1.0), ("J-1001", "Paint", "S. Reyes", D(2026, 8, 24), 3, 0.5),
            ("J-1001", "Fixtures, trim & punch list", "M. Ortiz", D(2026, 9, 7), 4, 0.0),
            ("J-1003", "Remove old deck", "D. Lee", D(2026, 9, 15), 2, 0.0), ("J-1003", "Footings (sub)", "J. Owner", D(2026, 9, 17), 2, 0.0),
            ("J-1003", "Framing", "M. Ortiz", D(2026, 9, 21), 4, 0.0), ("J-1003", "Decking & rails", "M. Ortiz", D(2026, 9, 25), 3, 0.0),
            ("J-1003", "Final inspection", "J. Owner", D(2026, 9, 30), 1, 0.0)]
# date, worker, job, code, hours, rate override
TIMESHEET = [(D(2026, 7, 7), "M. Ortiz", "J-1001", "02", 8, None), (D(2026, 7, 7), "D. Lee", "J-1001", "02", 8, None),
             (D(2026, 7, 8), "M. Ortiz", "J-1001", "05", 8, None), (D(2026, 7, 9), "M. Ortiz", "J-1001", "05", 8, None),
             (D(2026, 7, 9), "D. Lee", "J-1001", "05", 8, None), (D(2026, 7, 10), "D. Lee", "J-1001", "05", 4, None),
             (D(2026, 7, 11), "D. Lee", "J-1001", "05", 4, 57), (D(2026, 8, 4), "M. Ortiz", "J-1001", "14", 10, None),
             (D(2026, 8, 5), "M. Ortiz", "J-1001", "14", 10, None), (D(2026, 8, 11), "D. Lee", "J-1001", "14", 9, None),
             (D(2026, 8, 12), "D. Lee", "J-1001", "14", 9, None), (D(2026, 8, 13), "D. Lee", "J-1001", "14", 8, None),
             (D(2026, 8, 24), "S. Reyes", "J-1001", "14", 8, None), (D(2026, 8, 25), "S. Reyes", "J-1001", "14", 8, None),
             (D(2026, 8, 20), "J. Owner", "J-1001", "01", 22, None),
             (D(2026, 6, 2), "D. Lee", "J-1002", "02", 10, None), (D(2026, 6, 9), "M. Ortiz", "J-1002", "14", 6, None),
             (D(2026, 6, 10), "M. Ortiz", "J-1002", "14", 10, None), (D(2026, 6, 11), "M. Ortiz", "J-1002", "14", 10, None),
             (D(2026, 6, 12), "M. Ortiz", "J-1002", "14", 8, None), (D(2026, 6, 15), "M. Ortiz", "J-1002", "14", 6, None),
             (D(2026, 6, 19), "D. Lee", "J-1002", "14", 15, None), (D(2026, 6, 22), "J. Owner", "J-1002", "01", 12, None)]
# date, job, code, vendor, desc, qty, unit cost, paid, method, receipt
JOB_COSTS = [(D(2026, 7, 8), "J-1001", "21", "Metro Dumpster", "20 yd dumpster pull 1", 1, 425, "Yes", "Card", "MD-4471"),
             (D(2026, 7, 9), "J-1001", "06", "Builders Supply", "Framing lumber & blocking", 1, 1012.40, "Yes", "Card", "BS-20913"),
             (D(2026, 7, 14), "J-1001", "10", "Bright Plumbing", "Rough-in (50%)", 1, 2100, "Yes", "Check", "#2201"),
             (D(2026, 7, 15), "J-1001", "11", "Volt Electric", "Rough-in (50%)", 1, 1950, "Yes", "Check", "#2202"),
             (D(2026, 7, 21), "J-1001", "11", "Volt Electric", "Gas line reroute (CO-1)", 1, 620, "Yes", "Check", "#2205"),
             (D(2026, 7, 22), "J-1001", "20", "City of Dover", "Building permit", 1, 450, "Yes", "Card", "P-77120"),
             (D(2026, 7, 28), "J-1001", "13", "Smooth Drywall", "Patch & finish", 1, 1600, "No", "Check", "INV 1187"),
             (D(2026, 8, 4), "J-1001", "17", "CabinetDirect", "Cabinets delivered", 1, 9800, "Yes", "ACH / Bank", "CD-55821"),
             (D(2026, 8, 11), "J-1001", "16", "Tile Outlet", "Porcelain tile 230 sq ft", 230, 6.80, "Yes", "Card", "TO-3310"),
             (D(2026, 8, 18), "J-1001", "17", "Stone & Co", "Quartz tops installed", 1, 4300, "No", "Check", "SC-0917"),
             (D(2026, 8, 24), "J-1001", "11", "Volt Electric", "Under-cabinet lighting (CO-2)", 1, 520, "No", "Check", "V-1190"),
             (D(2026, 8, 25), "J-1001", "15", "Paint Pro", "Paint & supplies", 1, 362.50, "Yes", "Card", "PP-8802"),
             (D(2026, 6, 3), "J-1002", "10", "Bright Plumbing", "Plumbing complete", 1, 2600, "Yes", "Check", "#2188"),
             (D(2026, 6, 4), "J-1002", "11", "Volt Electric", "Electrical complete", 1, 1350, "Yes", "Check", "#2189"),
             (D(2026, 6, 5), "J-1002", "20", "City of Dover", "Permit", 1, 250, "Yes", "Card", "P-76802"),
             (D(2026, 6, 8), "J-1002", "06", "Builders Supply", "Subfloor plywood (CO-4)", 1, 410, "Yes", "Card", "BS-20640"),
             (D(2026, 6, 10), "J-1002", "16", "Tile Outlet", "Tile 150 sq ft", 150, 7.10, "Yes", "Card", "TO-3188"),
             (D(2026, 6, 16), "J-1002", "18", "Fixture Depot", "Vanity, toilet, faucet", 1, 2240, "Yes", "Card", "FD-1022"),
             (D(2026, 6, 18), "J-1002", "15", "Paint Pro", "Paint & supplies", 1, 172, "Yes", "Card", "PP-8710")]
# date, job, code, from, to, purpose, miles
MILEAGE = [(D(2026, 7, 9), "J-1001", "22", "Shop", "Builders Supply → 118 Maple St", "Lumber pickup", 24),
           (D(2026, 8, 4), "J-1001", "22", "Shop", "118 Maple St", "Cabinet delivery day", 18),
           (D(2026, 6, 16), "J-1002", "22", "Shop", "Fixture Depot → 42 Orchard Ave", "Fixture pickup", 15)]
# date, job, weather, crew, work, issues, inspections
DAILY_LOG = [(D(2026, 7, 7), "J-1001", "Sunny", 2, "Demo complete; cabinets and flooring removed; hauled to dumpster.", "Found corroded gas line behind range — flagged for CO.", ""),
             (D(2026, 7, 17), "J-1001", "Cloudy", 1, "Rough plumbing/electrical inspection.", "", "City inspector — rough-in PASSED"),
             (D(2026, 8, 12), "J-1001", "Rain", 1, "Tile floor: 60% laid.", "Tile delivery short 10 sq ft; reorder placed.", ""),
             (D(2026, 6, 8), "J-1002", "Sunny", 1, "Subfloor replaced after rot found (CO-4).", "Half day lost to subfloor rot.", "")]
# co#, job, date, desc, status, markup, cost, sched days, approved date
CHANGE_ORDERS = [(1, "J-1001", D(2026, 7, 20), "Move gas line for range", "Approved", None, 900, 1, D(2026, 7, 21)),
                 (2, "J-1001", D(2026, 8, 3), "Add under-cabinet lighting", "Approved", 0.25, 640, 0, D(2026, 8, 4)),
                 (3, "J-1001", D(2026, 8, 18), "Upgrade to pot filler", "Pending", None, 480, 1, None),
                 (4, "J-1002", D(2026, 6, 8), "Replace subfloor (rot found)", "Approved", None, 750, 1, D(2026, 6, 8))]
# inv#, job, date, due, amount, desc
INVOICES = [("INV-2026-014", "J-1001", D(2026, 7, 6), D(2026, 7, 13), 12750, "Deposit — kitchen remodel"),
            ("INV-2026-019", "J-1001", D(2026, 8, 5), D(2026, 8, 19), 15000, "Progress billing 1 — rough-in complete"),
            ("INV-2026-023", "J-1001", D(2026, 8, 28), D(2026, 9, 11), 8000, "Progress billing 2 — cabinets & tile"),
            ("INV-2026-009", "J-1002", D(2026, 6, 1), D(2026, 6, 8), 6000, "Deposit — bathroom remodel"),
            ("INV-2026-012", "J-1002", D(2026, 6, 24), D(2026, 7, 8), 9137.50, "Final — bathroom remodel incl. CO-4")]
# date, job, inv#, amount, method, ref
PAYMENTS = [(D(2026, 7, 8), "J-1001", "INV-2026-014", 12750, "ACH / Bank", "Deposit"), (D(2026, 8, 12), "J-1001", "INV-2026-019", 15000, "Check", "#2210"),
            (D(2026, 6, 3), "J-1002", "INV-2026-009", 6000, "Zelle", "Deposit"), (D(2026, 7, 1), "J-1002", "INV-2026-012", 9137.50, "Check", "#1187")]
SELECTORS = {"Proposal!I2": "J-1001", "Job Profit Report!C4": "J-1001", "CO Form!I2": 2, "Invoice Print!I2": "INV-2026-023",
             "Schedule!B2": D(2026, 7, 6), "Timesheet!Q3": D(2026, 7, 12), "Client Statement!L2": "Alvarez residence"}
