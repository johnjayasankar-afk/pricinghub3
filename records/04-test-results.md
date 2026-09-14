# Test results — final build 2026-09-08

## Product: Contractor-Job-Costing-SAMPLE.xlsx / -BLANK.xlsx (24 sheets)

| Test | Method | Result |
|---|---|---|
| Formula correctness, 893 output cells across every sheet, including v4 additions (This-Month tiles, quarterly tax table, vendor unpaid balances, overtime column, library usage counts, proposal number, remittance slip, revised completion date on the CO form, overpaid-invoice check) | `tests/test_workbook.py`: recomputes every value in pure Python from `product/sample_data.py`, then drives Microsoft Excel 16.112 via AppleScript | **893/893 pass** |
| Formula-error scan on 22 sheet ranges | Health Check diagnostics read from Excel | 0 on SAMPLE, 0 on BLANK |
| Data-integrity checks (13 Check columns, incl. new "active job with no estimate or contract") | Health Check totals | 0 issues |
| Blank workbook with no selectors chosen | Excel recalc; forms blank cleanly | Pass |
| All 24 sheets render to PDF/PNG; every sheet visually reviewed at full size at least once | Excel export, page split, 2800 px raster | Pass |
| Listing images 01–10 | Visual review after final render | Pass |
| Quick-Start Guide PDF (4 pages) | Rebuilt; text verified | Pass |
| Financial model at $29 list / 25% launch sale | Excel recalc | Net per Etsy sale $18.26 (83.9%) |
| Google Sheets / Excel 2016 static compatibility audit of all 22,600 formulas, every conditional-format rule and every data validation | `tests/sheets_compat_audit.py`: 27 functions used, all on the Sheets-supported list; no dynamic arrays, XLOOKUP, IFS, LET, structured references or `_xlfn` prefixes; validations are plain lists | Pass |
| Adversarial data test: 18 bad rows injected into the SAMPLE (duplicate job/client/crew/vendor/CO/invoice IDs, unknown jobs, workers and cost codes, missing dates and quantities, an approved CO without a client date, an overpaid invoice, an 82-day-old unpaid invoice, a cost that pushes a job over budget), recalculated in Excel | `tests/test_adversarial.py`: 46 probes across all 13 Health Check integrity lines, all 8 attention items, the Check columns, aging bucket, overdue status, negative balance, rate fallback, and the Dashboard alert tiles | **46/46 pass**; Health Check total 21, formula errors 0 |
| Sheet-protection audit: every light-blue input cell unlocked, every formula cell locked, on SAMPLE and BLANK | `tests/protection_audit.py` over 70,544 input cells and 45,202 formula cells | 0 problems |
| Version stamp (Start Here B3 "Version 1.0 (2026-09)", workbook title/subject/description properties) | Rebuilt after stamp; harness rerun | 893/893 pass |

## Client-document pass (2026-09-08)
1. **Defect found while publishing sample PDFs on the landing page**: the print areas of Proposal, Invoice Print, CO Form and Client Statement included row 2, so a client would have received the "← pick a job…" picker and its hint printed at the top of every document. The pickers now live in columns H–I (K–L on the statement), outside every print area, with the hint beneath them. Layout selectors, sample data, first-open cursors and the Quick-Start Guide were updated. Excel harness after the change: 893/893. Protection audit 0 problems. Sheets audit clean. All ten listing images and the four sample PDFs re-rendered and inspected: no picker text on any client document.

## Storefront pass
1. **Etsy listing copy was over Etsy's limits**: the title ran 157 characters (limit 140) and four tags ran 21–22 characters (limit 20). Etsy rejects both at upload, so this would have stopped the owner at step 2. Rewritten to 137 characters and 13 tags of 20 or fewer; a new `tests/listing_copy_audit.py` checks title, tag count, tag length, duplicate tags, file sizes against the 20 MB limit and the ten 2000×2000 images.
2. Hero image re-laid so the dashboard card and the feature checklist both sit inside Etsy's 4:3 search-thumbnail crop (rows 250–1750); verified by cropping the image the way Etsy does.
3. Start Here: the first setup step had a phantom blank row under it (wrap height applied to a line that fits). Fixed.

## Pixel pass on the listing images
Every crop now snaps to a table gridline (no row is ever sliced), every canvas is filled deliberately with captioned panels (Schedule + Daily Log + Mileage; Timesheet + Crew + Job Costs; Reports + Payments), colliding captions were shortened, the Health Check and Vendors panels end on complete rows, and a new `tests/listing_image_audit.py` verifies that no card touches the title band, footer or canvas edges. One Settings header column was widened so "Payment method" no longer clips.

## Listing-image pass
All ten listing images re-checked after the layout changes. Four had panels cut mid-row or text spilling out of a card (Estimate, Timesheet/Job Costs, Change Orders, What-you-get). Crops now end on row boundaries and cards size to their text. The financial model workbook opens and recalculates in Excel; its PDF export hangs Excel's automation, which is irrelevant to buyers (owner-only file) and noted in the runbook.

## Robustness pass
No defects found. The adversarial test and the protection audit passed on their first run, which is the evidence that the Check columns and Health Check behave as documented under bad data.

## Defects found and fixed in the previous pass (BLANK workbook opened as a buyer would)
1. With Settings blank or no selector picked, Proposal, Invoice Print, CO Form and Client Statement showed stray "0" values in the company block, "0" for the invoice number and a lone dash on the job line. All echoes now stay blank. Verified in Excel on the BLANK file.
2. First-open polish: every sheet now opens with the cursor on its selector or first input cell, wide tables open at 90 % zoom, Start Here is the selected tab. Quick-Start Guide carries the version stamp.
3. Clients, Cost Library and Payments sheets rendered and reviewed for the first time: no defects.

## Defects found and fixed in the visual pass (every sheet re-rendered from Excel and reviewed at full size)
1. Dashboard charts rendered without axes or labels in current Excel builds (openpyxl leaves the axis "delete" flag unset). Axes, $-formatted scales and rotated category labels added; legend moved clear of the title.
2. Dashboard charts overlapped the Jobs table header. Table moved down two rows.
3. Clipped text: Invoice Print and CO Form job line, Proposal client email, Client Statement job column, Settings tax-scope note and proposal terms, Job Profit Report "Unpaid bills" label, Vendors email and trade columns. All widened, wrapped or merged.
4. Proposal detail table was narrower than the summary above it; now shares the same right edge (line prices in column F).
5. Daily Log wrapped rows had misaligned hours. Formula cells top-aligned.
6. CO Form selector showed a right-aligned number; now left-aligned with an input prompt.
7. Invoice Print gains a "PAID IN FULL" line that appears only when the selected invoice is fully paid.

## Defects found and fixed in the previous pass
1. Settings: tall document-text rows shared rows with the cost-code table, stretching codes 17–24; notes clipped. Layout reordered and columns widened.
2. Mileage "To" column clipped. Widened.
3. CO Form total formula referenced the wrong rows after the new "Revised target completion" line was added (caught by the harness: 4 failures, fixed).

## Cumulative defects caught by testing and visual review across all versions: 21 (see earlier entries in git-less history within this file's previous versions: CF fills, #VALUE! on blank selectors, false red rows, stretched timesheet row, truncated labels, negative days-to-target, narrow columns, clipped chart, PDF export omission, sample-data column collision, and the two above).

## Not verified
- Google Sheets *execution*. Two live attempts this session: (1) `tests/sheets_probe.xlsx` (20 formulas covering every function family the product uses, a named range, a formula conditional format and a list validation) uploaded through the Drive connector converted to a native Google Sheet with no error; (2) the same probe imported from CSV also converted cleanly. Neither could be *read back* with computed values: Google evaluates a never-opened converted file lazily and the connector's export returns the formula cells empty, and the in-app browser stops at a Google sign-in that only the owner may perform. Conversion is therefore verified; recalculation in Sheets is not. Both temporary files were trashed. The owner's 2-minute check stays in the approval packet.
- Excel for Windows / Excel 2016 specifically.
- Sheet protection flow in Google Sheets.
- Etsy upload and buyer download path.
