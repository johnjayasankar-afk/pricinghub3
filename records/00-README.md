# Project index — Contractor Job Costing Spreadsheet, final build 2026-09-08 (Etsy)

Status labels (evidence-based): **Researched · Prototype complete · Tested · Ready for approval.** Not yet: Published, Payment collection verified, First genuine sale, Funds available, Bank payout confirmed.

| File | What it is |
|---|---|
| records/01-research-log.md | Sources, access dates, observations, uncertainties for 25 platforms/mechanisms |
| records/02-scorecard-and-decision.md | 12-opportunity weighted scorecard, decision, stop/pivot rules |
| records/04-test-results.md | 893-check Excel cross-check, error scans, defects found and fixed, what remains unverified |
| records/05-operations-runbook.md | What runs automatically, owner time, support playbook, pause switch, update procedure |
| records/06-approval-packet.md | Go/no-go packet: what publishes, where, price, spend, accounts, owner-only steps |
| tests/test_adversarial.py | Injects 18 bad rows into the SAMPLE, recalculates in Excel, asserts 46 Health Check / Check-column / Dashboard-alert behaviours |
| tests/protection_audit.py | Static audit that inputs are unlocked and formulas locked (Protect Sheet works as documented) |
| tests/sheets_probe.xlsx | 20-formula probe used for the Google Drive conversion test (upload → converts to a Google Sheet; open it to see 'ALL 20 OK' in L4) |
| tests/listing_copy_audit.py | Etsy limits: title ≤140, 13 tags ≤20 chars, files ≤20 MB, ten 2000×2000 images |
| tests/listing_image_audit.py | Pixel audit of the ten composed listing images (cards inside the content zone, margins intact) |
| tests/sheets_compat_audit.py | Static Google Sheets / Excel 2016 compatibility audit of every formula, CF rule and validation (exit 0 = clean) |
| product/build_workbook.py + layout.py, wb_common.py, sheets_setup.py, sheets_jobs.py, sheets_costs.py, sheets_reports.py, sample_data.py | Generator for both workbooks (24 sheets). Edit → rebuild → `python3 tests/test_workbook.py` |
| product/Contractor-Job-Costing-SAMPLE.xlsx, -BLANK.xlsx, Quick-Start-Guide.pdf | The three files buyers download |
| product/build_guide.py | Generates the PDF guide |
| listing/etsy-listing.md | Title, tags, description, FAQ, policies, buyer message, Payhip copy |
| listing/make_images.py, listing/images/01–10 | Renders real sheets from Excel and composes the ten 2000×2000 listing images (`--render` to re-export) |
| model/financial-model.xlsx | Editable fees/scenarios/cash-timing model ($29 list) |
| tests/test_workbook.py, tests/excel_recalc.applescript | Excel cross-check harness |
| deliverables/, contractor-job-costing-launch-package.zip | Everything the owner needs to publish |

## Next action to resume
Owner completes steps 1–5 in records/06-approval-packet.md. Then report the listing URL; next milestone is **Published**, then **First genuine sale** within the 45-day window.

## Fallback (not built)
Apify pay-per-event website-audit actor. Account exists. Build only if the pivot trigger fires.
