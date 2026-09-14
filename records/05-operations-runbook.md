# Operations runbook — Contractor Job Costing Spreadsheet

## What runs automatically (no owner action)
- Etsy hosts the three files, charges the buyer, delivers the download instantly, emails the buyer a link, collects and remits sales tax, and deposits funds weekly (Monday) once each sale is 14 days old.
- Payhip (if enabled) delivers the files and pays PayPal instantly.
- Listing auto-renews for $0.20 on each sale and every 4 months if unsold.

## Recurring owner work (honest estimate)
| Task | Frequency | Time |
|---|---|---|
| Read and answer buyer messages (Etsy app notifications) | As they arrive; expect 0–2 per 10 sales | 5 min each |
| Check Etsy Stats: views, favorites, orders | Weekly | 5 min |
| Approve or refund a "not as described" case | Rare | 5 min |
| Review 1099-K / income for taxes | Yearly | Depends on accountant |

Total: roughly 15–30 minutes per month at low volume. This is not zero, but there is no content treadmill, no fulfillment, and no infrastructure.

## Support playbook (paste-ready replies are in listing/etsy-listing.md)
1. "File won't open" → confirm they downloaded .xlsx from the Etsy purchases page (not the email preview) and are opening in Excel 2016+/Sheets. Send the Google Sheets steps.
2. "Formulas show #NAME?" → they overwrote a formula; tell them to press Undo or re-download the BLANK file.
3. "Can you customize it?" → decline custom work politely (keeps the business passive), point to Settings and the FAQ.
4. "I want a refund" → Etsy digital items are non-returnable, but a goodwill refund on a $18 sale costs less than a 1-star review. Refund if the complaint is plausible. Budget: 3% of sales.

## Pausing sales
- Etsy: Shop Manager → Settings → Options → Vacation mode (listing stays but cannot be bought), or deactivate the listing.
- Payhip: Products → toggle "Unpublished".

## Health checks
- Monthly: buy-path check by opening the listing logged out (files upload state, price, images). Etsy shows "This listing has files" in Shop Manager.
- If Etsy flags the listing for policy review, respond within the deadline they give (they email).

## Updating the product
- Edit `product/build_workbook.py`, rebuild SAMPLE and BLANK, run `python3 tests/test_workbook.py`, `python3 tests/test_adversarial.py` (both need Excel installed), `python3 tests/protection_audit.py`, `python3 tests/sheets_compat_audit.py`, and after any listing change `python3 tests/listing_copy_audit.py` and `python3 tests/listing_image_audit.py`, regenerate `Quick-Start-Guide.pdf`, upload the new files to the listing (Etsy → Edit listing → Files). Buyers who already purchased can re-download the latest files from their purchases page.

## Records to keep
- Etsy monthly statements (Finances → Monthly statements) are the revenue evidence. Payhip transactions export as CSV.
- Keep `records/` in this workspace for the research trail and decisions.

## Kill switch and pivot
- If 45 days pass with under 150 views or zero sales after two title/tag revisions, execute the fallback in `records/02-scorecard-and-decision.md` (Apify pay-per-event actor) rather than building another template blind.

## Known automation quirk
model/financial-model.xlsx: do not export it to PDF through AppleScript; Excel raises a parameter error and leaves the file open, which blocks later automated exports until Excel is quit. Read it in Excel directly.
