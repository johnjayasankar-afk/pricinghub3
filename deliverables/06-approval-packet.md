# Approval packet — ready for your go/no-go (final build 2026-09-08)

## What will be published
One Etsy instant-download listing: **Contractor Job Costing Spreadsheet (Excel + Google Sheets, 24 sheets)** — two .xlsx workbooks (SAMPLE and BLANK) plus a 4-page PDF Quick-Start Guide. Copy, tags, ten 2000×2000 images and shop policies are finished in `listing/etsy-listing.md` and `listing/images/`.

Optional second channel: the same product on a free Payhip storefront (instant PayPal payouts, no traffic of its own).

## Where it appears
Etsy search results and your shop page. Category: Design & Templates › Templates.

## Price
$29.00 list, 25% launch sale for 30 days ($21.75 realized). Net to you per Etsy sale ≈ $18.26 (model in `model/financial-model.xlsx`).

## Exact spending
| Item | Amount | When |
|---|---|---|
| Etsy one-time shop set-up fee | $15–$29 (Etsy shows the exact amount during setup) | At shop opening |
| Etsy listing fee | $0.20 | At publish, then $0.20 per sale / per 4 months |
| Everything else | $0 | — |

Nothing has been spent. No paid ads are planned.

## Accounts and terms involved
- Etsy seller account (new): Seller Policy, Creativity Standards (AI-assisted disclosure is on in the listing copy), Etsy Payments terms.
- Payhip (optional, new): Payhip terms; connects to your existing PayPal.
- No Stripe, Apify or other account is touched for the primary plan.

## The only actions you can perform (about 30–40 minutes)
1. **Open the Etsy shop** at https://www.etsy.com/sell → "Open your Etsy shop". Etsy asks for: shop name (suggestion: "TradeSheetWorks" or your own), country/currency, identity verification through Persona (photo ID + selfie), date of birth, address, partial SSN for US individuals, bank account for deposits, a card for fees, and the set-up fee.
2. **Create the listing**: Shop Manager → Listings → Add a listing → Digital files. Paste title, description, tags and FAQ from `listing/etsy-listing.md`; upload the 10 images from `listing/images/` in order; upload the 3 files from `product/`; price $29, quantity 999; run a 25% sale via Marketing → Sales and discounts. Tick the AI-assisted disclosure.
3. **Set the auto message to buyers** (Shop Manager → Settings → Info & Appearance → Message to buyers) from the same file.
4. **Google Sheets 2-minute check** (recommended): upload `Contractor-Job-Costing-SAMPLE.xlsx` to Google Drive, open it, File → Save as Google Sheets, confirm the Dashboard shows $79,609 revised contract and $12,644 projected profit, and that the Schedule sheet shows Gantt bars. If anything errors, tell me the sheet and cell.
5. Optional: **Payhip** at https://payhip.com → sign up with your email, connect PayPal, add the product using the Payhip section of the listing file.

## Remaining material uncertainties
- Discovery: a new Etsy shop with one listing may take weeks to its first sale; evidence shows one-person shops sell in this niche, not that this listing will.
- Exact set-up fee until Etsy displays it.
- Google Sheets execution not yet verified (step 4). Every formula was statically audited against the Sheets function list, and a probe workbook with every construct the product uses converted to a native Google Sheet without error. Only a live recalculation remains unobserved, because reading it needs your Google sign-in.
- Etsy may hold a new seller's first funds beyond 14 days at its discretion.

## Stop / pivot rule
45 days after publishing: under 150 views or zero sales after two title/tag revisions → switch to the Apify fallback.

## Reply options
- "Go" → everything is ready; you complete steps 1–5.
- "Change price to $X" / "rename shop to Y" → I update copy and model.
- "Payhip only" → I trim the packet to the Payhip flow.
