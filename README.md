# PricingHub

PricingHub is an Etsy digital product (a job-costing workbook for contractors, Excel + Google Sheets, 24 sheets) plus its landing page, generator, tests and launch records.

**Keep this repository private.** `product/` and `deliverables/` contain the paid workbooks. A public repo would give them away.

## Repository layout

| Folder | What it is | Deployed to Vercel? |
|---|---|---|
| `index.html`, `assets/`, `vercel.json`, `robots.txt` | Static landing page at the repository root | **Yes — the only files Vercel serves** |
| `product/` | Workbook generator (`build_workbook.py` and sheet modules), the built SAMPLE and BLANK workbooks, the Quick-Start Guide | No |
| `listing/` | Etsy listing copy, the ten 2000×2000 listing images, the image compositor | No |
| `deliverables/` | Exactly the files to upload to Etsy, plus the approval packet | No |
| `tests/` | Excel harness (893 checks), adversarial Health Check test (46 probes), protection / Sheets / listing audits | No |
| `records/` | Research log, scorecard, test results, operations runbook, approval packet | No |
| `model/` | Editable financial model (`financial-model.xlsx`) | No |

## Deploy the landing page to Vercel

1. Push this repository to GitHub as a **private** repo.
2. In Vercel: **Add New → Project → Import** the repo.
3. Leave every setting at its default: Framework preset **Other**, Root Directory **empty**, no build command, no output directory. The page is plain static HTML at the repository root.
4. Deploy. `.vercelignore` keeps the product, tests and records out of the upload, so only `index.html`, `404.html`, `assets/`, `links.json`, `robots.txt` and `vercel.json` are ever served. It is a hidden file: in Finder press Cmd+Shift+. to see it, and include it whenever you upload through GitHub's web uploader. As a second lock, `vercel.json` redirects the private folders to the 404 page in case the ignore file ever goes missing.
5. The `og:image` and `twitter:image` meta tags near the top of `index.html` use the absolute URL `https://pricing-hub-seven.vercel.app/assets/og-pricinghub.jpg`. If the site moves to another domain, update both (link previews in messages and social apps need absolute URLs).
6. After the Etsy listing (and optionally Payhip) exist, edit `links.json` (three fields: `etsy`, `payhip`, `price`) directly on GitHub, commit, and Vercel redeploys; no HTML editing needed. Links must start with `https://`. Until then the buy buttons read "Coming soon".

If you already created the Vercel project with **Root Directory** set to `site`, clear that field in Project Settings → General and redeploy.

## Check the site before a deploy

```bash
python3 tests/site_audit.py
```

Verifies every referenced asset exists, `srcset` widths match the files, `links.json` and `vercel.json` parse (links must be `https://`, price like `$29`), the ignore rules still exclude the product, metadata lengths, the Product JSON-LD, ids used by the script, and the script's syntax. Exit code 1 means do not deploy.

## Rebuild the product

```bash
cd product
python3 build_workbook.py Contractor-Job-Costing-SAMPLE.xlsx
python3 build_workbook.py --blank Contractor-Job-Costing-BLANK.xlsx
python3 build_guide.py
cd ..
python3 tests/test_workbook.py          # needs Microsoft Excel (macOS AppleScript)
python3 tests/test_adversarial.py       # needs Microsoft Excel
python3 tests/protection_audit.py
python3 tests/sheets_compat_audit.py
python3 listing/make_images.py --render # needs Microsoft Excel; re-exports every sheet and recomposes the ten images
python3 tests/listing_image_audit.py
python3 tests/listing_copy_audit.py
```

Python 3.11 with `openpyxl`, `Pillow`, `numpy`, `pypdf`, `reportlab`.

## Status

See `records/06-approval-packet.md`. Nothing is published, no money has been spent and no third party has been contacted until the owner says so.
