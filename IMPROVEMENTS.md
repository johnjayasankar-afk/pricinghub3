# Landing page improvement log

The site is the single static page at `index.html` (plus `assets/`), deployed to Vercel from the repository root. It sells a 24-sheet contractor job-costing workbook. Its users are small contractors deciding whether to buy; their jobs on this page are: understand what the workbook does, see the real sheets, confirm it runs on their spreadsheet app, and buy.

## Cycle 1 — 2026-09-07

### Delivered
- **Visual identity.** Manrope for display type, JetBrains Mono for every figure (tabular numerals), system sans for body. Navy/teal palette carried over from the product itself, with an amber accent reserved for warnings and focus rings. Blueprint-grid texture in the hero and buy band, a tilted hero card that settles on hover, consistent 16/22 px radii, one shadow scale.
- **Profit check (new feature).** A live calculator that runs the same arithmetic as the workbook: cost × markup, contingency on the subtotal, overhead as a share of revenue, net profit after an overrun, break-even overrun, and a dollar-split bar. Tiles turn amber or red as the job stops paying for itself. Example numbers are labelled as such; nothing is stored or sent.
- **Sheet preview (new feature).** Every tour image opens in a native `<dialog>` with previous/next buttons, arrow-key navigation, a counter, Escape and backdrop close, focus returned to the card, and next-image preloading.
- **"Will it run?" chooser (new feature).** Segmented control for Microsoft 365, Excel 2016–2021, Excel for Mac, iPad/web and Google Sheets, each with honest what-works / caveats / how-to-start content drawn from the test records (including that Sheets recalculation is not yet observed in our own tests).
- **Interaction quality.** Focus-visible rings, skip link, `aria-live` regions on the results and chooser, `aria-pressed` on the segmented control, reduced-motion support (no reveal animation, no transitions), scroll-reveal with a 2.5 s safety net so nothing can stay hidden, sticky mobile buy bar that appears once the hero scrolls away, hover/active states on all buttons and cards.
- **Honesty.** With no shop links configured the buy band shows one "Coming soon on Etsy" button and an explanatory line; the PayPal button appears only once its link exists. Proof numbers (24 sheets, 893 checks, 46 probes) come from the test records, not marketing.

### Verified
- Served locally from the repo root. Page, images, `robots.txt` and `vercel.json` return 200.
- Second critique fixed two defects: the cost input showed a red invalid border because a `step="100"` constraint rejected 36,860 (all inputs now accept any decimal), and the amber caveat lines in the platform chooser were too faint (now amber-ink text).
- Calculator outputs for the default inputs match an independent Python recomputation exactly ($46,831 proposal, 21.3 % gross margin, 11.3 % net, $1,602 after a 10 % overrun, 14.3 % break-even). A 20 % overrun turns the overrun tile red.
- Lightbox opens on click, advances with the right arrow, shows "4 of 9", closes cleanly.
- Chooser switches panels; Google Sheets panel renders its caveats.
- No horizontal overflow at 375 px or 1280 px; sticky bar displays at mobile width; Google Fonts loaded.
- No console errors after reload (earlier connection-refused entries came from a stale local server, not the page).

### Limitations / dependencies
- Buy links depend on the Etsy listing and optional Payhip product existing; fill `LINKS` in `index.html`.
- Google Fonts is an external request; the page falls back to system fonts if it is blocked.
- The calculator uses one blended markup; the workbook applies markup per category.
- Nothing has been deployed by this cycle; deployment is the owner's action.

### Next-cycle opportunities as of cycle 1
1. A comparison table using only verifiable facts. 2. A downloadable sheet map. 3. Self-hosted fonts. 4. A Gantt video.

## Cycle 2 — 2026-09-07

### Delivered
- **Interactive sheet map (new feature)**, replacing the static pill row. All 23 working sheets are laid out by stage (Setup, Inputs, Documents, Reports). Selecting one turns it navy, turns the sheets that feed it amber and the sheets that use it teal, dims the rest, and opens a detail panel with a one-line description and both lists. Hovering previews a sheet when nothing is selected; click toggles; Clear resets.
- **Truthful data.** The links are not hand-written: a script read every cross-sheet formula reference out of the built workbook and generated the map data, so "Estimate feeds Proposal, Jobs, Job Profit Report, Reports, Cost Library and Health Check" is exactly what the formulas say.
- **Printable export.** "Print the map" applies a print stylesheet that prints only the map (chips undimmed, detail panel, a title and key), which covers the "downloadable sheet map" idea from cycle 1 without shipping a separate PDF.
- **Keyboard and screen-reader support.** Chips are real buttons with `aria-pressed`; arrow keys, Home and End move between them; the detail panel is a live region; the map has a group label describing the keys.
- Second critique: the legend wording ("Feeds it" / "Fed by it") was ambiguous against the detail panel; it now reads "Feeds the selected sheet" / "Uses the selected sheet", and the detail lists are "Fed by" / "Used by".

### Verified
- Fresh load: 23 chips render in four stage columns. Selecting Estimate marks Settings, Cost Library, Jobs and Proposal as feeders and Cost Library, Jobs, Proposal, Job Profit Report, Reports and Health Check as users; the detail panel title updates. Arrow-down moves focus to the next chip. The print media rule is present in the stylesheet.
- Desktop (1280 px) and mobile (375 px) renders inspected; no horizontal overflow at either width; two-column chip layout on mobile.
- Profit check still calculates ($46,831 default). No failed resource loads.

### Limitations / dependencies
- Print output was verified by rule presence, not by an actual print preview in this environment.
- Health Check is fed by every input sheet, so it lights up for most selections; that is accurate but visually busy.
- Cycle-1 limitations still apply: buy links, Google Fonts, blended markup in the calculator.

### Next-cycle opportunities as of cycle 2
1. Self-hosted fonts. 2. Fact-only comparison. 3. First-ten-minutes stepper. 4. Gantt video.

## Cycle 3 — 2026-09-07

### Delivered
- **First ten minutes (new feature).** The Quick-Start Guide's setup path as an eight-step walkthrough between the sheet map and the tour: Settings, Vendors & Crew, Jobs, Estimate, Proposal, Timesheet / Job Costs, Invoices / Payments, Dashboard. Each step shows a crop of the real sheet from the sample workbook (eight new images generated from the Excel renders, 668 KB total), a sheet tag, a one-sentence instruction, a progress bar and a step counter. Steps are real buttons with `aria-current="step"`; arrow keys, previous/next buttons and clicks all work; the next image is preloaded; the crossfade is skipped under reduced motion.
- **Why this over the alternatives.** Self-hosting fonts requires downloading font files from a third party, which this cycle does not do without the owner's approval. A comparison table would need claims about other products that cannot be verified here. The walkthrough answers the buyer's biggest unspoken question, "will I actually be able to set this up?", with the product's own screens.
- Second critique: on phones the horizontal step cards wrapped long titles into tall boxes, so at small widths the list becomes compact numbered pills labelled by sheet name while the title lives in the stage card; on desktop the stage card is sticky so it stays beside the list while scrolling. Navigation gained a "First ten minutes" link.

### Verified
- Fresh load at 1280 px: eight steps render; Next moves to step 2 of 8 with the matching title; clicking step 5 then pressing the right arrow selects and focuses step 6 with the progress bar at 75 %; two Previous clicks return to step 4.
- Mobile (375 px): no horizontal overflow; the step strip scrolls horizontally; the stage card stacks below.
- Script syntax checked with Node; every id referenced by the script exists.
- Cycle 1 and 2 features unaffected (calculator and sheet map still present on the same page).

### Limitations / dependencies
- Step images are static crops of the sample workbook, not live spreadsheets.
- Self-hosted fonts remain undone pending approval to download the font files.
- Print output of the sheet map is still verified by rule presence only.

### Next-cycle opportunities as of cycle 3
1. Self-hosted fonts (needs approval). 2. Cost-code picker by trade. 3. No feedback widget without a backend. 4. Gantt video.

## Cycle 4 — 2026-09-07

### Delivered
- **Removed the eight-card feature grid.** It was the last generic template block and duplicated what the sheet map and the walkthrough now show better. The page is shorter and every section earns its place.
- **Trade fit check (new feature).** Eight trades (remodeler, handyman, painter, roofer, electrician, plumber, landscaper, flooring & tile). Picking one lists all 24 cost codes from the Settings sheet, shows which to keep and which to rename with a suggested name, keeps the category (and therefore the default markup) visible, and counts keeps vs renames. "Copy this list" puts a tab-separated code list on the clipboard for pasting into Settings, with a visible "Copied" state, an execCommand fallback, and an honest message if the browser blocks the clipboard.
- **Truthful data.** The 24 codes and categories are verified against the BLANK workbook's Settings sheet by script (exact match). The renames are labelled as suggestions, not features.
- Navigation label "Will it run?" became "Fit check" now that the area covers both app and trade.

### Verified
- Code table equals the workbook's Settings sheet (24 of 24, names and categories).
- Script syntax checked with Node; every referenced id exists.
- Browser (1280 px): Remodeler shows 24 rows and "the defaults"; Painter shows 13 renamed rows, "11 keep · 13 rename", and the first rename reads Site work & excavation → Pressure washing & prep. The feature grid is gone; the nav label reads "Fit check".
- Copy button: the sandboxed preview has no trusted click gesture, so the clipboard call was refused and the fallback message "Copy blocked: select the list instead" appeared as designed; a real click in a normal browser takes the clipboard path.
- Mobile (375 px): no horizontal overflow. Second critique: the struck-through default and the bold rename ran together on one line at phone width; they now stack, with a smaller category chip.

### Limitations / dependencies
- Suggested renames are editorial guidance from common trade scopes, not derived from customer data.
- Clipboard writes need a secure context (https or localhost); Vercel serves https, so this works in production.
- Self-hosted fonts remain pending approval.

### Next-cycle opportunities as of cycle 4
1. Self-hosted fonts (approval needed). 2. Deep links. 3. Link-preview image.

## Cycle 5 — 2026-09-07

### Delivered
- **Shareable state (new feature).** The page keeps its interactive state in the URL hash: `#calc=cost~markup~cont~oh~over`, `trade=<name>`, `step=<n>`, `sheet=<name>`. Opening such a link restores the profit check numbers, the chosen trade, the walkthrough step and the selected map sheet, then scrolls to the first thing the link names. State is written with `history.replaceState` (no history spam) as you interact, and a "Copy link to this check" button under the profit check puts the exact URL on the clipboard with a visible confirmation and an honest fallback when the clipboard is blocked. The separator is `~` so links stay readable without percent-encoding.
- **Link-preview image.** `assets/og.jpg` (1200×630) composed from the hero: headline, flow line, four checks, price, and the dashboard card. Open Graph and Twitter card meta tags point at it, with width and height. README explains that the two image URLs must be made absolute after the first deploy because scrapers do not resolve relative paths.
- Second critique: the first draft of the preview image ran the headline into the dashboard card, and the first share links carried `%2C` between numbers; both fixed.

### Verified
- Loading `#calc=50000,25,5,12,15&trade=Roofer&step=3&sheet=Invoices` restored cost 50,000 and overrun 15 % (proposal $65,625), selected Roofer (15 renames), step 3 of 8, and the Invoices sheet on the map. Editing the overrun to 30 % rewrote the hash within 150 ms. The share button hit the clipboard block in the sandbox and showed "Copy the address bar" as designed.
- Preview image inspected at full size after the fix. Script syntax checked with Node.

### Limitations / dependencies
- The scroll-to-target on load could not be observed in the hidden preview pane (scrolling does not render there); the call is a plain `scrollIntoView`.
- Link previews will only work once the meta URLs are absolute on the deployed domain.
- Fonts still load from Google, pending approval to fetch the files.

### Next-cycle opportunities as of cycle 5
1. Self-hosted fonts (approval needed). 2. Printable profit check. 3. Section shortcuts.

## Cycle 6 — 2026-09-07

### Delivered
- **Mobile navigation (usability fix).** Below 900 px the header had collapsed to the brand alone, leaving phone visitors, most of Etsy's traffic, with no way to jump between sections. A disclosure button now opens a full-width menu with six section links (each with a one-line hint) and the buy button. It moves focus to the first link, closes on Escape, backdrop tap, link tap or resizing to desktop, returns focus to the button, and locks page scroll while open. Reduced motion disables its slide-in.
- **Printable profit check (new feature).** "Print this check" next to the share button prints only the calculator, in black on white with a dated title line, using a `body[data-print]` switch shared with the sheet-map print. Handy for handing a margin check to a partner or client.
- **Accessibility pass.** An in-page audit (alt text, accessible names, label associations, heading order, contrast on 24 sampled text styles) found three small-text runs at 3.0:1 and one heading-order jump; all fixed (muted grey instead of faint grey; the print title is no longer a heading). Images without alt: 0. Unnamed buttons or links: 0. Inputs without labels: 0.
- **Performance.** Hero image preloaded with high priority; every tour image carries explicit dimensions and async decoding so the gallery no longer shifts layout as images arrive. Page HTML is 80 KB with everything inline.
- Defect caught and fixed during verification: the new menu code ran before the script's element helper was declared, which threw and would have halted every later feature. The helper is now declared first.

### Verified
- Mobile (375 px): menu button visible; opening sets `aria-expanded="true"`, shows the panel, focuses the first link and locks scroll; Escape closes and focuses the button; tapping a link closes it and unlocks scroll. Screenshot inspected.
- Desktop (1280 px): menu button hidden, inline nav shown, print button present, preload link present. Calculator still computes.
- Script syntax checked with Node; every referenced id exists.

### Limitations / dependencies
- Print output verified by rule presence and DOM switch, not an actual print preview in this environment.
- Fonts still load from Google, pending approval to fetch the files.

### Next-cycle opportunities as of cycle 6
1. Self-hosted fonts (approval needed). 2. Version note near the buy band. 3. Trim the tour.

## Cycle 7 — 2026-09-07

### Delivered
- **Coherence pass instead of a new feature.** A full top-to-bottom review at 1280 px measured the page at 9,412 px tall with the tour alone at 1,712 px, and found three of its nine images duplicating content now covered elsewhere (timesheet in the walkthrough, tax/vendors/health in the map, "what you get" in the Included section). The tour is trimmed to the six strongest sheets and its intro says so; the lightbox counter follows.
- **Version and updates line** in Included, from the workbook's own version stamp and Etsy's re-download behaviour, so returning buyers know updates arrive on the same listing.
- **Real download facts in the buy band**: three files, about 0.9 MB in total (measured from the product files), no installer, no account.
- Second critique: at phone width the six square tour cards still stacked into about 2,700 px of scrolling, so below 760 px the gallery is a horizontal snapping row (84 % wide cards, the next one peeking) with a swipe hint; the lightbox is unchanged. Phone page height went from 16,301 px to 14,198 px.

### Verified
- Gallery renders 6 cards; opening the sixth shows "6 of 6" and the right arrow wraps to "1 of 6".
- Buy band and Included text updated as intended; no horizontal overflow at 375 px.
- Desktop page height 9,412 px → 9,040 px; phone 16,301 px → 14,198 px. Gallery scroll width 1,726 px inside a 335 px viewport confirms the row scrolls. Script syntax checked with Node; every referenced id exists.

### Limitations / dependencies
- The three trimmed images remain in `assets/` and the Etsy listing set; only the page stopped showing them.
- Fonts still load from Google, pending approval to fetch the files.

### Next-cycle opportunities as of cycle 7
1. Self-hosted fonts (approval needed). 2. Critique and fix, do not add.

## Cycle 8 — 2026-09-07

### Delivered
- **Branded 404 page.** Vercel serves a root `404.html` automatically for static sites; the page now has one in the site's own identity ("That cell is empty."), showing the missed path, with a way back and a shortcut to the profit check. `noindex` set.
- **Phone hero proof strip.** The four proof figures wrapped raggedly at 375 px (two on one row, then one, then one). Below 520 px they are a 2×2 grid.
- **Robustness.** Browsers without `<dialog>` support open the tour image directly instead of throwing; a step image that fails to load no longer leaves the walkthrough card blank (the fade-in also clears on error).
- Deliberately not done: a dark-mode theme. The identity is a fixed light composition with navy sections; a second palette would double the CSS surface for little buyer value.

### Verified
- 404 page rendered locally: title, missed path echoed, both links present; screenshot inspected.
- Phone (375 px) hero proof strip re-checked as a 2×2 grid. Script syntax checked with Node.

### Limitations / dependencies
- The `<dialog>` fallback and the image-error path were exercised by reading the code, not by simulating an old browser or a broken image.
- Fonts still load from Google, pending approval to fetch the files.

### Next-cycle opportunities as of cycle 8
1. Self-hosted fonts (approval needed). 2. Critique and fix; re-run the accessibility audit.

## Cycle 9 — 2026-09-07

### Delivered
- **Buy links without touching HTML.** `links.json` at the repository root now holds `etsy`, `payhip` and `price`. The page fetches it on load (no-store) and activates the hero, sticky-bar and buy-band buttons, sets the price everywhere, and clears the "being set up" note; only `https://` links are accepted; the inline values stay as the fallback if the fetch fails. The owner can flip the shop live by editing one file in GitHub's web editor.
- **Mobile menu interaction.** Tab and Shift+Tab now cycle within the open menu (button and links) instead of wandering into the scroll-locked page, and the sticky buy bar slides away while the menu is open so it no longer overlaps the last menu item on short phones.
- **Accessibility re-audit** across every visible text style after the content changes of cycles 7 and 8 (results below).

### Verified
- With a test `links.json` (an Etsy URL and price $31): all three Etsy buttons became live links with the right labels, the price read $31 in every place, and the note cleared. Restored to empty afterwards; the buttons read "Coming soon on Etsy" again.
- Menu: with focus on the last link, Tab wrapped to the menu button; the sticky bar's transform moved it off-screen while open.
- Audit over 56 distinct visible text styles: 0 images without alt, 0 unnamed buttons or links, 0 unlabeled inputs, 0 heading jumps. Two real contrast misses fixed (section kickers 4.3:1 → darker teal; the tour zoom glyph 4.0:1 → darker teal). Two reports were false positives from translucent backgrounds the audit cannot composite (step numbers and sheet names on the navy walkthrough, which are white on navy). Script syntax checked with Node.

### Limitations / dependencies
- `links.json` is fetched client-side, so a buyer with JavaScript disabled sees the inline fallback ("Coming soon" until the owner also updates the inline values, which the README no longer requires but still allows).
- Fonts still load from Google, pending approval to fetch the files.

### Next-cycle opportunities as of cycle 9
1. Self-hosted fonts (approval needed). 2. Critique and fix only.

## Cycle 10 — 2026-09-07

### Delivered
- **Read the guide before you buy.** The four-page Quick-Start Guide (the same PDF that ships with the product, 11 KB) is now published at `assets/quick-start-guide.pdf` and linked from the Included list. It lets a buyer see the setup path, the sheet-by-sheet walkthrough and the FAQ before paying, without giving away the workbooks. The Vercel ignore rules were tightened so the guide deploys while every workbook stays out.
- Second critique: the "read it first" link sat inside a bullet where a scanning buyer would miss it; it is now a visible button under the Included list with the format and length in the label.
- **Product structured data.** A Product JSON-LD block carries the factual description, image, sheet count and formats. An Offer (URL, price parsed from `links.json`, USD, in stock) is added at runtime only once a real `https://` shop link exists, so search engines never see a price for a listing that is not live.

### Verified
- With a test `links.json` containing a shop link: the JSON-LD gained an Offer with url, price 29.00, USD and InStock; the guide PDF returned 200 as `application/pdf`; the Included link opens in a new tab with `noopener`. `links.json` restored to empty afterwards.
- Vercel upload simulation: 27 files deployed, the guide included, no workbook or product folder leaks. JSON-LD parsed as valid JSON. Script syntax checked with Node.

### Limitations / dependencies
- The runtime Offer depends on JavaScript, which Google executes but some other crawlers do not; the static block is still valid Product data without it.
- Fonts still load from Google, pending approval to fetch the files.

### Next-cycle opportunities as of cycle 10
1. Self-hosted fonts (approval needed). 2. Critique and fix only.

## Cycle 11 — 2026-09-07

### Delivered
- **Editorial pass over all 193 lines of visible copy** plus the metadata search engines read. Findings and fixes: the meta description ran 173 characters (search results cut around 155) and is now 155; the tour heading still said "Every sheet" after cycle 7 trimmed it to six and now says "Six sheets"; the platform chooser said "Excel iPad / web" beside "Excel for Mac" and now reads "Excel for iPad / web"; one FAQ answer joined two claims with a comma and now uses sentences. Terminology probes confirmed consistent use of "cost code", "change order", "one-time", "Google Sheets" and "Job Profit Report" throughout; no double spaces or stray punctuation (the one "space before period" hit was the filename ".xlsx").
- **Regression sweep** of every interactive feature after ten cycles of change (results below).

### Verified
- Desktop: description length 155; new tour heading and chooser label rendered; calculator recomputes on input; sheet map selects Invoices; walkthrough advances to step 2; trade chooser switches to Plumber; lightbox opens the third card with "3 of 6"; the URL hash carries the state (written after a 150 ms debounce). The console tool accumulates entries across loads and still lists the cycle-6 `$` ordering error under its old `?c6` URL; no entry belongs to this load, and every feature working confirms no load-time exception.
- Phone: menu opens and closes with Escape; gallery is the horizontal row; no horizontal overflow.

### Limitations / dependencies
- Fonts still load from Google, pending approval to fetch the files.

### Next-cycle opportunities as of cycle 11
1. Self-hosted fonts (approval needed). 2. Critique and fix only.

## Cycle 12 — 2026-09-07

### Delivered
- **Responsive images.** The hero and the six tour images now ship a 700 px variant (476 KB for all seven) alongside the 1400 px original, with `srcset` and `sizes` that describe the real slot (540 px on desktop, 92 vw on phones for the hero; 360 px or 84 vw for tour cards). The hero preload carries the same candidates. Phones and 1× screens download roughly half the image bytes; 2× desktop screens still get the full-resolution file.
- **No reflow when the web fonts arrive late or fail.** Two fallback faces declared with `@font-face` map Manrope to local Arial and JetBrains Mono to local Menlo/Consolas with `size-adjust`, `ascent-override` and `descent-override`. The values are measured, not guessed: a 40 px sample line was rendered in each web font and its fallback in the browser and the adjustments tuned until widths matched (results below). This removes the layout shift that the Google Fonts dependency could cause, without downloading any font file.

### Verified
- Font metrics after tuning (40 px sample line, measured in the browser): Manrope vs fallback −0.2 % at weight 800, +1.5 % at 700, −1.9 % at 400; line height 54.5 px vs 55.5 px. JetBrains Mono vs fallback +0.3 % at 500 and 700; line height identical. The first attempt was 7 % and 29 % off because `local()` needs exact face names; the fallbacks are now declared per weight range with the bold faces named explicitly.
- Hero `srcset`/`sizes` present and parsed by the browser; deployment simulation ships the seven new variants with no workbook leaks; script syntax checked with Node.
- Candidate selection could only be partly observed here: the pane runs at 2× and had the 1400 px file cached, so it kept it; the `sizes` maths for a 375 px phone (345 CSS px × 2 = 690 device px) selects the 700 px file on a fresh load.

### Limitations / dependencies
- The fallback faces assume Arial (or Liberation Sans) and Menlo/Consolas/DejaVu Sans Mono exist locally; where none does, the stack falls through to the system font as before.
- Fonts still load from Google, pending approval to fetch the files; the fallbacks now make that dependency far less visible.

### Next-cycle opportunities as of cycle 12
1. Self-hosted fonts (approval needed). 2. Critique and fix only.

## Cycle 13 — 2026-09-07

### Delivered
- **A test harness for the site itself** (`tests/site_audit.py`), matching the workbook's testing culture. It checks that every local asset the page references exists (17 references), that each `srcset` descriptor matches the real image width, that `links.json` parses with the expected keys, `https://` links and a `$29`-style price, that `vercel.json` parses and `.vercelignore` still excludes every product folder and workbook, that the title and meta description lengths are in range, that the Product JSON-LD parses and the preview image declares its size, that there is one `h1`, no duplicate ids, no id used by the script that is missing from the page, no `http://` resource, that `404.html` exists, and that the page script passes Node's syntax check. Exit code 1 means do not deploy. README documents it as the pre-deploy step.
- The page itself did not change this cycle, so no browser verification was needed; the value is that future owner edits (especially to `links.json`) and future cycles are caught before they reach production.

### Verified
- Clean run: 17 references checked, script syntax OK, 0 problems.
- Negative test: with an `http://` shop link, a bare "29" price, a wrong `srcset` width and a renamed id injected, the audit reported each and exited 1; all injected faults were reverted and the clean run repeated.

### Limitations / dependencies
- The audit is static; it does not render the page or measure layout. Browser checks remain the manual part of each cycle.
- Fonts still load from Google, pending approval to fetch the files.

### Next-cycle opportunities as of cycle 13
1. Self-hosted fonts (approval needed). 2. Critique and fix; run the site audit each cycle.

## Cycle 14 — 2026-09-08

### Delivered
- **A product defect found and fixed through the site.** Preparing to publish the real client documents exposed that the print areas of the Proposal, Invoice Print, CO Form and Client Statement sheets included row 2, so every printed client document carried the "← pick a job…" picker and its hint at the top. The pickers now sit in columns H–I (K–L on the statement), outside every print area, with the hint beneath them. Layout selectors, sample data, first-open cursors and the Quick-Start Guide were updated; the Excel harness passes 893/893 after the change; protection and Sheets audits are clean; all ten listing images and the walkthrough's proposal image were re-rendered without the picker row. The Etsy deliverables and launch package were rebuilt, and the test record notes the defect as number 21.
- **The documents your clients receive (new on the site).** Four PDFs exported from the sample workbook with its own print areas (proposal, invoice, change order, client statement) are linked from the tour as document cards. A buyer can open exactly what their own client would get. They are produced by the render pipeline in a `SAMPLE_DOCS=1` mode (`listing/export_samples.py` wraps it) that swaps the four documents' crop areas for their full print areas, because the normal image crops would have cut the proposal's terms and signature lines and the statement's payment instructions. A separate whole-workbook export was tried first and abandoned: fitting the 1,000-row input sheets to single pages stalls Excel, and an AppleScript `repeat` over worksheets raises a parameter error in this Excel build.

### Verified
- Workbook: 893/893 harness, 0 protection problems, Sheets audit clean, listing-copy and listing-image audits pass; the re-rendered Proposal shows no picker row.
- Site: the four document cards render with `target=_blank`; HEAD requests to each PDF return 200 as `application/pdf`; the site audit passes (0 problems). Mobile: no horizontal overflow.
- Documents: text extraction confirms the proposal carries Terms, Exclusions and both signature lines; the change order carries Terms and signatures; the invoice and statement carry payment instructions; none contains the picker hint.
- Second critique: the full-page proposal made a weak listing image (empty detail rows dominate), so the listing keeps its tight crop and only the PDF uses the full page.

### Limitations / dependencies
- The sample documents show the fictional Northline Remodeling data and are labelled as such on the page.
- Fonts still load from Google, pending approval to fetch the files.

### Next-cycle opportunities as of cycle 14
1. Self-hosted fonts (approval needed). 2. Critique and fix only.

## Cycle 15 — 2026-09-08

### Delivered
- **Touch targets on phones.** A scripted audit at 375 px measured every visible link, button, input and summary: 44 of 81 were under the 44 px guideline (segmented chooser buttons at 40 px, soft buttons such as Copy, Print and Clear at 32–34 px, the 23 sheet-map chips at 32 px, the lightbox and step arrows at 40 px, the brand link at 28 px). Below 900 px, every control now has a 44 px minimum height (and width for icon buttons); the sheet-map chips use 40 px so the 23-chip grid still fits one screen. Desktop is unchanged.
- The local site is left running at http://localhost:8765 for inspection, as requested.

### Verified
- Re-measured at 375 px after the change: 81 controls checked, 23 under 44 px, all of them the map chips at exactly 40 px; every other control is 44 px or taller. Site audit passes; no horizontal overflow; the trade section's Copy button inspected visually.

### Limitations / dependencies
- The 40 px chips are above WCAG's 24 px minimum but below the 44 px guideline by design.
- Fonts still load from Google, pending approval to fetch the files.

### Next-cycle opportunities (highest value first)
1. Self-hosted fonts (approval needed).
2. Critique and fix only.

## Cycle 16 — 2026-09-13

### Delivered
- **The site is PricingHub.** Every place the page named itself now says PricingHub instead of "Contractor Job Costing":
  - the header wordmark and its label;
  - the page title, the Open Graph title and the Product JSON-LD name;
  - the footer and the 404 page title;
  - the printed sheet-map heading, the copied cost-code list and the printed profit check.
- **A new hero image.** The old hero image had the old name baked into its banner. The hero now shows a clean crop of the same dashboard render (`assets/hero-dashboard.jpg`, at 1400 and 700 px). It uses a new file name because `/assets/` is served with an immutable cache.
- **Labs family look.**
  - Inter and IBM Plex Mono are self-hosted in `assets/fonts/`, copied from the owner's own site files. This retires the Google Fonts request that every cycle since the first listed as pending.
  - A porcelain ground with a dot grid, and forest night with mint for the hero, the walkthrough and the buy band. PricingHub's teal stays as the accent.
  - Pill buttons and pickers, rounded cards and mono labels.
  - The favicon and the header mark share the family tile with the three-bar glyph, and the 404 page matches.
- **New link-preview card:** `assets/og-pricinghub.jpg` (1200×630). `og:image` and `twitter:image` are absolute URLs on pricing-hub-seven.vercel.app, and `og:url` is set.
- **The product folders were being served; now they stay off Vercel.**
  - The repository had no `.vercelignore`. The live site therefore returned both workbooks from `deliverables/` and `product/`, plus `records/`, `tests/`, `listing/`, `model/` and this log.
  - `.vercelignore` is restored. It is a hidden file and easy to lose when uploading through Finder.
  - As a second lock, `vercel.json` redirects those folders, and any `.xlsx`, `.md` or `.py` path, to the 404 page. Vercel applies redirects before it serves a file.
- **A dead buy button, fixed.** `.btn` set `display`, which overrode the `hidden` attribute. So whenever no PayPal link was set, "Buy with PayPal" still appeared and linked to `#`. A `[hidden]` rule restores the intended single button.
- A footer credit to John Jayasankar and Labs.

### Verified
- **Audits:**
  - The site audit reports 0 problems across 22 local references.
  - Every selector in the old stylesheet is still present, and no layout property was removed (360 selectors before, 371 after).
- **Contrast, over the whole page at 1440 and 390 px:**
  - Before, three text styles sat at or under 4.5:1: the sheet map's stage label (4.3), the net-margin figure (4.49) and white on the teal buttons (4.64).
  - After, the only text under 4.5:1 is the disabled "Coming soon on Etsy" button (3.43, up from 2.51).
  - Every other text is at least 5.6:1. The menu, lightbox, walkthrough and 404 page are at least 6.2:1.
- **Layout and behaviour:** there is no horizontal overflow at 1440 or 390 px, and no console errors. These work as before:
  - the profit check (a 30 % overrun turns the tile red);
  - the sheet map, walkthrough and lightbox;
  - the fit and trade pickers and the mobile menu.
- **Buy buttons:**
  - With `links.json` empty, there is one disabled "Coming soon on Etsy" button, PayPal is hidden, and the JSON-LD has no Offer.
  - With test `https://` links, every Etsy button points at the link, "Buy with PayPal" appears, the setup note clears and the Offer is added.
  - `links.json` was restored afterwards, with no diff.
- **Redirects, on a local server applying `vercel.json`:** the workbook, records, tests, listing, model and Markdown paths return 307 to `/404`. The page, assets, `links.json` and the guide PDF return 200.

### Limitations / dependencies
- The GitHub repository is public, so the workbooks can still be downloaded from GitHub. Only the owner can make it private.
- Copies downloaded while the files were exposed cannot be recalled.
- The redirects were checked against a local approximation of Vercel's router; confirm them on the live site after the deploy.
- The Quick-Start Guide PDF, the workbook file names and the Etsy listing still say "Contractor Job Costing".

### Next-cycle opportunities (highest value first)
1. After the deploy, confirm that `/deliverables/Contractor-Job-Costing-BLANK.xlsx` on the live site no longer downloads.
2. Decide whether the product itself (the guide, the workbook file names, the listing) also takes the PricingHub name.
