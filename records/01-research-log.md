# Research log

Access date for all entries: 2026-09-04 (US Eastern). "Uncertain" notes what the source did not settle.

## Environment (verified locally)

- macOS, US Eastern timezone, en_US locale. Node 22-era toolchain, Python 3.11, Microsoft Excel installed. No LibreOffice, no Apify CLI, no gh.
- Existing repos in ~/Documents/GitHub (AgentFit, RailDrop, RideLens, jobSearchOS, portfolio) are prototypes needing paid infrastructure (Supabase, Parse, Playwright). Not ready-to-sell assets.
- Gmail (read-only search, task-relevant queries only):
  - Apify account exists (welcome email 2026-04-27).
  - A Stripe login exists (password-set notice 2026-08-20). Activation/verification status unknown.
  - PayPal account exists (legal-agreement notice 2026-07-24).
  - No evidence of Gumroad, Payhip, Lemon Squeezy, Ko-fi, itch.io, or an Etsy seller account.

## Platforms and payout rules

| Platform | Source | Observation | Uncertain |
|---|---|---|---|
| Etsy fees | https://help.etsy.com/hc/en-us/articles/115014483627 | One-time non-refundable set-up fee, "cost varies by location"; $0.20 listing; $0.20 auto-renew on each sale of a digital listing; 6.5% transaction fee; payment processing fee "set rate plus percent" varies by country (US: 3% + $0.25 per third-party summaries); Offsite Ads 15% if opted in / under $10k. | Exact US set-up fee (third-party sources say $15–$29). |
| Etsy deposits | https://help.etsy.com/hc/en-us/articles/360046998234 (via search summary) | New sellers: funds eligible for deposit 14 days after a sale; default weekly Monday deposit; 5-day hold after changing bank details; deposit minimum applies. | Reserve policy for new shops is case-by-case. |
| Etsy identity | https://help.etsy.com/hc/en-us/articles/22481159004567 ; .../360001980067 | Persona ID + selfie; name, DOB, address, partial SSN (US); bank account in same name. Shop paused if not completed. | None material. |
| Etsy digital items | https://help.etsy.com/hc/en-us/articles/115015628347 ; https://www.etsy.com/legal/creativity | Must be made/designed by seller; AI-assisted creations must disclose AI; up to 5 files, 20MB each; digital listings cannot be returned; sellers cannot set return policies on digital listings; Purchase Protection covers "not as described". | How strictly AI disclosure is enforced for spreadsheets. |
| Etsy demand (live search, in-app browser) | etsy.com/search for "budget spreadsheet excel template", "rental property spreadsheet excel", "freelancer tax spreadsheet quarterly estimated", "airbnb income expense tracker spreadsheet", "contractor job costing spreadsheet" | Bestseller badges on many digital spreadsheet listings; personal budget templates saturated with $0.99–$2 offers and 50–75% discounts; rental/Airbnb trackers $1.35–$15 with heavy discounting; contractor/job-costing templates priced $9.97–$97 with Bestseller badges at $47 and fewer near-duplicate listings. | Etsy does not show sales counts on search pages; Bestseller badge is a relative signal, not a number. |
| Etsy first-sale timing | https://closo.co/blogs/beginner-guides-how-tos/how-to-sell-digital-downloads-on-etsy-2 ; Etsy community threads "Shops Waiting On First Sale" | Third-party guides say first sales commonly within 30–90 days; community threads show many new shops waiting weeks. | Anecdotal; no distribution data. |
| Apify payouts | https://help.apify.com/en/articles/10057167-how-developer-payouts-work ; https://docs.apify.com/actors/publishing/monetize/monthly-payouts | Invoices generated 11th of month for prior month; released days 21–25; min $20 PayPal/Wise, $100 other; KYC (ID photo) required; only paid-plan user usage counts. | Time from KYC submission to approval. |
| Apify monetization | https://docs.apify.com/platform/actors/publishing/monetize/pay-per-event ; store publishing terms https://docs.apify.com/legal/store-publishing-terms-and-conditions | profit = 0.8 × revenue − platform costs; free-plan usage pays nothing; must respond to issues within 14 days or payouts halt; rental model retired Oct 1 2026. Apify recommends ~2 h/week maintenance. | None. |
| Apify demand (public Store API, api.apify.com/v2/store) | Queries: seo audit, broken link, lighthouse, pdf, rss, dns, structured data, etc. | Utility niches are crowded with near-duplicate actors; niche leaders have 100–600 lifetime users and roughly 10–100 users/30 days; dozens of clones with 0–10 users. Demand concentrates in social/maps scrapers (100k+ users) that carry ToS and brittleness risk. | Revenue per user not public. |
| Gumroad | https://gumroad.com/help/article/13-getting-paid (+ third-party 2026 summaries) | Weekly payouts; identity verification required for bank payouts; third-party sources report $100 minimum for unverified accounts, $10 verified, 10% + $0.50 fee, 30% on Discover sales. | Official page body not retrievable; figures from third parties. |
| Payhip | https://payhip.com/pricing | Free plan, 5% fee + PayPal/Stripe processing; funds deposited immediately to seller's PayPal/Stripe after each sale. | Marketplace traffic is small. |
| Ko-fi | third-party 2026 summaries | 5% shop fee free tier; direct to PayPal/Stripe instantly. | Official page not fetched. |
| itch.io | https://itch.io/docs/creators/payments | "Direct to you" = instant to PayPal/Stripe; "Collected" = 7-day hold + review; default 10% share. | Demand for non-game assets. |
| Notion Marketplace | https://www.notion.com/help/selling-on-marketplace (via search summary) | 10% + $0.40; 14-day hold; biweekly payouts; needs Notion account to build. | Not buildable in this environment. |
| Figma Community | https://help.figma.com/hc/en-us/articles/12067637274519 | Paid selling is invite-only. | Rejected. |
| Framer Marketplace | https://www.framer.com/help/articles/how-the-creator-program-works/ | 100% revenue on templates; needs Framer to build. | Not buildable here. |
| Canva Creators | https://www.canva.com/help/canva-creators-program/ | Application takes months; Element creators closed. | Rejected. |
| Chrome Web Store | https://developer.chrome.com/docs/webstore/register/ | $5 one-time; no native payments; ExtensionPay 5%. | Demand unproven. |
| WordPress.org | https://make.wordpress.org/updates/2026/08/31/plugins-team-31-aug-2026/ | 4,715 plugins in review queue on 2026-08-31. Freemius pays 10th monthly, $100 min. | Rejected on speed. |
| Shopify apps | community/third-party | 1–4 week review; monthly payout. | Rejected on speed. |
| Unity Asset Store | https://assetstore.unity.com/publishing/publish-and-sell-assets | 5+ business day review; 70/30; monthly PayPal. | No game-asset production capability here. |
| Adobe Stock | https://helpx.adobe.com/stock/contributor/help/getting-paid.html | $25 min; 45 days after first sale; gen-AI vectors accepted with label. | Rejected on speed and quality. |
| Redbubble | https://help.redbubble.com/hc/en-us/articles/50959863016724 | Standard tier 50% platform fee; monthly payout; $10 threshold from 2026-07-01. | Rejected on margin. |
| Creative Market | https://support.creativemarket.com/hc/en-us/articles/115004015634 | 40–50% commission; monthly; application. | Rejected. |
| Envato | https://author.envato.com/hub/changes-to-envato-market-revenue-share-and-exclusivity-what-you-need-to-know/ | 50% author fee from 2026-07-01; monthly 15th; $50 min. | Rejected on margin/review. |
| Whop | third-party 2026 summaries | ~3% platform + 2.7% + $0.30 processing; paid payouts ($2.50 ACH). | No discovery for new seller. |
| RapidAPI | https://docs.rapidapi.com/docs/monetizing-your-api-on-rapidapicom ; https://www.buildmvpfast.com/alternatives/rapidapi | Still operating under Nokia; declining developer activity. | Rejected. |
| Algora bounties | https://algora.io/ ; https://algora.io/challenges/prettier | Public board moved; challenge pages show large completed prizes ($25k Prettier). No open-bounty list retrievable; Stripe Connect KYC needed. | Whether small open bounties currently exist. Not passive. |
| Bug bounties | https://trainingcamp.com/articles/the-best-bug-bounty-websites-in-2026-... | 2–8 weeks from report to payment; triage congested. | Not passive; rejected. |
| Fiverr | https://help.fiverr.com/hc/en-us/articles/4402267122449-Early-Payout (+ summaries) | 14-day clearance for new sellers; then PayPal within ~2 days. | Not passive; a bridge at best. |
| Broken-link/SEO SaaS pricing (for Apify fallback) | https://securitybot.dev/blog/best-automated-link-checker-tools ; https://bulkurlchecker.com/compare/ahrefs-broken-link-checker | Paid tools exist ($5–$99/mo), Screaming Frog free to 500 URLs. | Willingness to pay per crawl on Apify unknown. |
