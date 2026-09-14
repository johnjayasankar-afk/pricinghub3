"""Static audit of the landing page (index.html, 404.html, links.json, vercel.json, assets/).
Run after any edit or before a deploy. Exit 1 on any problem. Needs only Python; uses Node for script syntax if available."""
import json, os, re, shutil, subprocess, sys
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
problems = []
def P(msg): problems.append(msg)
html = open(os.path.join(ROOT, "index.html"), encoding="utf-8").read()

# 1. Every local asset referenced (src, srcset, href, preload, og image, JSON-LD image) must exist
refs = set(re.findall(r'(?:src|href|imagesrcset|srcset|content)="([^"]+)"', html))
paths = set()
for r in refs:
    for cand in r.split(","):
        cand = cand.strip().split(" ")[0]
        if cand.startswith("assets/") or cand in ("links.json", "vercel.json", "robots.txt"): paths.add(cand)
for m in re.findall(r'"image":"([^"]+)"', html): paths.add(m)
for p in sorted(paths):
    if not os.path.exists(os.path.join(ROOT, p)): P(f"referenced file missing: {p}")
print(f"local references checked: {len(paths)}")

# 2. srcset candidates: the descriptor width must match the file's real width
try:
    from PIL import Image
    for srcset in re.findall(r'srcset="([^"]+)"', html):
        for cand in srcset.split(","):
            path, w = cand.strip().split(" "); w = int(w.rstrip("w"))
            real = Image.open(os.path.join(ROOT, path)).size[0]
            if real != w: P(f"srcset width {w} but {path} is {real}px wide")
except ImportError: print("Pillow not installed; srcset width check skipped")

# 3. links.json: valid, expected keys, https-only links, price format
try:
    cfg = json.load(open(os.path.join(ROOT, "links.json")))
    for k in ("etsy", "payhip", "price"):
        if k not in cfg: P(f"links.json missing key {k}")
    for k in ("etsy", "payhip"):
        v = cfg.get(k, "")
        if v and not v.startswith("https://"): P(f"links.json {k} must start with https:// (got {v!r})")
    if not re.fullmatch(r"\$\d+(\.\d{2})?", str(cfg.get("price", ""))): P(f"links.json price should look like $29 or $29.00 (got {cfg.get('price')!r})")
except Exception as e: P(f"links.json unreadable: {e}")

# 4. vercel.json valid; .vercelignore keeps the product out
json.load(open(os.path.join(ROOT, "vercel.json")))
ign = open(os.path.join(ROOT, ".vercelignore")).read()
for d in ("product/", "deliverables/", "tests/", "records/", "model/", "listing/"):
    if d not in ign: P(f".vercelignore no longer excludes {d}")
if "*.xlsx" not in ign: P(".vercelignore should keep *.xlsx")

# 5. Metadata: title, description length, JSON-LD parses, og image dimensions declared
title = re.search(r"<title>(.*?)</title>", html).group(1)
if not (20 <= len(title) <= 70): P(f"title length {len(title)} outside 20–70")
desc = re.search(r'<meta name="description" content="([^"]+)"', html).group(1)
if not (80 <= len(desc) <= 160): P(f"meta description length {len(desc)} outside 80–160")
ld = re.search(r'<script type="application/ld\+json" id="productLd">\s*(\{.*?\})\s*</script>', html, re.S)
if not ld: P("Product JSON-LD block missing")
else:
    try: json.loads(ld.group(1))
    except Exception as e: P(f"JSON-LD invalid: {e}")
if 'property="og:image:width"' not in html: P("og:image:width missing")

# 6. Structure: one h1, no duplicate ids, every id used by the script exists, no http:// resources
if html.count("<h1") != 1: P(f"{html.count('<h1')} h1 elements")
ids = re.findall(r'\sid="([^"]+)"', html)
dups = {i for i in ids if ids.count(i) > 1}
if dups: P(f"duplicate ids: {sorted(dups)}")
used = set(re.findall(r'\$\("([^"]+)"\)', html)) | set(re.findall(r'getElementById\("([^"]+)"\)', html))
missing = sorted(used - set(ids))
if missing: P(f"script references missing ids: {missing}")
if re.search(r'(src|href)="http://', html): P("insecure http:// resource reference")
for f in ("404.html",):
    if not os.path.exists(os.path.join(ROOT, f)): P(f"{f} missing")

# 7. Script syntax via Node when available
i, j = html.index("<script>") + 8, html.index("</script>", html.index("<script>"))
if shutil.which("node"):
    tmp = os.path.join(ROOT, "tests", "_page.js"); open(tmp, "w").write(html[i:j])
    r = subprocess.run(["node", "--check", tmp], capture_output=True, text=True); os.remove(tmp)
    if r.returncode: P("page script has a syntax error: " + r.stderr.strip()[:200])
    else: print("page script syntax OK")
else: print("node not found; script syntax check skipped")

print(f"problems: {len(problems)}")
for p in problems: print("  PROBLEM:", p)
sys.exit(1 if problems else 0)
