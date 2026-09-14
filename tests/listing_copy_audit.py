"""Checks the Etsy listing package against Etsy's published limits: title <= 140 chars, exactly 13 tags of <= 20 chars, no duplicate tags,
upload files under 20 MB each, exactly ten 2000x2000 images. Exit 1 on any violation."""
import os, re, sys
from PIL import Image
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
md = open(os.path.join(ROOT, "listing", "etsy-listing.md")).read()
title = md.split("## Title (≤140 characters)")[1].strip().split("\n")[0].strip()
tags = [t.strip() for t in md.split("## Tags (13, each ≤20 chars)")[1].strip().split("\n")[0].split("|")]
problems = []
if len(title) > 140: problems.append(f"title is {len(title)} chars")
if len(tags) != 13: problems.append(f"{len(tags)} tags, Etsy allows 13")
for t in tags:
    if len(t) > 20: problems.append(f"tag too long: {t!r} ({len(t)})")
if len(set(tags)) != len(tags): problems.append("duplicate tags")
for f in ("Contractor-Job-Costing-SAMPLE.xlsx", "Contractor-Job-Costing-BLANK.xlsx", "Quick-Start-Guide.pdf"):
    sz = os.path.getsize(os.path.join(ROOT, "product", f))
    if sz > 20 * 1024 * 1024: problems.append(f"{f} is {sz} bytes, over Etsy's 20 MB file limit")
imgs = sorted(f for f in os.listdir(os.path.join(ROOT, "listing", "images")) if f.endswith(".png"))
if len(imgs) != 10: problems.append(f"{len(imgs)} listing images, expected 10")
for f in imgs:
    if Image.open(os.path.join(ROOT, "listing", "images", f)).size != (2000, 2000): problems.append(f"{f} is not 2000x2000")
print(f"title {len(title)}/140 chars; {len(tags)} tags, longest {max(len(t) for t in tags)}/20; {len(imgs)} images; files under 20 MB")
for p in problems: print(" PROBLEM:", p)
sys.exit(1 if problems else 0)
