#!/usr/bin/env python3
"""Build a watermarked PDF of the DPR from the Markdown source.

- Converts DPR-Biochar-Project.md to HTML (tables, fenced code).
- Inlines plot-layout.svg into Annexure A.
- Applies a repeating "Arun" watermark on every page.
- Renders to DPR-Biochar-Project.pdf via headless Chrome.
"""
import base64
import subprocess
import pathlib
import markdown

ROOT = pathlib.Path(__file__).parent
MD = ROOT / "DPR-Biochar-Project.md"
SVG = ROOT / "plot-layout.svg"
HTML = ROOT / "dpr.html"
PDF = ROOT / "DPR-Biochar-Project.pdf"
WATERMARK_TEXT = "Arun"

body_html = markdown.markdown(
    MD.read_text(encoding="utf-8"),
    extensions=["tables", "fenced_code", "sane_lists"],
)

# Inline the plot layout SVG just before the ASCII fallback block in Annexure A.
svg_markup = SVG.read_text(encoding="utf-8")
svg_block = f'<div class="svgwrap">{svg_markup}</div>'
if "<pre>" in body_html:
    body_html = body_html.replace("<pre>", svg_block + "<pre>", 1)
else:
    body_html += svg_block

# Repeating diagonal watermark tile as a base64 SVG background.
tile = (
    '<svg xmlns="http://www.w3.org/2000/svg" width="340" height="230">'
    '<text x="170" y="130" fill="#c8c8c8" fill-opacity="0.35" font-size="54" '
    'font-family="Arial, sans-serif" font-weight="bold" text-anchor="middle" '
    'transform="rotate(-30 170 130)">' + WATERMARK_TEXT + "</text></svg>"
)
tile_b64 = base64.b64encode(tile.encode("utf-8")).decode("ascii")

html = f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<style>
  @page {{ size: A4; margin: 0; }}
  * {{ box-sizing: border-box; }}
  html, body {{ margin: 0; padding: 0; }}
  .watermark {{
    position: fixed; top: 0; left: 0; width: 100%; height: 100%;
    z-index: 0; pointer-events: none;
    background-image: url("data:image/svg+xml;base64,{tile_b64}");
    background-repeat: repeat;
    -webkit-print-color-adjust: exact; print-color-adjust: exact;
  }}
  .content {{
    position: relative; z-index: 1;
    padding: 16mm 15mm; max-width: 100%;
    font-family: "DejaVu Sans", Arial, Helvetica, sans-serif;
    font-size: 10.5px; line-height: 1.5; color: #1a1a1a;
  }}
  h1 {{ font-size: 20px; border-bottom: 3px solid #38761d; padding-bottom: 6px; }}
  h2 {{ font-size: 15px; margin-top: 20px; color: #0b5394; border-bottom: 1px solid #ccc; padding-bottom: 3px; }}
  h3 {{ font-size: 12.5px; margin-top: 14px; color: #333; }}
  table {{ border-collapse: collapse; width: 100%; margin: 8px 0; font-size: 9.5px; }}
  th, td {{ border: 1px solid #bbb; padding: 4px 6px; text-align: left; vertical-align: top; }}
  th {{ background: #eef3f7; }}
  tr:nth-child(even) td {{ background: #fafafa; }}
  code, pre {{ font-family: "DejaVu Sans Mono", monospace; }}
  pre {{ background: #f5f5f5; border: 1px solid #ddd; padding: 8px; font-size: 8px; line-height: 1.15; overflow: hidden; white-space: pre; }}
  .svgwrap {{ border: 1px solid #ccc; padding: 6px; margin: 10px 0; background: #fff; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
  .svgwrap svg {{ width: 100%; height: auto; }}
  hr {{ border: none; border-top: 1px solid #ddd; margin: 16px 0; }}
</style></head>
<body>
  <div class="watermark"></div>
  <div class="content">{body_html}</div>
</body></html>"""

HTML.write_text(html, encoding="utf-8")

for chrome in ("google-chrome", "chromium", "chromium-browser"):
    if subprocess.run(["which", chrome], capture_output=True).returncode == 0:
        subprocess.run([
            chrome, "--headless=new", "--no-sandbox", "--disable-gpu",
            "--disable-dev-shm-usage", "--disable-setuid-sandbox",
            "--user-data-dir=/tmp/chrome-dpr-profile",
            f"--print-to-pdf={PDF}", "--no-pdf-header-footer",
            HTML.as_uri(),
        ], check=True)
        break
print("PDF written:", PDF)
