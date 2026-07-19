#!/usr/bin/env python3
"""Build a PDF of the private Function/Marriage Hall alternate-use layout.

- Converts hall/Marriage-Hall-Layout.md to HTML (tables, fenced code).
- Inlines marriage-hall-layout.svg before the ASCII fallback block.
- Applies a repeating "PRIVATE" watermark.
- Renders to hall/Marriage-Hall-Layout.pdf via headless Chrome.
"""
import base64
import subprocess
import pathlib
import markdown

HERE = pathlib.Path(__file__).parent
MD = HERE / "Marriage-Hall-Layout.md"
SVG = HERE / "marriage-hall-layout.svg"
HTML = HERE / "hall.html"
PDF = HERE / "Marriage-Hall-Layout.pdf"
WATERMARK_TEXT = "PRIVATE"

body_html = markdown.markdown(
    MD.read_text(encoding="utf-8"),
    extensions=["tables", "fenced_code", "sane_lists"],
)

svg_markup = SVG.read_text(encoding="utf-8")
svg_block = f'<div class="svgwrap">{svg_markup}</div>'
if "<pre>" in body_html:
    body_html = body_html.replace("<pre>", svg_block + "<pre>", 1)
else:
    body_html += svg_block

tile = (
    '<svg xmlns="http://www.w3.org/2000/svg" width="360" height="240">'
    '<text x="180" y="140" fill="#d88" fill-opacity="0.30" font-size="46" '
    'font-family="Arial, sans-serif" font-weight="bold" text-anchor="middle" '
    'transform="rotate(-30 180 140)">' + WATERMARK_TEXT + "</text></svg>"
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
  h1 {{ font-size: 19px; border-bottom: 3px solid #b45309; padding-bottom: 6px; }}
  h2 {{ font-size: 14px; margin-top: 18px; color: #0b5394; border-bottom: 1px solid #ccc; padding-bottom: 3px; }}
  table {{ border-collapse: collapse; width: 100%; margin: 8px 0; font-size: 9.5px; }}
  th, td {{ border: 1px solid #bbb; padding: 4px 6px; text-align: left; vertical-align: top; }}
  th {{ background: #f6efe6; }}
  blockquote {{ background: #fff8f0; border-left: 4px solid #b45309; margin: 8px 0; padding: 6px 12px; }}
  code, pre {{ font-family: "DejaVu Sans Mono", monospace; }}
  pre {{ background: #f5f5f5; border: 1px solid #ddd; padding: 8px; font-size: 8px; line-height: 1.15; white-space: pre; }}
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
            "--user-data-dir=/tmp/chrome-hall-profile",
            f"--print-to-pdf={PDF}", "--no-pdf-header-footer",
            HTML.as_uri(),
        ], check=False)
        break
print("PDF written:", PDF)
