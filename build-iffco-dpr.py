#!/usr/bin/env python3
"""Build a watermarked PDF of the IFFCO Chittur DPR from the Markdown source.

- Converts DPR-IFFCO-Chittur-Kisan-Samriddhi.md to HTML (tables, fenced code).
- Applies a repeating "Arun" watermark on every page.
- Renders to DPR-IFFCO-Chittur-Kisan-Samriddhi.pdf via headless Chrome.
"""
import base64
import pathlib
import subprocess
import tempfile
import time

import markdown

ROOT = pathlib.Path(__file__).parent
MD = ROOT / "DPR-IFFCO-Chittur-Kisan-Samriddhi.md"
HTML = ROOT / "iffco-dpr.html"
PDF = ROOT / "DPR-IFFCO-Chittur-Kisan-Samriddhi.pdf"
WATERMARK_TEXT = "Arun"

body_html = markdown.markdown(
    MD.read_text(encoding="utf-8"),
    extensions=["tables", "fenced_code", "sane_lists"],
)

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
  pre {{ background: #f5f5f5; border: 1px solid #ddd; padding: 8px; font-size: 8.5px; line-height: 1.2; overflow: hidden; white-space: pre; }}
  hr {{ border: none; border-top: 1px solid #ddd; margin: 16px 0; }}
</style></head>
<body>
  <div class="watermark"></div>
  <div class="content">{body_html}</div>
</body></html>"""

HTML.write_text(html, encoding="utf-8")

for chrome in ("google-chrome", "chromium", "chromium-browser"):
    if subprocess.run(["which", chrome], capture_output=True).returncode == 0:
        with tempfile.TemporaryDirectory(prefix="chrome-iffco-dpr-") as profile:
            if PDF.exists():
                PDF.unlink()
            proc = subprocess.Popen([
                chrome, "--headless=new", "--no-sandbox", "--disable-gpu",
                "--disable-dev-shm-usage", "--disable-setuid-sandbox",
                "--no-first-run", "--disable-extensions",
                "--disable-background-networking", "--disable-sync",
                "--disable-crash-reporter",
                f"--user-data-dir={profile}",
                f"--print-to-pdf={PDF}", "--no-pdf-header-footer",
                HTML.as_uri(),
            ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            deadline = time.time() + 60
            while time.time() < deadline:
                if proc.poll() is not None:
                    break
                if PDF.exists() and PDF.stat().st_size > 0:
                    time.sleep(1)
                    break
                time.sleep(0.5)
            if proc.poll() is None:
                proc.terminate()
                try:
                    proc.wait(timeout=10)
                except subprocess.TimeoutExpired:
                    proc.kill()
        break

if not (PDF.exists() and PDF.stat().st_size > 0):
    raise SystemExit("Failed to render PDF")
print("PDF written:", PDF, f"({PDF.stat().st_size} bytes)")
