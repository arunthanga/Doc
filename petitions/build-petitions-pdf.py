#!/usr/bin/env python3
"""Build print-ready A4 PDFs of the Chittur delegation petitions.

- Converts each numbered Markdown petition in this folder to its own PDF.
- Also produces a single combined PDF (cover note + all five petitions),
  with each petition starting on a new page.
- Rendered via headless Chrome (no extra Python deps beyond `markdown`).
"""
import pathlib
import subprocess
import tempfile
import time

import markdown

HERE = pathlib.Path(__file__).parent
OUT = HERE / "pdf"
OUT.mkdir(exist_ok=True)

STYLE = """
  @page { size: A4; margin: 18mm 16mm; }
  * { box-sizing: border-box; }
  html, body { margin: 0; padding: 0; }
  body {
    font-family: "DejaVu Serif", Georgia, "Times New Roman", serif;
    font-size: 11.5px; line-height: 1.55; color: #111;
  }
  h1 { font-size: 17px; text-align: center; border-bottom: 2px solid #333;
       padding-bottom: 8px; margin: 0 0 14px; }
  h2 { font-size: 13px; margin-top: 16px; color: #0b3d0b; }
  h3 { font-size: 12px; margin-top: 12px; }
  table { border-collapse: collapse; width: 100%; margin: 8px 0; font-size: 10.5px; }
  th, td { border: 1px solid #999; padding: 5px 7px; text-align: left; vertical-align: top; }
  th { background: #eee; }
  hr { border: none; border-top: 1px solid #bbb; margin: 14px 0; }
  ol, ul { margin: 6px 0 6px 18px; }
  .page { page-break-after: always; }
  .page:last-child { page-break-after: auto; }
"""


def md_to_html_fragment(path: pathlib.Path) -> str:
    return markdown.markdown(
        path.read_text(encoding="utf-8"),
        extensions=["tables", "sane_lists"],
    )


def render(html_body: str, out_pdf: pathlib.Path) -> None:
    html = (
        '<!DOCTYPE html><html lang="en"><head><meta charset="utf-8">'
        f"<style>{STYLE}</style></head><body>{html_body}</body></html>"
    )
    html_file = out_pdf.with_suffix(".html")
    html_file.write_text(html, encoding="utf-8")
    for chrome in ("google-chrome", "chromium", "chromium-browser"):
        if subprocess.run(["which", chrome], capture_output=True).returncode == 0:
            with tempfile.TemporaryDirectory(prefix="chrome-petitions-") as profile:
                if out_pdf.exists():
                    out_pdf.unlink()
                proc = subprocess.Popen(
                    [
                        chrome, "--headless=new", "--no-sandbox", "--disable-gpu",
                        "--disable-dev-shm-usage", "--disable-setuid-sandbox",
                        "--no-first-run", "--disable-extensions",
                        "--disable-background-networking", "--disable-sync",
                        "--disable-crash-reporter",
                        f"--user-data-dir={profile}",
                        f"--print-to-pdf={out_pdf}", "--no-pdf-header-footer",
                        html_file.as_uri(),
                    ],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                )
                # Headless Chrome writes the PDF but may not always self-exit;
                # wait until the file is produced, then stop the process.
                deadline = time.time() + 60
                while time.time() < deadline:
                    if proc.poll() is not None:
                        break
                    if out_pdf.exists() and out_pdf.stat().st_size > 0:
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
    html_file.unlink(missing_ok=True)
    if not (out_pdf.exists() and out_pdf.stat().st_size > 0):
        raise RuntimeError(f"Failed to render {out_pdf}")
    print("PDF written:", out_pdf, f"({out_pdf.stat().st_size} bytes)")


sources = sorted(p for p in HERE.glob("*.md"))
combined_parts = []
for src in sources:
    frag = md_to_html_fragment(src)
    combined_parts.append(f'<div class="page">{frag}</div>')
    render(frag, OUT / f"{src.stem}.pdf")

render("".join(combined_parts), OUT / "Chittur-Petitions-Combined.pdf")
