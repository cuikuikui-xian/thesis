from pathlib import Path
import re
import sys

from pypdf import PdfReader

sys.stdout.reconfigure(encoding="utf-8", errors="replace")


SOURCE_DIR = Path(r"G:\荣誉\中国科协青年科技人才培育工程博士生专项计划\论文")
OUT_DIR = Path(r"C:\Users\32933\Documents\ChatGPT\论文\outputs\contribution_update_20260903\pdf_text")
OUT_DIR.mkdir(parents=True, exist_ok=True)

for pdf_path in sorted(SOURCE_DIR.glob("0*.pdf")):
    reader = PdfReader(str(pdf_path))
    pages = []
    for i, page in enumerate(reader.pages, start=1):
        text = page.extract_text() or ""
        pages.append(f"\n\n===== PAGE {i} =====\n{text}")
    full_text = "".join(pages)
    out_path = OUT_DIR / f"{pdf_path.stem}.txt"
    out_path.write_text(full_text, encoding="utf-8")

    metric_lines = []
    for line in full_text.splitlines():
        if re.search(r"%|fold|times|W/\(m|W/cm|thermal resistance|thermal conductivity|maximum thermal load|samples|references", line, re.I):
            metric_lines.append(line.strip())
    print(f"\n### {pdf_path.name} | pages={len(reader.pages)}")
    for line in metric_lines[-80:]:
        print(line)
