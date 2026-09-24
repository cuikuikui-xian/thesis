from pathlib import Path
from pypdf import PdfReader

root = Path(r"C:\Users\32933\Downloads")
needles = {
    "paper1": "lightweight miniature loop heat pipe",
    "paper2": "bending the heat",
    "paper3": "tilt angle and base plate design",
    "paper4": "bi-porous wick for passive phase-change devices",
    "paper5": "from fundamental researches to applications",
}

for pdf in root.rglob("*.pdf"):
    try:
        reader = PdfReader(str(pdf))
        text = " ".join(" ".join((page.extract_text() or "").split()) for page in reader.pages).lower()
        hits = [key for key, needle in needles.items() if needle in text]
        if hits:
            print("|".join(hits), pdf, pdf.stat().st_size, sep="\t")
    except Exception:
        pass
