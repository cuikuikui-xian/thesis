import json
import re
from pathlib import Path
from pypdf import PdfReader

base = Path(r"G:\荣誉\基本科研业务费申请\证明材料")
files = {
    "paper1_review": base / "论文1 A comprehensive review of loop heat pipe From fundamental researches to applications.pdf",
    "paper2_bending": base / "论文2 Bending the heat Innovative ultrathin flexible loop heat pipes for enhanced mobile device cooling.pdf",
    "paper3_device": base / "论文3 A thin and lightweight miniature loop heat pipe for cooling mobile electronic devices.pdf",
    "paper4_tilt": base / "论文4 Effect of tilt angle and base plate design on the performance of a loop heat pipe driven by vapor–liquid injector.pdf",
    "paper5_biporous": base / "论文5 Experimental and theoretical study on performance of biporous wick for passive phasechange devices.pdf",
}

result = {}
for key, file in files.items():
    reader = PdfReader(str(file))
    text = "\n".join((page.extract_text() or "") for page in reader.pages[:3])
    match = re.search(
        r"\bABSTRACT\b\s*(.*?)(?=\n\s*(?:1\.\s*Introduction|I\.\s*INTRODUCTION|NOMENCLATURE)\b)",
        text,
        flags=re.IGNORECASE | re.DOTALL,
    )
    abstract = match.group(1) if match else ""
    abstract = re.sub(r"([A-Za-z])\s+-\s*\n\s*([A-Za-z])", r"\1\2", abstract)
    abstract = re.sub(r"\s+", " ", abstract).strip()
    result[key] = {"path": str(file), "pages": len(reader.pages), "abstract": abstract}

print(json.dumps(result, ensure_ascii=True, indent=2))
