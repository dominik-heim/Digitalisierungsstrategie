"""Gesamtauswertung: Workshop-Input einlesen → analysieren → PowerPoint, Excel, BPMN erzeugen."""
from __future__ import annotations

import datetime as dt
import json
from pathlib import Path
from typing import Any

from .analysis import analyze, management_summary
from .bpmn import export_bpmn
from .project import Project
from .report_excel import build_workbook
from .report_pptx import build_deck


def run(project: Project, use_ai: bool = False, log=print) -> dict[str, Any]:
    data = project.load()
    log(f"Eingelesen: {', '.join(f'{k} ({len(v)})' for k, v in data.tables.items() if v)}"
        + (f", Matrizen: {', '.join(data.wide)}" if data.wide else ""))
    results = analyze(data, project.config)
    summary = management_summary(results, project.config)
    if use_ai:
        try:
            from .ai import ai_management_summary
            summary = ai_management_summary(results, project.config) or summary
            log("KI-Management-Summary erstellt.")
        except Exception as e:  # KI ist optional – regelbasierte Summary bleibt erhalten
            log(f"KI-Summary nicht möglich ({e}); regelbasierte Summary wird verwendet.")
    stamp = dt.date.today().strftime("%Y-%m-%d")
    out = project.output_dir
    files = {
        "pptx": build_deck(results, project.config, out / f"{stamp}_Digitalisierungsstrategie.pptx", summary),
        "xlsx": build_workbook(results, project.config, out / f"{stamp}_Auswertung.xlsx", summary),
    }
    for name, proc in results.get("prozess", {}).get("prozesse", {}).items():
        files[f"bpmn:{name}"] = export_bpmn(name, proc, out / "BPMN" / f"{_safe(name)}.bpmn")
    (out / "ergebnisse.json").write_text(json.dumps(results, default=str, ensure_ascii=False, indent=1),
                                         encoding="utf-8")
    for issue in results.get("issues", [])[:15]:
        log(f"  ⚠ {issue}")
    for k, f in files.items():
        log(f"→ {k}: {f}")
    return {"results": results, "files": files, "summary": summary}


def _safe(s: str) -> str:
    return "".join(c if c.isalnum() or c in "-_ " else "_" for c in s).strip().replace(" ", "_")
