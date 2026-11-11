from pathlib import Path

from openpyxl import load_workbook
from pptx import Presentation

from digistrat.analysis import analyze
from digistrat.demo import create_demo
from digistrat.pipeline import run
from digistrat.workshop import plan, render


def test_demo_end_to_end(tmp_path: Path):
    p = create_demo(tmp_path / "demo")
    data = p.load()
    assert not data.issues, data.issues
    res = analyze(data, p.config)
    for key in ("stakeholder", "prozess", "vsm", "painpoints", "userstories", "handlungsfelder", "nwa",
                "impact_effort", "fitgap", "businesscase", "roadmap", "raci", "vpc", "bmc"):
        assert key in res, key
    assert res["nwa"]["items"][0]["id"] == "UC-01"
    assert res["raci"]["fehler"], "Demo enthält absichtlich eine Zeile mit 2×A"
    out = run(p, log=lambda *a: None)
    prs = Presentation(out["files"]["pptx"])
    assert len(prs.slides) > 30
    wb = load_workbook(out["files"]["xlsx"])
    assert {"Nutzwertanalyse", "BusinessCase", "Roadmap", "RACI"} <= set(wb.sheetnames)
    assert out["files"]["bpmn:Auftragsabwicklung"].read_text().count("<bpmn:task") >= 5


def test_workshops(tmp_path: Path):
    p = create_demo(tmp_path / "demo")
    plan(p, start="2027-01-11")
    files = render(p)
    assert files["pptx"].exists() and files["xlsx"].exists()
    assert any(k.startswith("leitfaden:") for k in files)
    wb = load_workbook(files["vorlage:WS2"])
    assert {"PainPoints", "5Whys", "Ishikawa"} <= set(wb.sheetnames)


def test_empty_project(tmp_path: Path):
    from digistrat.project import Project
    p = Project.create(tmp_path / "leer")
    out = run(p, log=lambda *a: None)
    assert out["files"]["pptx"].exists()
