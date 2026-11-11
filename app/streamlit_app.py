"""Web-Oberfläche für den Methodenkoffer Digitalisierungsstrategie.

Start:  digistrat app   (oder: streamlit run app/streamlit_app.py)
"""
from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st
import yaml

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from digistrat.analysis import analyze, management_summary  # noqa: E402
from digistrat.model import METHODS, PHASES, methods_by_phase  # noqa: E402
from digistrat.project import Project, write_input_templates  # noqa: E402
from digistrat.themes import PRESETS  # noqa: E402

st.set_page_config(page_title="Digitalisierungsstrategie", page_icon="🏭", layout="wide")
ROOT = Path(os.environ.get("DIGISTRAT_PROJEKTE", Path.cwd() / "projekte"))
ROOT.mkdir(parents=True, exist_ok=True)
KEY_FILE = Path.home() / ".digistrat" / "anthropic_key"
if not os.environ.get("ANTHROPIC_API_KEY") and KEY_FILE.exists():
    os.environ["ANTHROPIC_API_KEY"] = KEY_FILE.read_text(encoding="utf-8").strip()

st.markdown("""<style>
h1, h2, h3 {color:#152C5A;} .stButton>button {background:#004F9F;color:white;border:0;}
.stButton>button:hover {background:#00AFE9;color:white;}
</style>""", unsafe_allow_html=True)


# ---------------------------------------------------------------- Projekt-Auswahl
def sidebar() -> Project | None:
    st.sidebar.title("🏭 Digitalisierungsstrategie")
    projects = sorted(p.parent.name for p in ROOT.glob("*/projekt.yaml"))
    with st.sidebar.expander("➕ Neues Projekt", expanded=not projects):
        name = st.text_input("Ordnername", placeholder="kunde-2027")
        kunde = st.text_input("Kunde")
        branche = st.text_input("Branche", "Produzierendes Unternehmen")
        c1, c2 = st.columns(2)
        if c1.button("Anlegen", disabled=not name):
            Project.create(ROOT / name, kunde=kunde or name, branche=branche)
            st.session_state.projekt = name
            st.rerun()
        if c2.button("Demo-Projekt"):
            from digistrat.demo import create_demo
            create_demo(ROOT / "demo-maschinenbau")
            st.session_state.projekt = "demo-maschinenbau"
            st.rerun()
    projects = sorted(p.parent.name for p in ROOT.glob("*/projekt.yaml"))
    if not projects:
        st.info("Bitte links ein Projekt anlegen oder das Demo-Projekt laden.")
        return None
    cur = st.session_state.get("projekt", projects[0])
    sel = st.sidebar.selectbox("Projekt", projects, index=projects.index(cur) if cur in projects else 0)
    st.session_state.projekt = sel
    key_ok = bool(os.environ.get("ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_AUTH_TOKEN"))
    with st.sidebar.expander("🟢 KI aktiv" if key_ok else "⚪ KI einrichten (API-Key)", expanded=False):
        st.caption("Key von console.anthropic.com → API Keys. Wird nur lokal auf diesem Rechner gespeichert.")
        k = st.text_input("Anthropic API-Key", type="password", placeholder="sk-ant-…")
        remember = st.checkbox("Auf diesem Rechner merken", value=True)
        if st.button("Key übernehmen", disabled=not k):
            os.environ["ANTHROPIC_API_KEY"] = k.strip()
            if remember:
                KEY_FILE.parent.mkdir(parents=True, exist_ok=True)
                KEY_FILE.write_text(k.strip(), encoding="utf-8")
                try:
                    KEY_FILE.chmod(0o600)
                except OSError:
                    pass
            st.rerun()
        if key_ok and KEY_FILE.exists() and st.button("Gespeicherten Key löschen"):
            KEY_FILE.unlink()
            os.environ.pop("ANTHROPIC_API_KEY", None)
            st.rerun()
    proj = Project.open(ROOT / sel)
    if not proj.config.get("ki", {}).get("aktiv", True):
        st.sidebar.markdown("🔒 **KI für dieses Projekt deaktiviert**")
    st.sidebar.caption(f"Projekte in `{ROOT}`")
    return proj


def ki_on(p: Project) -> bool:
    return bool(p.config.get("ki", {}).get("aktiv", True))


def consent(lines: list[str], key: str) -> bool:
    """Hinweis vor jedem KI-Aufruf: zeigt, was gesendet wird, und verlangt eine Bestätigung."""
    personal = any(line.startswith("⚠") for line in lines)
    box = st.warning if personal else st.info
    box("**Diese Daten werden an Anthropic (Claude) gesendet:**\n\n" + "\n".join(f"- {x}" for x in lines)
        + ("\n\n⚠ = enthält personenbezogene Angaben" if personal else ""))
    return st.checkbox("Ich bestätige die Übermittlung dieser Daten an Anthropic", key=f"consent_{key}")


def ki_off_note() -> None:
    st.info("🔒 KI ist für dieses Projekt deaktiviert – es werden keine Daten an Anthropic gesendet. "
            "Änderbar unter „Projekt“.")


def page_projekt(p: Project) -> None:
    st.header("Projekt & Einstellungen")
    c = p.config
    c1, c2 = st.columns(2)
    c["kunde"] = c1.text_input("Kunde", c.get("kunde", ""))
    c["branche"] = c2.text_input("Branche", c.get("branche", ""))
    c["projekt"] = c1.text_input("Projekttitel", c.get("projekt", ""))
    c["berater"] = c2.text_input("Berater/Team", c.get("berater", ""))
    c["roadmap_start"] = c1.text_input("Roadmap-Start (JJJJ-Qn)", c.get("roadmap_start", ""))
    c["zinssatz"] = c2.number_input("Kalkulationszins", 0.0, 0.5, float(c.get("zinssatz", 0.08)), 0.005, format="%.3f")
    th = c.setdefault("theme", {})
    th["preset"] = c1.selectbox("Farbschema", list(PRESETS), index=list(PRESETS).index(th.get("preset", "ingenics")))
    th["pptx_template"] = c2.text_input("Optional: PowerPoint-Master (.pptx-Pfad)", th.get("pptx_template", ""))
    ki = c.setdefault("ki", {})
    st.subheader("KI-Nutzung")
    ki["aktiv"] = st.toggle("KI-Funktionen für dieses Projekt erlauben", value=bool(ki.get("aktiv", True)),
                            help="Ausschalten, wenn der Kunde den Einsatz externer KI-Dienste ausschließt. "
                                 "Dann werden keinerlei Daten an Anthropic übertragen.")
    if not ki["aktiv"]:
        st.caption("🔒 Deaktiviert: KI-Import, KI-Synthese, KI-Leitfragen und KI-Summary sind gesperrt.")
    ki["model"] = c1.text_input("KI-Modell", ki.get("model", "claude-opus-5-5"))
    ki["effort"] = c2.selectbox("KI-Gründlichkeit (effort)", ["low", "medium", "high", "xhigh", "max"],
                                index=["low", "medium", "high", "xhigh", "max"].index(ki.get("effort", "high")))
    if st.button("Speichern"):
        p.save_config()
        st.success("Gespeichert.")
        st.rerun()
    st.subheader("Vorgehensmodell")
    cols = st.columns(3)
    for col, (ph, ms) in zip(cols, methods_by_phase().items()):
        col.markdown(f"**{PHASES[ph]}**")
        col.markdown("\n".join(f"- {m.title}" for m in ms))


def page_workshops(p: Project) -> None:
    from digistrat import workshop
    st.header("Workshop-Planung")
    f = p.workshop_dir / "workshops.yaml"
    c1, c2, c3 = st.columns(3)
    start = c1.date_input("Erster Workshop")
    interval = c2.number_input("Abstand [Tage]", 1, 60, 7)
    opt = c3.checkbox("Fit-Gap-Workshop (Standardsoftware) einplanen", True)
    if st.button("Workshop-Serie (neu) planen"):
        if f.exists():
            f.unlink()
        workshop.plan(p, start.isoformat(), int(interval), opt)
        st.rerun()
    if not f.exists():
        st.info("Noch keine Planung vorhanden.")
        return
    txt = st.text_area("Planung bearbeiten (YAML: Termine, Ort, Blöcke, Minuten)", f.read_text(encoding="utf-8"),
                       height=320)
    use_ai = ki_on(p) and st.checkbox("Kundenspezifische Leitfragen per KI erzeugen")
    ctx = st.text_input("Kontext für die KI (Ziele, Besonderheiten des Kunden)") if use_ai else ""
    ok = True
    if use_ai:
        from digistrat.ai import describe_payload
        ok = consent(describe_payload("questions", p.config, context=ctx), "questions")
    if st.button("Agenden, Leitfäden & Erfassungsvorlagen erzeugen", disabled=not ok):
        f.write_text(txt, encoding="utf-8")
        questions = None
        if use_ai:
            from digistrat.ai import workshop_questions
            with st.spinner("KI formuliert Leitfragen …"):
                questions = workshop_questions(p.config, workshop.load(p), ctx)
        files = workshop.render(p, questions)
        st.success(f"{len(files)} Dateien erzeugt in {p.workshop_dir}")
    ws = workshop.load(p)
    st.dataframe(pd.DataFrame([{"ID": w["id"], "Workshop": w["titel"], "Datum": w["datum"],
                                "Zeit": f"{w['beginn']}–{w['ende']}", "Ziel": w["ziel"]} for w in ws]),
                 hide_index=True, width="stretch")
    _downloads(sorted(p.workshop_dir.glob("*.xlsx")) + sorted(p.workshop_dir.glob("*.pptx")) +
               sorted((p.workshop_dir / "Erfassungsvorlagen").glob("*.xlsx")) +
               sorted((p.workshop_dir / "Moderationsleitfaeden").glob("*.md")))


def page_input(p: Project) -> None:
    st.header("Workshop-Ergebnisse erfassen")
    st.caption("Ausgefüllte Excel-Vorlagen hochladen – oder Fotos, Notizen, Transkripte per KI strukturieren lassen.")
    data = p.load()
    up = st.file_uploader("Ausgefüllte Erfassungsdatei(en) (.xlsx) hochladen", type=["xlsx"], accept_multiple_files=True)
    if up and st.button("Dateien übernehmen"):
        for u in up:
            (p.input_dir / u.name).write_bytes(u.getvalue())
        st.success(f"{len(up)} Datei(en) gespeichert.")
        st.rerun()
    st.write("Vorhandene Dateien: " + ", ".join(f.name for f in sorted(p.input_dir.glob("*.xlsx"))))
    if not ki_on(p):
        ki_off_note()
    for ph, ms in methods_by_phase().items():
        st.subheader(PHASES[ph])
        for m in ms:
            n = len(data.get(m.key)) or len(data.wide.get(m.key, ([], []))[1])
            with st.expander(f"{'✅' if n else '⬜'} {m.title} – {n} Einträge"):
                st.markdown(f"**Ziel:** {m.purpose}\n\n**Ablauf:** {m.workshop}\n\n**Ergebnis:** {m.output}")
                if n:
                    if m.wide:
                        cols, rows = data.wide[m.key]
                        st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch")
                    else:
                        st.dataframe(pd.DataFrame(data.get(m.key)).rename(
                            columns={f.key: f.header for f in m.fields}), hide_index=True, width="stretch")
                tmp = Path(tempfile.gettempdir()) / f"Vorlage_{m.sheet}.xlsx"
                write_input_templates(tmp, [m.key])
                st.download_button("Leere Vorlage", tmp.read_bytes(), file_name=tmp.name, key=f"dl_{m.key}")
                if not m.wide and ki_on(p):
                    _ai_import(p, m.key)


def _ai_import(p: Project, key: str) -> None:
    st.markdown("**KI-Import** aus Fotos (Stickies, Flipcharts), PDF, Word, PowerPoint, Miro-Export oder Notizen")
    files = st.file_uploader("Rohdaten", accept_multiple_files=True, key=f"up_{key}",
                             type=["png", "jpg", "jpeg", "webp", "pdf", "docx", "pptx", "xlsx", "txt", "md", "csv"])
    text = st.text_area("Notizen/Transkript", key=f"tx_{key}", height=100)
    ok = False
    if files or text:
        from digistrat.ai import describe_payload
        ok = consent(describe_payload("import", p.config, files=files, text=text), f"imp_{key}")
    if st.button("Mit KI strukturieren", key=f"ai_{key}", disabled=not ok):
        from digistrat.ai import extract_rows
        tmpdir = Path(tempfile.mkdtemp())
        paths = []
        for u in files or []:
            fp = tmpdir / u.name
            fp.write_bytes(u.getvalue())
            paths.append(fp)
            (p.raw_dir / u.name).write_bytes(u.getvalue())  # Rohdaten archivieren
        with st.spinner("Claude liest die Workshop-Ergebnisse …"):
            rows, hint = extract_rows(key, text or None, paths, p.config)
        st.session_state[f"rows_{key}"] = rows
        st.session_state[f"hint_{key}"] = hint
    _review_and_save(p, key)


def _review_and_save(p: Project, key: str) -> None:
    rows = st.session_state.get(f"rows_{key}")
    if not rows:
        return
    m = METHODS[key]
    if st.session_state.get(f"hint_{key}"):
        st.info("Hinweise der KI: " + st.session_state[f"hint_{key}"])
    df = pd.DataFrame(rows).rename(columns={f.key: f.header for f in m.fields})
    edited = st.data_editor(df, num_rows="dynamic", width="stretch", key=f"ed_{key}")
    if st.button("Geprüft – übernehmen", key=f"save_{key}"):
        from digistrat.ai import save_rows
        back = edited.rename(columns={f.header: f.key for f in m.fields}).to_dict("records")
        f = save_rows(p, key, back)
        st.success(f"{len(back)} Zeilen gespeichert in {f.name}")
        del st.session_state[f"rows_{key}"]


def page_ableiten(p: Project) -> None:
    st.header("KI-Synthese: Folge-Artefakte vorschlagen")
    st.caption("Claude leitet aus den bisherigen Ergebnissen Vorschläge ab – Sie prüfen und übernehmen.")
    if not ki_on(p):
        ki_off_note()
        return
    from digistrat.ai import DERIVE_PROMPTS, describe_payload
    target = st.selectbox("Was soll abgeleitet werden?", list(DERIVE_PROMPTS),
                          format_func=lambda k: f"{METHODS[k].title} – {DERIVE_PROMPTS[k][:70]}…")
    hint = st.text_input("Hinweis an die KI (optional)")
    res = analyze(p.load(), p.config)
    ok = consent(describe_payload("derive", p.config, results=res), f"derive_{target}")
    if st.button("Vorschläge erzeugen", disabled=not ok):
        from digistrat.ai import derive_rows
        with st.spinner("Claude analysiert …"):
            rows, h = derive_rows(target, res, p.config, hint)
        st.session_state[f"rows_{target}"], st.session_state[f"hint_{target}"] = rows, h
    _review_and_save(p, target)


def page_auswertung(p: Project) -> None:
    st.header("Prüfen & Auswerten")
    data = p.load()
    res = analyze(data, p.config)
    if res["issues"]:
        with st.expander(f"⚠ {len(res['issues'])} Hinweise zur Datenqualität"):
            st.write("\n".join(f"- {i}" for i in res["issues"]))
    k = res["kennzahlen"]
    cols = st.columns(min(6, max(1, len(k))))
    for i, (lab, v) in enumerate(list(k.items())[:6]):
        cols[i].metric(lab.replace(" [EUR]", ""), f"{v / 1000:,.0f} T€" if "EUR" in lab else
                       (f"{v:.0%}" if isinstance(v, float) else v))
    st.subheader("Management Summary (regelbasiert)")
    st.markdown("\n".join(f"- {b}" for b in management_summary(res, p.config)))
    c1, c2 = st.columns(2)
    if "painpoints" in res:
        c1.markdown("**Top Pain Points**")
        c1.dataframe(pd.DataFrame(res["painpoints"]["top"])[["id", "beschreibung", "kategorie", "score"]],
                     hide_index=True, width="stretch")
    if "nwa" in res:
        c2.markdown("**Nutzwert-Ranking**")
        c2.bar_chart(pd.DataFrame({"Nutzwert": [i["prozent"] for i in res["nwa"]["items"]]},
                                  index=[f"{i['id']} {i['titel']}"[:35] for i in res["nwa"]["items"]]),
                     horizontal=True)
    if "impact_effort" in res:
        c1.markdown("**Impact/Effort**")
        c1.scatter_chart(pd.DataFrame(res["impact_effort"]["items"]), x="effort", y="impact", color="quadrant")
    if "roadmap" in res:
        c2.markdown("**Roadmap**")
        c2.dataframe(pd.DataFrame(res["roadmap"]["items"])[["id", "titel", "horizont", "start", "dauer"]],
                     hide_index=True, width="stretch")
    use_ai = ki_on(p) and st.checkbox("Management Summary per KI formulieren")
    ok = True
    if use_ai:
        from digistrat.ai import describe_payload
        ok = consent(describe_payload("summary", p.config, results=res), "summary")
    if st.button("📊 PowerPoint & Excel erzeugen", type="primary", disabled=not ok):
        from digistrat.pipeline import run
        log = []
        with st.spinner("Erzeuge Ergebnisse …"):
            out = run(p, use_ai=use_ai, log=log.append)
        st.success("Fertig.")
        st.code("\n".join(map(str, log)))
        st.session_state.files = [str(f) for f in out["files"].values()]
    _downloads([Path(f) for f in st.session_state.get("files", [])] or sorted(p.output_dir.glob("*.*")))


def _downloads(files: list[Path]) -> None:
    files = [f for f in files if f.exists() and f.is_file() and f.suffix != ".json"]
    if not files:
        return
    st.markdown("**Downloads**")
    cols = st.columns(3)
    for i, f in enumerate(files):
        cols[i % 3].download_button(f"⬇ {f.name}", f.read_bytes(), file_name=f.name, key=f"dl_{f}")


PAGES = {"Projekt": page_projekt, "1 · Workshop-Planung": page_workshops, "2 · Workshop-Input": page_input,
         "3 · KI-Synthese": page_ableiten, "4 · Auswertung & Export": page_auswertung}

proj = sidebar()
if proj:
    page = st.sidebar.radio("Arbeitsschritt", list(PAGES))
    try:
        PAGES[page](proj)
    except Exception as e:  # Fehler verständlich anzeigen statt Absturz
        st.error(f"{type(e).__name__}: {e}")
        if "API" in type(e).__name__ or "auth" in str(e).lower():
            st.info("Für KI-Funktionen ANTHROPIC_API_KEY setzen (siehe README).")
