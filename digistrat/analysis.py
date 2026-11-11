"""Automatisierte Auswertung aller Methoden-Tools.

`analyze(data, config)` liefert ein Ergebnis-Dict, das von PowerPoint-/Excel-Reports und der
KI-Zusammenfassung verwendet wird. Alle Funktionen sind rein (keine Dateizugriffe) und testbar.
"""
from __future__ import annotations

import math
import re
from collections import Counter, OrderedDict, defaultdict
from typing import Any

from .model import METHODS
from .project import Data

NON_DIGITAL = {"Papier", "Excel", "E-Mail", "Mündlich/Telefon", "Manuell/ohne Medium"}


def analyze(data: Data, config: dict[str, Any] | None = None) -> dict[str, Any]:
    cfg = config or {}
    r: dict[str, Any] = {"issues": list(data.issues)}
    if data.has("stakeholder"):
        r["stakeholder"] = stakeholder(data.get("stakeholder"))
    if data.has("sipoc"):
        r["sipoc"] = sipoc(data.get("sipoc"))
    if data.has("prozess"):
        r["prozess"] = prozess(data.get("prozess"))
    if data.has("vsm"):
        r["vsm"] = vsm(data.get("vsm"))
    if data.has("painpoints"):
        r["painpoints"] = painpoints(data.get("painpoints"))
    if data.has("whys") or data.has("ishikawa"):
        r["ursachen"] = ursachen(data.get("whys"), data.get("ishikawa"))
    if data.has("userstories"):
        r["userstories"] = userstories(data.get("userstories"))
        r["storymap"] = storymap(data.get("userstories"))
    if data.has("userstories") or data.has("usecases") or data.has("handlungsfelder"):
        r["handlungsfelder"] = handlungsfelder(data, r.get("painpoints"))
    if data.has("usecases"):
        r["usecases"] = {"items": data.get("usecases")}
    if data.has("vpc"):
        r["vpc"] = vpc(data.get("vpc"))
    if data.has("bmc"):
        r["bmc"] = bmc(data.get("bmc"))
    if data.has("fitgap"):
        r["fitgap"] = fitgap(data.get("fitgap"))
    if data.has("nwa"):
        cols, rows = data.wide["nwa"]
        r["nwa"] = nwa(data.get("kriterien"), cols, rows, data.get("usecases"))
    if data.has("usecases"):
        r["impact_effort"] = impact_effort(data.get("usecases"), r.get("nwa"))
    if data.has("raci"):
        cols, rows = data.wide["raci"]
        r["raci"] = raci(cols, rows)
    if data.has("businesscase"):
        r["businesscase"] = businesscase(data.get("businesscase"), float(cfg.get("zinssatz", 0.08)))
    if data.has("usecases"):
        r["roadmap"] = roadmap(data.get("usecases"), data.get("roadmap"), r.get("impact_effort"),
                               r.get("nwa"), cfg.get("roadmap_start") or "2027-Q1")
    r["kennzahlen"] = kennzahlen(r)
    return r


# ====================================================================== Phase 1
def stakeholder_quadrant(einfluss: float, interesse: float) -> str:
    hi_e, hi_i = einfluss >= 3, interesse >= 3
    if hi_e and hi_i:
        return "Eng einbinden"
    if hi_e:
        return "Zufriedenstellen"
    if hi_i:
        return "Informieren"
    return "Beobachten"


def stakeholder(rows: list[dict]) -> dict:
    items = []
    for s in rows:
        e, i = s.get("einfluss") or 1, s.get("interesse") or 1
        q = stakeholder_quadrant(e, i)
        risk = s.get("haltung") in ("Skeptiker", "Blockierer") and e >= 3
        items.append({**s, "einfluss": e, "interesse": i, "quadrant": q, "risiko": risk,
                      "workshop": q in ("Eng einbinden", "Informieren") or (q == "Zufriedenstellen")})
    order = ["Eng einbinden", "Zufriedenstellen", "Informieren", "Beobachten"]
    return {
        "items": items,
        "quadranten": {q: [x["name"] for x in items if x["quadrant"] == q] for q in order},
        "risiken": [x for x in items if x["risiko"]],
        "teilnehmer": [x for x in items if x["quadrant"] in ("Eng einbinden", "Informieren")],
    }


def sipoc(rows: list[dict]) -> dict:
    out: dict[str, dict[str, list[str]]] = OrderedDict()
    for r in rows:
        p = out.setdefault(r["prozess"], {k: [] for k in ("supplier", "input", "process", "output", "customer")})
        for k in p:
            v = r.get(k)
            if v:
                for part in re.split(r"[;\n]", str(v)):
                    part = part.strip()
                    if part and part not in p[k]:
                        p[k].append(part)
    return {"prozesse": out}


def _successors(steps: list[dict]) -> dict[int, list[int]]:
    nrs = [s["nr"] for s in steps]
    succ = {}
    for idx, s in enumerate(steps):
        if s.get("nachfolger"):
            succ[s["nr"]] = [int(float(x)) for x in s["nachfolger"] if re.match(r"^\d+(\.0)?$", str(x))]
        elif s.get("typ") == "Ende" or idx == len(steps) - 1:
            succ[s["nr"]] = []
        else:
            succ[s["nr"]] = [nrs[idx + 1]]
    return succ


def _is_break(a: dict, b: dict) -> str | None:
    if a.get("typ") == "Entscheidung" or b.get("typ") == "Entscheidung":
        return None  # Gateways tragen kein Medium – Bruch wird zwischen den Aufgaben bewertet
    ma, mb = a.get("medium"), b.get("medium")
    sa, sb = (a.get("system") or "").strip(), (b.get("system") or "").strip()
    if ma and mb and ma != mb and (ma in NON_DIGITAL or mb in NON_DIGITAL):
        return f"{ma} → {mb}"
    if sa and sb and sa.lower() != sb.lower() and (ma in NON_DIGITAL or mb in NON_DIGITAL or not ma or not mb):
        return f"{sa} → {sb}"
    return None


def prozess(rows: list[dict]) -> dict:
    groups: dict[str, list[dict]] = OrderedDict()
    for r in rows:
        groups.setdefault(r["prozess"], []).append(r)
    out = {}
    for name, steps in groups.items():
        steps = sorted(steps, key=lambda s: s["nr"] or 0)
        by_nr = {s["nr"]: s for s in steps}
        succ = _successors(steps)
        breaks, handovers = [], 0
        for a_nr, targets in succ.items():
            for b_nr in targets:
                a, b = by_nr.get(a_nr), by_nr.get(b_nr)
                if not a or not b:
                    continue
                if a.get("lane") != b.get("lane"):
                    handovers += 1
                why = _is_break(a, b)
                if why:
                    breaks.append({"von": a["schritt"], "nach": b["schritt"], "von_nr": a_nr,
                                   "nach_nr": b_nr, "art": why})
        work = [s for s in steps if s.get("typ") in (None, "Aufgabe")]
        digital = [s for s in work if s.get("medium") == "Digital integriert"]
        lanes = list(OrderedDict.fromkeys(s.get("lane") or "–" for s in steps))
        out[name] = {
            "steps": steps, "succ": succ, "lanes": lanes, "medienbrueche": breaks,
            "uebergaben": handovers, "digitalgrad": (len(digital) / len(work)) if work else 0.0,
            "anzahl_schritte": len(work),
            "medien": dict(Counter(s.get("medium") or "unbekannt" for s in work)),
        }
    return {"prozesse": out}


def vsm(rows: list[dict]) -> dict:
    groups: dict[str, list[dict]] = OrderedDict()
    for r in rows:
        groups.setdefault(r["prozess"], []).append(r)
    out = {}
    for name, steps in groups.items():
        steps = sorted(steps, key=lambda s: s["nr"] or 0)
        bz = sum(s.get("bz") or 0 for s in steps)
        wz = sum(s.get("wz") or 0 for s in steps)
        rz = sum(s.get("rz") or 0 for s in steps)
        va = sum((s.get("bz") or 0) for s in steps if (s.get("wertschoepfend") or "Ja") == "Ja")
        dlz = bz + wz + rz
        rty = 1.0
        for s in steps:
            if s.get("qr") is not None:
                rty *= max(0.0, min(100.0, s["qr"])) / 100
        breaks = []
        for a, b in zip(steps, steps[1:]):
            why = _is_break(a, b)
            if why:
                breaks.append({"von": a["schritt"], "nach": b["schritt"], "art": why})
        bottleneck = max(steps, key=lambda s: (s.get("wz") or 0)) if steps else None
        out[name] = {
            "steps": steps, "bearbeitungszeit": bz, "wartezeit": wz, "ruestzeit": rz,
            "durchlaufzeit": dlz, "wertschoepfung": va,
            "flussgrad": (va / dlz) if dlz else 0.0, "rty": rty, "medienbrueche": breaks,
            "engpass": bottleneck["schritt"] if bottleneck and (bottleneck.get("wz") or 0) > 0 else None,
            "wartezeitanteil": (wz / dlz) if dlz else 0.0,
        }
    return {"prozesse": out}


def pp_score(p: dict) -> int:
    return (p.get("haeufigkeit") or 3) * (p.get("auswirkung") or 3)


def painpoints(rows: list[dict]) -> dict:
    items = [{**p, "score": pp_score(p), "stimmen": p.get("stimmen") or 0} for p in rows]
    items.sort(key=lambda p: (p["score"], p["stimmen"]), reverse=True)
    for i, p in enumerate(items, start=1):
        p["rang"] = i
    cat_score: Counter = Counter()
    cat_count: Counter = Counter()
    heat: dict[str, Counter] = defaultdict(Counter)
    for p in items:
        k = p.get("kategorie") or "Sonstige"
        cat_score[k] += p["score"]
        cat_count[k] += 1
        heat[p.get("prozess") or "–"][k] += p["score"]
    total = sum(cat_score.values()) or 1
    pareto, cum = [], 0
    for k, v in cat_score.most_common():
        cum += v
        pareto.append({"kategorie": k, "score": v, "anzahl": cat_count[k], "kumuliert": cum / total})
    return {"items": items, "pareto": pareto, "heatmap": {k: dict(v) for k, v in heat.items()},
            "kategorien": [x["kategorie"] for x in pareto], "top": items[:10]}


def ursachen(whys: list[dict], ish: list[dict]) -> dict:
    chains = []
    for w in whys:
        steps = [w.get(f"w{i}") for i in range(1, 6) if w.get(f"w{i}")]
        chains.append({**w, "kette": steps, "tiefe": len(steps),
                       "grundursache": w.get("grundursache") or (steps[-1] if steps else None)})
    fish: dict[str, dict[str, list[str]]] = OrderedDict()
    for r in ish:
        fish.setdefault(r["problem"], OrderedDict()).setdefault(r["kategorie"], []).append(
            r["ursache"] + (" ✓" if r.get("bestaetigt") == "Ja" else ""))
    return {"whys": chains, "ishikawa": fish,
            "flach": [c for c in chains if c["tiefe"] < 3]}


GWT = re.compile(r"(given|gegeben).*(when|wenn).*(then|dann)", re.I | re.S)
SOLUTION_WORDS = ["button", "datenbank", "sql", "api", "dropdown", "excel-makro", "popup", "tabelle anlegen"]


def invest_check(s: dict) -> dict:
    ziel = (s.get("ziel") or "").lower()
    ac = s.get("akzeptanz") or ""
    lines = [ln for ln in re.split(r"\n|(?<=\.)\s+(?=(?:given|gegeben))", ac, flags=re.I) if ln.strip()]
    sp = s.get("story_points")
    checks = OrderedDict()
    checks["I"] = (not re.search(r"\b(und|sowie|außerdem)\b", ziel), "Ziel enthält mehrere Anliegen ('und') – aufteilen")
    checks["N"] = (not any(w in ziel for w in SOLUTION_WORDS), "Lösung vorweggenommen – Ziel statt Umsetzung formulieren")
    checks["V"] = (len((s.get("nutzen") or "").strip()) >= 10, "Nutzen ('damit …') fehlt oder zu unkonkret")
    checks["E"] = (sp is not None, "Keine Schätzung (Story Points)")
    checks["S"] = (sp is None or sp <= 8, "Zu groß (>8 SP) – als Epic behandeln und schneiden")
    checks["T"] = (bool(lines) and all(GWT.search(ln) for ln in lines),
                   "Akzeptanzkriterien nicht im Given/When/Then-Format")
    ok = sum(1 for v, _ in checks.values() if v)
    return {"checks": {k: v for k, (v, _) in checks.items()},
            "hinweise": [msg for v, msg in checks.values() if not v], "score": ok / 6,
            "satz": f"Als {s.get('rolle')} möchte ich {s.get('ziel')}, damit {s.get('nutzen')}."}


def userstories(rows: list[dict]) -> dict:
    items = [{**s, "invest": invest_check(s)} for s in rows]
    avg = sum(i["invest"]["score"] for i in items) / len(items) if items else 0
    return {"items": items, "invest_quote": avg,
            "nachbessern": [i for i in items if i["invest"]["score"] < 1],
            "story_points": sum(i.get("story_points") or 0 for i in items)}


RELEASES = ["MVP", "Release 2", "Release 3", "Backlog"]


def storymap(rows: list[dict]) -> dict:
    acts = list(OrderedDict.fromkeys(r.get("aktivitaet") or "Sonstiges" for r in rows))
    rels = [x for x in RELEASES if any((r.get("release") or "Backlog") == x for r in rows)]
    grid = {a: {rl: [] for rl in rels} for a in acts}
    for r in rows:
        grid[r.get("aktivitaet") or "Sonstiges"][r.get("release") or "Backlog"].append(r)
    sp = {rl: sum(r.get("story_points") or 0 for r in rows if (r.get("release") or "Backlog") == rl) for rl in rels}
    return {"aktivitaeten": acts, "releases": rels, "grid": grid, "story_points": sp}


def handlungsfelder(data: Data, pp_res: dict | None) -> dict:
    pp = {p["id"]: p for p in (pp_res or {}).get("items", [])}
    names = [h["name"] for h in data.get("handlungsfelder")]
    for r in data.get("userstories") + data.get("usecases"):
        if r.get("handlungsfeld") and r["handlungsfeld"] not in names:
            names.append(r["handlungsfeld"])
    meta = {h["name"]: h for h in data.get("handlungsfelder")}
    out, covered = [], set()
    for n in names:
        stories = [s for s in data.get("userstories") if s.get("handlungsfeld") == n]
        ucs = [u for u in data.get("usecases") if u.get("handlungsfeld") == n]
        pids = {pid for x in stories + ucs for pid in (x.get("pain_point_ids") or [])}
        covered |= pids
        out.append({
            "name": n, **{k: v for k, v in meta.get(n, {}).items() if k != "name"},
            "stories": len(stories), "use_cases": [u["id"] for u in ucs],
            "story_points": sum(s.get("story_points") or 0 for s in stories),
            "pain_points": sorted(pids), "potenzial": sum(pp[p]["score"] for p in pids if p in pp),
        })
    out.sort(key=lambda x: x["potenzial"], reverse=True)
    uncovered = [p for pid, p in pp.items() if pid not in covered]
    return {"items": out, "nicht_adressiert": uncovered}


def vpc(rows: list[dict]) -> dict:
    out: dict[str, dict] = OrderedDict()
    for r in rows:
        c = out.setdefault(r["use_case_id"], {"segment": r.get("segment"), "bereiche": defaultdict(list), "relievers": []})
        c["bereiche"][r["bereich"]].append(r["eintrag"])
        if r["bereich"] in ("Pain Reliever", "Gain Creator") and r.get("bezug"):
            c["relievers"].append((r["bereich"], r["bezug"].strip().lower()))
    for c in out.values():
        refs = {b for _, b in c["relievers"]}
        c["offene_pains"] = [p for p in c["bereiche"].get("Pain", []) if not _matched(p, refs)]
        c["offene_gains"] = [g for g in c["bereiche"].get("Gain", []) if not _matched(g, refs)]
        n = len(c["bereiche"].get("Pain", [])) + len(c["bereiche"].get("Gain", []))
        c["fit"] = 1 - (len(c["offene_pains"]) + len(c["offene_gains"])) / n if n else 0
        c["bereiche"] = dict(c["bereiche"])
    return {"canvases": out}


def _matched(text: str, refs: set[str]) -> bool:
    t = text.lower()
    return any(r and (r in t or t in r) for r in refs)


def bmc(rows: list[dict]) -> dict:
    out: dict[str, dict[str, list[str]]] = OrderedDict()
    for r in rows:
        out.setdefault(r["use_case_id"], defaultdict(list))[r["baustein"]].append(r["eintrag"])
    return {"canvases": {k: dict(v) for k, v in out.items()}}


def fitgap(rows: list[dict]) -> dict:
    cats = ["Standard", "Konfiguration", "Erweiterung", "Gap"]
    cnt = Counter(r.get("abdeckung") for r in rows)
    matrix = {p: Counter() for p in ["Muss", "Soll", "Kann"]}
    for r in rows:
        matrix.setdefault(r.get("prioritaet") or "Soll", Counter())[r.get("abdeckung")] += 1
    effort = Counter()
    for r in rows:
        effort[r.get("abdeckung")] += r.get("aufwand_pt") or 0
    n = len(rows) or 1
    by_sys: dict[str, Counter] = defaultdict(Counter)
    for r in rows:
        by_sys[r.get("system") or "–"][r.get("abdeckung")] += 1
    return {
        "items": rows, "kategorien": cats, "anzahl": {c: cnt.get(c, 0) for c in cats},
        "fit_quote": (cnt.get("Standard", 0) + cnt.get("Konfiguration", 0)) / n,
        "matrix": {p: {c: m.get(c, 0) for c in cats} for p, m in matrix.items()},
        "aufwand": {c: effort.get(c, 0) for c in cats}, "aufwand_gesamt": sum(effort.values()),
        "kritische_gaps": [r for r in rows if r.get("abdeckung") == "Gap" and r.get("prioritaet") == "Muss"],
        "systeme": {k: dict(v) for k, v in by_sys.items()},
    }


# ====================================================================== Phase 3
def _criteria(kriterien: list[dict], cols: list[str]) -> list[dict]:
    if kriterien:
        crit = [dict(c) for c in kriterien if c["kriterium"] in cols]
    else:
        from .model import DEFAULT_CRITERIA
        lookup = {c[0]: c for c in DEFAULT_CRITERIA}
        crit = [{"kriterium": c, "kategorie": lookup.get(c, (c, "Nutzen"))[1],
                 "gewicht": lookup.get(c, (c, "", 100 / len(cols)))[2]} for c in cols]
    s = sum(c.get("gewicht") or 0 for c in crit) or 1
    for c in crit:
        c["w"] = (c.get("gewicht") or 0) / s
    return crit


def _nwa_scores(crit: list[dict], rows: list[dict]) -> dict[str, float]:
    return {r["_id"]: sum(c["w"] * (r.get(c["kriterium"]) or 0) for c in crit) for r in rows}


def nwa(kriterien: list[dict], cols: list[str], rows: list[dict], usecases: list[dict]) -> dict:
    crit = _criteria(kriterien, cols)
    titles = {u["id"]: u.get("titel") for u in usecases}
    scores = _nwa_scores(crit, rows)
    ranking = sorted(scores, key=scores.get, reverse=True)
    rank = {uc: i + 1 for i, uc in enumerate(ranking)}
    # Sensitivität: jedes Gewicht ±20 % variieren, Rangspanne je Use Case
    span = {uc: [rank[uc], rank[uc]] for uc in scores}
    for c in crit:
        for f in (0.8, 1.2):
            mod = [dict(x, w=x["w"] * (f if x is c else 1)) for x in crit]
            tot = sum(x["w"] for x in mod)
            for x in mod:
                x["w"] /= tot
            sc = _nwa_scores(mod, rows)
            for i, uc in enumerate(sorted(sc, key=sc.get, reverse=True), start=1):
                span[uc][0], span[uc][1] = min(span[uc][0], i), max(span[uc][1], i)
    cats = list(OrderedDict.fromkeys(c["kategorie"] for c in crit))
    items = []
    for uc in ranking:
        row = next(r for r in rows if r["_id"] == uc)
        sub = {k: sum(c["w"] * (row.get(c["kriterium"]) or 0) for c in crit if c["kategorie"] == k) for k in cats}
        items.append({"id": uc, "titel": titles.get(uc, uc), "score": scores[uc], "prozent": scores[uc] / 5,
                      "rang": rank[uc], "rang_min": span[uc][0], "rang_max": span[uc][1],
                      "teil": sub, "werte": {c["kriterium"]: row.get(c["kriterium"]) for c in crit},
                      "fehlend": [c["kriterium"] for c in crit if row.get(c["kriterium"]) is None]})
    return {"kriterien": crit, "kategorien": cats, "items": items}


def _scale10(avg5: float) -> float:
    return 1 + (avg5 - 1) / 4 * 9


def quadrant(impact: float, effort: float) -> str:
    if impact >= 5.5:
        return "Quick Win" if effort < 5.5 else "Großes Vorhaben"
    return "Lückenfüller" if effort < 5.5 else "Nicht machen"


def impact_effort(usecases: list[dict], nwa_res: dict | None) -> dict:
    by_id = {i["id"]: i for i in (nwa_res or {}).get("items", [])}
    crit = (nwa_res or {}).get("kriterien", [])
    items = []
    for u in usecases:
        imp, eff, src = u.get("impact"), u.get("effort"), "Workshop"
        n = by_id.get(u["id"])
        if (imp is None or eff is None) and n:
            src = "Nutzwertanalyse"
            ben = [c for c in crit if c["kategorie"] in ("Nutzen", "Strategie")]
            fea = [c for c in crit if c["kategorie"] in ("Machbarkeit", "Risiko")]
            if imp is None and ben:
                w = sum(c["w"] for c in ben)
                imp = _scale10(sum(c["w"] * (n["werte"].get(c["kriterium"]) or 1) for c in ben) / w)
            if eff is None and fea:
                w = sum(c["w"] for c in fea)
                eff = 11 - _scale10(sum(c["w"] * (n["werte"].get(c["kriterium"]) or 1) for c in fea) / w)
        if imp is None or eff is None:
            continue
        items.append({"id": u["id"], "titel": u.get("titel"), "impact": round(float(imp), 1),
                      "effort": round(float(eff), 1), "quadrant": quadrant(imp, eff), "quelle": src})
    groups = OrderedDict((q, [i for i in items if i["quadrant"] == q])
                         for q in ["Quick Win", "Großes Vorhaben", "Lückenfüller", "Nicht machen"])
    return {"items": items, "quadranten": groups}


def raci(roles: list[str], rows: list[dict]) -> dict:
    items, load = [], Counter()
    for r in rows:
        vals = {k: r.get(k, "") for k in roles}
        a = [k for k, v in vals.items() if "A" in v]
        rr = [k for k, v in vals.items() if "R" in v]
        problems = []
        if len(a) != 1:
            problems.append(f"{len(a)}× A (genau 1 nötig)")
        if not rr:
            problems.append("kein R")
        for k, v in vals.items():
            if v:
                load[k] += 1
        items.append({"aufgabe": r["_id"], "werte": vals, "probleme": problems})
    return {"rollen": roles, "items": items, "fehler": [i for i in items if i["probleme"]],
            "auslastung": dict(load)}


SIGN = {"Investition": -1, "Laufende Kosten": -1, "Einsparung": 1, "Mehrumsatz": 1}


def npv(rate: float, flows: list[float]) -> float:
    return sum(cf / (1 + rate) ** t for t, cf in enumerate(flows))


def irr(flows: list[float]) -> float | None:
    if not any(f < 0 for f in flows) or not any(f > 0 for f in flows):
        return None
    lo, hi = -0.99, 10.0
    f_lo, f_hi = npv(lo, flows), npv(hi, flows)
    if f_lo * f_hi > 0:
        return None
    for _ in range(200):
        mid = (lo + hi) / 2
        fm = npv(mid, flows)
        if f_lo * fm <= 0:
            hi, f_hi = mid, fm
        else:
            lo, f_lo = mid, fm
    return (lo + hi) / 2


def payback(flows: list[float]) -> float | None:
    cum = 0.0
    for t, cf in enumerate(flows):
        prev = cum
        cum += cf
        if cum >= 0 and t > 0 and prev < 0:
            return t - 1 + (-prev / cf if cf else 0)
        if cum >= 0 and t == 0 and cf >= 0:
            return 0.0
    return None


def businesscase(rows: list[dict], rate: float) -> dict:
    horizon = max((r.get("jahr") or 0) for r in rows)
    ucs: dict[str, dict] = OrderedDict()
    for r in rows:
        u = ucs.setdefault(r["use_case_id"], {"flows": [0.0] * (horizon + 1), "kosten": 0.0, "nutzen": 0.0,
                                              "invest": 0.0, "positionen": []})
        amt = abs(r.get("betrag") or 0)
        sign = SIGN.get(r.get("typ"), -1)
        u["flows"][r.get("jahr") or 0] += sign * amt
        u["positionen"].append(r)
        if sign < 0:
            u["kosten"] += amt
        else:
            u["nutzen"] += amt
        if r.get("typ") == "Investition":
            u["invest"] += amt
    for u in ucs.values():
        f = u["flows"]
        u["npv"] = npv(rate, f)
        u["roi"] = (u["nutzen"] - u["kosten"]) / u["kosten"] if u["kosten"] else None
        u["irr"] = irr(f)
        u["payback"] = payback(f)
        cum, c = [], 0.0
        for x in f:
            c += x
            cum.append(c)
        u["kumuliert"] = cum
    total = [sum(u["flows"][t] for u in ucs.values()) for t in range(horizon + 1)]
    tk = sum(u["kosten"] for u in ucs.values())
    return {"zinssatz": rate, "jahre": horizon, "use_cases": ucs,
            "portfolio": {"flows": total, "npv": npv(rate, total), "irr": irr(total), "payback": payback(total),
                          "invest": sum(u["invest"] for u in ucs.values()),
                          "roi": (sum(u["nutzen"] for u in ucs.values()) - tk) / tk if tk else None}}


# ---------------------------------------------------------------- Roadmap
def q_parse(q: str) -> tuple[int, int]:
    m = re.match(r"(\d{4})\s*-?\s*Q([1-4])", str(q).upper())
    if not m:
        raise ValueError(f"Quartal '{q}' nicht im Format JJJJ-Qn")
    return int(m.group(1)), int(m.group(2))


def q_add(q: str, n: int) -> str:
    y, k = q_parse(q)
    idx = y * 4 + (k - 1) + n
    return f"{idx // 4}-Q{idx % 4 + 1}"


def q_diff(a: str, b: str) -> int:
    ya, ka = q_parse(a)
    yb, kb = q_parse(b)
    return (yb * 4 + kb) - (ya * 4 + ka)


HORIZON_OFFSET = {"Now": 0, "Next": 2, "Later": 4}
DEFAULT_DUR = {"Quick Win": 3, "Großes Vorhaben": 9, "Lückenfüller": 4, "Nicht machen": 6}


def suggest_horizon(quad: str | None, rank: int | None, n: int) -> str:
    if quad == "Quick Win":
        return "Now"
    if quad == "Großes Vorhaben":
        return "Now" if rank is not None and rank <= max(1, n // 4) else "Next"
    if quad == "Lückenfüller":
        return "Next"
    if quad == "Nicht machen":
        return "Later"
    return "Next" if rank is None or rank <= n // 2 else "Later"


def roadmap(usecases: list[dict], plan: list[dict], ie: dict | None, nwa_res: dict | None, start: str) -> dict:
    quad = {i["id"]: i["quadrant"] for i in (ie or {}).get("items", [])}
    rank = {i["id"]: i["rang"] for i in (nwa_res or {}).get("items", [])}
    given = {p["use_case_id"]: p for p in plan}
    n = len(usecases)
    items: dict[str, dict] = OrderedDict()
    order = sorted(usecases, key=lambda u: (rank.get(u["id"], 999), u["id"]))
    for u in order:
        g = given.get(u["id"], {})
        q = quad.get(u["id"])
        hz = g.get("horizont") or suggest_horizon(q, rank.get(u["id"]), n)
        dur = g.get("dauer") or DEFAULT_DUR.get(q or "", 6)
        st = g.get("start") or q_add(start, HORIZON_OFFSET[hz])
        items[u["id"]] = {"id": u["id"], "titel": u.get("titel"), "horizont": hz, "start": st,
                          "dauer": int(dur), "mvp": g.get("mvp"), "pilot": g.get("pilot"), "kpi": g.get("kpi"),
                          "abhaengigkeiten": g.get("abhaengigkeiten") or [], "quadrant": q,
                          "rang": rank.get(u["id"]), "vorschlag": not g.get("horizont"),
                          "empfohlen": q != "Nicht machen", "handlungsfeld": u.get("handlungsfeld")}
    # Abhängigkeiten: Start nicht vor Ende der Vorgänger (max. 10 Durchläufe)
    for _ in range(10):
        changed = False
        for it in items.values():
            for d in it["abhaengigkeiten"]:
                pre = items.get(d)
                if not pre:
                    continue
                end_q = q_add(pre["start"], math.ceil(pre["dauer"] / 3))
                if q_diff(it["start"], end_q) > 0 and not given.get(it["id"], {}).get("start"):
                    it["start"], changed = end_q, True
        if not changed:
            break
    for it in items.values():
        it["start_idx"] = q_diff(start, it["start"])
        it["ende_idx"] = it["start_idx"] + max(1, math.ceil(it["dauer"] / 3))
    horizon_groups = {h: [i for i in items.values() if i["horizont"] == h] for h in ("Now", "Next", "Later")}
    n_q = max([i["ende_idx"] for i in items.values()] + [8])
    return {"start": start, "items": list(items.values()), "horizonte": horizon_groups,
            "quartale": [q_add(start, k) for k in range(n_q)]}


# ---------------------------------------------------------------- Kennzahlen
def kennzahlen(r: dict) -> dict:
    k: dict[str, Any] = {}
    if "stakeholder" in r:
        k["Stakeholder"] = len(r["stakeholder"]["items"])
    if "prozess" in r:
        p = r["prozess"]["prozesse"]
        k["Aufgenommene Prozesse"] = len(p)
        k["Medienbrüche"] = sum(len(x["medienbrueche"]) for x in p.values())
        steps = sum(x["anzahl_schritte"] for x in p.values()) or 1
        k["Digitalgrad Prozessschritte"] = sum(x["digitalgrad"] * x["anzahl_schritte"] for x in p.values()) / steps
    if "vsm" in r:
        v = list(r["vsm"]["prozesse"].values())
        k["Ø Flussgrad Wertstrom"] = sum(x["flussgrad"] for x in v) / len(v)
    if "painpoints" in r:
        k["Pain Points"] = len(r["painpoints"]["items"])
    if "userstories" in r:
        k["User Stories"] = len(r["userstories"]["items"])
        k["INVEST-Erfüllung"] = r["userstories"]["invest_quote"]
    if "usecases" in r:
        k["Use Cases"] = len(r["usecases"]["items"])
    if "impact_effort" in r:
        k["Quick Wins"] = len(r["impact_effort"]["quadranten"]["Quick Win"])
    if "fitgap" in r:
        k["Fit-Quote Standard"] = r["fitgap"]["fit_quote"]
    if "businesscase" in r:
        k["Kapitalwert Portfolio [EUR]"] = r["businesscase"]["portfolio"]["npv"]
        k["Investition Portfolio [EUR]"] = r["businesscase"]["portfolio"]["invest"]
    return k


def management_summary(r: dict, cfg: dict) -> list[str]:
    """Regelbasierte Kernaussagen (Fallback, wenn keine KI verfügbar)."""
    out = []
    if "painpoints" in r and r["painpoints"]["pareto"]:
        top = r["painpoints"]["pareto"][:2]
        out.append(f"{len(r['painpoints']['items'])} Pain Points erfasst; die Kategorien "
                   f"{' und '.join(t['kategorie'] for t in top)} vereinen "
                   f"{top[-1]['kumuliert']:.0%} der Problemlast.")
    if "prozess" in r:
        mb = sum(len(x["medienbrueche"]) for x in r["prozess"]["prozesse"].values())
        out.append(f"In {len(r['prozess']['prozesse'])} Kernprozessen wurden {mb} Medienbrüche identifiziert.")
    if "vsm" in r:
        worst = min(r["vsm"]["prozesse"].items(), key=lambda kv: kv[1]["flussgrad"])
        out.append(f"Wertstrom '{worst[0]}': nur {worst[1]['flussgrad']:.0%} der Durchlaufzeit sind wertschöpfend "
                   f"(DLZ {worst[1]['durchlaufzeit'] / 60:.1f} h).")
    if "handlungsfelder" in r and r["handlungsfelder"]["items"]:
        hf = r["handlungsfelder"]["items"][:3]
        out.append("Größtes Potenzial in den Handlungsfeldern: " + ", ".join(h["name"] for h in hf) + ".")
    if "impact_effort" in r:
        qw = r["impact_effort"]["quadranten"]["Quick Win"]
        if qw:
            out.append(f"{len(qw)} Quick Wins sofort startbar: " + ", ".join(i["titel"] for i in qw[:4]) + ".")
    if "nwa" in r and r["nwa"]["items"]:
        top = r["nwa"]["items"][:3]
        out.append("Top-3 laut Nutzwertanalyse: " + ", ".join(f"{t['titel']} ({t['prozent']:.0%})" for t in top) + ".")
    if "businesscase" in r:
        p = r["businesscase"]["portfolio"]
        pb = f", Amortisation nach {p['payback']:.1f} Jahren" if p["payback"] is not None else ""
        out.append(f"Portfolio-Business-Case: Investition {p['invest'] / 1000:,.0f} T€, Kapitalwert "
                   f"{p['npv'] / 1000:,.0f} T€ bei {r['businesscase']['zinssatz']:.0%} Kalkulationszins{pb}.")
    if "roadmap" in r:
        now = r["roadmap"]["horizonte"]["Now"]
        out.append(f"Roadmap: {len(now)} Vorhaben im Horizont NOW, Start {r['roadmap']['start']} mit MVP und Pilotierung.")
    return out
