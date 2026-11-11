"""Projektordner, Einlesen der Workshop-Erfassungen und Excel-Eingabevorlagen."""
from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml
from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from .model import (DEFAULT_CRITERIA, DEFAULT_RACI_ROLES, METHODS, PHASES, Method,
                    coerce, methods_by_phase, sheet_to_method)

DEFAULT_CONFIG: dict[str, Any] = {
    "kunde": "Muster GmbH",
    "branche": "Produzierendes Unternehmen",
    "projekt": "Digitalisierungsstrategie",
    "berater": "",
    "startdatum": dt.date.today().isoformat(),
    "roadmap_start": f"{dt.date.today().year + 1}-Q1",
    "zinssatz": 0.08,
    "theme": {
        "preset": "ingenics",  # siehe digistrat/themes.py; Einzelwerte (primary, accent, font) überschreibbar
        "pptx_template": "",  # optional: Pfad zu Corporate-Design-Master (.pptx)
        "logo": "",
    },
    "ki": {"model": "claude-opus-5-5", "effort": "high"},
}

HEADER_FILL = PatternFill("solid", fgColor="152C5A")
HEADER_FONT = Font(color="FFFFFF", bold=True)


@dataclass
class Project:
    root: Path
    config: dict[str, Any] = field(default_factory=dict)

    # ---------------------------------------------------------- Ordner
    @property
    def input_dir(self) -> Path:
        return self.root / "01_Workshop-Input"

    @property
    def workshop_dir(self) -> Path:
        return self.root / "00_Workshopplanung"

    @property
    def output_dir(self) -> Path:
        return self.root / "02_Ergebnisse"

    @property
    def raw_dir(self) -> Path:
        return self.root / "01_Workshop-Input" / "Rohdaten"

    @classmethod
    def create(cls, root: str | Path, **overrides: Any) -> "Project":
        root = Path(root)
        cfg = _deep_merge(DEFAULT_CONFIG, overrides)
        p = cls(root, cfg)
        for d in (p.input_dir, p.workshop_dir, p.output_dir, p.raw_dir):
            d.mkdir(parents=True, exist_ok=True)
        p.save_config()
        write_input_templates(p.input_dir / "Workshop-Erfassung.xlsx")
        return p

    @classmethod
    def open(cls, root: str | Path) -> "Project":
        root = Path(root)
        cfg_file = root / "projekt.yaml"
        if not cfg_file.exists():
            raise FileNotFoundError(f"{cfg_file} fehlt – mit 'digistrat init' anlegen.")
        cfg = _deep_merge(DEFAULT_CONFIG, yaml.safe_load(cfg_file.read_text(encoding="utf-8")) or {})
        return cls(root, cfg)

    def save_config(self) -> None:
        self.root.mkdir(parents=True, exist_ok=True)
        (self.root / "projekt.yaml").write_text(
            yaml.safe_dump(self.config, allow_unicode=True, sort_keys=False), encoding="utf-8")

    # ---------------------------------------------------------- Einlesen
    def load(self) -> "Data":
        return load_inputs(sorted(self.input_dir.glob("*.xlsx")))


@dataclass
class Data:
    """Alle eingelesenen Workshop-Ergebnisse. tables[method_key] = Liste von Zeilen-Dicts."""
    tables: dict[str, list[dict[str, Any]]] = field(default_factory=dict)
    wide: dict[str, tuple[list[str], list[dict[str, Any]]]] = field(default_factory=dict)
    issues: list[str] = field(default_factory=list)
    sources: dict[str, list[str]] = field(default_factory=dict)

    def get(self, key: str) -> list[dict[str, Any]]:
        return self.tables.get(key, [])

    def has(self, key: str) -> bool:
        return bool(self.tables.get(key)) or bool(self.wide.get(key, ([], []))[1])


def load_inputs(files: list[Path]) -> Data:
    data = Data()
    for f in files:
        if f.name.startswith("~$"):
            continue
        wb = load_workbook(f, data_only=True, read_only=True)
        for ws in wb.worksheets:
            m = sheet_to_method(ws.title)
            if not m:
                continue
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                continue
            headers = [str(h).strip() if h is not None else "" for h in rows[0]]
            body = [r for r in rows[1:] if any(v not in (None, "") for v in r)]
            if not body:
                continue
            data.sources.setdefault(m.key, []).append(f"{f.name}:{ws.title}")
            if m.wide:
                _read_wide(m, headers, body, data)
            else:
                _read_long(m, headers, body, data, f.name)
    _validate(data)
    return data


def _read_long(m: Method, headers: list[str], body: list[tuple], data: Data, fname: str) -> None:
    lookup = m.field_by_header()
    keys = {f.key: f for f in m.fields}
    colmap: dict[int, Any] = {}
    for i, h in enumerate(headers):
        hl = h.lower()
        f = lookup.get(hl) or keys.get(hl)
        if f is None:  # Toleranz: Präfix-Match ("Einfluss" -> "Einfluss (1-5)")
            f = next((ff for ff in m.fields if ff.header.lower().startswith(hl) and hl), None)
        if f:
            colmap[i] = f
    missing = [f.header for f in m.fields if f.required and f not in colmap.values()]
    if missing:
        data.issues.append(f"{fname}/{m.sheet}: Pflichtspalten fehlen: {', '.join(missing)}")
    out = data.tables.setdefault(m.key, [])
    for n, r in enumerate(body, start=2):
        row: dict[str, Any] = {f.key: None for f in m.fields}
        for i, f in colmap.items():
            if i < len(r):
                row[f.key] = coerce(f, r[i])
        for f in m.fields:
            v = row[f.key]
            if f.required and v in (None, [], ""):
                data.issues.append(f"{m.sheet} Zeile {n}: '{f.header}' ist leer")
            if isinstance(v, (int, float)) and f.min is not None and v < f.min:
                data.issues.append(f"{m.sheet} Zeile {n}: '{f.header}'={v} < {f.min}")
            if isinstance(v, (int, float)) and f.max is not None and v > f.max:
                data.issues.append(f"{m.sheet} Zeile {n}: '{f.header}'={v} > {f.max}")
            if f.type == "choice" and v and f.choices and v not in f.choices:
                data.issues.append(f"{m.sheet} Zeile {n}: '{v}' ist kein gültiger Wert für '{f.header}'")
        out.append(row)


def _read_wide(m: Method, headers: list[str], body: list[tuple], data: Data) -> None:
    cols = [h for h in headers[1:] if h]
    rows = []
    for r in body:
        key = r[0]
        if key in (None, ""):
            continue
        vals = {}
        for h, v in zip(headers[1:], r[1:]):
            if not h:
                continue
            if m.key == "nwa":
                try:
                    vals[h] = float(str(v).replace(",", ".")) if v not in (None, "") else None
                except ValueError:
                    vals[h] = None
            else:
                vals[h] = str(v).strip().upper() if v not in (None, "") else ""
        rows.append({"_id": str(key).strip(), **vals})
    prev_cols, prev_rows = data.wide.get(m.key, ([], []))
    data.wide[m.key] = (prev_cols + [c for c in cols if c not in prev_cols], prev_rows + rows)


def _validate(data: Data) -> None:
    ids = {r["id"] for r in data.get("painpoints") if r.get("id")}
    uc_ids = {r["id"] for r in data.get("usecases") if r.get("id")}
    for key in ("userstories", "usecases"):
        for r in data.get(key):
            for pid in r.get("pain_point_ids") or []:
                if ids and pid not in ids:
                    data.issues.append(f"{METHODS[key].sheet} {r.get('id')}: unbekannte Pain-Point-ID {pid}")
    for key in ("vpc", "bmc", "roadmap", "businesscase"):
        for r in data.get(key):
            if uc_ids and r.get("use_case_id") and r["use_case_id"] not in uc_ids:
                data.issues.append(f"{METHODS[key].sheet}: unbekannte Use-Case-ID {r['use_case_id']}")
    crit = data.get("kriterien")
    if crit:
        s = sum(c.get("gewicht") or 0 for c in crit)
        if abs(s - 100) > 0.5:
            data.issues.append(f"NWA_Kriterien: Gewichte summieren sich auf {s:.1f} % statt 100 % (wird normiert)")


# ---------------------------------------------------------------- Vorlagen
def write_input_templates(path: Path, methods: list[str] | None = None,
                          criteria: list[str] | None = None, roles: list[str] | None = None,
                          use_case_ids: list[str] | None = None) -> Path:
    """Erzeugt die Excel-Erfassungsvorlage (ein Blatt je Methode) inkl. Dropdowns und Hilfetexten."""
    wb = Workbook()
    guide = wb.active
    guide.title = "Anleitung"
    _write_guide(guide)
    keys = methods or list(METHODS)
    for key in keys:
        m = METHODS[key]
        ws = wb.create_sheet(m.sheet)
        if m.wide:
            if key == "nwa":
                heads = [m.fields[0].header] + (criteria or [c[0] for c in DEFAULT_CRITERIA])
            else:
                heads = [m.fields[0].header] + (roles or DEFAULT_RACI_ROLES)
            _header(ws, heads)
            dv = (DataValidation(type="whole", operator="between", formula1="1", formula2="5",
                                 showErrorMessage=True, error="Bitte 1-5")
                  if key == "nwa" else
                  DataValidation(type="list", formula1='"R,A,C,I,A/R"', allow_blank=True))
            ws.add_data_validation(dv)
            dv.add(f"B2:{get_column_letter(len(heads))}300")
            if key == "nwa":
                for i, uc in enumerate(use_case_ids or [], start=2):
                    ws.cell(i, 1, uc)
            continue
        _header(ws, m.headers)
        for i, f in enumerate(m.fields, start=1):
            col = get_column_letter(i)
            c = ws.cell(1, i)
            txt = f.help or ""
            if f.type == "choice":
                txt = (txt + "\n" if txt else "") + "Werte: " + ", ".join(f.choices or [])
            if f.required:
                txt = "PFLICHTFELD\n" + txt
            if txt.strip():
                c.comment = Comment(txt.strip(), "digistrat")
            if f.type == "choice" and f.choices:
                dv = DataValidation(type="list", formula1='"' + ",".join(f.choices) + '"', allow_blank=True)
                ws.add_data_validation(dv)
                dv.add(f"{col}2:{col}500")
            elif f.type in ("int", "float") and f.min is not None and f.max is not None:
                dv = DataValidation(type="decimal", operator="between", formula1=str(f.min),
                                    formula2=str(f.max), showErrorMessage=True,
                                    error=f"Wert zwischen {f.min:g} und {f.max:g}")
                ws.add_data_validation(dv)
                dv.add(f"{col}2:{col}500")
            ws.column_dimensions[col].width = 42 if f.type == "text" else max(14, len(f.header) + 4)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def _header(ws, heads: list[str]) -> None:
    for i, h in enumerate(heads, start=1):
        c = ws.cell(1, i, h)
        c.fill, c.font = HEADER_FILL, HEADER_FONT
        c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[get_column_letter(i)].width = max(14, len(h) + 4)
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 32


def _write_guide(ws) -> None:
    ws["A1"] = "Digitalisierungsstrategie – Workshop-Erfassung"
    ws["A1"].font = Font(size=16, bold=True, color="152C5A")
    ws["A2"] = ("Je Methode ein Tabellenblatt. Zeile 1 = Kopfzeile (nicht umbenennen), ab Zeile 2 erfassen. "
                "Kommentare an den Spaltenköpfen erklären die Felder. Blätter dürfen auch auf mehrere Dateien "
                "verteilt werden – maßgeblich ist der Blattname.")
    ws["A2"].alignment = Alignment(wrap_text=True)
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 90
    r = 4
    for phase, ms in methods_by_phase().items():
        ws.cell(r, 1, PHASES[phase]).font = Font(bold=True, size=12, color="E07A1F")
        r += 1
        for m in ms:
            ws.cell(r, 1, m.title)
            ws.cell(r, 2, f"Blatt: {m.sheet}")
            ws.cell(r, 3, m.purpose).alignment = Alignment(wrap_text=True)
            r += 1
        r += 1


def _deep_merge(a: dict, b: dict) -> dict:
    out = dict(a)
    for k, v in (b or {}).items():
        out[k] = _deep_merge(out[k], v) if isinstance(v, dict) and isinstance(out.get(k), dict) else v
    return out


def append_rows(path: Path, method_key: str, rows: list[dict[str, Any]]) -> Path:
    """Hängt (z.B. KI-extrahierte) Zeilen an das Methodenblatt einer Erfassungsdatei an."""
    m = METHODS[method_key]
    if path.exists():
        wb = load_workbook(path)
    else:
        write_input_templates(path, [method_key])
        wb = load_workbook(path)
    if m.sheet not in wb.sheetnames:
        ws = wb.create_sheet(m.sheet)
        _header(ws, m.headers if not m.wide else [m.fields[0].header] + list(rows[0].keys())[1:])
    ws = wb[m.sheet]
    heads = [c.value for c in ws[1]]
    if m.wide:
        for r in rows:
            vals = list(r.values())
            for h in list(r.keys())[1:]:
                if h not in heads:
                    heads.append(h)
                    ws.cell(1, len(heads), h)
            ws.append([vals[0]] + [r.get(h) for h in heads[1:]])
    else:
        by_header = {f.header: f.key for f in m.fields}
        for r in rows:
            ws.append([_cell(r.get(by_header.get(h, h))) for h in heads])
    wb.save(path)
    return path


def _cell(v: Any) -> Any:
    if isinstance(v, list):
        return ", ".join(map(str, v))
    return v
