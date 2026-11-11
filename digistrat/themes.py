"""Farbschemata (Corporate Design) für PowerPoint- und Excel-Ausgaben.

Auswahl in projekt.yaml:  theme: {preset: ingenics}
Einzelwerte (primary, accent, …) in projekt.yaml überschreiben das Preset.

'ingenics': Farbwerte aus den CSS-Variablen von www.ingenics.com (Stand 11/2026):
  --color-primary-indigo #152C5A, --color-primary-blue #004F9F, --color-primary-cyan #00AFE9,
  --color-secondary-cyan-dark #0089D0, --color-secondary-cyan-light #83D0F5,
  --color-secondary-anthracite #F0F0F2. Website-Schrift 'Helvetica Now Display' -> Office-Ersatz Arial.
"""
from __future__ import annotations

from typing import Any

PRESETS: dict[str, dict[str, Any]] = {
    "ingenics": {
        "primary": "152C5A",    # Primary Indigo – Headlines, Tabellenköpfe, Titelfolie
        "accent": "00AFE9",     # Primary Cyan – Akzentlinie, Hervorhebungen
        "secondary": "004F9F",  # Primary Blue – Diagrammreihen, Buttons
        "tertiary": "0089D0",   # Secondary Cyan dark
        "soft": "83D0F5",       # Secondary Cyan light
        "light": "F0F0F2",      # Secondary Anthracite – Flächen
        "font": "Arial",        # Website: Helvetica Now Display
    },
    "neutral": {
        "primary": "1F3A5F",
        "accent": "E07A1F",
        "secondary": "2F6DB5",
        "light": "F3F4F6",
        "font": "Calibri",
    },
}


def resolve(theme: dict[str, Any] | None) -> dict[str, Any]:
    theme = dict(theme or {})
    base = dict(PRESETS.get(theme.get("preset", "ingenics"), PRESETS["ingenics"]))
    for k, v in theme.items():
        if v not in (None, "") and k != "preset":
            base[k] = v
    return base
