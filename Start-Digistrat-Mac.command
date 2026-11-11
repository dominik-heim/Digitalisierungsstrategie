#!/bin/bash
# Digistrat starten (macOS) - per Doppelklick. Beim ersten Start wird alles automatisch installiert.
cd "$(dirname "$0")" || exit 1
if ! command -v python3 >/dev/null 2>&1 || ! python3 -c 'import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)'; then
  echo "Python 3.10 oder neuer fehlt. Bitte installieren: https://www.python.org/downloads/"
  open "https://www.python.org/downloads/"
  read -r -p "Enter zum Schliessen" _
  exit 1
fi
if [ ! -x .venv/bin/python ]; then
  echo "Erster Start: Installation laeuft, das dauert 1-3 Minuten ..."
  python3 -m venv .venv || exit 1
  .venv/bin/python -m pip install --upgrade pip >/dev/null
  .venv/bin/python -m pip install -e ".[app]" || { read -r -p "Installation fehlgeschlagen. Enter zum Schliessen" _; exit 1; }
fi
# Streamlit-Abfrage nach E-Mail-Adresse beim ersten Start unterdruecken
mkdir -p "$HOME/.streamlit"
[ -f "$HOME/.streamlit/credentials.toml" ] || printf '[general]\nemail = ""\n' > "$HOME/.streamlit/credentials.toml"
echo "Digistrat startet - der Browser oeffnet sich gleich (http://localhost:8501). Zum Beenden dieses Fenster schliessen."
.venv/bin/python -m streamlit run app/streamlit_app.py --browser.gatherUsageStats false
