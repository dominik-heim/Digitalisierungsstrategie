@echo off
REM Digistrat starten (Windows) - per Doppelklick. Beim ersten Start wird alles automatisch installiert.
cd /d "%~dp0"
set PY=
where py >nul 2>nul && set PY=py -3
if "%PY%"=="" ( where python >nul 2>nul && set PY=python )
if "%PY%"=="" (
  echo.
  echo Python ist nicht installiert.
  echo Bitte Python 3.10 oder neuer installieren: https://www.python.org/downloads/
  echo WICHTIG: Im Installer "Add python.exe to PATH" ankreuzen. Danach diese Datei erneut starten.
  start https://www.python.org/downloads/
  pause
  exit /b 1
)
if not exist ".venv\Scripts\python.exe" (
  echo Erster Start: Installation laeuft, das dauert 1-3 Minuten ...
  %PY% -m venv .venv || ( echo Fehler beim Anlegen der Umgebung & pause & exit /b 1 )
  ".venv\Scripts\python.exe" -m pip install --upgrade pip >nul
  ".venv\Scripts\python.exe" -m pip install -e ".[app]" || ( echo Installation fehlgeschlagen & pause & exit /b 1 )
)
REM Streamlit-Abfrage nach E-Mail-Adresse beim ersten Start unterdruecken
if not exist "%USERPROFILE%\.streamlit" mkdir "%USERPROFILE%\.streamlit"
if not exist "%USERPROFILE%\.streamlit\credentials.toml" (
  echo [general]> "%USERPROFILE%\.streamlit\credentials.toml"
  echo email = "">> "%USERPROFILE%\.streamlit\credentials.toml"
)
echo Digistrat startet - der Browser oeffnet sich gleich (http://localhost:8501).
echo Zum Beenden dieses Fenster schliessen.
".venv\Scripts\python.exe" -m streamlit run app\streamlit_app.py --browser.gatherUsageStats false
pause
