# Red Ink Usage Server – Installation mit Python

Das ZIP vollständig entpacken und ein Terminal im Ordner `RedInkUsageServer`
öffnen. Benötigt werden Python 3.11 oder neuer (empfohlen ist Python 3.12),
`venv`, `pip` und Zugriff auf den für pip konfigurierten Paketindex. Das
Paket enthält den Servercode. Die Laufzeitabhängigkeiten werden pro Rechner
in einer virtuellen Umgebung installiert.

Der Datenordner muss beschreibbar sein und sollte ausserhalb des entpackten
Programmordners liegen. Alle Befehle einer Installation verwenden denselben
Datenordner. `init` nur bei einer Neuinstallation ausführen. Es erstellt
die Konfiguration und zeigt das Administratorpasswort sowie die beiden
Tokens einmalig an. Diese Angaben sicher aufbewahren.

## Windows (PowerShell)

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m pip check
$DataDir = Join-Path $env:LOCALAPPDATA 'RedInkUsageServer'
.\.venv\Scripts\python.exe run_server.py init --data-dir "$DataDir"
.\.venv\Scripts\python.exe run_server.py check --data-dir "$DataDir"
.\.venv\Scripts\python.exe run_server.py serve --no-browser --data-dir "$DataDir"
```

Falls Python ohne `py` installiert ist, für die erste Zeile `python -m venv
.venv` verwenden. Für einen anderen kompatiblen, im Launcher registrierten
Interpreter kann `-3.12` entsprechend angepasst werden.

## Linux und macOS (Bash oder Zsh)

```bash
python3 -m venv .venv
./.venv/bin/python -m pip install -r requirements.txt
./.venv/bin/python -m pip check
DATA_DIR="$HOME/RedInkUsageServer-data"
./.venv/bin/python run_server.py init --data-dir "$DATA_DIR"
./.venv/bin/python run_server.py check --data-dir "$DATA_DIR"
./.venv/bin/python run_server.py serve --no-browser --data-dir "$DATA_DIR"
```

Falls das Linux-System `venv` separat liefert, beispielsweise unter
Debian/Ubuntu, das passende Paket installieren (typischerweise `python3-venv`).

## Start, Betrieb und Updates

- Das Dashboard ist mit der Standardkonfiguration unter <http://127.0.0.1:8765>
  erreichbar. Zum Beenden des Vordergrundprozesses `Ctrl+C` drücken.
- Für Netzwerkbetrieb, TLS und die Red Ink Add-ins den
  [Administrator Guide](docs/ADMINISTRATOR_GUIDE.md) verwenden. In seinen
  CLI-Beispielen den EXE-Aufruf durch den Python-Aufruf aus dem jeweiligen
  Abschnitt oben ersetzen. Die aktive Konfiguration ist `server.toml` im
  Datenordner. `config/server.example.toml` im Paket dient als Referenz.
- Beim Update den Server stoppen, den Datenordner sichern und das neue ZIP in
  einen neuen Programmordner entpacken. Die virtuelle Umgebung neu erstellen,
  die Abhängigkeiten installieren und `check` sowie `serve` mit dem
  bisherigen Datenordner ausführen. Beim Start erfolgen gegebenenfalls
  Datenbankmigrationen.
- Die [Benutzeranleitung](docs/USER_GUIDE.md) und
  [Hinweise zu Drittanbieterabhängigkeiten](docs/THIRD_PARTY_LICENSES.md)
  liegen im Ordner `docs`.
