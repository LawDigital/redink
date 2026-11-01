# Red Ink Prompt Picker – QA / Folgeupdate 4 (16.10.2026)

## Ausgangsbasis und Verträge

Die vollständige Projektbasis `PromptPicker` wurde aus dem kumulativen Folgeupdate v3 übernommen. **Alle 17 Dateien des letzten v3-Änderungs-ZIP wurden vor der Arbeit bytegenau gegen diese Basis abgeglichen.** Keine Red-Ink-Gen2-, Ask-Inky- oder `.inky`-Datei wird verändert. `redink.ini` wird niemals geschrieben; ohne einen konfigurierten `PromptLibLocal` wird nicht gespeichert oder eine Bibliothek erstellt; `PromptLib` bleibt unverändert. Ursprünglicher Hotkey, Einfügebehandlung, WebView2-UI und Fallback bleiben unverändert.

## Änderungen

1. **Startbreite 50 % grösser**: `PickerSettings.PickerClientWidth` von 760 auf **1140** Clientpixel erhöht; Höhe 540 unverändert. Wenn das gespeicherte **alte Standardmass genau 760** ist, wird es einmalig beim Lesen auf 1140 migriert. `PickerLayoutVersion=2` verhindert erneute Migration nach dem nächsten Speichern. Andere gespeicherte Breiten (auch unter 760) werden niemals migriert; X/Y, Höhe, PreviewPercent und Hotkey bleiben unverändert. Für Fenster ohne verwendbare gespeicherte Position bzw. auf verkleinerten Displays wird das neue Fenster bei Bedarf auf den aktuellen Arbeitsbereich begrenzt. Fenstergrösse und -position bleiben persistiert unter `%AppData%\Red Ink\PromptPicker\settings.json`.
2. **Tray-Lizenzhinweise**: Dezenter Menüpunkt `Lizenzhinweise …`, abgetrennt vom Konfigurationsbereich und vor `Beenden`. `LicenseInfoForm.vb` verwendet nur read-only WinForms-Controls, die gemeinsamen Symbole und normale skalierbare Layouts; vollständige Lizenztexte werden offline aus der bereits eingebetteten `LICENSE.txt` gelesen (`LegalResources.vb`). Keine zusätzliche Zeile im Haupt-Picker, keine Internetanfrage, keine Schreiboperation.
3. **Lizenztext**: Das einleitende, von Ask Inky kopierte Komponentenblatt wurde auf Prompt Picker zugeschnitten (WebView2/.NET). **Die eigentlichen Red-Ink-Lizenzbedingungen ab `RED INK` sind bytegenau unverändert** (Bytevergleich zum Ausgangsstand erfolgreich).

## Durchgeführte Prüfungen

- Bestehender kompletter `QA/run_static_checks.py`-Katalog einschliesslich neuer spezifischer Invarianten für Standardgrösse/Migration, Tray-Menü, eingebettete Lizenzressourcen und read-only Darstellung: **bestanden**. Das Prüfskript selbst liegt bei.
- `QA/test_browser_ui.py` in Chromium: **bestanden** (Suchfokus, Pfeile/Enter, Doppelklick genau einmal, Klick+Enter, Splitter, Wiederherstellungsnachrichten, mehrere Fenstergrössen, keine Schreiboption ohne lokale Library).
- Node-JavaScript-Syntaxparser: **bestanden** (in statischen Prüfungen).
- Vergleich gegen ZIP der unmittelbar vorherigen Version v3 sowie Bytevergleich des unveränderten Red-Ink-Lizenztexts: **bestanden**.
- Changed-only-Manifest und Unified Diff gegen v3 sowie SHA-256 der geänderten und aller kumulierten Quelldateien: **erstellt**. Vollständige kumulative Projektbasis intern verwendet, nicht als redundantes zweites ZIP ausgeliefert.

## Grenzen der Prüfungen

Es steht hier keine Windows/.NET-8-WinForms-Build- und Laufzeitumgebung zur Verfügung. **Weder tatsächliche Visual-Studio-Kompilierung noch Tray-Menü-/Dialog-/DPI-Laufzeittest oder Einfügetests in Notepad/Word/Outlook wurden ausgeführt.** Die vorhandenen statischen und Chromium-Tests ersetzen diese nicht.

## Empfohlene Windows-Gegenprobe

1. In Visual Studio 2022 neu bauen; Tray mit Kontextmenü öffnen: `Lizenzhinweise …` und Scroll-/Schliessen-Funktion prüfen.
2. Alte `settings.json` mit v3-Standardbreite 760 öffnen: Picker muss mit 1140 Pixeln starten (soweit Monitor erlaubt); danach erneut öffnen und stabil dieselbe Grösse erhalten.
3. `settings.json` mit individuell gewählter Breite (z. B. 900) öffnen: Breite und Position müssen unverändert bleiben; Hotkey und Vorschau-Splitter unverändert.
4. Ohne `settings.json` erststarten: 1140×540 Clientpixel; kleineres Display: Fenster darf nicht vom Bildschirm abgeschnitten sein.
5. Bestehenden Notepad-/Office-Einfügetest und Libraries ohne `PromptLibLocal` wiederholen.
