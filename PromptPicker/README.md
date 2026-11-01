# Prompt Picker – Standalone Tray-App für Red Ink

Eigenständiges **VB.NET / WinForms / .NET 8 (Windows)**-Projekt im Design von Ask Inky. Verwendet denselben WebView2-NuGet-Stand und dasselbe Red-Ink-Icon wie Ask Inky. Es gibt **keine Projekt- oder Laufzeitabhängigkeit von Ask Inky, Office oder der grossen SharedLibrary**. Das Projekt ist in der mit der Erstlieferung gelieferten angepassten `Red Ink.sln` bereits eingetragen; in einer anderen Solution kann `PromptPicker.vbproj` über **Vorhandenes Projekt hinzufügen** eingefügt werden. **Folgeupdate (Fenster und Einfügen):** Das Änderungs-ZIP enthält ausschliesslich neue/geänderte Dateien, die über den bisherigen `PromptPicker`-Projektordner kopiert werden. Die übrige Gen2-/Ask-Inky-Codebasis bleibt unverändert. Die Solution und die Red-Ink-/Ask-Inky-Projekte bleiben unverändert.

## Start

1. Bei Erstinstallation den Ordner `PromptPicker` neben `AskInky` ablegen und bei Bedarf die Solution aus der Erstlieferung verwenden. Bei diesem UI-Update die Dateien aus dem Änderungs-ZIP mit gleicher Verzeichnisstruktur im bestehenden Projekt überschreiben.
2. In **Visual Studio 2022**, mit .NET-8-Windows-Desktop-Workload, `PromptPicker` als Startprojekt setzen und erstellen. NuGet stellt `Microsoft.Web.WebView2` 1.0.2903.40 wieder her.
3. `PromptPicker.exe` starten; die App läuft ausschliesslich im Tray. Standardmässig öffnet **Ctrl+Alt+P** den Picker, separat von den Ask-Inky-Hotkeys Ctrl+Alt+K und Ctrl+Alt+O.
4. Per Rechtsklick auf das Tray-Icon: Picker öffnen, globalen Hotkey anpassen, **Lizenzhinweise anzeigen** oder beenden. Hotkey, Picker-Fensterposition, Fenstergrösse und Breitenverhältnis der Vorschau liegen gemeinsam unter `%AppData%\Red Ink\PromptPicker\settings.json`, nicht in `redink.ini`. Bestehende `settings.json`-Dateien mit nur der Hotkey-Einstellung bleiben kompatibel. Die Fensterkoordinaten werden beim Wiederöffnen anhand der aktuell angeschlossenen Bildschirme geprüft; ein entferntes Monitorlayout führt zur zentrierten Darstellung.

`dotnet publish PromptPicker.vbproj -c Release -r win-x64` erstellt gemäss `.vbproj` eine Self-contained Single-file-Veröffentlichung. Eine Authenticode-Signatur und Windows-/VSTO-Kompilierung sind **nicht** Bestandteil dieser Lieferung.

## Konfiguration – exakt die Word-Suchpriorität

Quellbasis: `SharedMethods.Settings.vb:GetActiveConfigSource`, `SharedMethods.Constants.vb:DefaultINIPaths`, `SharedMethods.LoadConfig.vb` und `SharedMethods.PromptLibrary.vb` vom bereitgestellten Gen2-Stand (16.10.2026).

1. Die Word-Standarddatei `%AppData%\Microsoft\Word\redink.ini` hat Vorrang. `RegPath_IniPrio=False`, wie in Red Ink.
2. Nur wenn sie fehlt, wird `HKCU\Software\Red Ink\IniPath` berücksichtigt. Red Ink interpretiert dessen Dateisystemwert als **Ordner** und ergänzt `redink.ini`. HTTPS wird als bereits vollständige URL interpretiert. Ohne beides wird der Word-Standardpfad verwendet und ein Ladefehler gemeldet.
3. Die Ini wird **nur gelesen**: Zeilen `;` und Leerzeilen werden ignoriert, Schlüssel vor `=` sind nicht case-sensitive. Es werden `PromptLib` (zentral) und `PromptLibLocal` (lokal) verwendet.
4. Dateipfade werden wie in `SharedMethods.VarHelpers.ExpandEnvironmentVariables` erweitert: Umgebungsvariablen und Red-Ink-Aliasse (`%DOCUMENTS%`, `%DOWNLOADS%`, `%APPSTARTUPPATH%` usw.), umgebende Anführungszeichen, absolute und UNC-Pfade.
5. Wie in Red Ink erfolgt die Reihenfolge der Bibliotheksabfrage: **lokal vor zentral**. Kategorien aus `<Name>`/`</Name>`, Einträge `Titel|Prompt`, Rest des Textes nach dem ersten Pipe-Zeichen bleibt enthalten. Lokale Einträge tragen für die Anzeige ` (local)`.
6. Die Ini-Datei und die zentrale Library werden unter keinen Umständen geschrieben. Fehlende/unlesbare Bibliotheken werden wie beim Add-in übersprungen und im Log protokolliert.

**HTTPS-Einschränkung:** Die Quell-URL wird mit einem .NET-HTTPS-Client (ohne Redirects, mit automatischer Client-Zertifikatswahl und Windows-Standardanmeldung) gelesen. Add-in-spezifische konfigurierbare HTTP-Stacks der konfigurierbare Auswahlmechanismus des Add-in-HTTP-Stacks aus `ConfigurationResourceLoader` sind in dieser eigenständigen App nicht enthalten. Wird eine Client-Zertifikats-authentifizierte Konfigurations-URL verwendet, muss der Transport entsprechend dem vorhandenen Unternehmens-Setup erweitert werden. Der Auflösungs-/Parsing-Algorithmus selbst entspricht Word.


## Lizenzhinweise im Tray

Der Tray-Menüpunkt **„Lizenzhinweise …“** (zwischen Hotkey-Einstellungen und Beenden) öffnet einen schlichten, skalierbaren Nur-Lese-Dialog mit Copyrightangabe und den vollständigen eingebetteten Red-Ink-Lizenzbedingungen sowie Drittanbieterhinweisen. Dafür ist weder eine Internetverbindung noch Zugriff auf `redink.ini` nötig. `LICENSE.txt` wurde im einleitenden Komponentenhinweis von Ask Inky auf Prompt Picker umgestellt; der eigentliche Red-Ink-Lizenztext ab der Überschrift `RED INK` bleibt bytegenau unverändert. Die Lizenzhinweise werden bewusst nicht als zusätzlicher Header im Picker angezeigt.

## Auswahl und lokales Speichern

Mit markiertem Text in der Zielanwendung den Hotkey drücken. Der Picker versucht, die Markierung durch **Ctrl+C mit anschliessender Wiederherstellung der Zwischenablage** zu ermitteln. Nach Auswahl kann der Benutzer entweder einen vorhandenen Prompt in das vorher aktive Fenster einfügen oder die Textmarkierung mit eigenem Titel und optionaler Kategorie **in `PromptLibLocal`** speichern. Beim Speichern wird der Picker danach neu geladen; die aktuelle Selektion bleibt für weitere Aktionen verfügbar. Das Dateiformat verwendet wie Red Ink `\n` für Zeilenumbrüche und `¦` als Ersatz für Pipes im gespeicherten Prompt.

**Verbindliche Einschränkung:** Ist `PromptLibLocal` in `redink.ini` **nicht definiert oder leer**, ist die Speicherfunktion deaktiviert. Es wird weder ein Pfad abgefragt, eine lokale Datei angelegt, ein Verzeichnis eingerichtet noch `redink.ini` verändert. Nur wenn ein gültiger lokaler Zielpfad explizit konfiguriert ist, darf die dort bezeichnete Library bei einem vom Benutzer bestätigten Speichervorgang angelegt bzw. ergänzt werden. Niemals wird stattdessen in `PromptLib` gespeichert.

**Einfügen:** Der Zielprozess und – soweit Win32 dies zugänglich macht – das fokussierte Kindfenster werden vor dem Öffnen festgehalten. **Doppelklick auf einen Prompt oder Enter** löst die Auswahl aus. Im Browser werden sowohl Doppelklick- als auch Pointer-up-Ereignisse auf denselben einmaligen Commit geleitet; die WinForms-Seite schliesst den WebView-Dialog nicht synchron innerhalb der Nachrichtenzustellung. Nach dem Loslassen von Tasten/Maustasten und dem Schliessen des Dialogs wird ausschliesslich die ursprüngliche Anwendung aktiviert, deren Fokus nach Möglichkeit wiederhergestellt und der Text per Unicode-Clipboard und Win32-`SendInput` (Ctrl+V) eingefügt. Zwischen Tasteneingabe und Zwischenablage-Wiederherstellung liegt ein Sicherheitsintervall für asynchron einfügende Editoren. Die vorherige Zwischenablage wird, soweit Windows sie bereitstellt, zurückgesetzt; falls ein anderes Programm sie zwischenzeitlich geändert hat, wird diese Änderung nicht überschrieben. Der Vorgang kann in erhöhten Anwendungen, geschützten Eingabefeldern, Remote-Desktops oder Anwendungen, die Ctrl+C/Ctrl+V abfangen, eingeschränkt sein. Das Fehlen einer nachweisbaren Selektion deaktiviert lediglich **Speichern**, nicht die Prompt-Auswahl. Beim Tray-Menü-Aufruf ist die Markierungserkennung möglicherweise nicht möglich, wenn Windows den Fokus auf die Taskleiste verschoben hat; der globale Hotkey ist dafür vorgesehen.

## Oberfläche und Fallback

WebView2 stellt eine kompakte, kopfzeilenfreie Ansicht dar: gleich zu Beginn „Prompts durchsuchen“, daneben Kategorien, darunter links die Prompt-Liste und rechts die standardmässig breitere Vorschau (55 %). Ein verschiebbarer Trenner – mit Maus oder Links/Rechts bei fokussiertem Trenner – passt das Verhältnis an. Die Fensterhöhe wurde von 415 auf 540 Clientpixel (+30 %) erhöht, die Breite von 690 auf 760 Clientpixel. Beim Öffnen hat die Suche den Tastaturfokus; Auf/Ab bewegen die Auswahl und Enter fügt den markierten Prompt ein. Fensterposition, Fenstergrösse und Vorschauverhältnis werden gespeichert. Die Oberfläche verwendet 32px-SVG-Symbol-Buttons, Doppelklick/Enter und einen Speicherknopf. Der Picker und die beiden Nebenfenster sind TopMost; auch im nativen Fallback gibt es einen verschiebbaren SplitContainer. Der Hotkey-Dialog ist für längere Texte und DPI-Skalierung breiter und höher ausgelegt; der native Speichern- und Hotkey-Dialog benutzt skalierungsfreundliche Zeilenlayouts und native Vektorsymbole. Die HTML-/CSS-/JS-Dateien sind eingebettete Ressourcen; es werden keine Webseiten oder Bibliotheksinhalte ans Internet übermittelt. Nachrichten aus der Webansicht werden im WinForms-Host validiert. Fehlt WebView2 oder ist es defekt, öffnet der Picker eine native WinForms-Liste mit denselben Kernfunktionen. Die Auflösung wird bei jedem Aufruf neu gelesen; nach dem Speichern wird die Liste erneut geladen.

## Architektur / QA

- `WordConfigReader.vb`, `RedInkPathExpander.vb`: Priorität, HTTPS/Datei, Ini-Parser, kompatible Variablen-/Pfadauflösung, ausschliesslich lesen
- `PromptLibraryStore.vb`: Kategorien, Format, Reihenfolge, lokal begrenztes Schreiben
- `SelectionClipboardService.vb`: Markierung lesen, Zielaktivierung und Einfügen
- `PickerContext.vb`, `ShortcutService.vb`: Tray, globaler Hotkey, Reconnect, Fokus
- `PickerForm.vb`, `SavePromptForm.vb`, `ShortcutSettingsForm.vb`, `PickerUi.vb`, `Web/*`: Darstellung, native Vektor-Symbol-Buttons, Interaktionen
- `QA/`: Prüfprotokoll, statische Regressionstests und Chromium-UI-Test; **kein** Windows-Build und kein echter Notepad-/Office-Einfügetest wurden hier durchgeführt

Quellcodebasis von Red Ink und Ask Inky wurde **nicht verändert**; nur die Solution-Datei wurde um das neue Projekt ergänzt. Es sind keine Änderungen an `SharedMethods.DialogOwner.vb` erforderlich: Es gibt keinen Cross-Process-Dialog-Owner; die WinForms-Dialoge verwenden nur eigene Same-Process-Formulare.

## Regressionstests für Windows

Diese Lieferung wurde statisch geprüft und die eingebettete HTML-Oberfläche in Chromium getestet. **Die WinForms-WebView2-Integration, tatsächliche .NET-Windows-Kompilierung und Notepad-Einfügung wurden in dieser Linux-Umgebung nicht ausgeführt.** Für eine lokale Gegenprobe: Notepad öffnen, Cursor in die Texteingabe setzen, Picker per Hotkey öffnen. Einmal per Doppelklick einfügen; danach erneut mit Klick auf einen Eintrag und Enter; zuletzt mit Suchfokus Auf/Ab und Enter. Anschliessend Grösse, Position und Trenner ändern, schliessen, erneut öffnen und die Wiederherstellung prüfen. In der Hotkey-Konfiguration auch eine 150–200-%-DPI-Skalierung testen.

Bei nicht erfolgreicher Einfügung bitte die relevante Zeitspanne aus `%AppData%\Red Ink\PromptPicker\promptpicker.log` prüfen bzw. bereitstellen. `picker.selection.commit` belegt den Empfang der Auswahl im Host; `insertion.activate.attempt`, `insertion.dispatch` und `insertion.keys.dispatched` unterscheiden Aktivierungs-, Fokus- und Eingabeschritte. Letztere Meldung bestätigt die an Windows übergebenen Tastenereignisse **nicht** die tatsächliche Übernahme des Textes durch die Zielanwendung.

Die Ordner und Regeln `PromptLib` und `PromptLibLocal` sind unverändert: `redink.ini` wird nie geschrieben; ohne konfigurierten lokalen Pfad wird nichts angelegt oder gespeichert.
