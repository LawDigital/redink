# Ask Inky – Windows-Desktopprojekt

Eigenständiges VB.NET-Projekt für Visual Studio 2022, Windows x64 und .NET 8. Die vorhandenen Gen2-/VSTO-Projekte bleiben unabhängig. Die beiden Dialoge verwenden einen nativen Windows-Fensterrahmen und eine vollständig lokale WebView2-Oberfläche.

## Aktualisierung

Die geänderten Dateien unter `AskInky/` über den bestehenden Projektordner kopieren; dabei die Unterordner beibehalten. Die eigene Solution-Datei und die bestehenden Assets nicht ersetzen oder entfernen. Bei Nachfrage das geänderte Projekt neu laden. Das SDK-Projekt nimmt neue `.vb`-Dateien automatisch auf; HTML, CSS, JavaScript und Übersetzungen sind explizit als Assembly-Ressourcen in `AskInky.vbproj` eingebettet. Anschließend **AskInky → Build** und **Debug → Start**. Bestehende Konfiguration und geschützte Registrierung bleiben erhalten.

## Bedienung

Text in einer anderen Anwendung markieren und **Ctrl + Alt + K** drücken. Der Desktop-Client wartet auf das Loslassen der Modifikatortasten und versucht anschließend, die Auswahl über Strg+C zu übernehmen. Bei einem Fokuswechsel wird die Erfassung abgebrochen statt an die falsche Anwendung zu senden. Nicht alle Anwendungen unterstützen diesen Zugriff; insbesondere geschützte Felder und höher privilegierte Anwendungen benötigen einen gesonderten Windows-Test.

Das Hauptfenster enthält ausschließlich Auswahl, Frage und HTML-Antwort. Rechts stehen Einstellungen, Senden und Kopieren als SVG-Symbole mit lokalisierten Tooltips. Es gibt keine zusätzliche Titelzeile, kein Logo im Textbereich und keine Feldüberschriften. Titel und Original-Icon bleiben im nativen Fensterrahmen; das Icon bleibt auch im Tray und trägt dort die Bezeichnung **Ask Inky**.

**Enter** sendet ohne Rückfrage; **Shift+Enter** fügt einen Zeilenumbruch ein. Während einer Anfrage sind Aktionsbuttons und Textfelder deaktiviert. Bei einem Fehler wird wieder freigegeben; die vorhandene Antwort geht nicht verloren. Das Schließen des Hauptfensters blendet es aus; zum vollständigen Beenden dient der Tray-Eintrag.

Ein einfacher Klick auf das Tray-Symbol zeigt die Einträge Aufrufen, Screenshot-OCR, Einstellungen und Beenden – in der gespeicherten Sprache. Der Aufruf aus dem Tray öffnet das vorhandene Fenster, ohne aus dem inzwischen aktiven Tray-Menü eine vermeintliche Textauswahl zu kopieren. Nur der globale Tastaturaufruf erfasst Text aus der fremden Anwendung. Es gibt keinen globalen Maus-Hook und keine Rechtsklick-Geste mehr. Einfügen und Eintippen sind entfernt.

## Einstellungen

Übernommen sind die Verbindungs-, Registrierungs-, Lizenz- und Darstellungseinstellungen des mitgelieferten Office-Web-Add-ins: Sprache (E/D/F/IT), Hell-/Dunkelmodus, Helper-URL, Helper-API-Schlüssel, Verbindungstest, einmaliger Registrierungslink, lokale Aufhebung der Registrierung, Lizenz-Benutzername, Schriftgröße, Temperatur und optionaler Request-/Response-Debugbereich. Dazu kommen der konfigurierbare Desktop-Hotkey und der schon zuvor verwendete, konfigurierbare Systemprompt.

Der **Helper-API-Schlüssel** ist der Zugangsschlüssel zum Red Ink Helper, nicht der API-Schlüssel des eigentlichen KI-Anbieters. Der KI-Anbieter bleibt auf dem Helper konfiguriert. Schlüssel werden im Passwortfeld erfasst und mit Windows DPAPI für den aktuellen Benutzer geschützt gespeichert. Beim Wechsel des Server-Ursprungs werden bestehende Gerätezugangsdaten und ein unverändert übernommener alter API-Schlüssel nicht an den neuen Server weitergegeben; einen neuen Schlüssel ausdrücklich eingeben.

Für den Hotkey entweder in das Erfassungsfeld klicken und die Kombination drücken oder Modifikatoren und Taste auswählen. **Standard** setzt Ctrl+Alt+K. Speichern versucht zuerst, die neue Kombination global zu registrieren; erst nach erfolgreicher Speicherung wird die alte Registrierung freigegeben. Bei einem Konflikt bleibt die bisherige Kombination bestehen. Windows meldet globale Reservierungen; anwendungsinterne Belegungen lassen sich damit nicht vollständig erkennen. Eine bereits gespeicherte individuelle Kombination wird beim Update nicht überschrieben.

Die Office-spezifischen Optionen für Änderungsnachverfolgung, Kommentar-Präfix und Office-Schreibprotokoll sind nicht vorhanden. Ebenso ist der MCP-/Agenten-Quelleneditor nicht portiert: Dieses Desktopprojekt verwendet weiterhin `/llm/chat`, nicht den erweiterten Agenten-Chat. Es gibt keine wirkungslosen Platzhalter für diese Funktionen.

## Localhost und Authentifizierung

Leer oder `localhost` wird zu `http://localhost`; einen abweichenden Port ausdrücklich angeben, beispielsweise `http://localhost:3000`. HTTP ohne Portangabe bedeutet Port 80. HTTP ist nur für Loopback-Adressen zugelassen; externe Verbindungen benötigen HTTPS. Lokal besteht keine zusätzliche clientseitige Registrierungssperre. Eine serverseitige Authentifizierungs- oder Lizenzanforderung wird nicht umgangen.

**Registry-HelperLink und Erste Einrichtung (V121026, 12. Oktober 2026):** Bei jedem Programmstart prüft Ask Inky, sofern nicht deaktiviert, unter `HKEY_CURRENT_USER\Software\Red Ink` zuerst den Zeichenfolgenwert `HelperLink`, dann kompatible alte Registrierungs-/Einladungswertnamen. Auch wenn Host, API-Schlüssel oder Gerätezugangsdaten bereits lokal ausgefüllt sind, hat **ein neuer gültiger Registry-Link Vorrang**. `https://host/?api_key=...` ist ein zulässiges URL-Format; optional können `uid` (Lizenzbenutzer) und `lang` (Oberflächensprache) angegeben werden. Ein Konfigurationslink ohne `auth` übernimmt die Werte lokal ohne HTTP-Anfrage; ausschließlich Auth-Einladungslinks registrieren das Gerät durch einmaligen JSON-GET. Schlägt der Versuch fehl, bleibt eine vorhandene Konfiguration unverändert.

Ein übernommener Link wird durch einen SHA-256-Fingerabdruck in `config.json` identifiziert; zusätzlich schützt `%LOCALAPPDATA%\AskInky\registry-registration-attempt.sha256` gegen erneute GET-Aufrufe möglicherweise verbrauchter Links. Bei unverändertem Registry-Link wird **keine zweite Registrierung** gesendet. Unter **Erweiterte Einstellungen** kann die standardmäßig deaktivierte Option **„Registry-Wert nicht übernehmen“** aktiviert werden. Dann dürfen die Verbindungs-/Registrierungsfelder wieder manuell bearbeitet werden und beim nächsten Start erfolgt kein Registry-Import. Eine manuelle Änderung löst die Zuordnung zum bisherigen Registry-Link; bei erneut gewünschter Registry-Registrierung kann wegen Einmalverwendung ein **neuer** Link nötig sein. Ist der erfolgreich übernommene Registry-Link noch vorhanden und die Option deaktiviert, sind die entsprechenden Felder vor versehentlichem Überschreiben geschützt. Fehlen Host und Zugangsdaten ganz und gelingt keine automatische Einrichtung, öffnet sich die Einstellungsseite mit fokussiertem Link-Feld und lokalisierter Anweisung. Beschädigte bestehende Konfigurationen werden nicht überschrieben.

Die Geräteregistrierung verwendet den tatsächlich gelieferten Office-Web-Vertrag: exakt ein GET auf die vollständige Einmal-URL, `Accept: application/json`, keine Redirects, keine automatische Wiederholung und kein Aufruf beim bloßen Einfügen des Links. Erfolg erfordert `status=ok`, `registered=true`, `userId` und `deviceToken`. `api_key`, `uid` und eine unterstützte Sprache aus dem Link werden übernommen. Erhaltene Zugangsdaten werden sofort gespeichert, unabhängig von einem noch ungespeicherten Hotkey. Bei einem Speicherfehler bleibt der Dialog mit den erhaltenen Zugangsdaten offen; erneut Speichern verwenden, nicht den verbrauchten Link erneut aufrufen. „Registrierung aufheben“ entfernt die lokale Gerätezuordnung, widerruft aber kein serverseitiges Gerät.

Normale Anfragen verwenden die vorhandenen Header `Authorization: Bearer`, `X-User-Id`, `X-Device-Token`, `X-License-User` und `X-RedInk-Session`, soweit konfiguriert. Der Verbindungstest prüft `/health` und danach den Zugriff auf `/config`; Erreichbarkeit und Berechtigung werden getrennt angezeigt. Die Lizenzinformation wird aus `config.license` gelesen. KI-Anfragen gehen unverändert an `/llm/chat` mit `prompt_system`, `prompt_user` und `temperature`.

## Darstellung, Zwischenablage und Sicherheit

`Web/office-web.css` ist die unveränderte Original-CSS-Datei aus dem gelieferten Add-in. `Web/desktop.css` ergänzt nur das Desktoplayout und die Symbolgeometrie. Die Oberfläche lädt keine CDN-Skripte, Schriftdateien oder externen UI-Assets.

Markdown wird mit der bestehenden Markdig-Version in HTML gewandelt. Die lokale Web-Oberfläche bereinigt dieses HTML deterministisch vor der Anzeige. Aus genau diesem bereinigten Inhalt werden formatierte HTML-Kopie und Unicode-Klartext erzeugt. Zusätzlich kann nur der vom Benutzer **innerhalb der Antwort markierte Teil** kopiert werden; das eigene Symbol wird ausschließlich bei einer passenden aktuellen Textauswahl eingeblendet. Andere Eingabefelder oder eine Auswahl aus früheren Antworten aktivieren es nicht. CF_HTML verwendet UTF-8 und Bytepositionen; Windows-UnicodeText verwendet UTF-16. Überschriften, Tabellen, Listen, Hervorhebungen und Code bleiben erhalten; im Klartext fehlen die Markdown-Steuerzeichen. Literale Zeichen und Einrückungen innerhalb von Code werden nicht als Markdown entfernt.

Externe Bilder werden nicht automatisch geladen, sondern sichtbar durch ihren Alternativtext ersetzt. Aktive HTML-Inhalte werden entfernt. Beide Fallbacks werden in der Statuszeile kenntlich gemacht und ohne Textinhalte protokolliert. Links öffnen sich erst nach einem bewussten Klick im Standardbrowser; das WebView selbst navigiert nicht zu fremden Seiten.

HTTP, DPAPI, globale Hotkeys und Zwischenablage bleiben im nativen Host. Die Web-Dialoge erhalten nur die für ihre jeweilige Aufgabe erforderlichen Daten. API-Schlüssel werden ausschließlich im eigenen Einstellungsdialog angezeigt; Geräte-Tokens werden nie an die Web-Oberfläche übergeben. Nachrichten erfordern die erwartete Seitenadresse und Vertragsversion. Die gemeinsame WebView-Shell verweigert fremde Navigationen, Downloads, Host-Objekte, Frames und Berechtigungsanfragen. Es wird kein fremder HWND als modaler Dialog-Owner verwendet.

## Lokale Daten und Verteilung

Konfiguration (`config.json`), automatischer Einladungsversuchs-Fingerprint, WebView2-Profil und Diagnoseprotokolle (`diagnostics.log`, `diagnostics.log.previous`) liegen ausschließlich unter `%LOCALAPPDATA%\AskInky`; keine Migration. Die DPAPI-Entropie wurde gegenüber den bisherigen Versionen nicht verändert. Alte `UseRightClick`-/`ConfirmSend`-Einträge werden ignoriert. Die Diagnose enthält Vorgang, Ausnahmetyp und HResult, keine Fragen, Antworten, URLs oder Schlüssel. Die optionalen detaillierten KI-Daten bleiben nur im Prozessspeicher und werden beim Beenden verworfen.

Die Single-EXE-Konfiguration bleibt erhalten; siehe `PUBLISH.md`. Eine separate Microsoft Edge WebView2 Runtime ist weiterhin erforderlich. Bei fehlender oder ausgefallener WebView2-Oberfläche erscheint eine verständliche native Fehlermeldung statt einer stillen Rückkehr zum alten Design.

Prüfumfang und noch nötige Windows-Abnahme: siehe `QA.md`. Externe Prüfberichte und Diffs sind nicht Bestandteil des Produktquellbaums.

## VDI / Remote Desktop – Hotkey troubleshooting (V091026)

On startup, Ask Inky registers the configured global shortcut. If registration fails, it logs the Win32 error and retries every 15 seconds while the app is running. A session unlock or remote reconnect also triggers a UI-thread re-registration. If the host rejects `MOD_NOREPEAT` (error 87), Ask Inky tries the same combination without this flag. A genuine hotkey conflict is **not** bypassed.

If the shortcut does nothing, click the Ask Inky tray icon and choose **Open** (localized). Check `%LOCALAPPDATA%\AskInky\diagnostics.log`:

- `hotkey.registered.*` = Windows accepted registration.
- `hotkey.registration.failed.win32.1409` = another app has registered the shortcut; choose another in Settings.
- `hotkey.message.received` = Windows delivered the shortcut to Ask Inky.
- If registration succeeded but no `hotkey.message.received` appears after a keypress, investigate VDI client's keyboard redirection, keyboard layout/AltGr, or shortcuts intercepted before reaching the guest.
- `assistant.open.failed` / `assistant.foreground.topmost-fallback` = inspect foreground/window activation or WebView2 behavior.

The tray command is the reliable fallback when the VDI client captures the shortcut. The hotkey remains configurable; previously saved choices are not overwritten. Clipboard-based text selection still requires a foreground application that accepts Ctrl+C.

## macOS

The supplied Windows EXE does not run natively on macOS. A Mac edition would need a separate native host for global shortcuts, Accessibility-based selection, Keychain storage, clipboard, WKWebView, tray/menu bar, and macOS signing/notarization. The existing web UI and Red Ink Helper API can be reused. A VDI session accessed **from a Mac** can still run the Windows edition **inside the Windows VDI** if keyboard shortcuts are forwarded.

Screenshot OCR: configurable second hotkey (default Ctrl+Alt+O), drag to select, Escape to cancel, uses existing Office Web Helper /files/upload; no Helper modifications.


#### Registration diagnostics

Registration failures in `%LOCALAPPDATA%\AskInky\diagnostics.log` now include `ResourceKey` and numeric `HttpStatus` if the exception is `HelperFailure`. `HttpStatus=0` means no HTTP status was available (for example a response-format failure). The diagnostic intentionally excludes the link, API key, server response and other credentials.

## 2026-10-12 — Konfigurationslinks und einheitliches Datenverzeichnis (aktueller Vertrag)

- Konfigurationslink `https://host/?api_key=...&uid=...&lang=de`: nur lokales Übernehmen von Helper-Host, API-Key, optionalem Lizenzbenutzer und unterstützter UI-Sprache; **kein Netzwerk-Registrierungs-GET**.
- Auth-Einladungslink mit nicht leerem Query-Parameter `auth`: einmaliger, unveränderter HTTP-GET mit JSON-Antwort. Auth-Tokens nicht loggen. Derselbe Pfad wird für Registry `HKCU\Software\Red Ink\HelperLink` und manuelle Eingabe genutzt. Registry hat Vorrang außer bei aktivierter Opt-out-Checkbox.
- Fehlt eine Verbindung und ein gültiger Registry-Link, Settings mit lokalisierter Anleitung und fokussierter Link-Eingabe.
- Sämtliche Daten (config.json, Geräte-Link-Marker, WebView2-Profil, diagnostics.log und previous) ab jetzt unter `%LOCALAPPDATA%\AskInky`; keine Migration, keine Speicherung im früheren Roaming-/RedInk-Unterordner. DPAPI-Entropie bleibt erhalten.
- Regression: Konfigurationslink darf keinen GET auslösen und darf nicht als Geräteanmeldung erscheinen; Einladung darf nicht doppelt abgerufen werden. Windows-/VSTO-Kompilierung und reale Registry-/WebView2-Tests stehen aus.

## 2026-10-12 – Antwortstil, Verarbeitung, UI-Stall-Diagnose

- Neuer Systemprompt: kurz, unmittelbar, ohne Floskeln; bestehende individuell gesetzte Systemprompts bleiben erhalten. Der exakt frühere Standardprompt erhält denselben neuen Standard bei Anfragen.
- Selektions-Ctrl+C wird mit `SendKeys.Send` statt UI-blockierendem `SendWait` ausgelöst; die vorhandene begrenzte Zwischenablage-Abfrage und Wiederherstellung bleiben erhalten.
- UI-Heartbeat (2 s) und Worker-Monitor (5 s) melden `ui.message-pump.unresponsive.over-15s` einmal je beobachtetem Stillstand; kein Nachweis der alleinigen Ursache, keine Nutzdaten im Log.
- Busy-Spinner behält die CSS-Animation und bekommt eine Compositor-Transform-Optimierung. Windows-/RDP-WebView2-Abnahme erforderlich.
