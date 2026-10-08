# Abnahmehinweise für Ask Inky

## Entwicklungsprüfungen dieser Fassung

Die eigenständige Browser-Prüfung umfasst DOM-Rendering, HTML-/Klartext-Kopierinhalt, Unicode und Code-Einrückungen, Enter/Shift+Enter, IME-Eingabe, Busy-/Fehlerzustände, Sprach- und Themenwechsel, API-Key-Nachrichten, Shortcut-Erfassung sowie verschiedene Fenstergrößen und Pixeldichten. JavaScript-Syntax, eingebettete Ressourcen, Übersetzungsschlüssel, unveränderte Office-Web-CSS und die Quellverträge zum Helper werden zusätzlich statisch geprüft.

Diese Prüfungen laufen mit einer nachgebildeten WebView2-Nachrichtenbrücke außerhalb des Produktprojekts. Wegen einer Browser-Umgebungsrichtlinie werden die tatsächlichen UI-Dateien in ein lokales Testdokument injiziert, nicht über die native Ressourcenroute geladen. Dies ist kein WebView2-/Windows-Ende-zu-Ende-Test und kein Ausführen des VB-Codes. Ein .NET SDK für einen Build war nicht verfügbar. Externe Ergebnisberichte und Diffs werden getrennt geliefert.

## Noch auf Windows zu prüfen

1. Projekt nach Dateiaustausch neu laden, NuGet wiederherstellen, **AskInky → Build**. Im Debug-Modus starten. Beide Dialoge müssen ihre eingebetteten Ressourcen korrekt laden. Native Titel: „Ask Inky“ und „Ask Inky – Einstellungen“ bzw. Übersetzung. Kein weiterer Anwendungstitel oder Logo im Inhalt.
2. In Notepad, Word, Outlook und Browsern Text markieren und Ctrl+Shift+K testen. Modifikatortasten loslassen; Auswahl muss vor dem Öffnen erfasst werden. Ein Fokuswechsel während der Erfassung darf keine Eingabe in ein anderes Fenster auslösen. Zwischenablage vor/nach Erfassung mit Unicode, HTML, Bild und Dateiliste prüfen; Wiederherstellung bleibt bestmöglich, nicht für alle fremden Clipboard-Provider garantiert.
3. Hotkey ändern, Neustart, Konflikt mit einer anderen globalen Registrierung und Schreibfehler beim Konfigurationsspeichern prüfen. Vorherige Kombination muss bei einem Fehlschlag wirksam bleiben. Es darf keine globale Rechtsklick-Geste mehr geben. Ein Tray-Klick bietet drei übersetzte Einträge.
4. API-Schlüssel eingeben, speichern, neu starten, Verbindung testen. HTTP-Loopback ohne Gerätesperre prüfen; bei externem HTTPS-Helper API-Key, Geräte- und Lizenzanforderungen nach dessen Konfiguration prüfen. Bestehende DPAPI-Zugangsdaten müssen weiterhin lesbar sein. Beim Ändern des Servers keine alten Geheimnisse weitergeben.
5. Einen frischen Registrierungslink genau einmal bewusst aufrufen. Das bloße Einfügen löst keinen Request aus. Erfolgreiche Zugangsdaten müssen sofort persistiert werden. Lokales Aufheben der Registrierung und danach API-Key-Betrieb testen. Fehler bei Registrierung/HTTP/JSON/Speichern dürfen nicht zum automatischen Wiederverwenden des Einmal-Links führen.
6. Erfolgreiche und fehlgeschlagene `/llm/chat`-Anfragen testen: keine Versand-Rückfrage, keine Doppelanfragen bei mehrfachem Enter, gesperrte Aktionen während des Wartens, Freigabe nach Ende und Erhalt der vorigen Antwort bei einem Fehler. Alle vier Sprachen und Hell-/Dunkelmodus prüfen.
7. Antwort mit Überschriften, Tabelle, Liste, Code, Umlauten, Akzenten und Emoji kopieren. In Word/Outlook muss HTML-Formatierung, in Notepad Klartext ohne Markdown-Steuerzeichen ankommen. Windows-Clipboard-Sperre und Wiederholung prüfen. Einfügen/Eintippen dürfen keine eigenen Funktionen mehr sein.
8. Windows-DPI 100/150/200 %, Monitorwechsel, minimale Fenstergröße, lange Fehlertexte und Tastaturbedienung testen. In dieser Fassung werden SVG-Symbole statt Raster-/Emoji-Icons verwendet. Fehlende bzw. ausgefallene WebView2 Runtime muss den nativen Fehlerhinweis zeigen.
9. Self-contained Single-EXE tatsächlich veröffentlichen und aus einem anderen Ordner sowie auf einem zweiten Windows-PC starten. Keine externen UI-Dateien dürfen erforderlich sein. Die WebView2 Runtime bleibt eine Voraussetzung.

Bestehende Gen2-, Helper- und Office-Web-Dateien wurden für diese Änderung nicht verändert. Deren Windows-/VSTO-Workflows wurden nicht ausgeführt; Quellgleichheit ist kein Laufzeitnachweis.
