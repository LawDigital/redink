# Ask Inky – Aufgaben und Änderungen

## 2026-10-12 – Localhost ohne verpflichtenden API-Key
- [x] Prüfung: `CreateRequest` übermittelt ohne hinterlegten API-Key keinen `Authorization`-Header.
- [x] Bei HTTP 401/403 vom Loopback-Helper neutrale, lokalisierte Server-Ablehnung anzeigen, statt einen fehlenden API-Key oder eine Geräteregistrierung zu behaupten; externe Helper unverändert.
- [x] Gleiche Behandlung für den Verbindungstest (`/config`) und reguläre HTTP-Anfragen.
- [x] Ask-Inky-Dokumentation liegt innerhalb von `AskInky/` (`CONTRIBUTING.md`, `Tasks.md`).
- [ ] Native VS-2022-/WebView2-Laufzeitprüfung: localhost ohne Credentials mit erfolgreich konfiguriertem Server; 401/403 lokal; 401/403 extern; HTTP-Header prüfen.

Bestehende Aufgaben des ursprünglichen Repositories liegen in dessen unveränderter Root-`Tasks.md`; diese Ask-Inky-Aufgabenliste wird ab jetzt innerhalb des Projektverzeichnisses geführt.
