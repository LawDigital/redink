# CONTRIBUTING – Ask Inky

## Release and versioning

- Source of truth: `Program.vb`, `Program.Version` (`VDDMMYY`); currently **V121026** (12 October 2026). Update the date-derived version on **every new dated release/delivery**; propagate it to release notes and any displayed program version as appropriate. Do not modify stored user settings merely to update the version.
- Create changed-only ZIP preserving `AskInky/` relative paths and maintain a cumulative latest source baseline between iterations. Never replace earlier fixes with older versions. Track changed-only list, cumulative diff, static tests and unverified Windows tests. Only claim a Visual Studio / Windows build when actually run.
- Preserve the currently supported standalone **win-x64, self-contained, single-file** publish settings, and embedded Office Web assets/licensing. WebView2 Runtime is still a machine prerequisite.

## Functional invariants

- Default shortcut: **Ctrl + Alt + K** (Control+Alt = 3, virtual key K = 75). User-configured shortcut in the persisted JSON **takes precedence** and is not silently overwritten. Settings reset must use the same default as `ClientConfig` and `SettingsForm` defaults. Global hotkey registration must keep the old hotkey if a new one fails. No mouse hook or right-click gesture.
- Original selected text stays available for follow-up questions. Maintain conversation pair navigation, deletion, rich HTML preview and formatted HTML + Unicode clipboard copying. Ask only when user presses Send; never add a pre-send confirmation.
- The native titlebar owns the `Ask Inky` title. No redundant logo/title inside selected text. Tray menu and tooltips stay localized in DE/EN/FR/IT. Settings include server API key, configurable model list, licenses and Red Ink copyright/attribution.
- Window starts at a usable size with a lower bound that keeps the bottom action row visible. HTML answer area must shrink/scroll when the **user manually resizes**. Automatic growth may respond to new content until the user **manually shrinks** the current session; after that, do not override their selected height. A new selected text/session may re-enable content-driven growth. Do not trigger a WebView2 size feedback loop on every browser `resize` event.
- `http://localhost` is supported (standard HTTP port semantics), no client-enforced device registration; server-side authorization remains authoritative. Protect existing DPAPI entropy and encrypted user secrets.
- Third-party attribution must remain bundled and accessible from Settings even when published as a single EXE. Do not alter the original Red Ink license text silently. Preserve Office Web's original CSS and existing licenses.

## Coding and QA

- Follow Red Ink VB headers (copyright, file, purpose, architecture) and entry-point attribution header. Use fully qualified `System.Windows.Forms`/`System.Drawing` references; use `System.Exception`, `System.Threading.Tasks.Task`, and `IsNot` syntax.
- Verify model retrieval/parsing end-to-end (including JSON property casing, authentication/permissions and explicit error display); preserve manual fallback for unavailable model catalog.
- Regression-check persisted custom hotkeys, default reset, sending, navigation, clipboard, settings, licenses, WebView2 assets, minimum size, manual shrink and new-session automatic sizing. Static checks are not substitutes for Windows/VS2022 runtime checks.
- Deliver only new or edited source files (same directory structure) as ZIP, plus a brief QA report; store cumulative baseline internally for the next change.

## VDI and remote desktop resilience (9 October 2026, V091026)

- The current hotkey uses Windows `RegisterHotKey` on a message-only window. Never introduce a system-wide keyboard hook or keyboard polling to bypass VDI policy. An RDP/Citrix/VMware client can consume a shortcut **before Windows receives it**; no client-side registration can guarantee delivery in that case. Users must be able to change the shortcut and use the localized tray **Open** command as a fallback.
- Keep registration errors in `%LOCALAPPDATA%\AskInky\diagnostics.log` (Win32 codes), success markers and `hotkey.message.received` without logging user text, credentials or window titles. Compare the log after an attempted shortcut to distinguish registration failures from shortcuts not delivered to the guest. The recovery timer retries only when unregistered; session reconnect/unlock requests a rebind on the UI thread. Do not run synthetic keystrokes during registration recovery.
- On VDI reconnect, allow registration to recover without restarting. If `MOD_NOREPEAT` is unsupported (`ERROR_INVALID_PARAMETER=87`), retry the same shortcut without that flag. Do **not** fall back on a different shortcut automatically or steal a conflicting shortcut (`1409`).
- Foreground activation after async clipboard capture can be rejected by Windows. Use the native activation path and only a transient topmost fallback, never a permanently topmost dialog.
- VDI QA must include: initial registration, shortcut not forwarded by client, conflict (1409), remote disconnect/reconnect, session lock/unlock, hotkey change and restart, native tray fallback, repeated hotkey while dialog is open, clipboard capture and no unintended input in source application.
- **Versioning:** same-date changes on 9 October 2026 retain `V091026`; the next dated delivery must update `Program.Version` and related documentation. Preserve Authenticode signing after Publish and verify signature before distribution.

## macOS feasibility

- The present WinForms/.NET 8 Windows application **cannot** be published as a native macOS app. WinForms, Win32 `RegisterHotKey`/foreground APIs, WebView2, DPAPI, Windows clipboard formats, tray and code-signing targets are platform-specific.
- A macOS port should share the API/domain contracts and HTML/CSS/JS where possible, but use a macOS host (e.g. Swift/AppKit + WKWebView or a cross-platform desktop host), native global shortcut registration, selection capture through macOS Accessibility/clipboard with explicit permissions, Keychain credentials, and notarized macOS signing. Never promise universal selection access on either platform.
- Keep backend contracts and licensing independent of OS and preserve the Office Web visual language; test model list, authentication, markdown, copy and history on both platforms.

## Screenshot OCR (V091026)

- On upgrade migrate **only the precise old default pair** (Ctrl+Shift+K and Ctrl+Alt+Shift+K) to Ctrl+Alt+K / Ctrl+Alt+O. Never rewrite any other persisted shortcut combinations. If the user deliberately saved a single legacy default, require them to use Settings → Reset instead.

- No Helper changes: reuse Office Web `POST /files/upload` multipart fields `file` (PNG) and `allow_ocr=true`; server response `content` is OCR text. Existing device / API credentials and server licensing remain authoritative.
- Secondary hotkey defaults to Ctrl+Alt+O, independently editable. It must never be identical to the primary shortcut, and both registrations must remain recoverable after VDI reconnect. Treat helper `/config` `object_capable` as an advisory UI hint only, not a promise of OCR success for every selected model.
- After screenshot selection, display a non-blocking WinForms progress ring with the original Red Ink logo until the OCR request succeeds or fails; always close it in Finally. Capture is strictly user initiated; Escape/right-click cancels, do not persist screenshots or log image contents, and bound screenshot size. In VDI, screen capture may be black/blocked by remote desktop policy; report errors rather than bypassing policy.
- Once OCR succeeds, place the text in the selected-text field and start a new conversation if its content differs. Preserve the ordinary selected-text shortcut and all prior dialogs, licenses, and single-file signing behavior.

## Settings window placement and OCR shortcut layout (V091026)

- When invoked from the visible assistant, Settings is owned by the assistant and positioned `CenterParent`; when opened from the tray while the assistant is hidden, use `CenterScreen` explicitly. Do not assign third-party application HWNDs as owners. Bring Settings to foreground on open; only a *transient* TopMost toggle is permitted as a VDI fallback.
- Keep the regular shortcut help text and `Screenshot OCR shortcut` label visually separate with a dedicated CSS margin. Never override the original `office-web.css` for this.
- Regression-check the Settings window from both the tray and the assistant, repeated invocation, VDI focus, 100%/150% DPI, the OCR shortcut spacing, and preservation of the settings footer/save actions.

## Persistent window placement / VDI monitor changes (V091026)

- `WindowPlacement.vb` owns all native window geometry; settings must be positioned explicitly above the **visible** assistant rather than assuming `CenterParent` works with modeless `Show(owner)`. Tray Settings must center on the current pointer monitor. Keep owner relationship for z-order.
- Persist only the assistant's **user-chosen** position (`AssistantLeft`, `AssistantTop`) in existing `config.json` (optional nullable integers). Do not reset settings, credentials, or hotkeys. No saved position means center on the active monitor.
- At creation/reopen, always clamp restored coordinates within a real attached monitor's **WorkingArea**, including negative coordinates, VDI monitor removal, changed resolution and changed taskbar size. The window must never be off-screen; foreground activation must remain best-effort without permanent TopMost.
- Do not persist automatic window centering/reflow every frame or use a save timer. Save after user move/resize and when the assistant is hidden by the user. Reopen with existing answer/history intact. Regression check Windows at 100% / 150% scaling, primary/secondary monitors, remote reconnect, settings from tray and from visible assistant.

## OCR wait indicator (V091026)

- `OcrProgressForm.vb` uses an **opaque rectangular white background** (`BackColor = White` and white paint clear); do not enable `TransparencyKey` or round the form/background. This intentionally supersedes the earlier transparency experiment after user feedback.
- Keep the animated ring and visible-alpha-cropped original Red Ink logo centered in the same `ClientRectangle` coordinate system, including at 100% and 150% DPI/VDI.
- Keep the indicator non-blocking and dispose it on both successful and failed OCR requests. Regression-check no transparent/cutout pixels or off-center mark after reopening; do not change OCR transport or UI state transitions.

- OCR progress popup uses a square white background with centered Red Ink logo and animated ring.

## Initial configuration, registry bootstrap and answer selections (12 October 2026, V121026)

- **Current policy (supersedes the first-run-only import rule of V121026-01):** At each application start, unless `IgnoreRegistryHelperLink=True` or config loading failed, inspect only `HKCU\Software\Red Ink`; use the explicit string `HelperLink` ahead of compatible legacy registration/invitation values. A new valid link can take precedence **even over an already persisted Helper host/API/registered device**. Do not expand registry values or query unknown host URLs; a named `HelperLink` may be `https://host/?api_key=...` with an empty registration path. A missing/invalid/already attempted registry link must never clear existing credentials or cause a server request.
- Use the existing one-time Helper registration contract from Office Web with the exact original URL and JSON GET; `api_key`, `uid`, `lang` are parsed only by the shared registration path. Save the successful link SHA-256 fingerprint in `ClientConfig.AppliedRegistryLinkFingerprint` (no secret URL), in addition to recording an attempt fingerprint **before** GET. If the fingerprint matches, reuse previously stored credentials without replaying the one-use URL. A changed link may replace previously stored credentials **only after confirmed successful server registration and config persistence**; on failure preserve the prior configuration. Never claim a previous link succeeded just because its pre-request attempt marker exists.
- The advanced Settings option `IgnoreRegistryHelperLink` defaults to `False` and is user-persisted. When `True`, registry import is skipped on the next launch, and connection/registration fields can be edited. With the successfully applied matching registry link present and opt-out unchecked, disable those fields and enforce their authoritative values in the native settings handler too (do not depend only on JS disabling). When manual opt-out replaces credentials, remove the obsolete applied-fingerprint marker; a consumed link cannot be replayed if the user enables registry import again without a **new** link. Manual registration/unregistration must not mislabel its credentials as belonging to an older registry link.
- The unconfigured first launch still opens Settings with the registration field focused and localized guidance when no usable automatic import succeeded. Explicit tray/shortcut actions before Application.Idle must serialize behind the same registration attempt and preserve the requested action (including selection vs. tray open). If successful registration cannot be saved, retain protected credentials and open non-dismissible Settings for a save retry, including when prior credentials existed. No registry writes, server retries, secret-bearing diagnostics, or change to DPAPI entropy. Preserve hotkeys, placement, locale except an explicitly supplied `lang`, and other preferences.
- Keep diagnostics in `%LOCALAPPDATA%\AskInky\diagnostics.log` (LocalAppData, without legacy migration) without URLs, API keys, user text or token values. Windows registry, WebView2 and helper integration require separate real-platform acceptance checks; source tests alone are insufficient.
- Tray icon and notification titles identify **Ask Inky**; the branded Red Ink asset and copyright remain unchanged.
- Whole-answer HTML/Unicode clipboard copying remains unchanged. Selection copy is separate, visible/enabled only for a nonempty DOM selection wholly within the **current rendered answer**. Reuse the existing sanitizer and CF_HTML/Unicode writer; reject missing/stale/oversized payloads, preserve selection on icon press, hide after answer replacement, busy state or Settings. Test code, lists, tables, keyboard/mouse selection and four localized tooltips.
- Repeated-shortcut crashes require logs/event traces for a proven causal attribution. A source-level native exception risk exists during WebView2 response creation if the browser process ends; resource event errors must be contained and unhanded streams disposed. A code-level risk is not yet a demonstrated user crash. Guard the entire shortcut UI activation (including the unconfigured Settings path) against disposal races and contain managed dispatch failures in the Win32 callback; retain detailed no-secret diagnostics. Test repeated Ctrl+Alt+K/O, VDI reconnect, clipboard contention, WebView2 process loss on real Windows.
- Release regressions: with/without registry entry, valid/invalid/consumed link, inaccessible Helper, explicit saved localhost/remote host, API-only config, damaged config file and registry denial. Never automatically re-run registration on routine restarts. VS2022/WinForms runtime acceptance remains an unexecuted separate gate.

### HelperFailure diagnostic contract (2026-10-12)

- `ClientLog.Write(operation, failure)` logs `ResourceKey` (only a bounded identifier matching `[a-zA-Z0-9_.-]`) and numeric `HttpStatus` for `HelperFailure` failures; `HttpStatus=0` indicates that there was no HTTP response status available, **not** HTTP 0. Do not serialize `Exception.Message`, response content, full `HelperLink`, user IDs, API keys or device credentials.
- The destination remains `%LOCALAPPDATA%\AskInky\diagnostics.log`; no migration from legacy paths. Preserve existing rotation, stack-frame-only diagnostics and handling of all other exception types.
- Validate a 403 failure and a non-HTTP response failure, plus generic-exception compatibility, before shipping. Actual registration endpoint behavior must be validated on Windows; never consume a one-time link just for a diagnostic test.

## Default assistant window width (2026-10-12)

- AssistantForm default client width is 600 pixels (previously 540) in both the WebDialog constructor and the post-working-area client-size assignment. Keep both widths equal; allow manual resize and retain the existing 470-pixel minimum width.
- Preserve saved assistant *position*, active-monitor centering and off-screen clamping. The application does not persist an independent user-selected width, so do not add/reset a width setting.
- Check copy-whole-answer and selection-copy icon visibility together with previous/next and other action buttons at 100%/150% DPI and smaller screens. Real Windows/WebView2 verification remains pending.

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

## 2026-10-12 – Busy-Indikator bei reduzierten Animationen

- Die CSS-Regel `prefers-reduced-motion: reduce` in `desktop.css` darf den funktionsrelevanten Kreis `body.is-busy .desktop-status::before` nicht abschalten. Die Rotation wird nur für diesen Indikator wieder aktiviert; andere Animationen bleiben deaktiviert.
- Browser-/Windows-Abnahme mit eingeschalteter und ausgeschalteter Einstellung „Animationen reduzieren“ prüfen.
