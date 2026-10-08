# Ask Inky – eine EXE veröffentlichen

## Voraussetzungen

Windows x64 und Visual Studio 2022 mit .NET-Desktopentwicklung/.NET 8 SDK. Auf dem Ziel-PC ist die Microsoft Edge WebView2 Runtime erforderlich. Die Self-contained-Veröffentlichung enthält die .NET-Laufzeit, nicht die WebView2-Browser-Runtime.

## Veröffentlichung

Im Verzeichnis mit `AskInky.vbproj`:

```powershell
dotnet publish .\AskInky.vbproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:PublishTrimmed=false -p:DebugType=embedded
```

Erwarteter Zielpfad: `bin\Release\net8.0-windows\win-x64\publish\AskInky.exe`.

Alternativ **AskInky → Publish → Folder**, Release, win-x64, Self-contained, Produce single file, kein Trimming. Das vorhandene Projekt in der Solution muss dafür nicht mit anderen Projekten zusammengeführt werden.

Das Icon und alle HTML-/CSS-/JavaScript-/Sprachressourcen sind eingebettet. Es ist kein separater Web-Ordner neben der veröffentlichten EXE erforderlich. Native Bibliotheken können beim Start in ein Benutzerverzeichnis extrahiert werden. Einstellungen und WebView2-Profildaten liegen separat unter `%LOCALAPPDATA%\RedInk\AskInky`; das Programm benötigt dort Schreibrechte, nicht im EXE-Verzeichnis.

Vor Verteilung den tatsächlichen Publish-Ordner prüfen und die EXE aus einem anderen Verzeichnis auf einem zweiten Windows-PC testen. Kein Adminstart als Standard; erhöhte Zielanwendungen sind ein eigener Kompatibilitätsfall. Projekt- und Drittsoftware-Lizenzbedingungen bleiben zu beachten.

Ein Windows-Release-Publish und ein Start der fertigen EXE wurden in der Erstellungsumgebung nicht ausgeführt. Die vorliegende Lieferung ist Quellcode, keine vorab kompilierte EXE.
