# Native Windows UI regression probe

This source-only project accompanies the Red Ink add-in threading corrections.
It targets .NET Framework 4.8. Its selected production code was compiled against
Microsoft's reference assemblies with **zero errors and zero warnings** in the
delivery environment. **It was not executed on Windows or inside Office.**

## Build and run

Use a Windows machine with .NET Framework 4.8 and a .NET SDK that can build the
SDK-style VB project, such as the SDK installed with Visual Studio 2022. Package
restore requires access to NuGet. Run in an interactive desktop session, from
this directory:

```powershell
dotnet build .\HangProbe.vbproj --configuration Release
if ($LASTEXITCODE -ne 0) { throw 'Probe build failed.' }
.\bin\Release\net48\HangProbe.exe
if ($LASTEXITCODE -ne 0) { throw 'Native UI probe failed.' }
```

The probe opens and closes real WinForms windows. Do not run it as an unattended
server service. It uses temporary files for the RTF cases and is separate from
the server's Python tests and from the production VSTO solution.

## What it checks

The project includes complete selected UI classes and extracted production
members for STA/RTF conversion, the actual host dispatcher and awaitable,
modeless registry, legacy Office-reader routing, LogWindow, both splashes,
progress/AskUser UI and OCR status lifetime. It exercises native RichEdit
Unicode/paragraph/tab/table behavior, invalid RTF, MTA-to-STA conversion, caller
context preservation, dispatch and shutdown cancellation, window registration,
startup/close races and owned resource disposal.

`source-hashes.json` identifies the selected snapshots. `ActualHelpers.vb` and
`ActualOfficeReaderDispatch.vb` contain extracted production sections;
`Stubs.vb` explicitly replaces unrelated resources, native Office owner lookups
and actual Office file-reading bodies. Those reader doubles check routing, not
Word/Excel/PowerPoint automation. The normal Office applications and their COM
objects are not exercised by this standalone probe.

After this probe, build the full add-in solution and execute the real Office
acceptance cases in `docs/THREADING_REVIEW.md` and
`docs/validation/STA_LICENSING_REVIEW.md` at the application root. The server
delivery's `docs/QA.md` separates completed checks from outstanding Windows
runtime checks. No compiled probe, SDK, Office component, certificate or test
database is distributed with this project.
