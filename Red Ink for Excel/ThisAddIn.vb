' Part of "Red Ink for Excel"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' 9.1.2026
'
' The compiled version of Red Ink also ...
'
' Includes DiffPlex in unchanged form; Copyright (c) 2023 Matthew Manela; licensed under the Apache-2.0 license (http://www.apache.org/licenses/LICENSE-2.0) at GitHub (https://github.com/mmanela/diffplex).
' Includes Newtonsoft.Json in unchanged form; Copyright (c) 2023 James Newton-King; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://www.newtonsoft.com/json
' Includes HtmlAgilityPack in unchanged form; Copyright (c) 2024 ZZZ Projects, Simon Mourrier,Jeff Klawiter,Stephan Grell; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://html-agility-pack.net/
' Includes Bouncycastle.Cryptography in unchanged form; Copyright (c) 2024 Legion of the Bouncy Castle Inc.; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://www.bouncycastle.org/download/bouncy-castle-c/
' Includes PdfPig in unchanged form; Copyright (c) 2024 UglyToad, EliotJones PdfPig, BobLd; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/UglyToad/PdfPig
' Includes MarkDig in unchanged form; Copyright (c) 2024 Alexandre Mutel; licensed under the BSD 2 Clause (Simplified) license (https://licenses.nuget.org/BSD-2-Clause) at https://github.com/xoofx/markdig
' Includes NAudio and components in unchanged form; Copyright (c) 2020 Mark Heath; licensed under a proprietary open source license (https://www.nuget.org/packages/NAudio/2.2.1/license) at https://github.com/naudio/NAudio
' Includes Vosk in unchanged form; Copyright (c) 2022 Alpha Cephei Inc.; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://alphacephei.com/vosk/
' Includes Whisper.net in unchanged form; Copyright (c) 2024 Sandro Hanea; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/sandrohanea/whisper.net
' Includes Grpc.core/Grpc.net in unchanged form; Copyright (c) 2023/2025 The gRPC Authors; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/grpc/grpc
' Includes Google Speech V1 library and related API libraries in unchanged form; Copyright (c) 2024 Google LLC; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/googleapis/google-cloud-dotnet
' Includes Google Protobuf in unchanged form; Copyright (c) 2025 Google Inc.; licensed under the BSD-3-Clause license (https://licenses.nuget.org/BSD-3-Clause) at https://github.com/protocolbuffers/protobuf
' Includes Google.Api in unchanged form; Copyright (c) 2025 Google LLC; licensed under the BSD-3-Clause license (https://licenses.nuget.org/BSD-3-Clause) at https://github.com/googleapis/gax-dotnet
' Includes Google.Apis in unchanged form; Copyright (c) 2025 Google LLC; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/googleapis/google-api-dotnet-client
' Includes Google.Longrunning in unchanged form; Copyright (c) 2025 Google LLC; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/googleapis/google-cloud-dotnet
' Includes MarkdownToRTF in modified form; Copyright (c) 2025 Gustavo Hennig; original licensed under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/GustavoHennig/MarkdownToRtf
' Includes Nito.AsyncEx in unchanged form; Copyright (c) 2021 Stephen Cleary; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/StephenCleary/AsyncEx
' Includes NetOffice libraries in unchanged form; Copyright (c) 2020 Sebastian Lange, Erika LeBlanc; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/netoffice/NetOffice-NuGet
' Includes NAudio.Lame in unchanged form; Copyright (c) 2019 Corey Murtagh; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/Corey-M/NAudio.Lame
' Includes PdfiumViewer in unchanged form; Copyright (c) 2017 Pieter van Ginkel; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/pvginkel/PdfiumViewer
' Includes PDFsharp in unchanged form; Copyright (c) 2025 PDFSharp Team; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://docs.pdfsharp.net/
' Includes System.Interactive.Async in unchanged form; Copyright (c) 2025 by .NET Foundation and Contributors; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/dotnet/reactive
' Includes also various Microsoft distributables and libraries copyrighted by Microsoft Corporation and available, among others, under the Microsoft EULA, the Visual Studio Community 2022 License, the Microsoft.Web.WebView2 License (for Microsoft.Web.WebView2, see license on https://www.nuget.org/packages/Microsoft.Web.WebView2/ and below) and the MIT License (including Microsoft.Bcl.*, Microsoft.Extensions.*, System.*, System.Security.*, System.CodeDom, DocumentFormat.OpenXml.*, Microsoft.ml.*, CommunityToolkit.HighPerformance licensed under MIT License) (https://licenses.nuget.org/MIT); Copyright (c) 2016- Microsoft Corp.
'
' Licenses of Red Ink and of third-party components and further legal terms/notices are available in the installation folder and via https://redink.ai.
'
' Documentation for developers: See at the end of this file, throughout the code and the manual (https://redink.ai).

Option Strict On
Option Explicit On

Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary

Module ModuleGetAsyncKeyState

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Function GetAsyncKeyState(ByVal vKey As Integer) As Short
    End Function

End Module

Partial Public Class ThisAddIn

    ' Hardcoded config values

    Public Shared Version As String = "V.090126" & SharedMethods.VersionQualifier

    Public Const AN As String = "Red Ink"
    Public Const AN2 As String = "redink"
    Public Const AN5 As String = "RI"

    Private Const ShortenPercent As Integer = 20
    Private Const TextPrefix As String = "TextOnly:"
    Private Const TextPrefix2 As String = "Text:"
    Private Const BatchPrefix As String = "Batch:"
    Private Const CellByCellPrefix As String = "CellByCell:"
    Private Const CellByCellPrefix2 As String = "CBC:"
    Private Const PurePrefix As String = "Pure:"
    Private Const PanePrefix As String = "Pane:"
    Private Const BubblesPrefix As String = "Bubbles:"
    Private Const ExtTrigger As String = "{doc}"
    Private Const ExtDirTrigger As String = "{dir}"
    Private Const ExtTriggerFixed As String = "{[path]}"
    Private Const ExtWSTrigger As String = "(addws)"
    Private Const ObjectTrigger As String = "(file)"
    Private Const ObjectTrigger2 As String = "(clip)"
    Private Const ColorTrigger As String = "(color)"
    Private Const RIMenu = AN
    Private Const MinHelperVersion = 1  ' Minimum version of the helper file that is required
    Public Const LargeWorksheetSize As Integer = 2500

    Public RemoveMenu As Boolean = False  ' Legacy: If true, the old context menu will be removed (Gen1)

    ' Publicly declared variables so that InterpolateAtRuntime can access them; case-sensitive

    Public TranslateLanguage As String
    Public OutputLanguage As String
    Public FileNameBody As String
    Public FileDate As String
    Public CurrentDate As String = "(Current Date: " & DateTime.Now.ToString("dd-MMM-yyyy", CultureInfo.GetCultureInfo("en-US")) & ")"
    Public ShortenLength As Double
    Public SummaryLength As Integer
    Public OtherPrompt As String = ""
    Public Separator As String = ""
    Public LineNumber As Integer = 0
    Public Context As String
    Public SysPrompt As String
    Public OldParty, NewParty As String
    Public SelectedText As String

    Public Shared DragDropFormLabel As String = ""
    Public Shared DragDropFormFilter As String = ""

    Public Shared allowedExtensions As System.Collections.Generic.HashSet(Of System.String) =
                        New System.Collections.Generic.HashSet(Of System.String)(
                            New System.String() {".txt", ".ini", ".csv", ".log", ".json", ".xml", ".html", ".htm", ".rtf", ".doc", ".docx", ".pdf"},
                            System.StringComparer.OrdinalIgnoreCase
                        )

    Public Structure CellState
        Public WorksheetName As String
        Public CellAddress As String
        Public OldValue As Object
        Public HadFormula As Boolean
        Public OldFormula As String
    End Structure

    Public Shared undoStates As New List(Of CellState)

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function FindWindow(
                                ByVal lpClassName As String,
                                ByVal lpWindowName As String
                            ) As IntPtr
    End Function

    Private Shared _uiContext As SynchronizationContext
    Private Shared _uiScheduler As TaskScheduler
    Private mainThreadControl As New System.Windows.Forms.Control()

    ' Lightweight UI switch helper (no exception risk).
    Private Shared Function EnsureUIThread() As Task
        ' If already on UI thread (or UI context not captured), do nothing.
        If _uiContext Is Nothing OrElse SynchronizationContext.Current Is _uiContext Then
            Return Task.CompletedTask
        End If

        Dim tcs As New TaskCompletionSource(Of Object)()

        ' Post back to the captured UI context.
        _uiContext.Post(
            Sub(state As Object)
                tcs.SetResult(Nothing)
            End Sub,
            Nothing)

        Return tcs.Task
    End Function

    Private chatForm As frmAIChat

    Private automationObject As BridgeSubs

    Protected Overrides Function RequestComAddInAutomationService() As Object
        If automationObject Is Nothing Then
            automationObject = New BridgeSubs()
        End If
        Return automationObject
    End Function

    Private Sub ThisAddIn_Startup() Handles Me.Startup

        ' Necessary for Update Handler to work correctly

        ' 1) Force the creation of the Control's handle on the Office UI thread
        Dim dummy = mainThreadControl.Handle

        ' 2) Capture synchronization context & scheduler exactly once.
        _uiContext = SynchronizationContext.Current
        If _uiContext Is Nothing Then
            ' Fallback: install a WindowsFormsSynchronizationContext explicitly.
            _uiContext = New WindowsFormsSynchronizationContext()
            SynchronizationContext.SetSynchronizationContext(_uiContext)
        End If
        _uiScheduler = TaskScheduler.FromCurrentSynchronizationContext()

        ' 3) Give that Control to the UpdateHandler so it can Invoke on it
        UpdateHandler.MainControl = mainThreadControl

        ' 4) Capture the host window’s HWND (Word / Excel / Outlook)
        Dim hwnd As IntPtr
        Dim progId = Me.Application.GetType().Name.ToLowerInvariant()
        If progId.Contains("word") OrElse progId.Contains("excel") Then
            hwnd = New IntPtr(CInt(Me.Application.Hwnd))
        Else
            hwnd = FindWindow("rctrl_renwnd32", Nothing)
        End If

        ' Other startup tasks

        SharedMethods.Initialize(Me.CustomTaskPanes)
        InitializeAddInFeatures()
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        RemoveOldContextMenu()
    End Sub

    Public Sub InitializeAddInFeatures()
        InitializeConfig(True, True)
        AddContextMenu()
        UpdateHandler.PeriodicCheckForUpdates(INI_UpdateCheckInterval, RDV, INI_UpdatePath, _context)
    End Sub

    ' Bridge to SharedLibrary

    Public Sub InitializeConfig(FirstTime As Boolean, Reload As Boolean)
        _context.InitialConfigFailed = False
        _context.RDV = "Excel (" & Version & ")"
        SharedMethods.InitializeConfig(_context, FirstTime, Reload)
    End Sub
    Private Function INIValuesMissing() As Boolean
        Return SharedMethods.INIValuesMissing(_context)
    End Function
    Public Shared Async Function PostCorrection(inputText As String, Optional ByVal UseSecondAPI As Boolean = False) As Task(Of String)
        Return Await SharedMethods.PostCorrection(_context, inputText, UseSecondAPI)
    End Function
    Public Shared Async Function LLM(ByVal promptSystem As String, ByVal promptUser As String, Optional ByVal Model As String = "", Optional ByVal Temperature As String = "", Optional ByVal Timeout As Long = 0, Optional ByVal UseSecondAPI As Boolean = False, Optional ByVal Hidesplash As Boolean = True, Optional ByVal AddUserPrompt As String = "", Optional ByVal FileObject As String = "") As Task(Of String)
        Dim Response = Await SharedMethods.LLM(_context, promptSystem, promptUser, Model, Temperature, Timeout, UseSecondAPI, Hidesplash, AddUserPrompt, FileObject)
        Await EnsureUIThread().ConfigureAwait(False)
        Return Response
    End Function
    Private Sub ShowSettingsWindow(Settings As Dictionary(Of String, String), SettingsTips As Dictionary(Of String, String))
        SharedMethods.ShowSettingsWindow(Settings, SettingsTips, _context)
    End Sub
    Private Function ShowPromptSelector(filePath As String, filepathlocal As String, enableMarkup As Boolean, enableBubbles As Boolean) As (String, Boolean, Boolean, Boolean)
        Return SharedMethods.ShowPromptSelector(filePath, filepathlocal, enableMarkup, enableBubbles, _context)
    End Function


End Class


' =========================================================================================
' Red Ink Excel Add-in: Maintenance & Security Review Overview
' =========================================================================================
' PURPOSE
'   Provides AI-assisted text, document, and data processing inside Excel. Central entry
'   point is ThisAddIn.vb which wires UI elements, configuration, update checks, helper
'   modules, and cross-process automation bridges.
'
' CORE EXECUTION FLOW
'   1. ThisAddIn_Startup initializes threading context, update handler, custom task panes,
'      menus/ribbon bindings, configuration (via SharedLibrary), and feature modules.
'   2. User actions (Ribbon / context menu / drag & drop / VBA helper calls / pane UI)
'      funnel into command handlers (see ThisAddIn.Commands.vb) which route to processing
'      helpers (LLM calls, file/text transforms, worksheet insertion, CSV analysis, etc.).
'   3. Long-running or AI calls use async wrappers (LLM / PostCorrection) and marshal back
'      to the UI thread via EnsureUIThread().
'
' ARCHITECTURAL COMPONENTS (FILES)
'   ThisAddIn.vb
'       - Host add-in partial class; lifecycle (Startup/Shutdown); hardcoded constants;
'         global state; async LLM bridge; configuration bootstrap; exposes COM automation
'         entry via BridgeSubs; maintains undo cell states; initializes menus & panes.
'   VBA Helper
'       - A VBA-accessible module (external helper file/workbook) that invokes exposed
'         automation methods (e.g. via Application.COMAddIns(...).Object) for macro-based
'         integration; version checked against MinHelperVersion.
'   BridgeSubs.vb
'       - COM-visible automation surface (late-bound calls from VBA or external processes).
'         Provides restricted, parameterized access to safe subsets of add-in functionality.
'         SECURITY: Validate inputs (paths, prompts) to avoid injection or unsafe file IO.
'   DragDropForm.Designer.vb / DragDropForm.vb
'       - WinForms UI for drag & drop ingestion of files. Filters extensions (see
'         allowedExtensions) to constrain processing surface. SECURITY: Enforce whitelist,
'         reject large/binary/unsupported types; sanitize textual content before AI calls.
'   Form1.Designer.vb / Form1.vb (Excel)
'       - Main configuration/interaction window (settings, manual prompt execution, status).
'         Likely hosts controls for model selection, temperature, timeouts, etc. SECURITY:
'         Persist only necessary settings; avoid logging secrets/API keys in plain text.
'   Ribbon1.vb
'       - Defines custom Ribbon buttons and callbacks, dispatching user actions to command
'         handlers. Keep callbacks thin; delegate logic to Commands/Helpers modules.
'   ThisAddIn.Commands.vb
'       - Central command dispatcher; orchestrates user operations (summarize, shorten,
'         translate, markup insertion, batch processing). SECURITY: Validate spreadsheet
'         ranges & user-selected text; guard against over-large prompts (memory / API cost).
'   ThisAddIn.CSVAnalyzer.vb
'       - CSV profiling (delimiter detection, column stats, sampling). SECURITY: Stream
'         files instead of loading entire large datasets; handle malformed input robustly.
'   ThisAddIn.ExcelHelpers.vb
'       - Utility functions for range enumeration, cell value extraction, formula handling,
'         undo stack management, safe write operations. SECURITY: Avoid formula injection;
'         preserve previous state for rollback (undoStates).
'   ThisAddIn.FileHelpers.vb
'       - File IO abstraction (load/save text, PDF/RTF/HTML parsing). SECURITY: Enforce
'         allowedExtensions; normalize paths; avoid directory traversal; prefer read-only
'         access; treat external content as untrusted.
'   ThisAddIn.Helpers.vb
'       - General shared helpers (string transforms, prompt construction, token limits,
'         color parsing, extension-based routing). Keep pure & testable. SECURITY: Encode
'         user-provided text before embedding in markup/HTML contexts.
'   ThisAddIn.MenuContext.vb
'       - Builds legacy or modern context menus. SECURITY: Ensure commands are disabled
'         when selection state invalid; prevent accidental mass operations.
'   ThisAddIn.Pane.vb
'       - Manages CustomTaskPane instances (chat panel / insights). Handles lifecycle,
'         visibility, and cross-thread UI marshalling.
'   ThisAddIn.Processing.InsertIntoWorksheet.vb
'       - Writes AI-generated output back into a worksheet (cell-by-cell, batch, pane).
'         SECURITY: Strip control characters; guard against formula payloads; size checks
'         (LargeWorksheetSize).
'   ThisAddIn.Processing.vb
'       - Core transformation pipeline (invoke LLM, apply shortening, diffing, annotation,
'         bubble markup, context expansion). SECURITY: Rate-limit; timeout long calls;
'         segregate model/system prompts vs. user input; handle cancellation.
'   ThisAddIn.Properties.vb
'       - Centralized strongly-typed property/state access (current prompt mode, user
'         language, flags, feature toggles). SECURITY: Avoid storing secrets; thread-safety
'         for mutable shared state if accessed concurrently.
'
' EXTERNAL DEPENDENCIES
'   - Multiple third-party libraries (DiffPlex, Newtonsoft.Json, HtmlAgilityPack, PdfPig,
'     MarkDig, NAudio, Whisper/Vosk, gRPC, Google Cloud Speech, etc.) used unmodified except
'     MarkdownToRTF (modified). Review their licenses & update cadence; pin versions; ensure
'     no dynamic code execution surfaces (e.g., HTML parsing, audio transcription).
'
' THREADING & ASYNC
'   - Synchronization via WindowsFormsSynchronizationContext; EnsureUIThread() used after
'     async LLM calls to safely update UI. REVIEW: Confirm no deadlocks; prefer ConfigureAwait(False)
'     in library calls where UI not needed.
'
' SECURITY & REVIEW CHECKLIST
'   - Input Validation: All file paths, selected text, worksheet ranges, prompts.
'   - Resource Limits: Prevent oversized prompt or worksheet operations (LargeWorksheetSize).
'   - External Calls: Centralize API keys/config (do not hardcode); use timeouts/cancellation.
'   - Deserialization: Newtonsoft.Json usage must restrict to safe types (avoid TypeNameHandling.All).
'   - Temp / Disk IO: Sanitize filenames; avoid exposing arbitrary filesystem locations.
'   - Undo Integrity: Ensure undoStates cleared after destructive ops to avoid memory growth.
'   - Logging: Avoid writing sensitive data (user content, API tokens).
'
' MAINTENANCE NOTES
'   - Keep processing pure/business logic separate from UI (Commands vs Forms).
'   - Centralize constants (prefix triggers) to avoid divergence across modules.
'   - Consider extracting common code to SharedLibrary for Word/Outlook parity.
'   - Add unit tests for Helpers / ExcelHelpers / Processing transformations.
'
' QUICK START FOR REVIEWERS
'   1. Start at ThisAddIn_Startup to see initialization order.
'   2. Trace a Ribbon command -> Commands -> Processing -> InsertIntoWorksheet.
'   3. Inspect BridgeSubs for external surface & VBA entry points.
'   4. Audit FileHelpers & CSVAnalyzer for untrusted input handling.
'   5. Verify LLM wrapper (ThisAddIn.vb) enforces model/timeout policies.
'
' =========================================================================================