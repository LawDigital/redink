' Part of "Red Ink for Excel"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.vb
' Purpose:
'   Main Excel VSTO add-in entry point. Owns startup/shutdown, configuration
'   initialization, feature setup, and the host-level Excel add-in lifecycle.
'
' Architecture:
'   Root partial ThisAddIn lifecycle module for the Excel host. It initializes shared
'   configuration and add-in features, while worksheet processing, commands, helpers,
'   analysis, panes, and file workflows are implemented in the other ThisAddIn.* files.
' =============================================================================
'
' 22.8.2026
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
' Includes Google Speech V1/V2 library and related API libraries in unchanged form; Copyright (c) 2024 Google LLC; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/googleapis/google-cloud-dotnet
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
' Includes Microsoft.Playwright (Playwright for .NET) in unchanged form; Copyright (c) 2020 Darío Kondratiuk and other contributors, with modifications copyright (c) Microsoft Corporation where stated in the source files; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/microsoft/playwright-dotnet
' Includes also various Microsoft distributables and libraries copyrighted by Microsoft Corporation and available, among others, under the Microsoft EULA, the Visual Studio Community 2022 License, the Microsoft.Web.WebView2 License (for Microsoft.Web.WebView2, see license on https://www.nuget.org/packages/Microsoft.Web.WebView2/ and below) and the MIT License (including Microsoft.Bcl.*, Microsoft.Extensions.*, Microsoft.Identity.Client, Microsoft.Identity.Client.Extensions.Msal, System.*, System.Security.*, System.CodeDom, DocumentFormat.OpenXml.*, Microsoft.ml.*, CommunityToolkit.HighPerformance licensed under MIT License) (https://licenses.nuget.org/MIT); Copyright (c) 2016- Microsoft Corp.
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
Imports Microsoft.Office.Tools.Ribbon
Imports SharedLibrary.SharedLibrary

Module ModuleGetAsyncKeyState

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Function GetAsyncKeyState(ByVal vKey As Integer) As Short
    End Function

End Module

Partial Public Class ThisAddIn

    ' Hardcoded config values

    Public Shared Version As String = "V.220826" & SharedMethods.VersionQualifier

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
    Private Const NoFormulasTrigger As String = "(noformulas)"
    Private Const KBTrigger As String = "(kb)"
    Private Const RIMenu = AN
    Private Const MinHelperVersion = 1  ' Minimum version of the helper file that is required
    Public Const LargeWorksheetSize As Integer = 2500

    Public RemoveMenu As Boolean = False  ' Legacy: If true, the old context menu will be removed (Gen1)

    ' Publicly declared variables so that InterpolateAtRuntime can access them; case-sensitive

    Public TranslateLanguage As String
    Public OutputLanguage As String
    Public FormulaInstruction As String
    Public Dictionary As String = ""
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

    Private Shared _uiContext As SynchronizationContext
    Private Shared _uiScheduler As TaskScheduler
    Private mainThreadControl As New System.Windows.Forms.Control()

    ' Lightweight UI switch helper (no exception risk).
    Private Shared Function EnsureUIThread() As Task
        ' If already on UI thread (or UI context not captured), do nothing.
        If _uiContext Is Nothing OrElse SynchronizationContext.Current Is _uiContext Then
            Return Task.CompletedTask
        End If

        Dim tcs As New TaskCompletionSource(Of Object)(TaskCreationOptions.RunContinuationsAsynchronously)

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

        ' Crash diagnostics: enabled only when the user's My.Settings.CrashLog is True.
        ' This flag is reconciled from INI_Crashlog after config load, so it takes effect
        ' on the next host launch. When False, no handlers are installed (zero overhead).
        Try
            RiCrashLogger.Initialize(
                "RedInk Excel Add-in",
                Me.GetType().Assembly,
                My.Settings.CrashLog,
                True,
                RDV)
        Catch
        End Try

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

        ' 4) Capture the host window’s HWND.
        Dim hwnd As IntPtr = New IntPtr(CInt(Me.Application.Hwnd))
        UpdateHandler.HostHandle = hwnd

        ' Other startup tasks

        SharedMethods.Initialize(Me.CustomTaskPanes)

        Try
            If System.Threading.SynchronizationContext.Current Is Nothing Then
                System.Threading.SynchronizationContext.SetSynchronizationContext(
                    New System.Windows.Forms.WindowsFormsSynchronizationContext())
            End If

            Using anchor As New System.Windows.Forms.Control()
                Dim h = anchor.Handle
            End Using

            SharedLibrary.SharedLibrary.SharedMethods.CleanupOrphanedWebView2Profiles()
            SharedLibrary.Agents.WebView2JsSandbox.Initialize(
                System.Threading.SynchronizationContext.Current)
        Catch
            ' URL import will report "sandbox_uninitialized" if this failed.
        End Try

        InitializeAddInFeatures()
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        Try
            RiCrashLogger.Shutdown("ThisAddIn_Shutdown was called.")
        Catch
        End Try
        RemoveOldContextMenu()
    End Sub

    Public Sub InitializeAddInFeatures()
        InitializeConfig(True, True)

        ' Reconcile the persisted CrashLog switch with the INI parameter. Any change
        ' takes effect on the next host launch (the INI is read after ThisAddIn_Startup).
        Try
            If My.Settings.CrashLog <> INI_Crashlog Then
                My.Settings.CrashLog = INI_Crashlog
                My.Settings.Save()
            End If
        Catch
        End Try

        ' Restore the previously selected primary model (if multi-model is configured)
        If _context.INIloaded Then
            Try
                Dim saved = PrimaryModelManager.LoadSavedModelNumber()
                If PrimaryModelManager.GetAvailableModels().Count > 0 Then
                    ' Try saved selection first; fall back to model 1 if it no longer exists
                    If Not PrimaryModelManager.SelectModel(_context, saved) Then
                        PrimaryModelManager.SelectModel(_context, PrimaryModelManager.GetAvailableModels()(0))
                    End If
                End If
            Catch
                ' non-critical
            End Try
        End If

        AddContextMenu()

        If _context IsNot Nothing AndAlso _context.INIloaded Then
            SharedMethods.RegisterLicenseCounterUsageAndReport(_context, "EX")
        End If

        UpdateHandler.PeriodicCheckForUpdates(INI_UpdateCheckInterval, RDV, INI_UpdatePath, _context)

        ' Initialize model menu buttons on the ribbon
        Try
            If Globals.Ribbons.Ribbon1 IsNot Nothing Then
                Globals.Ribbons.Ribbon1.UpdateModelsMenu()
                Globals.Ribbons.Ribbon1.ApplyRibbonVisibilityConfiguration()
            End If
        Catch
            ' non-critical
        End Try

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


    Public Shared Sub SelectModel(modelNumber As Integer)
        Try
            If PrimaryModelManager.SelectModel(_context, modelNumber) Then
                Try
                    If Globals.Ribbons.Ribbon1 IsNot Nothing Then
                        Globals.Ribbons.Ribbon1.UpdateModelsMenu()
                    End If
                Catch
                    ' non-critical
                End Try
            Else
                SharedMethods.ShowCustomMessageBox($"Model {modelNumber} is not configured.")
            End If
        Catch ex As Exception
            SharedMethods.ShowCustomMessageBox($"Error switching model: {ex.Message}")
        End Try
    End Sub


End Class


' =================================================================================================
' Red Ink for Excel - Architecture Overview
'
' ROLE OF THIS FILE
'   ThisAddIn.vb is the Excel VSTO composition root. It owns startup/shutdown, shared
'   configuration/context initialization, UI-thread marshaling, update checks and the
'   host bridge to shared LLM/model services. Worksheet/business features live in the
'   other partial ThisAddIn.* files.
'
' PRIMARY ARCHITECTURAL AREAS
'   - UI/commands: Ribbon1.vb, Form1.vb, DragDropForm.vb, MenuContext and Pane.
'   - Processing: ThisAddIn.Processing.vb orchestrates transformations; the
'     Processing.InsertIntoWorksheet partial performs controlled worksheet writes.
'   - Excel data helpers: ExcelHelpers handles range/cell/formula concerns; CSVAnalyzer
'     and DirectoryAnalyzer analyze external data; FactExtractor bridges structured fact
'     extraction; FileHelpers/Helpers/Properties provide host utilities and state.
'   - File operations: FileRenamer and related helpers implement explicit file workflows.
'   - Automation: BridgeSubs exposes the bounded COM automation surface used by approved
'     external/VBA callers.
'
' SHARED-LIBRARY BOUNDARY
'   Configuration/licensing, LLM/HTTP/model selection, common text handling, knowledge/
'   M365 services, agent/tool infrastructure and shared UI/utilities belong in
'   SharedLibrary. Excel code should remain focused on workbook/range semantics and host UI.
'
' EXECUTION FLOW
'   Excel/VSTO event or Ribbon/context command -> Commands/Processing -> shared LLM or
'   host helper -> controlled range/workbook update -> UI result. Long-running work must
'   avoid blocking Excel and marshal workbook/UI mutations back to the Excel UI thread.
'
' MAINTENANCE HOTSPOTS
'   COM object lifetime and UI-thread ownership; formula/value distinction and formula
'   injection; range-size/resource limits; undo/state restoration; file/path validation;
'   external/VBA automation inputs; and keeping cross-host policy in SharedLibrary.
' =================================================================================================
