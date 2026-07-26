' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.Processing.Tooling.vb
' Purpose: Implements a model-agnostic "tooling loop" for LLM tool/function
'          calling, including tool selection, tool call detection/extraction,
'          per-tool execution, and response injection back into subsequent LLM
'          iterations.
'
' Architecture:
'  - Tooling execution loop (`ExecuteToolingLoop`):
'      - Builds system prompt augmentation via `BuildToolInstructionsPrompt`.
'      - Injects model-specific tool definitions into `INI_APICall_ToolInstructions_2`.
'      - Calls `LLM(...)` iteratively until no tool calls are detected or
'        `MaxIterations` is reached.
'      - Detects, extracts, executes, and feeds back tool responses between iterations.
'  - Tool execution:
'      - Internal tools include web retrieval, internet search, knowledge-store
'        search, and Microsoft 365 helpers.
'      - External tools execute via model-driven `ModelConfig` definitions.
'      - Tool errors are handled according to `ModelConfig.ToolErrorHandling`.
'  - Tool selection and persistence:
'      - Loads tool-capable services from `INI_SpecialServicePath`.
'      - Adds built-in internal tools and restores persisted user selections.
'  - Diagnostics:
'      - `ToolingFileLogger` records per-run diagnostics and raw-response stubs.
'      - Optional `LogWindow` output provides user-visible progress.
'
' External Dependencies:
'  - SharedLibrary.SharedMethods and shared context/config helpers.
'  - Newtonsoft.Json for parsing and formatting tool calls and responses.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods

''' <summary>
''' Provides tooling support helpers for model-agnostic tool/function calling in LLM interactions.
''' </summary>
Partial Public Class ThisAddIn


    Private _activeToolingContext As ToolExecutionContext = Nothing


    Private Const SubAgentLargeToolResponseThresholdChars As Integer = 30000
    Private Const SubAgentLargeToolResponseExcerptChars As Integer = 8000

    Private Const InternalKnowledgeToolNamePrefix As String = "knowledge_search_store_"


#Region "Tooling Data Classes"

    ''' <summary>
    ''' Represents a single tool call extracted from an LLM response.
    ''' </summary>
    Public Class ToolCall

        ''' <summary>Tool call identifier used to correlate call and response objects.</summary>
        Public Property CallId As String

        ''' <summary>Name of the tool to execute.</summary>
        Public Property ToolName As String

        ''' <summary>Arguments passed to the tool.</summary>
        Public Property Arguments As Dictionary(Of String, Object)

        ''' <summary>Raw JSON representation of the tool call token.</summary>
        Public Property RawJson As String

        ''' <summary>
        ''' Initializes a new tool call instance with an empty arguments dictionary.
        ''' </summary>
        Public Sub New()
            Arguments = New Dictionary(Of String, Object)()
        End Sub
    End Class

    ''' <summary>
    ''' Represents the outcome of executing a single tool call.
    ''' </summary>



#End Region




#Region "Execute Tooling"

    ''' <summary>
    ''' Executes an iterative tool-enabled LLM loop until either no tool calls are detected, the maximum iteration count is reached,
    ''' or the user cancels. Tool call detection/extraction and response injection are controlled by the active tooling model config.
    ''' </summary>
    ''' <param name="sysCommand">Base system command prompt text.</param>
    ''' <param name="userText">User prompt text (used only if fullPromptOverride is empty).</param>
    ''' <param name="selectedTools">Tool configurations available to the model.</param>
    ''' <param name="useSecondAPI">Whether to route LLM calls through the secondary API.</param>
    ''' <param name="fileObject">Optional file object payload passed to <c>LLM</c>.</param>
    ''' <param name="doTPMarkup">Unused parameter in this method body (passed through by caller signature).</param>
    ''' <param name="bubblesText">Optional bubble text appended to the user prompt.</param>
    ''' <param name="noFormatting">Unused parameter in this method body (passed through by caller signature).</param>
    ''' <param name="keepFormat">Unused parameter in this method body (passed through by caller signature).</param>
    ''' <param name="slideDeck">Unused parameter in this method body (passed through by caller signature).</param>
    ''' <param name="doMyStyle">Unused parameter in this method body (passed through by caller signature).</param>
    ''' <param name="myStyleInsert">Unused parameter in this method body (passed through by caller signature).</param>
    ''' <param name="addDocs">If True, <paramref name="insertDocs"/> is appended to the user prompt (when provided).</param>
    ''' <param name="insertDocs">Optional inserted document text appended to the user prompt when <paramref name="addDocs"/> is True.</param>
    ''' <param name="slideInsert">Optional slide text appended to the user prompt.</param>
    ''' <param name="otherPrompt">Optional additional user prompt passed to <c>LLM</c>.</param>
    ''' <param name="fullPromptOverride">When provided, uses this as the complete user prompt instead of building one internally. 
    ''' This ensures tooling calls receive the same context as non-tooling LLM calls.</param>
    ''' <param name="hideSplash">When True, suppresses the splash/progress indicator during LLM calls.</param>
    ''' <param name="hideLogWindow">When True, suppresses the tooling log window (useful for chat integration).</param>
    ''' <param name="DoChart">When True, adds charting instructions to the system prompt.</param>
    ''' <returns>The final LLM response string returned by the last iteration.</returns>
    Public Async Function ExecuteToolingLoop(
        sysCommand As String,
        userText As String,
        selectedTools As List(Of ModelConfig),
        useSecondAPI As Boolean,
        Optional fileObject As String = "",
        Optional doTPMarkup As Boolean = False,
        Optional bubblesText As String = "",
        Optional noFormatting As Boolean = False,
        Optional keepFormat As Boolean = False,
        Optional slideDeck As String = "",
        Optional doMyStyle As Boolean = False,
        Optional myStyleInsert As String = "",
        Optional addDocs As Boolean = False,
        Optional insertDocs As String = "",
        Optional slideInsert As String = "",
        Optional otherPrompt As String = "",
        Optional fullPromptOverride As String = "",
        Optional hideSplash As Boolean = False,
        Optional hideLogWindow As Boolean = False,
        Optional DoChart As Boolean = False,
        Optional cancellationToken As System.Threading.CancellationToken = Nothing,
        Optional subAgentMode As Boolean = False,
        Optional subAgentAllowedToolNames As IReadOnlyList(Of String) = Nothing,
        Optional subAgentSpecialModelKey As String = Nothing,
        Optional subAgentAuthoritativeRegistry As SharedLibrary.Agents.ToolRegistry = Nothing,
        Optional subAgentRegistrySource As String = Nothing,
        Optional subAgentParentRunId As String = Nothing,
        Optional subAgentInvocationIndex As Integer = 0,
        Optional subAgentAgentInvocationCount As Integer = 0,
        Optional subAgentName As String = Nothing,
        Optional userLanguage As String = "",
        Optional workflowId As String = "",
        Optional memoryGroundingMode As SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingMode = SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingMode.None,
        Optional memoryGroundingModeIsExplicit As Boolean = False,
        Optional finalResponseContract As SharedLibrary.Agents.ToolingFinalResponseContract = SharedLibrary.Agents.ToolingFinalResponseContract.UserFacingTaskStatus) As Task(Of String)


        ToolingFileLogger.StartSession()

        Dim parentToolingContext = _activeToolingContext
        Dim workflowScope As IDisposable = Nothing
        Dim acceptedFinalStatus As String = ""

        ' Initialize the Word agent workspace (persisted; falls back to active doc folder, then Desktop).
        Try
            Dim ws = SharedLibrary.Agents.WorkspaceStore.Load("word")
            If String.IsNullOrWhiteSpace(ws.RootPath) OrElse Not Directory.Exists(ws.RootPath) Then
                Try
                    Dim doc As Microsoft.Office.Interop.Word.Document = Nothing
                    Try : doc = Globals.ThisAddIn.Application.ActiveDocument : Catch : End Try
                    If doc IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(doc.Path) AndAlso Directory.Exists(doc.Path) Then
                        ws = New SharedLibrary.Agents.WorkspaceState() With {
                            .RootPath = doc.Path,
                            .PersistUntilRevoked = False,
                            .AllowRead = True, .AllowWrite = True,
                            .AllowMoveCopyRename = True, .AllowDelete = False
                        }
                    End If
                Catch
                End Try
            End If
            SharedLibrary.Agents.WorkspaceTools.SetActive(ws)
        Catch
            SharedLibrary.Agents.WorkspaceTools.SetActive(New SharedLibrary.Agents.WorkspaceState())
        End Try

        Dim fullAllowedTools As List(Of ModelConfig) =
            If(selectedTools, New List(Of ModelConfig)()).
                Where(Function(t) t IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(t.ToolName)).
                GroupBy(Function(t) t.ToolName, StringComparer.OrdinalIgnoreCase).
                Select(Function(g) g.First()).
                ToList()

        selectedTools = fullAllowedTools

        Try
            SharedLibrary.Agents.HostToolRegistration.RegisterResolvedInternalTools(
                SharedLibrary.Agents.ToolingHostKind.Word,
                fullAllowedTools)
        Catch ex As Exception
            ToolingFileLogger.LogWarn("Failed to register selected Word tooling tools.", ex:=ex)
        End Try

        Dim context As New ToolExecutionContext() With {
            .MaxIterations = INI_ToolingMaximumIterations
        }

        context.FinalResponseContract = finalResponseContract
        context.Log(
            "Final response contract initialized: " &
            SharedLibrary.Agents.ToolingFinalResponseContractHelpers.FormatToolingFinalResponseContract(context.FinalResponseContract),
            "diag")

        context.SequencingState.UserLanguage =
            If(Not String.IsNullOrWhiteSpace(userLanguage),
               userLanguage.Trim(),
               Await ResolveToolingUserLanguageAsync(
                   userText,
                   otherPrompt,
                   fullPromptOverride,
                   useSecondAPI,
                   hideSplash))

        context.LatestUserRequestRaw =
            ResolveLatestUserRequestRaw(
                userText,
                otherPrompt,
                fullPromptOverride)

        ' Created-deliverable enforcement is metadata-driven only.
        ' Do not classify the user's text to decide whether an artifact is required.

        context.HostTaskSummary =
            ResolveOptionalHostTaskSummary(
                context.LatestUserRequestRaw,
                userText,
                otherPrompt,
                fullPromptOverride,
                insertDocs,
                slideInsert,
                bubblesText)

        Dim toolSelectionHintText As String = BuildToolSelectionHintText(
            userText,
            fullPromptOverride,
            otherPrompt,
            insertDocs,
            slideInsert,
            bubblesText)

        Try
            SharedLibrary.Agents.AgentResources.SetPaths(INI_AgentResourcesPath, INI_AgentResourcesPathLocal)
            SharedLibrary.Agents.AgentResources.Refresh()
        Catch
        End Try

        ' Replace the authoritative-registry setup block in ExecuteToolingLoop(...)

        Dim authoritativeRegistrySource As SharedLibrary.Agents.ToolRegistry = Nothing

        If subAgentMode AndAlso subAgentAuthoritativeRegistry IsNot Nothing Then
            authoritativeRegistrySource = subAgentAuthoritativeRegistry.Snapshot()
        Else
            authoritativeRegistrySource = SharedLibrary.Agents.ToolRegistryBuilder.FromModelConfigs(fullAllowedTools, "selected")

            ' Only register the skills/agents the user actually selected.
            ' This keeps Word aligned with Outlook and prevents unselected skills/agents
            ' from becoming callable just because they exist on disk.
            Try
                SharedLibrary.Agents.AgentResources.Refresh()

                Dim selectedToolNames As New HashSet(Of String)(
            If(fullAllowedTools, New List(Of ModelConfig)()).
                Where(Function(t) t IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(t.ToolName)).
                Select(Function(t) t.ToolName.Trim()),
            StringComparer.OrdinalIgnoreCase)

                If selectedToolNames.Count > 0 Then
                    Dim selectedSkills =
                SharedLibrary.Agents.AgentResources.Skills.
                    Where(Function(sk)
                              Return sk IsNot Nothing AndAlso
                                     Not String.IsNullOrWhiteSpace(sk.Name) AndAlso
                                     selectedToolNames.Contains("skill_" & sk.Name.Trim())
                          End Function)

                    Dim selectedAgents =
                SharedLibrary.Agents.AgentResources.Agents.
                    Where(Function(ag)
                              Return ag IsNot Nothing AndAlso
                                     Not String.IsNullOrWhiteSpace(ag.Name) AndAlso
                                     selectedToolNames.Contains("agent_" & ag.Name.Trim())
                          End Function)

                    SharedLibrary.Agents.ToolRegistryBuilder.AddSkills(authoritativeRegistrySource, selectedSkills)
                    SharedLibrary.Agents.ToolRegistryBuilder.AddAgents(authoritativeRegistrySource, selectedAgents)
                End If
            Catch
            End Try
        End If

        Dim authoritativeRegistrySnapshot As SharedLibrary.Agents.ToolRegistry =
    If(authoritativeRegistrySource Is Nothing,
       Nothing,
       authoritativeRegistrySource.Snapshot())

        context.AuthoritativeToolRegistry = authoritativeRegistrySource
        context.AuthoritativeToolRegistrySnapshot = authoritativeRegistrySnapshot
        context.HostKind = "Word"
        context.EnforceAllowedToolScope = subAgentMode
        context.SequencingState.ActiveToolingSession = True
        context.SequencingState.HasOpenToolWorkflow = True
        context.SequencingState.ToolRequiredModeUsed = False
        context.SequencingState.MemoryGroundingMode = memoryGroundingMode
        context.SequencingState.UserLanguage =
            If(Not String.IsNullOrWhiteSpace(userLanguage),
               userLanguage.Trim(),
               context.SequencingState.UserLanguage)
        context.SequencingState.FinalCompleteRejectedForMissingMemoryAccess = False
        context.WorkflowId = ResolveToolingWorkflowId(workflowId, subAgentMode, parentToolingContext)

        If Not subAgentMode Then
            Dim clearedCount As Integer =
                SharedLibrary.Agents.SessionMemory.ClearTransientEntriesForHost(
                    context.HostKind,
                    context.WorkflowId)

            If clearedCount > 0 Then
                context.Log(
                    $"Cleared {clearedCount} transient session-memory entr{If(clearedCount = 1, "y", "ies")} for host '{context.HostKind}'.",
                    "diag")
            End If
        End If

        context.RuntimeState =
    If(subAgentMode,
       SharedLibrary.Agents.WorkflowContinuity.AttachWorkflow(context.WorkflowId, context.HostKind),
       SharedLibrary.Agents.WorkflowContinuity.StartWorkflow(context.WorkflowId, context.HostKind))

        workflowScope = SharedLibrary.Agents.WorkflowContinuity.BeginWorkflowScope(context.WorkflowId, context.HostKind)

        ToolingFileLogger.LogStep(
            SharedLibrary.Agents.WorkflowContinuity.ComposeWorkflowLogMessage(
                "Tooling workflow initialized.",
                context.WorkflowId,
                If(context.RuntimeState?.CurrentPhase, ""),
                hostName:=context.HostKind) &
            " [subAgentMode: " & If(subAgentMode, "true", "false") & "]")

        context.Log("Workflow continuity initialized.")
        context.Log("Memory grounding state initialized: " &
            SharedLibrary.Agents.ToolCallSequencing.BuildMemoryGroundingStateSummary(context.SequencingState))

        Dim initialSubAgentToolNames As HashSet(Of String) = Nothing

        If subAgentMode Then
            Dim expandedSubAgentAllowedToolNames As List(Of String) =
        ExpandAllowedToolNamesForRegistry(
            subAgentAllowedToolNames,
            context.AuthoritativeToolRegistrySnapshot)

            initialSubAgentToolNames = BuildToolNameSet(expandedSubAgentAllowedToolNames)
            context.AllowedToolNames = New HashSet(Of String)(initialSubAgentToolNames, StringComparer.OrdinalIgnoreCase)
            context.AllowedToolRegistry =
        If(context.AuthoritativeToolRegistrySnapshot Is Nothing,
           New SharedLibrary.Agents.ToolRegistry(),
           context.AuthoritativeToolRegistrySnapshot.Narrow(context.AllowedToolNames))
        Else
            context.AllowedToolNames = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            context.AllowedToolRegistry =
        If(context.AuthoritativeToolRegistrySnapshot Is Nothing,
           New SharedLibrary.Agents.ToolRegistry(),
           context.AuthoritativeToolRegistrySnapshot)
        End If

        ToolingFileLogger.LogStep(
            SharedLibrary.Agents.WorkflowContinuity.ComposeWorkflowLogMessage(
                "Tooling host context attached.",
                context.WorkflowId,
                If(context.RuntimeState?.CurrentPhase, ""),
                hostName:=context.HostKind))

        If subAgentMode Then
            ToolingFileLogger.LogStep(
                SharedLibrary.Agents.WorkflowContinuity.ComposeWorkflowLogMessage(
                    "Allowed tool scope initialized.",
                    context.WorkflowId,
                    If(context.RuntimeState?.CurrentPhase, ""),
                    hostName:=context.HostKind,
                    leadingMarker:="[subagent]") &
                " [allowedTools: " &
                If(context.AllowedToolNames.Count = 0,
                   "(none)",
                   String.Join(", ", context.AllowedToolNames.OrderBy(Function(n) n))) &
                "]")
        End If

        If subAgentMode Then
            context.LazyToolLoadingEnabled = False

            Dim scopeInit = SharedLibrary.Agents.SubAgentToolScopeInitializer.Initialize(
    context.AuthoritativeToolRegistrySnapshot,
    context.AllowedToolNames)

            context.AllowedToolRegistry = scopeInit.NarrowedRegistry
            context.SelectedTools = New List(Of ModelConfig)(scopeInit.ResolvedTools)

            Dim requestedNamesText As String =
                If(scopeInit.RequestedToolNames Is Nothing OrElse scopeInit.RequestedToolNames.Count = 0,
                   "(none)",
                   String.Join(", ", scopeInit.RequestedToolNames))

            Dim resolvedNamesText As String =
                If(scopeInit.ResolvedToolNames.Count = 0,
                   "(none)",
                   String.Join(", ", scopeInit.ResolvedToolNames))

            Dim missingNamesText As String =
                If(scopeInit.MissingToolNames.Count = 0,
                   "(none)",
                   String.Join(", ", scopeInit.MissingToolNames))

            Dim finalSelectedNamesText As String =
                If(scopeInit.FinalCallableToolNames.Count = 0,
                   "(none)",
                   String.Join(", ", scopeInit.FinalCallableToolNames))

            Dim repeatedInvocationLabel As String =
                If(subAgentAgentInvocationCount > 1, "later", "first")

            context.Log($"Sub-agent tool scope initialized from '{If(subAgentRegistrySource, "")}'.")

            ToolingFileLogger.LogStep(
                "[subagent-host] tool_scope_init: " &
                $"parent_run_id={If(subAgentParentRunId, "(none)")}; " &
                $"invocation_index={subAgentInvocationIndex}; " &
                $"agent='{If(subAgentName, "")}'; " &
                $"requested_allowed_tools={requestedNamesText}; " &
                $"registry_source={If(subAgentRegistrySource, "(unspecified)")}; " &
                $"resolved_tools={resolvedNamesText}; " &
                $"missing_tools={missingNamesText}; " &
                $"final_selected_tools={finalSelectedNamesText}; " &
                $"same_agent_invocation={repeatedInvocationLabel}")

            If scopeInit.HasRequestedTools AndAlso
               (scopeInit.HasMissingRequestedTools OrElse scopeInit.HasMissingFinalToolNames) Then

                Dim missingRequiredToolNames = scopeInit.MissingFinalToolNames

                Dim payload = SharedLibrary.Agents.SubAgentRuntimeHardening.BuildRequiredToolMissingPayload(
                    missingRequiredToolNames,
                    requestedToolNames:=scopeInit.RequestedToolNames,
                    resolvedToolNames:=scopeInit.ResolvedToolNames)

                context.LogWarn(
                    "Sub-agent required tools could not be resolved before model call.",
                    details:=$"host={context.HostKind}; agent={If(subAgentName, "")}; parentRunId={If(subAgentParentRunId, "(none)")}; invocationIndex={subAgentInvocationIndex}; registrySource={If(subAgentRegistrySource, "(unspecified)")}; requested={requestedNamesText}; resolved={resolvedNamesText}; missing={If(missingRequiredToolNames.Count = 0, "(none)", String.Join(", ", missingRequiredToolNames))}; finalSelected={finalSelectedNamesText}")

                ToolingFileLogger.EndSession(False, "Sub-agent required tools missing.")
                Return payload
            End If
        Else
            context.SelectedTools = BuildInitialToolExposure(fullAllowedTools, context.AllowedToolRegistry, toolSelectionHintText)
            context.LazyToolLoadingEnabled =
                context.SelectedTools IsNot Nothing AndAlso
                context.SelectedTools.Any(
                    Function(t)
                        Return t IsNot Nothing AndAlso
                               Not String.IsNullOrWhiteSpace(t.ToolName) AndAlso
                               t.ToolName.Equals(SharedLibrary.Agents.ToolLoaderTool.LoaderToolName, StringComparison.OrdinalIgnoreCase)
                    End Function)
        End If

        selectedTools = New List(Of ModelConfig)(context.SelectedTools)

        context.ToolingModel = GetCurrentConfig(_context)

        ' Start-of-run config logging (ONCE)
        ToolingFileLogger.LogModelConfigOnce(context.ToolingModel, "Tooling LLM ModelConfig")

        Dim internalSelected As Boolean =
            (selectedTools IsNot Nothing AndAlso selectedTools.Any(Function(t) t.ToolName IsNot Nothing AndAlso t.ToolName.Equals(InternalWebToolName, StringComparison.OrdinalIgnoreCase)))

        ToolingFileLogger.LogStep($"Internal web tool selected: {internalSelected}")

        If selectedTools IsNot Nothing Then
            For i = 0 To selectedTools.Count - 1
                ToolingFileLogger.LogModelConfigOnce(selectedTools(i), $"Selected Tool #{i + 1} ModelConfig")
            Next
        End If

        If INI_ToolingLogWindow AndAlso Not hideLogWindow Then
            ' The LogWindow is a WinForms Form and MUST be created on the Word UI/STA thread.
            ' Callers may await us with ConfigureAwait(False), which can land this code on a
            ' thread-pool worker without a message pump. In that case, creating/showing the
            ' Form on the wrong thread causes "frozen, unpositioned" windows and full Word
            ' hangs on subsequent invocations. Always marshal via the captured UI context.
            Dim createdForm As LogWindow = Nothing
            Dim createError As Exception = Nothing

            Dim createOnUi As System.Action =
                Sub()
                    Try
                        createdForm = New LogWindow()
                        createdForm.Show()
                        AddHandler createdForm.CancelRequested, Sub() context.IsCancelled = True
                    Catch ex As Exception
                        createError = ex
                    End Try
                End Sub

            If UiSyncContext IsNot Nothing AndAlso
               System.Threading.Thread.CurrentThread.ManagedThreadId <> UiThreadId Then
                ' Synchronous Send so the form exists before we continue logging into it.
                UiSyncContext.Send(Sub() createOnUi(), Nothing)
            Else
                createOnUi()
            End If

            If createError IsNot Nothing Then
                ToolingFileLogger.LogWarn("Failed to create LogWindow on UI thread.", ex:=createError)
            End If

            context.LogWindowForm = createdForm

        ElseIf hideLogWindow AndAlso
               parentToolingContext IsNot Nothing AndAlso
               parentToolingContext.LogWindowForm IsNot Nothing AndAlso
               Not parentToolingContext.LogWindowForm.IsDisposed Then

            context.LogPrefix = "[subagent] "
            context.ExternalLogSink =
                Sub(message As String, level As String)
                    parentToolingContext.LogWindowForm.AppendLog(message, level)
                End Sub
        End If

        _activeToolingContext = context

        context.Log("Starting tooling session...")
        If selectedTools IsNot Nothing Then
            context.Log($"Selected tools: {String.Join(", ", selectedTools.Select(Function(t) t.ToolName))}")
        Else
            context.Log("Selected tools: (none)")
        End If

        Try
            Await ResolveMemoryGroundingModeAsync(
                context,
                userText,
                otherPrompt,
                fullPromptOverride,
                useSecondAPI,
                hideSplash,
                memoryGroundingMode,
                memoryGroundingModeIsExplicit,
                subAgentMode)

            If Not subAgentMode Then
                Await TryPrimeRecentMemoryStubsAsync(context, cancellationToken)
                Await TryForceRequiredMemoryGetsAsync(context, cancellationToken)

                If context.AllToolResponses.Count > 0 Then
                    INI_APICall_ToolResponses_2 = BuildToolResponsesForModel(
                        context.AllToolResponses,
                        context.ToolingModel,
                        compactForSubAgent:=False)
                    context.Log("Initial tool responses prepared for model", "diag")
                End If
            End If

            ' Build System Prompt (matching direct LLM() call plus Tooling Instructions)
            ' Base system command
            Dim baseSysPrompt As String = sysCommand

            ' Add bubbles extract instruction if bubbles text is present
            If Not String.IsNullOrWhiteSpace(bubblesText) Then
                baseSysPrompt &= " " & SP_Add_BubblesExtract
            End If

            ' Add revisions instruction if TP markup is enabled
            If doTPMarkup Then
                baseSysPrompt &= " " & SP_Add_Revisions
            End If

            ' Add formatting instructions (slides vs. HTML/inline)
            If Not String.IsNullOrWhiteSpace(slideDeck) Then
                baseSysPrompt &= " " & SP_Add_Slides
            ElseIf Not noFormatting Then
                If keepFormat Then
                    baseSysPrompt &= " " & SP_Add_KeepHTMLIntact
                Else
                    baseSysPrompt &= " " & SP_Add_KeepInlineIntact
                End If
            End If

            ' Add MyStyle insert if enabled
            If doMyStyle AndAlso Not String.IsNullOrWhiteSpace(myStyleInsert) Then
                baseSysPrompt &= " " & myStyleInsert
            End If

            ' Add DoChart insert if enabled
            If DoChart Then
                baseSysPrompt &= " " & SP_Add_Chart
            End If

            ' Add privacy protection instructions when search privacy is enabled
            If INI_EnablePrivacyForSearch AndAlso Not String.IsNullOrWhiteSpace(SP_Add_PrivacyProtection) Then
                baseSysPrompt &= " " & SP_Add_PrivacyProtection
            End If

            ' Add tool instructions on top of the standard prompt additions
            ' Dim enhancedSysPrompt As String = baseSysPrompt & Environment.NewLine & Environment.NewLine & BuildToolInstructionsPrompt(selectedTools)

            If Not subAgentMode Then
                ' Agent layer: prepend Inky.md guidance + skill/agent availability summary, append agent-layer addendum (main loop only).
                Dim selectedSkillToolNames As New List(Of String)()
                Dim selectedAgentToolNames As New List(Of String)()
                If selectedTools IsNot Nothing Then
                    For Each __t In selectedTools
                        If __t Is Nothing OrElse String.IsNullOrWhiteSpace(__t.ToolName) Then Continue For
                        If __t.ToolName.StartsWith("skill_", StringComparison.OrdinalIgnoreCase) Then selectedSkillToolNames.Add(__t.ToolName)
                        If __t.ToolName.StartsWith("agent_", StringComparison.OrdinalIgnoreCase) Then selectedAgentToolNames.Add(__t.ToolName)
                    Next
                End If
                Dim inkyHeader As String = SharedLibrary.Agents.InkyPromptBuilder.Build(selectedSkillToolNames, selectedAgentToolNames)
                Dim agentLayerActive As Boolean =
                (selectedSkillToolNames.Count > 0) OrElse
                (selectedAgentToolNames.Count > 0) OrElse
                (selectedTools IsNot Nothing AndAlso selectedTools.Any(Function(t) t IsNot Nothing AndAlso SharedLibrary.Agents.MemoryTools.IsMemoryTool(t.ToolName)))
                If Not String.IsNullOrWhiteSpace(inkyHeader) Then
                    baseSysPrompt = inkyHeader & Environment.NewLine & Environment.NewLine & baseSysPrompt
                End If
                If agentLayerActive Then
                    baseSysPrompt &= Environment.NewLine & Environment.NewLine & Default_SP_Add_AgentLayer
                End If
            End If

            ' Language contract — final user-facing prose must be in the detected user
            ' language regardless of the language of guard prompts or tool output (P1 fix).
            Dim languageContractFragment As String = ""

            If SharedLibrary.Agents.ToolingFinalResponseContractHelpers.RequiresTaskStatusFooter(context.FinalResponseContract) Then
                languageContractFragment =
                    Agents.ToolingOrchestrator.BuildLanguageContractSystemPromptFragment(
                        If(context.SequencingState IsNot Nothing,
                           context.SequencingState.UserLanguage,
                           userLanguage))
            End If

            Dim enhancedSysPrompt As String =
                baseSysPrompt & Environment.NewLine & Environment.NewLine &
                BuildToolInstructionsPromptForSession(
                    context.SelectedTools,
                    subAgentMode,
                    context.FinalResponseContract)

            If Not String.IsNullOrWhiteSpace(languageContractFragment) Then
                enhancedSysPrompt = enhancedSysPrompt & Environment.NewLine & Environment.NewLine &
                                    languageContractFragment
            End If


            Dim toolDefinitions = BuildToolInstructionsForModel(context.SelectedTools, context.ToolingModel)
            INI_APICall_ToolInstructions_2 = toolDefinitions
            INI_APICall_ToolResponses_2 = ""

            context.Log("Tool definitions prepared for model", "diag")

            If Not subAgentMode Then
                Await TryPrimeRecentMemoryStubsAsync(context, cancellationToken)

                If context.AllToolResponses.Count > 0 Then
                    INI_APICall_ToolResponses_2 = BuildToolResponsesForModel(
                        context.AllToolResponses,
                        context.ToolingModel,
                        compactForSubAgent:=False)
                    context.Log("Initial tool responses prepared for model", "diag")
                End If
            End If

            If INI_ToolingDryRun Then
                Dim preview = $"The following tools will be made available to the model:{Environment.NewLine}{Environment.NewLine}"
                For Each tool In selectedTools
                    preview &= $"- {tool.ToolName}: {tool.ToolInstructionsPrompt}{Environment.NewLine}"
                Next

                Dim proceed = ShowCustomYesNoBox(preview & Environment.NewLine & "Do you want to proceed with the tool-enabled call?", "Proceed", "Abort")
                If proceed <> 1 Then
                    context.LogWarn("Dry run aborted by user")
                    ToolingFileLogger.EndSession(False, "Dry run aborted by user")
                    Return ""
                End If
            End If

            Dim currentResponse As String = ""
            Dim iteration As Integer = 0
            Dim fullUserPrompt As String = ""
            Dim abortDueToRepeatedToolLoop As Boolean = False

            ' Determine if usertext is empty/whitespace
            Dim noSelectedText As Boolean = String.IsNullOrWhiteSpace(userText)

            While iteration < context.MaxIterations AndAlso Not context.IsCancelled

                iteration += 1
                context.CurrentIteration = iteration
                context.Log($"--- Iteration {iteration} of {context.MaxIterations} ---")

                context.Log("Calling LLM...", "llm")
                Debug.WriteLine($"[WORD-TOOLING] BEFORE prompt-build subAgentMode={subAgentMode} iteration={iteration} gateBusy={SharedLibrary.Agents.AgentGate.IsBusy} thread={Threading.Thread.CurrentThread.ManagedThreadId}")

                ' Build user prompt - use override if provided, otherwise build internally
                If Not String.IsNullOrWhiteSpace(fullPromptOverride) Then
                    ' Use the caller-provided prompt (ensures same context as non-tooling calls)
                    fullUserPrompt = fullPromptOverride
                Else
                    ' Original internal prompt building logic
                    If noSelectedText Then
                        fullUserPrompt = If(addDocs AndAlso Not String.IsNullOrWhiteSpace(insertDocs), " " & insertDocs & " ", "")
                        If Not String.IsNullOrWhiteSpace(slideInsert) Then
                            fullUserPrompt &= slideInsert
                        End If
                    Else
                        fullUserPrompt = "<TEXTTOPROCESS>" & userText & "</TEXTTOPROCESS>"
                        If addDocs AndAlso Not String.IsNullOrWhiteSpace(insertDocs) Then
                            fullUserPrompt &= " " & insertDocs & " "
                        End If
                        If Not String.IsNullOrWhiteSpace(slideInsert) Then
                            fullUserPrompt &= slideInsert
                        End If
                        If Not String.IsNullOrWhiteSpace(bubblesText) Then
                            fullUserPrompt &= " " & bubblesText
                        End If
                    End If
                End If

                fullUserPrompt =
                    BuildPromptWithAuthoritativeLatestUserRequest(
                        context,
                        fullUserPrompt)

                LogLatestUserRequestDiagnostic(context, "main")

                enhancedSysPrompt =
                    baseSysPrompt & Environment.NewLine & Environment.NewLine &
                    BuildToolInstructionsPromptForSession(
                        context.SelectedTools,
                        subAgentMode,
                        context.FinalResponseContract)

                If Not String.IsNullOrWhiteSpace(languageContractFragment) Then
                    enhancedSysPrompt = enhancedSysPrompt & Environment.NewLine & Environment.NewLine &
                                        languageContractFragment
                End If

                INI_APICall_ToolInstructions_2 = BuildToolInstructionsForModel(context.SelectedTools, context.ToolingModel)

                ' Continuation-guard injection (sysprompt addendum + rejected-turn evidence on the user side).
                ' We append both to the user prompt so the model sees explicit conversational evidence that its
                ' previous attempt was rejected. System-prompt-only nudges proved insufficient because the user
                ' prompt and tool-response state are otherwise identical between iterations.
                Dim sysPromptForThisCall As String = enhancedSysPrompt
                Dim userPromptForThisCall As String = fullUserPrompt

                If Not subAgentMode Then
                    Dim runtimeContextBlock As String = BuildRuntimeContextPromptBlock(context)
                    If Not String.IsNullOrWhiteSpace(runtimeContextBlock) Then
                        sysPromptForThisCall &= Environment.NewLine & Environment.NewLine & runtimeContextBlock
                    End If

                    If SharedLibrary.Agents.ToolingFinalResponseContractHelpers.RequiresTaskStatusFooter(context.FinalResponseContract) Then
                        Dim postToolContinuationBlock As String =
                            BuildPostToolContinuationBlock(context)

                        If Not String.IsNullOrWhiteSpace(postToolContinuationBlock) Then
                            sysPromptForThisCall &= Environment.NewLine & Environment.NewLine & postToolContinuationBlock
                            LogLatestUserRequestDiagnostic(context, "continuation")
                        End If
                    End If
                End If

                If Not String.IsNullOrWhiteSpace(context.PendingContinuationGuardPrompt) Then
                    sysPromptForThisCall &= Environment.NewLine & Environment.NewLine & context.PendingContinuationGuardPrompt

                    Dim rejected As String = If(context.PendingRejectedAssistantTurn, "")
                    Dim guardTitle As String = If(context.PendingGuardTitle, "HOST CONTINUATION GUARD")
                    Dim rejectedExplanation As String = If(
        context.PendingRejectedTurnExplanation,
        "Your previous turn was rejected.")

                    Dim guardBlock As New System.Text.StringBuilder()
                    guardBlock.AppendLine()
                    guardBlock.AppendLine("[" & guardTitle & "]")
                    guardBlock.AppendLine(context.PendingContinuationGuardPrompt)
                    guardBlock.AppendLine()
                    If Not String.IsNullOrWhiteSpace(rejected) Then
                        guardBlock.AppendLine(rejectedExplanation)
                        guardBlock.AppendLine("<<<REJECTED_TURN")
                        guardBlock.AppendLine(rejected)
                        guardBlock.AppendLine("REJECTED_TURN>>>")
                        guardBlock.AppendLine()
                    End If
                    guardBlock.AppendLine("[/" & guardTitle & "]")

                    userPromptForThisCall = fullUserPrompt & Environment.NewLine & guardBlock.ToString()

                    LogLatestUserRequestDiagnostic(context, "repair")
                    context.LogWarn("Applying host-side continuation guard.", visibleToUser:=False)
                    context.PendingContinuationGuardPrompt = ""
                    context.PendingRejectedAssistantTurn = ""
                    context.PendingGuardTitle = ""
                    context.PendingRejectedTurnExplanation = ""
                End If

                LogFinalResponseContractDiagnostics(context, sysPromptForThisCall)

                Debug.WriteLine($"[WORD-TOOLING] BEFORE LLM subAgentMode={subAgentMode} iteration={iteration} sysLen={If(sysPromptForThisCall, "").Length} userLen={If(fullUserPrompt, "").Length} tools={If(selectedTools?.Count, 0)} gateBusy={SharedLibrary.Agents.AgentGate.IsBusy} thread={Threading.Thread.CurrentThread.ManagedThreadId}")

                currentResponse = Await LLM(
                    sysPromptForThisCall,
                    userPromptForThisCall,
                    "", "", 0,
                    useSecondAPI,
                    hideSplash,
                    otherPrompt,
                    fileObject,
                    True)

                Debug.WriteLine($"[WORD-TOOLING] AFTER LLM subAgentMode={subAgentMode} iteration={iteration} responseLen={If(currentResponse, "").Length} gateBusy={SharedLibrary.Agents.AgentGate.IsBusy} thread={Threading.Thread.CurrentThread.ManagedThreadId}")

                ToolingFileLogger.LogRawResponseStub("Main LLM()", currentResponse)

                If String.IsNullOrWhiteSpace(currentResponse) Then
                    Dim lastSuccess = GetLastSuccessfulToolResponse(context)

                    If Not subAgentMode AndAlso
                       lastSuccess IsNot Nothing AndAlso
                       SharedLibrary.Agents.ToolCallSequencing.IsSuccessfulDeliverableResult(
                           If(lastSuccess.Response, "")) Then

                        context.EmptyMainModelResponse = False
                        currentResponse = Await BuildLocalizedCreatedDeliverableSuccessMessageAsync(
                            context,
                            lastSuccess,
                            useSecondAPI,
                            hideSplash, cancellationToken)

                        If context.SequencingState IsNot Nothing Then
                            context.SequencingState.FinalResponseOrigin = "host_generated"
                            context.SequencingState.HasOpenToolWorkflow = False
                        End If

                        context.Log(
                            "Empty main-model response after successful deliverable creation; returning localized host-generated success message.",
                            "success")

                        Exit While
                    End If

                    If Not subAgentMode AndAlso lastSuccess IsNot Nothing Then
                        Dim requiresMissingOutputFileRecovery As Boolean =
                            context.SequencingState IsNot Nothing AndAlso
                            context.SequencingState.RequestRequiresCreatedDeliverable AndAlso
                            context.SequencingState.LastToolProducesIntermediateData AndAlso
                            Not SharedLibrary.Agents.ToolCallSequencing.HasProducedUserDeliverable(context.SequencingState)

                        If requiresMissingOutputFileRecovery AndAlso context.PrematureTextRetryCount = 0 Then
                            context.PrematureTextRetryCount = 1
                            context.PendingContinuationGuardPrompt =
                                BuildEmptyResponseAfterProgressRecoveryPrompt(context)
                            context.PendingGuardTitle = "HOST EMPTY-RESPONSE RECOVERY"
                            context.PendingRejectedTurnExplanation =
                                "Your previous turn was empty after intermediate data was produced, but the requested output file is still missing."
                            context.PendingRejectedAssistantTurn = ""

                            INI_APICall_ToolResponses_2 = BuildToolResponsesForModel(
                                context.AllToolResponses,
                                context.ToolingModel)

                            LogLatestUserRequestDiagnostic(context, "continuation")

                            context.LogWarn(
                                "Empty main-model response after intermediate progress; retrying once with strong missing-output-file recovery prompt.",
                                details:=$"host={context.HostKind}; lastTool={If(lastSuccess.ToolName, "")}")

                            Continue While
                        End If

                        If requiresMissingOutputFileRecovery Then
                            context.EmptyMainModelResponse = True
                            context.FinalizationBlocked = True
                            context.FinalizationBlockedReason = SharedLibrary.Agents.SubAgentRuntimeHardening.ModelEmptyResponseCode
                            currentResponse = Await LocalizeHostMessageIfNeededAsync(
                                "Something went wrong. I could not reliably create the requested output file. Please try again or narrow the request.",
                                ResolveBlockedFallbackUserLanguage(context),
                                useSecondAPI,
                                hideSplash, cancellationToken)

                            If context.SequencingState IsNot Nothing Then
                                context.SequencingState.FinalResponseOrigin = "host_generated"
                                context.SequencingState.HasOpenToolWorkflow = False
                            End If

                            context.LogWarn(
                                "Empty main-model response persisted after strong missing-output-file recovery prompt; returning simple host-generated message.",
                                details:=$"host={context.HostKind}; lastTool={If(lastSuccess.ToolName, "")}")

                            Exit While
                        End If

                        If context.PrematureTextRetryCount < ToolExecutionContext.MaxContinuationRetries Then
                            context.PrematureTextRetryCount += 1
                            context.PendingContinuationGuardPrompt =
                                BuildEmptyResponseAfterProgressRecoveryPrompt(context)
                            context.PendingGuardTitle = "HOST EMPTY-RESPONSE RECOVERY"
                            context.PendingRejectedTurnExplanation =
                                "Your previous turn was empty after successful partial progress."
                            context.PendingRejectedAssistantTurn = ""

                            INI_APICall_ToolResponses_2 = BuildToolResponsesForModel(
                                context.AllToolResponses,
                                context.ToolingModel)

                            LogLatestUserRequestDiagnostic(context, "continuation")

                            context.LogWarn(
                                $"Empty main-model response after successful partial progress; retrying with preserved current request ({context.PrematureTextRetryCount}/{ToolExecutionContext.MaxContinuationRetries}).",
                                details:=$"host={context.HostKind}; lastTool={If(lastSuccess.ToolName, "")}")

                            Continue While
                        End If
                    End If

                    If Not subAgentMode AndAlso
                       lastSuccess IsNot Nothing AndAlso
                       context.PrematureTextRetryCount < ToolExecutionContext.MaxContinuationRetries Then

                        context.PrematureTextRetryCount += 1
                        context.PendingContinuationGuardPrompt =
                            BuildEmptyResponseAfterProgressRecoveryPrompt(context)
                        context.PendingGuardTitle = "HOST EMPTY-RESPONSE RECOVERY"
                        context.PendingRejectedTurnExplanation =
                            "Your previous turn was empty after successful partial progress."
                        context.PendingRejectedAssistantTurn = ""

                        INI_APICall_ToolResponses_2 = BuildToolResponsesForModel(
                            context.AllToolResponses,
                            context.ToolingModel)

                        LogLatestUserRequestDiagnostic(context, "continuation")

                        context.LogWarn(
                            $"Empty main-model response after successful partial progress; retrying with preserved current request ({context.PrematureTextRetryCount}/{ToolExecutionContext.MaxContinuationRetries}).",
                            details:=$"host={context.HostKind}; lastTool={If(lastSuccess.ToolName, "")}")

                        Continue While
                    End If

                    If subAgentMode AndAlso
       lastSuccess IsNot Nothing AndAlso
       context.SubAgentEmptyResponseRetryCount < 1 Then

                        context.SubAgentEmptyResponseRetryCount += 1
                        context.PendingContinuationGuardPrompt = BuildSubAgentEmptyResponseRecoveryPrompt(context)
                        context.PendingGuardTitle = "SUB-AGENT EMPTY-RESPONSE RECOVERY"
                        context.PendingRejectedTurnExplanation =
            "Your previous turn was empty after a successful tool call."
                        context.PendingRejectedAssistantTurn = ""
                        INI_APICall_ToolResponses_2 = BuildToolResponsesForModel(
            context.AllToolResponses,
            context.ToolingModel,
            compactForSubAgent:=True)

                        context.LogWarn(
            $"Empty sub-agent model response after successful tool result; retrying with compact context ({context.SubAgentEmptyResponseRetryCount}/1).",
            details:=$"host={context.HostKind}; lastTool={If(lastSuccess.ToolName, "")}")

                        Continue While
                    End If

                    context.EmptyMainModelResponse = True
                    context.LogWarn("Empty response from LLM", details:="LLM() returned null/empty/whitespace.")
                    ToolingFileLogger.LogWarn("Empty main-model response.",
              details:=$"host={context.HostKind}; iteration={iteration}; subAgentMode={subAgentMode}")
                    Exit While
                End If

                context.Log($"Response received ({currentResponse.Length} chars)")

                Dim detectionPattern = context.ToolingModel.ToolCallDetectionPattern

                If ContainsToolCalls(currentResponse, detectionPattern) Then
                    context.Log("Tool calls detected in response")

                    Dim extractionMap = context.ToolingModel.ToolCallExtractionMap
                    Dim toolCalls = ExtractToolCalls(currentResponse, extractionMap)
                    context.Log($"Extracted {toolCalls.Count} tool call(s)")

                    If toolCalls.Count = 0 Then
                        context.LogWarn(
                            "Tool calls detected but none could be parsed; treating as text response",
                            details:=$"ToolCallExtractionMap='{extractionMap}'")
                        Exit While
                    End If

                    context.Log($"Tool-call batch execution mode: sequential_ordered_no_deferral; total={toolCalls.Count}", "diag")
                    ToolingFileLogger.LogStep($"Executing tool-call batch sequentially in listed order. host={context.HostKind}; count={toolCalls.Count}")

                    Dim stopCurrentBatchAfterTool As Boolean = False
                    Dim restartForRequiredMemoryGrounding As Boolean = False
                    Dim restartAfterToolLoader As Boolean = False
                    Dim restartAfterExposureMiss As Boolean = False
                    Dim turnVisibleToolNames As HashSet(Of String) =
                        BuildTurnVisibleToolNameSet(context.SelectedTools)

                    For Each tc In toolCalls
                        If context.IsCancelled Then Exit For

                        If Not subAgentMode AndAlso
                               context.SequencingState IsNot Nothing AndAlso
                               context.SequencingState.RequiresParentRecovery Then

                            Dim failedToolName As String = context.SequencingState.LastToolName
                            context.SequencingState.NoteRecoveryByLaterToolCall(tc.ToolName)

                            context.Log($"Recovered from prior skipped tool failure by issuing '{tc.ToolName}'.", "success")
                            ToolingFileLogger.LogStep(
                                    $"Skipped agent failure recovered by later parent tool call. host={context.HostKind}; failedTool={failedToolName}; recoveryTool={tc.ToolName}")
                        End If

                        If context.IsCancelled Then Exit For

                        If Not subAgentMode AndAlso
                           SharedLibrary.Agents.ToolCallSequencing.RequiresRequiredMemoryGroundingBeforeNonMemoryTool(
                               context.SequencingState,
                               tc.ToolName) Then

                            If context.PrematureTextRetryCount < ToolExecutionContext.MaxContinuationRetries Then
                                context.PrematureTextRetryCount += 1
                                context.PendingContinuationGuardPrompt =
                                    SharedLibrary.Agents.ToolCallSequencing.BuildRequiredMemoryGroundingRepairPrompt(context.SequencingState)
                                context.PendingGuardTitle = "HOST REQUIRED MEMORY GROUNDING"
                                context.PendingRejectedTurnExplanation =
                                    "Your previous turn attempted a non-memory tool before required Memory grounding was completed."
                                context.PendingRejectedAssistantTurn = If(currentResponse, "")

                                context.LogWarn(
                                    "Blocked non-memory tool before required Memory grounding.",
                                    details:=$"host={context.HostKind}; tool={tc.ToolName}; {SharedLibrary.Agents.ToolCallSequencing.BuildMemoryGroundingStateSummary(context.SequencingState)}")

                                restartForRequiredMemoryGrounding = True
                                Exit For
                            End If

                            context.FinalizationBlocked = True
                            context.FinalizationBlockedReason = SharedLibrary.Agents.ToolCallSequencing.MissingRequiredMemoryAccessCode
                            currentResponse = Await BuildBlockedToolingResultAsync(
                                context,
                                SharedLibrary.Agents.ToolCallSequencing.MissingRequiredMemoryAccessCode,
                                "The tooling run required Memory access before continuing, but Memory was not accessed successfully.", useSecondAPI, hideSplash, cancellationToken)

                            If context.SequencingState IsNot Nothing Then
                                context.SequencingState.FinalResponseOrigin = "host_generated"
                                context.SequencingState.HasOpenToolWorkflow = False
                            End If

                            context.LogWarn(
                                "Blocked tooling loop after repeated attempts to bypass required Memory grounding.",
                                details:=$"host={context.HostKind}; tool={tc.ToolName}; {SharedLibrary.Agents.ToolCallSequencing.BuildMemoryGroundingStateSummary(context.SequencingState)}")
                            Exit For
                        End If

                        context.Log($"Executing tool: {tc.ToolName} (ID: {tc.CallId})")

                        Dim normalizedToolName As String = If(tc.ToolName, "").Trim()
                        Dim toolWasVisibleAtTurnStart As Boolean =
                            turnVisibleToolNames.Contains(normalizedToolName)

                        If Not toolWasVisibleAtTurnStart Then
                            Dim preparedForNextTurn As Boolean =
                                TryPrepareToolForNextTurnAfterExposureMiss(tc.ToolName, context)

                            Dim hiddenResponse As ToolResponse =
                                BuildToolNotExposedThisTurnResponse(tc, turnVisibleToolNames)

                            context.AllToolResponses.Add(hiddenResponse)

                            If context.SequencingState IsNot Nothing Then
                                context.SequencingState.NoteToolFailure(
                                    tc.ToolName,
                                    If(hiddenResponse.ErrorCode, SharedLibrary.Agents.ToolCallSequencing.ToolNotExposedInCurrentTurnCode),
                                    If(hiddenResponse.ErrorMessage, "Tool schema was not exposed at the start of the turn."))
                            End If

                            context.PendingContinuationGuardPrompt =
                                BuildTurnExposureGuardPrompt(tc.ToolName, preparedForNextTurn)
                            context.PendingGuardTitle = "HOST TOOL EXPOSURE GUARD"
                            context.PendingRejectedTurnExplanation =
                                "Your previous turn attempted to call a tool whose full schema was not exposed at the start of that turn."
                            context.PendingRejectedAssistantTurn = If(currentResponse, "")
                            context.PrematureTextRetryCount = 0
                            stopCurrentBatchAfterTool = True
                            restartAfterExposureMiss = preparedForNextTurn

                            context.LogWarn(
                                "Blocked tool because it was not exposed at turn start.",
                                details:=$"host={context.HostKind}; tool={tc.ToolName}; preparedForNextTurn={If(preparedForNextTurn, "true", "false")}")

                            Exit For
                        End If

                        If subAgentMode AndAlso Not IsToolAllowedForCurrentContext(tc.ToolName, context) Then
                            Dim blockedResponse = BuildToolNotAllowedResponse(tc, context)
                            context.AllToolResponses.Add(blockedResponse)

                            If context.SequencingState IsNot Nothing Then
                                context.SequencingState.NoteToolFailure(tc.ToolName,
                                                    If(blockedResponse.ErrorCode, "tool_not_allowed"),
                                                    If(blockedResponse.ErrorMessage, "Tool call was rejected by the sub-agent runtime."))
                            End If

                            stopCurrentBatchAfterTool = True
                            context.LogWarn("Stopping current tool-call batch after blocked tool.",
                        details:=$"host={context.HostKind}; tool={tc.ToolName}; errorCode={If(blockedResponse.ErrorCode, "tool_not_allowed")}")
                            Exit For
                        End If

                        Dim toolConfig = context.SelectedTools.FirstOrDefault(
                            Function(t)
                                Return t IsNot Nothing AndAlso
                                       Not String.IsNullOrWhiteSpace(t.ToolName) AndAlso
                                       t.ToolName.Equals(tc.ToolName, StringComparison.OrdinalIgnoreCase)
                            End Function)

                        If toolConfig Is Nothing Then
                            toolConfig = EnsureVisibleToolLoaded(tc.ToolName, context)
                        End If

                        If toolConfig Is Nothing Then
                            If tc.ToolName.Equals(InternalWebToolName, StringComparison.OrdinalIgnoreCase) Then
                                toolConfig = GetInternalWebTool()
                                ToolingFileLogger.LogStep("Using internal web tool.")

                            ElseIf tc.ToolName.Equals(InternalDownloadWebFilesToolName, StringComparison.OrdinalIgnoreCase) Then
                                toolConfig = GetInternalDownloadWebFilesTool()
                                ToolingFileLogger.LogStep("Using internal web download tool.")

                            ElseIf tc.ToolName.Equals(InternalSearchToolName, StringComparison.OrdinalIgnoreCase) Then
                                toolConfig = GetInternalSearchTool(enforcePrivacy:=INI_EnablePrivacyForSearch)
                                ToolingFileLogger.LogStep("Using internal search tool.")

                            ElseIf IsInternalKnowledgeToolName(tc.ToolName) Then
                                toolConfig = GetInternalKnowledgeTool(tc.ToolName)
                                If toolConfig IsNot Nothing Then
                                    ToolingFileLogger.LogStep("Using store-specific internal knowledge tool.")
                                End If
                            End If

                            If toolConfig Is Nothing Then
                                context.LogError(
                                    $"Unknown tool: {tc.ToolName}",
                                    details:=$"CallId={tc.CallId}; Raw={tc.RawJson}")

                                Dim errorResp As New ToolResponse() With {
                                    .CallId = tc.CallId,
                                    .ToolName = tc.ToolName,
                                    .Success = False,
                                    .ErrorMessage = $"Unknown tool: {tc.ToolName}",
                                    .OriginalCallJson = tc.RawJson
                                }

                                context.AllToolResponses.Add(errorResp)

                                If context.SequencingState IsNot Nothing Then
                                    context.SequencingState.NoteToolFailure(tc.ToolName, "unknown_tool", errorResp.ErrorMessage)
                                End If

                                stopCurrentBatchAfterTool = True
                                Exit For
                            End If
                        End If

                        Dim normalizedToolCallSignature As String = BuildNormalizedToolCallSignature(tc)

                        Dim toolArgumentValidationError As String = ""

                        If Not TryValidateToolCallArguments(tc, toolConfig, toolArgumentValidationError) Then
                            Dim invalidArgsResponse As ToolResponse =
                                BuildInvalidToolArgumentsResponse(tc, toolArgumentValidationError)

                            context.AllToolResponses.Add(invalidArgsResponse)

                            If context.SequencingState IsNot Nothing Then
                                context.SequencingState.NoteToolFailure(
                                    tc.ToolName,
                                    If(invalidArgsResponse.ErrorCode, "invalid_tool_arguments"),
                                    If(invalidArgsResponse.ErrorMessage, "Tool arguments failed schema validation."))
                            End If

                            context.PendingContinuationGuardPrompt =
                                BuildInvalidToolArgumentsGuardPrompt(tc.ToolName, toolArgumentValidationError)
                            context.PendingGuardTitle = "HOST TOOL ARGUMENT VALIDATION"
                            context.PendingRejectedTurnExplanation =
                                "Your previous turn attempted to call a tool with invalid arguments."
                            context.PendingRejectedAssistantTurn = If(currentResponse, "")
                            context.PrematureTextRetryCount = 0
                            stopCurrentBatchAfterTool = True

                            context.LogWarn(
                                "Blocked tool call because arguments failed schema validation.",
                                details:=$"host={context.HostKind}; tool={tc.ToolName}; validationError={toolArgumentValidationError}")

                            Exit For
                        End If

                        Dim toolResponse As ToolResponse = Nothing

                        If TryBuildDuplicateSuccessfulToolReplay(tc, normalizedToolCallSignature, context, toolResponse) Then
                            context.LogWarn(
                                $"Skipped duplicate successful tool call for '{tc.ToolName}' and replayed the prior result.",
                                details:=$"CallId={tc.CallId}; Signature='{normalizedToolCallSignature}'")

                            ToolingFileLogger.LogWarn(
                                "Skipped duplicate successful tool call and replayed prior result.",
                                details:=$"ToolName='{tc.ToolName}'; CallId={tc.CallId}; Signature='{normalizedToolCallSignature}'")
                        Else
                            toolResponse = Await ExecuteToolCall(tc, toolConfig, context, cancellationToken)
                        End If
                        Dim skippedStructuredAgentFailure As Boolean =
                            IsSkippedStructuredAgentFailure(tc, toolResponse, toolConfig, subAgentMode)

                        If Not toolResponse.Success AndAlso String.IsNullOrWhiteSpace(toolResponse.ErrorMessage) Then
                            If Not String.IsNullOrWhiteSpace(toolResponse.Response) Then
                                toolResponse.ErrorMessage = BuildResultExcerpt(toolResponse.Response, 160)
                            Else
                                toolResponse.ErrorMessage = "Tool failed without returning an error message."
                            End If
                        End If

                        toolResponse.OriginalCallJson = tc.RawJson
                        toolResponse.NormalizedCallSignature = normalizedToolCallSignature
                        context.AllToolResponses.Add(toolResponse)

                        If Not toolResponse.WasDuplicateReplay Then
                            SharedLibrary.Agents.ToolCallSequencing.NoteToolExecutionMetadata(
                                context.SequencingState,
                                tc.ToolName,
                                tc.Arguments,
                                toolResponse.Success)

                            If toolResponse.Success Then
                                SharedLibrary.Agents.ToolCallSequencing.NoteToolResultForRepair(
                                    context.SequencingState,
                                    tc.ToolName,
                                    toolResponse.Response,
                                    toolResponse.ResultKind)
                            End If

                            If Not toolResponse.Success Then
                                If context.SequencingState IsNot Nothing Then
                                    context.SequencingState.NoteToolFailure(tc.ToolName,
                                                    If(toolResponse.ErrorCode, ""),
                                                    If(toolResponse.ErrorMessage, "Tool failed."),
                                                    skippedByPolicy:=skippedStructuredAgentFailure,
                                                    returnedToParent:=skippedStructuredAgentFailure)
                                End If
                            Else
                                If context.SequencingState IsNot Nothing Then
                                    context.SequencingState.NoteSuccessfulProgress()
                                End If
                            End If

                            UpdateWorkflowContinuityAfterToolExecution(context, tc, toolResponse)

                            SharedLibrary.Agents.ToolCallSequencing.NoteMemoryGroundingToolResult(
                                context.SequencingState,
                                tc.ToolName,
                                toolResponse.Response,
                                toolResponse.Success)

                            SharedLibrary.Agents.ToolCallSequencing.NoteMemoryGroundingToolResult(
                                context.SequencingState,
                                tc.ToolName,
                                toolResponse.Response,
                                toolResponse.Success)

                            context.Log("Memory grounding state updated: " &
                                SharedLibrary.Agents.ToolCallSequencing.BuildMemoryGroundingStateSummary(context.SequencingState), "diag")
                        Else
                            context.LogWarn(
                                "Duplicate successful tool call was replayed without re-execution.",
                                details:=$"host={context.HostKind}; tool={tc.ToolName}; signature={normalizedToolCallSignature}")
                        End If

                        If toolResponse.WasDuplicateReplay Then
                            context.PendingContinuationGuardPrompt = BuildDuplicateSuccessfulToolCallGuardPrompt(tc.ToolName)
                            context.PendingGuardTitle = "HOST DUPLICATE TOOL CALL GUARD"
                            context.PendingRejectedTurnExplanation =
                                "Your previous tool call had already completed successfully with the same arguments. Use that result instead of calling the tool again."
                            context.PendingRejectedAssistantTurn = ""
                            context.PrematureTextRetryCount = 0
                            stopCurrentBatchAfterTool = True

                            context.LogWarn(
                                "Stopping current tool-call batch after duplicate successful tool replay.",
                                details:=$"host={context.HostKind}; tool={tc.ToolName}; signature={normalizedToolCallSignature}")

                            Exit For
                        End If

                        If toolResponse.Success AndAlso
                           tc.ToolName.Equals(SharedLibrary.Agents.ToolLoaderTool.LoaderToolName, StringComparison.OrdinalIgnoreCase) Then

                            restartAfterToolLoader = True
                            context.PendingContinuationGuardPrompt = BuildToolLoaderBarrierGuardPrompt()
                            context.PendingGuardTitle = "HOST TOOL LOADER BARRIER"
                            context.PendingRejectedTurnExplanation =
                                "The previous turn successfully loaded tools. Newly loaded tool schemas are available only from the next turn onward."
                            context.PendingRejectedAssistantTurn = If(currentResponse, "")
                            context.PrematureTextRetryCount = 0
                            stopCurrentBatchAfterTool = True

                            context.LogWarn(
                                "Stopping current tool-call batch after successful tool_loader call.",
                                details:=$"host={context.HostKind}; tool={tc.ToolName}", visibleToUser:=False)

                            Exit For
                        End If


                        If RegisterToolFailureLoopState(tc, toolResponse, context) Then
                            context.LogWarn(
                                    $"Tool '{tc.ToolName}' failed {context.ConsecutiveFailedToolCount} consecutive times. Forcing a recovery round to reassess the tool choice.")

                            ToolingFileLogger.LogWarn(
                                    "Forcing recovery round after repeated consecutive tool failures.",
                                    details:=$"ToolName='{tc.ToolName}'; Count={context.ConsecutiveFailedToolCount}")

                            context.PendingContinuationGuardPrompt = BuildToolFailureReassessmentGuardPrompt(tc.ToolName)
                            context.PendingGuardTitle = "HOST TOOL FAILURE RECOVERY"
                            context.PendingRejectedTurnExplanation =
                                "Your previous tool step failed repeatedly. Reassess the tool choice before continuing."
                            context.PendingRejectedAssistantTurn = ""
                            context.PrematureTextRetryCount = 0
                            stopCurrentBatchAfterTool = True
                            Exit For
                        End If

                        Dim executedSignature As String = BuildExecutedToolSignature(tc, toolResponse)

                        If String.Equals(context.LastToolExecutionSignature, executedSignature, StringComparison.Ordinal) Then
                            context.LastToolExecutionRepeatCount += 1
                        Else
                            context.LastToolExecutionSignature = executedSignature
                            context.LastToolExecutionRepeatCount = 1
                        End If

                        If context.LastToolExecutionRepeatCount >= context.DuplicateToolExecutionAbortThreshold Then
                            context.LogWarn($"Detected repeated identical tool execution for '{tc.ToolName}'. Forcing a recovery round to reassess the tool choice.")
                            ToolingFileLogger.LogWarn(
                                    "Forcing recovery round after repeated identical tool execution.",
                                    details:=$"ToolName='{tc.ToolName}'; RepeatCount={context.LastToolExecutionRepeatCount}; Signature='{executedSignature}'")
                            context.PendingContinuationGuardPrompt = BuildToolFailureReassessmentGuardPrompt(tc.ToolName)
                            context.PendingGuardTitle = "HOST TOOL FAILURE RECOVERY"
                            context.PendingRejectedTurnExplanation =
                                "The same tool step repeated without progress. Reassess the tool choice before continuing."
                            context.PendingRejectedAssistantTurn = ""
                            context.PrematureTextRetryCount = 0
                            stopCurrentBatchAfterTool = True
                            Exit For
                        End If

                        If Not toolResponse.Success Then
                            context.LogError(
        $"Tool error ({tc.ToolName}): {toolResponse.ErrorMessage}",
        details:=$"CallId={tc.CallId}; RawCall={tc.RawJson}")

                            Select Case toolConfig.ToolErrorHandling?.ToLowerInvariant()
                                Case "abort"
                                    context.LogError("Aborting due to tool error (ToolErrorHandling=abort)")
                                    ShowCustomMessageBox($"Tool execution failed: {toolResponse.ErrorMessage}")
                                    ToolingFileLogger.EndSession(False, $"Tool error: {toolResponse.ErrorMessage}")
                                    Return ""

                                Case "retry"
                                    context.LogWarn("Will retry on next iteration (ToolErrorHandling=retry)")
                                    context.PendingContinuationGuardPrompt = BuildToolFailureReassessmentGuardPrompt(tc.ToolName)
                                    context.PendingGuardTitle = "HOST TOOL FAILURE RECOVERY"
                                    context.PendingRejectedTurnExplanation =
                                        "Your previous tool step failed. Reassess the tool choice before continuing."
                                    context.PendingRejectedAssistantTurn = ""
                                    context.PrematureTextRetryCount = 0

                                Case Else
                                    context.LogWarn("Skipping tool error (ToolErrorHandling=skip)")

                                    If Not skippedStructuredAgentFailure Then
                                        context.PendingContinuationGuardPrompt = BuildToolFailureReassessmentGuardPrompt(tc.ToolName)
                                        context.PendingGuardTitle = "HOST TOOL FAILURE RECOVERY"
                                        context.PendingRejectedTurnExplanation =
                                            "Your previous tool step failed. Reassess the tool choice before continuing."
                                        context.PendingRejectedAssistantTurn = ""
                                        context.PrematureTextRetryCount = 0
                                    End If

                                    If skippedStructuredAgentFailure Then
                                        context.PendingContinuationGuardPrompt = BuildSkippedToolFailureRecoveryGuardPrompt()
                                        context.PendingGuardTitle = "HOST TOOL FAILURE GUARD"
                                        context.PendingRejectedTurnExplanation =
                                            "Your previous turn was rejected because the prior skipped tool failure is still unresolved."
                                        context.PendingRejectedAssistantTurn = ""
                                        context.PrematureTextRetryCount = 0

                                        context.LogWarn(
                                            "Structured agent failure returned to parent as tool response and skipped by policy.",
                                            details:=$"host={context.HostKind}; tool={tc.ToolName}; errorCode={If(toolResponse.ErrorCode, "")}")

                                        ToolingFileLogger.LogWarn(
                                            "Agent failure returned to parent as structured error and skipped by policy.",
                                            details:=$"host={context.HostKind}; tool={tc.ToolName}; errorCode={If(toolResponse.ErrorCode, "")}; returnedToParent=true; skipped=true")
                                    End If
                            End Select

                            stopCurrentBatchAfterTool = True
                        Else
                            context.LastInvalidAssistantTurnSignature = ""
                            context.LastInvalidTurnReason = ""
                            context.LastInvalidTurnRepeatCount = 0
                            context.ForceNoToolFinalizationRequested = False
                            context.ForceNoToolFinalizationReason = ""

                            context.PrematureTextRetryCount = 0
                            context.Log($"Tool completed successfully ({toolResponse.Response?.Length} chars)", "success")
                        End If

                        If stopCurrentBatchAfterTool Then
                            context.LogWarn("Stopping current tool-call batch after failure.",
                    details:=$"host={context.HostKind}; tool={tc.ToolName}; errorCode={If(toolResponse.ErrorCode, "")}")
                            Exit For
                        End If
                    Next

                    If restartAfterExposureMiss Then
                        Dim preparedToolResponsesAfterExposureMiss As String = BuildToolResponsesForModel(
                            context.AllToolResponses,
                            context.ToolingModel,
                            compactForSubAgent:=subAgentMode)

                        INI_APICall_ToolResponses_2 = preparedToolResponsesAfterExposureMiss
                        context.Log("Tool was prepared after exposure miss; restarting next iteration with refreshed tool definitions.", "diag")
                        Continue While
                    End If

                    If restartAfterToolLoader Then
                        Dim preparedToolResponsesAfterLoader As String = BuildToolResponsesForModel(
                            context.AllToolResponses,
                            context.ToolingModel,
                            compactForSubAgent:=subAgentMode)

                        INI_APICall_ToolResponses_2 = preparedToolResponsesAfterLoader
                        context.Log("tool_loader completed; refreshed tool definitions will be exposed on the next iteration.", "diag")
                        Continue While
                    End If

                    If restartForRequiredMemoryGrounding Then
                        Dim preparedToolResponses As String = BuildToolResponsesForModel(
                            context.AllToolResponses,
                            context.ToolingModel,
                            compactForSubAgent:=subAgentMode)
                        INI_APICall_ToolResponses_2 = preparedToolResponses
                        context.Log("Tool responses prepared for next iteration", "diag")
                        Continue While
                    End If

                    If context.FinalizationBlocked Then
                        Exit While
                    End If

                    If abortDueToRepeatedToolLoop Then
                        Exit While
                    End If

                    Dim toolResponses = BuildToolResponsesForModel(
    context.AllToolResponses,
    context.ToolingModel,
    compactForSubAgent:=subAgentMode)
                    INI_APICall_ToolResponses_2 = toolResponses
                    context.Log("Tool responses prepared for next iteration", "diag")
                Else
                    If context.FinalResponseContract = SharedLibrary.Agents.ToolingFinalResponseContract.RawCallerText Then
                        context.PrematureTextRetryCount = 0

                        If context.SequencingState IsNot Nothing Then
                            context.SequencingState.HasOpenToolWorkflow = False
                            context.SequencingState.FinalResponseOrigin = "model_provided"
                        End If

                        context.Log($"Final raw-text response accepted for caller-defined non-user-facing contract ({currentResponse.Length} chars).")
                        Exit While
                    End If

                    If context.SequencingState IsNot Nothing Then
                        context.SequencingState.FinalCompleteRejectedForMissingMemoryAccess = False
                    End If

                    Dim turnValidation = SharedLibrary.Agents.ToolCallSequencing.ValidateActiveToolingTurn(
                        currentResponse,
                        hasToolCalls:=False,
                        hasUnresolvedToolFailure:=SharedLibrary.Agents.ToolCallSequencing.HasBlockingUnresolvedToolFailure(context.SequencingState),
                        runState:=context.SequencingState)

                    If context.SequencingState IsNot Nothing Then
                        context.SequencingState.LastDetectedTurnType = turnValidation.TurnKind.ToString()
                        context.SequencingState.LastInvalidTurnReason = If(turnValidation.InvalidReason, "")
                        context.SequencingState.FinalCompleteRejectedForMissingMemoryAccess =
                            SharedLibrary.Agents.ToolCallSequencing.IsMemoryGroundingRejectionReason(
                                turnValidation.InvalidReason)

                    End If

                    context.Log(
                        $"activeToolingSession={If(context.SequencingState IsNot Nothing AndAlso context.SequencingState.ActiveToolingSession, "true", "false")}; detectedModelTurnType={turnValidation.TurnKind}; taskStatusParseResult={turnValidation.TaskStatusSummary}; toolRequiredModeUsed={If(context.SequencingState IsNot Nothing AndAlso context.SequencingState.ToolRequiredModeUsed, "true", "false")}; {SharedLibrary.Agents.ToolCallSequencing.BuildMemoryGroundingStateSummary(context.SequencingState)}", "diag")

                    Select Case turnValidation.TurnKind
                        Case SharedLibrary.Agents.ToolCallSequencing.ActiveToolingTurnKind.FinalCompleteTurn
                            If Not SharedLibrary.Agents.ToolCallSequencing.IsRequiredMemoryGroundingSatisfied(
                                context.SequencingState,
                                turnValidation.TurnKind) Then

                                If context.SequencingState IsNot Nothing Then
                                    context.SequencingState.FinalCompleteRejectedForMissingMemoryAccess = True
                                End If

                                If context.PrematureTextRetryCount < ToolExecutionContext.MaxContinuationRetries Then
                                    context.PrematureTextRetryCount += 1
                                    context.PendingContinuationGuardPrompt = SharedLibrary.Agents.ToolCallSequencing.BuildRequiredMemoryGroundingRepairPrompt(context.SequencingState)
                                    context.PendingRejectedAssistantTurn = If(currentResponse, "")
                                    context.PendingGuardTitle = "HOST REQUIRED MEMORY GROUNDING REPAIR"
                                    context.PendingRejectedTurnExplanation =
                                        "Your previous turn was rejected because the current run explicitly required Memory access before finalization."

                                    context.LogWarn(
                                        $"Final complete turn rejected for missing required memory access. {SharedLibrary.Agents.ToolCallSequencing.BuildMemoryGroundingStateSummary(context.SequencingState)}; repairAttempt={context.PrematureTextRetryCount}/{ToolExecutionContext.MaxContinuationRetries}")

                                    Continue While
                                End If

                                context.FinalizationBlocked = True
                                context.FinalizationBlockedReason = SharedLibrary.Agents.ToolCallSequencing.MissingRequiredMemoryAccessCode
                                currentResponse = Await BuildBlockedToolingResultAsync(
                                    context,
                                    SharedLibrary.Agents.ToolCallSequencing.MissingRequiredMemoryAccessCode,
                                    "The tooling run required Memory access before finalization, but Memory was not accessed successfully.", useSecondAPI, hideSplash, cancellationToken)

                                If context.SequencingState IsNot Nothing Then
                                    context.SequencingState.FinalResponseOrigin = "host_generated"
                                    context.SequencingState.HasOpenToolWorkflow = False
                                End If

                                context.LogWarn(
                                    $"Active-tooling final complete turn blocked for missing required memory access. {SharedLibrary.Agents.ToolCallSequencing.BuildMemoryGroundingStateSummary(context.SequencingState)}")

                                Exit While
                            End If

                            context.PrematureTextRetryCount = 0

                            If context.SequencingState IsNot Nothing Then
                                context.SequencingState.FinalCompleteRejectedForMissingMemoryAccess = False
                                context.SequencingState.HasOpenToolWorkflow = False
                                context.SequencingState.FinalResponseOrigin = "model_provided"
                            End If

                            ' === Final-turn gate (P3a + P3b) — applies to 'complete' branch.
                            ' Even 'complete' must not announce future work (contract rule 3).
                            Dim _ftGateUntried_C As IReadOnlyList(Of String) = CollectUntriedDeliverableFallbackToolNames(context)
                            Dim _ftGateNeedsDeliverable_C As Boolean =
                                    context.SequencingState IsNot Nothing AndAlso
                                    context.SequencingState.RequestRequiresCreatedDeliverable
                            Dim _ftGateHasDeliverable_C As Boolean =
                                    context.SequencingState IsNot Nothing AndAlso
                                    SharedLibrary.Agents.ToolCallSequencing.HasProducedUserDeliverable(context.SequencingState)

                            Dim _ftGateEval_C As Agents.FinalTurnEvaluation =
                                    Agents.ToolingOrchestrator.EvaluateFinalTurn(
                                        currentResponse,
                                        _ftGateNeedsDeliverable_C,
                                        _ftGateHasDeliverable_C,
                                        _ftGateUntried_C)

                            If _ftGateEval_C.Decision <> Agents.FinalTurnDecision.Accept Then
                                If System.Convert.ToInt32(context.PrematureTextRetryCount) < ToolExecutionContext.MaxContinuationRetries Then
                                    context.PendingContinuationGuardPrompt = _ftGateEval_C.GuardPrompt
                                    context.PendingGuardTitle = _ftGateEval_C.GuardTitle
                                    context.PendingRejectedTurnExplanation = _ftGateEval_C.Reason
                                    context.PendingRejectedAssistantTurn = currentResponse
                                    context.PrematureTextRetryCount = System.Convert.ToInt32(context.PrematureTextRetryCount) + 1
                                    context.Log("Final-turn rejected at 'complete' branch: " & _ftGateEval_C.Reason, "warn")
                                    Continue While
                                End If
                                context.LogWarn(
                                        "Final-turn repair budget exhausted at 'complete' branch; accepting candidate.",
                                        details:="reason=" & _ftGateEval_C.Reason)
                            End If

                            SharedLibrary.Agents.ToolCallSequencing.ClearNonBlockingUnresolvedToolFailure(
                                context.SequencingState,
                                "host_exposure_recovery")

                            currentResponse = StripTaskStatus(currentResponse)
                            acceptedFinalStatus = "complete"
                            context.Log("Final complete response accepted.")
                            Exit While

                        Case SharedLibrary.Agents.ToolCallSequencing.ActiveToolingTurnKind.FinalBlockedTurn
                            context.PrematureTextRetryCount = 0

                            If context.SequencingState IsNot Nothing Then
                                If context.SequencingState.HasUnresolvedToolFailure Then
                                    context.SequencingState.NoteBlockedFinalHandled()
                                End If

                                context.SequencingState.HasOpenToolWorkflow = False
                                context.SequencingState.FinalResponseOrigin = "model_provided"
                            End If

                            ' === Final-turn gate (P3a + P3b) — applies to 'blocked' branch.
                            ' 'blocked' must not pair with prose announcing a fallback (rule 8), and must
                            ' only be declared after authorized fallback tools have been attempted (rule 9).
                            Dim _ftGateUntried_B As IReadOnlyList(Of String) = CollectUntriedDeliverableFallbackToolNames(context)
                            Dim _ftGateNeedsDeliverable_B As Boolean =
                                    context.SequencingState IsNot Nothing AndAlso
                                    context.SequencingState.RequestRequiresCreatedDeliverable
                            Dim _ftGateHasDeliverable_B As Boolean =
                                    context.SequencingState IsNot Nothing AndAlso
                                    SharedLibrary.Agents.ToolCallSequencing.HasProducedUserDeliverable(context.SequencingState)

                            Dim _ftGateEval_B As Agents.FinalTurnEvaluation =
                                    Agents.ToolingOrchestrator.EvaluateFinalTurn(
                                        currentResponse,
                                        _ftGateNeedsDeliverable_B,
                                        _ftGateHasDeliverable_B,
                                        _ftGateUntried_B)

                            If _ftGateEval_B.Decision <> Agents.FinalTurnDecision.Accept Then
                                If System.Convert.ToInt32(context.PrematureTextRetryCount) < ToolExecutionContext.MaxContinuationRetries Then
                                    context.PendingContinuationGuardPrompt = _ftGateEval_B.GuardPrompt
                                    context.PendingGuardTitle = _ftGateEval_B.GuardTitle
                                    context.PendingRejectedTurnExplanation = _ftGateEval_B.Reason
                                    context.PendingRejectedAssistantTurn = currentResponse
                                    context.PrematureTextRetryCount = System.Convert.ToInt32(context.PrematureTextRetryCount) + 1
                                    context.Log("Final-turn rejected at 'blocked' branch: " & _ftGateEval_B.Reason, "warn")
                                    Continue While
                                End If
                                context.LogWarn(
                                        "Final-turn repair budget exhausted at 'blocked' branch; accepting candidate.",
                                        details:="reason=" & _ftGateEval_B.Reason)
                            End If

                            SharedLibrary.Agents.ToolCallSequencing.ClearNonBlockingUnresolvedToolFailure(
                                context.SequencingState,
                                "host_exposure_recovery")

                            currentResponse = StripTaskStatus(currentResponse)
                            acceptedFinalStatus = "blocked"
                            context.Log("Final blocked response accepted.")
                            Exit While

                        Case Else
                            Dim memoryGroundingRepairRequired As Boolean =
                                SharedLibrary.Agents.ToolCallSequencing.IsMemoryGroundingRejectionReason(
                                    turnValidation.InvalidReason)

                            Dim invalidTurnSignature As String =
                                BuildInvalidTurnSignature(
                                    currentResponse,
                                    turnValidation.InvalidReason)

                            If String.Equals(
                                context.LastInvalidAssistantTurnSignature,
                                invalidTurnSignature,
                                StringComparison.Ordinal) Then

                                context.LastInvalidTurnRepeatCount += 1
                            Else
                                context.LastInvalidAssistantTurnSignature = invalidTurnSignature
                                context.LastInvalidTurnReason = If(turnValidation.InvalidReason, "")
                                context.LastInvalidTurnRepeatCount = 1
                            End If

                            Dim repeatedInvalidTurnAfterSuccessfulTool As Boolean =
                                Not memoryGroundingRepairRequired AndAlso
                                HasSuccessfulToolResponses(context) AndAlso
                                context.LastInvalidTurnRepeatCount >= 2

                            If repeatedInvalidTurnAfterSuccessfulTool Then
                                context.ForceNoToolFinalizationRequested = True
                                context.ForceNoToolFinalizationReason = If(turnValidation.InvalidReason, "invalid_turn")

                                context.LogWarn(
                                    $"Repeated identical invalid active-tooling turn detected; escalating to forced no-tool finalization. invalidTurnReason={context.ForceNoToolFinalizationReason}; repeatCount={context.LastInvalidTurnRepeatCount}; signature={invalidTurnSignature}; host={context.HostKind}")

                                Exit While
                            End If

                            If context.PrematureTextRetryCount < ToolExecutionContext.MaxContinuationRetries Then
                                context.PrematureTextRetryCount += 1
                                context.PendingContinuationGuardPrompt =
                                    If(
                                        memoryGroundingRepairRequired,
                                        SharedLibrary.Agents.ToolCallSequencing.BuildRequiredMemoryGroundingRepairPrompt(context.SequencingState),
                                        SharedLibrary.Agents.ToolCallSequencing.BuildActiveToolingRepairPrompt(
                                            context.SequencingState,
                                            turnValidation.InvalidReason))
                                context.PendingRejectedAssistantTurn = If(currentResponse, "")
                                context.PendingGuardTitle =
                                    If(
                                        memoryGroundingRepairRequired,
                                        "HOST REQUIRED MEMORY GROUNDING REPAIR",
                                        "HOST ACTIVE TOOLING CONTRACT REPAIR")
                                context.PendingRejectedTurnExplanation =
                                    If(
                                        memoryGroundingRepairRequired,
                                        "Your previous turn was rejected because the current run required a further Memory-grounding step before finalization.",
                                        "Your previous turn was rejected because it violated the active-tooling response contract.")

                                context.LogWarn(
                                    $"Invalid active-tooling turn rejected; invalidTurnReason={If(turnValidation.InvalidReason, "invalid_turn")}; repairAttempt={context.PrematureTextRetryCount}/{ToolExecutionContext.MaxContinuationRetries}; {SharedLibrary.Agents.ToolCallSequencing.BuildMemoryGroundingStateSummary(context.SequencingState)}; lastSuccessfulTool={If(context.SequencingState?.LastSuccessfulToolCall, "")}; unresolvedToolFailure={If(context.SequencingState IsNot Nothing AndAlso context.SequencingState.HasUnresolvedToolFailure, "true", "false")}; host={context.HostKind}")

                                Continue While
                            End If

                            context.FinalizationBlocked = True
                            context.FinalizationBlockedReason = If(turnValidation.InvalidReason, "invalid_turn")
                            currentResponse = Await BuildBlockedToolingResultAsync(
                                context,
                                If(
                                    memoryGroundingRepairRequired,
                                    If(turnValidation.InvalidReason, SharedLibrary.Agents.ToolCallSequencing.MissingRequiredMemoryAccessCode),
                                    SharedLibrary.Agents.ToolCallSequencing.InvalidTextOnlyFinalizationCode),
                                If(
                                    memoryGroundingRepairRequired,
                                    "The tooling run required Memory access before finalization, but the required Memory-grounding step was not completed successfully.",
                                    "The tooling run ended because the model did not return a valid next tool call or a valid final status."), useSecondAPI, hideSplash, cancellationToken)

                            If context.SequencingState IsNot Nothing Then
                                context.SequencingState.FinalResponseOrigin = "host_generated"
                                context.SequencingState.HasOpenToolWorkflow = False
                            End If

                            context.LogWarn(
                                $"Active-tooling repair exhausted; returning host-generated blocked message. invalidTurnReason={If(turnValidation.InvalidReason, "invalid_turn")}; {SharedLibrary.Agents.ToolCallSequencing.BuildMemoryGroundingStateSummary(context.SequencingState)}; unresolvedToolFailure={If(context.SequencingState IsNot Nothing AndAlso context.SequencingState.HasUnresolvedToolFailure, "true", "false")}; host={context.HostKind}")

                            Exit While
                    End Select
                End If

            End While

            If SharedLibrary.Agents.ToolingFinalResponseContractHelpers.RequiresTaskStatusFooter(context.FinalResponseContract) AndAlso
               context.ForceNoToolFinalizationRequested AndAlso
               Not context.IsCancelled AndAlso
               Not context.FinalizationBlocked Then

                Dim forcedFinalResponse As String =
                    Await RunForcedNoToolFinalizationAfterRepeatedInvalidTurnAsync(
                        context,
                        enhancedSysPrompt,
                        fullUserPrompt,
                        otherPrompt,
                        fileObject,
                        useSecondAPI,
                        hideSplash)

                If Not String.IsNullOrWhiteSpace(forcedFinalResponse) Then
                    currentResponse = forcedFinalResponse

                    Dim forcedValidation = SharedLibrary.Agents.ToolCallSequencing.ValidateActiveToolingTurn(
                        currentResponse,
                        hasToolCalls:=False,
                        hasUnresolvedToolFailure:=context.SequencingState IsNot Nothing AndAlso context.SequencingState.HasUnresolvedToolFailure,
                        runState:=context.SequencingState)

                    Select Case forcedValidation.TurnKind
                        Case SharedLibrary.Agents.ToolCallSequencing.ActiveToolingTurnKind.FinalCompleteTurn
                            acceptedFinalStatus = "complete"

                            If context.SequencingState IsNot Nothing Then
                                context.SequencingState.FinalResponseOrigin = "model_provided"
                                context.SequencingState.HasOpenToolWorkflow = False
                            End If

                            SharedLibrary.Agents.ToolCallSequencing.ClearNonBlockingUnresolvedToolFailure(
                                context.SequencingState,
                                "host_exposure_recovery")

                            context.Log("Forced no-tool final response accepted as complete.")

                        Case SharedLibrary.Agents.ToolCallSequencing.ActiveToolingTurnKind.FinalBlockedTurn
                            acceptedFinalStatus = "blocked"

                            If context.SequencingState IsNot Nothing Then
                                context.SequencingState.FinalResponseOrigin = "model_provided"
                                context.SequencingState.HasOpenToolWorkflow = False
                            End If

                            SharedLibrary.Agents.ToolCallSequencing.ClearNonBlockingUnresolvedToolFailure(
                                context.SequencingState,
                                "host_exposure_recovery")

                            context.Log("Forced no-tool final response accepted as blocked.")

                        Case Else
                            context.FinalizationBlocked = True
                            context.FinalizationBlockedReason = If(forcedValidation.InvalidReason, "invalid_forced_final")

                            currentResponse = Await BuildBlockedToolingResultAsync(
                                context,
                                SharedLibrary.Agents.ToolCallSequencing.InvalidTextOnlyFinalizationCode,
                                "The tooling run ended because the forced finalization response was still not valid.",
                                useSecondAPI,
                                hideSplash,
                                cancellationToken)

                            If context.SequencingState IsNot Nothing Then
                                context.SequencingState.FinalResponseOrigin = "host_generated"
                                context.SequencingState.HasOpenToolWorkflow = False
                            End If

                            context.LogWarn(
                                $"Forced no-tool finalization returned an invalid turn; invalidTurnReason={If(forcedValidation.InvalidReason, "invalid_turn")}; host={context.HostKind}")
                    End Select
                Else
                    context.EmptyMainModelResponse = True
                End If
            End If

            ' If we hit max iterations and the last response is either a pending tool call or is not
            ' presentable user-facing text (e.g. a raw JSON/structured tool echo, or empty), force a
            ' final text response so the user never receives raw protocol content.
            ' The tool results are already in INI_APICall_ToolResponses_2 from the last iteration.
            If iteration >= context.MaxIterations AndAlso
               Not context.IsCancelled AndAlso
               (ContainsToolCalls(currentResponse, context.ToolingModel.ToolCallDetectionPattern) OrElse
                Not SharedLibrary.Agents.ToolCallSequencing.IsUserPresentableFinalText(currentResponse)) Then

                context.Log("Forcing final response (max iterations reached with a pending tool call or non-user-facing content)...")

                ' Disable tool definitions to prevent further tool calls
                INI_APICall_ToolInstructions_2 = ""

                ' Append instruction to force synthesis
                Dim finalSysPrompt As String =
                    enhancedSysPrompt & Environment.NewLine & Environment.NewLine &
                    If(
                        SharedLibrary.Agents.ToolingFinalResponseContractHelpers.RequiresTaskStatusFooter(context.FinalResponseContract),
                        "IMPORTANT: You have reached the maximum number of tool iterations. Do NOT call any more tools. Based on all the information gathered from the tools so far, provide your final answer now.",
                        "IMPORTANT: You have reached the maximum number of tool iterations. Do NOT call any more tools. Based on all the information gathered from the tools so far, return only the final raw text payload in the exact caller-defined format.")

                ToolingFileLogger.LogStep("Forcing final LLM call without tools")
                ToolingFileLogger.LogPreMainLlmCallSnapshot()

                Try
                    Dim finalResponse As String = Await LLM(
                        finalSysPrompt,
                        fullUserPrompt,
                        "", "", 0,
                        useSecondAPI,
                        hideSplash,
                        otherPrompt,
                        fileObject,
                        True)

                    If Not String.IsNullOrWhiteSpace(finalResponse) Then
                        currentResponse = finalResponse
                        context.Log($"Final response received ({currentResponse.Length} chars)")
                        ToolingFileLogger.LogRawResponseStub("Main LLM() - Forced Final", currentResponse)
                    Else
                        context.EmptyMainModelResponse = True
                        context.LogWarn("Empty response from forced final LLM call")
                        ToolingFileLogger.LogWarn("Empty forced-final main-model response.",
                              details:=$"host={context.HostKind}; iteration={iteration}")
                    End If

                Catch ex As Exception
                    context.LogError($"Error during forced final call: {ex.Message}", ex:=ex)
                End Try
            End If

            If context.IsCancelled Then
                context.LogWarn("Session cancelled by user")
                ShowCustomMessageBox("Tooling session was cancelled.")
                ToolingFileLogger.EndSession(False, "Cancelled by user")
                Return ""
            End If

            If iteration >= context.MaxIterations Then
                context.LogWarn($"Maximum iterations ({context.MaxIterations}) reached")
                ShowCustomMessageBox($"Maximum tool iterations ({context.MaxIterations}) reached. The response may be incomplete.")
                ToolingFileLogger.LogWarn("Maximum iterations reached.", details:=$"MaxIterations={context.MaxIterations}")

                ' Guarantee user-understandable output: if the (possibly forced) final response is not
                ' presentable user-facing text, summarize what was achieved as a blocked result instead
                ' of returning raw tool/JSON content to the user.
                If Not context.FinalizationBlocked AndAlso
                   Not SharedLibrary.Agents.ToolCallSequencing.IsUserPresentableFinalText(currentResponse) Then

                    context.FinalizationBlocked = True
                    context.FinalizationBlockedReason = "max_iterations_no_user_facing_final"

                    currentResponse = Await BuildBlockedToolingResultAsync(
                        context,
                        SharedLibrary.Agents.ToolCallSequencing.InvalidTextOnlyFinalizationCode,
                        "The tool-assisted run reached its iteration limit before producing a user-facing result.",
                        useSecondAPI,
                        hideSplash,
                        cancellationToken)

                    If context.SequencingState IsNot Nothing Then
                        context.SequencingState.FinalResponseOrigin = "host_generated"
                        context.SequencingState.HasOpenToolWorkflow = False
                    End If

                    context.LogWarn(
                        "Max-iteration final response was not user-facing; converted to a summarized blocked message.",
                        details:=$"host={context.HostKind}")
                End If
            End If

            If context.EmptyMainModelResponse AndAlso String.IsNullOrWhiteSpace(currentResponse) Then
                context.FinalizationBlocked = True
                context.FinalizationBlockedReason = SharedLibrary.Agents.SubAgentRuntimeHardening.ModelEmptyResponseCode
                currentResponse = Await BuildBlockedToolingResultAsync(
                    context,
                    SharedLibrary.Agents.SubAgentRuntimeHardening.ModelEmptyResponseCode,
                    "The tooling run ended because the model returned no valid content.", useSecondAPI, hideSplash, cancellationToken)

                If context.SequencingState IsNot Nothing Then
                    context.SequencingState.FinalResponseOrigin = "host_generated"
                    context.SequencingState.HasOpenToolWorkflow = False
                End If

                context.LogWarn("Empty main-model response converted to a user-safe blocked message.",
                    details:=$"host={context.HostKind}")
            End If

            If Not context.FinalizationBlocked AndAlso
               context.SequencingState IsNot Nothing AndAlso
               SharedLibrary.Agents.ToolCallSequencing.HasBlockingUnresolvedToolFailure(context.SequencingState) AndAlso
               Not context.SequencingState.RequiresParentRecovery Then

                context.FinalizationBlocked = True
                context.FinalizationBlockedReason = "unresolved_tool_failure"
                currentResponse = Await BuildBlockedToolingResultAsync(
                            context,
                            SharedLibrary.Agents.ToolCallSequencing.UnresolvedToolFailureCode,
                            "The tooling run ended with an unresolved tool failure.", useSecondAPI, hideSplash, cancellationToken)
                context.LogWarn("Finalization blocked because unresolved tool failure remained.",
                                        details:=$"host={context.HostKind}; tool={context.SequencingState.LastToolName}; errorCode={context.SequencingState.LastErrorCode}")
            End If

            context.Log("=== Session Summary ===")
            context.Log($"Last successful tool: {If(context.SequencingState?.LastSuccessfulToolCall, "(none)")}")
            context.Log($"Unresolved tool failure: {If(context.SequencingState IsNot Nothing AndAlso context.SequencingState.HasUnresolvedToolFailure, "true", "false")}")
            context.Log($"Final response origin: {If(context.SequencingState IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(context.SequencingState.FinalResponseOrigin), context.SequencingState.FinalResponseOrigin, "model_provided")}")
            context.Log($"Total iterations: {iteration}")
            context.Log($"Total tool calls: {context.AllToolResponses.Count}")
            Dim successCount As Integer = context.AllToolResponses.Where(Function(r) r.Success).Count()
            Dim failedCount As Integer = context.AllToolResponses.Where(Function(r) Not r.Success).Count()
            context.Log($"Successful: {successCount}", If(failedCount = 0, "success", "step"))
            context.Log($"Failed: {failedCount}", If(failedCount = 0, "step", "warn"))

            If SharedLibrary.Agents.ToolingFinalResponseContractHelpers.RequiresTaskStatusFooter(context.FinalResponseContract) Then
                currentResponse =
                    SharedLibrary.Agents.ToolCallSequencing.StripTaskStatusBlocksFromUserFacingText(
                        StripTaskStatus(currentResponse))

                If Not context.FinalizationBlocked AndAlso
                   Not String.IsNullOrWhiteSpace(acceptedFinalStatus) AndAlso
                   Not SharedLibrary.Agents.ToolCallSequencing.HasSubstantiveUserFacingText(currentResponse) Then

                    Dim previousAcceptedFinalStatus As String = acceptedFinalStatus

                    context.FinalizationBlocked = True
                    context.FinalizationBlockedReason = "non_user_facing_final_text"

                    currentResponse = Await BuildBlockedToolingResultAsync(
                        context,
                        SharedLibrary.Agents.ToolCallSequencing.InvalidTextOnlyFinalizationCode,
                        "The tooling run ended because the accepted final response did not contain valid user-facing text.",
                        useSecondAPI,
                        hideSplash,
                        cancellationToken)

                    If context.SequencingState IsNot Nothing Then
                        context.SequencingState.FinalResponseOrigin = "host_generated"
                        context.SequencingState.HasOpenToolWorkflow = False
                    End If

                    acceptedFinalStatus = "blocked"

                    context.LogWarn(
                        "Accepted final response contained no substantive user-facing text after stripping the TASK_STATUS footer; converted to blocked message.",
                        details:=$"host={context.HostKind}; previousFinalStatus={previousAcceptedFinalStatus}")
                End If

                ' Post-egress localization of final prose when it does not match the
                ' detected user language. Sub-agents never localize.
                If Not subAgentMode Then
                    Dim _userLanguage As String =
                        If(context.SequencingState IsNot Nothing, context.SequencingState.UserLanguage, "")

                    If Agents.ToolingOrchestrator.ShouldPostLocalizeFinal(
                        currentResponse,
                        _userLanguage,
                        acceptedFinalStatus) Then

                        Try
                            currentResponse = Await LocalizeHostMessageIfNeededAsync(
                                currentResponse,
                                _userLanguage,
                                useSecondAPI,
                                hideSplash,
                                cancellationToken)
                        Catch ex As Exception
                            context.LogWarn(
                                "Final-response localization failed; returning original prose.",
                                details:=ex.Message)
                        End Try
                    End If
                End If

                currentResponse = AppendM365SourcesFooter(currentResponse, context.AllToolResponses)
            Else
                context.Log("Returning raw caller-defined final response without user-facing post-processing.", "diag")
            End If

            If Not subAgentMode Then
                Dim isBlockedFinal As Boolean =
                    context.FinalizationBlocked OrElse
                    context.EmptyMainModelResponse OrElse
                    String.Equals(acceptedFinalStatus, "blocked", StringComparison.OrdinalIgnoreCase)

                Dim finalCheckpointWritten As Boolean =
                    SharedLibrary.Agents.WorkflowContinuity.NoteFinalStatus(
                        context.WorkflowId,
                        context.HostKind,
                        isBlockedFinal)

                context.RuntimeState = SharedLibrary.Agents.WorkflowContinuity.GetState(context.WorkflowId)
                context.Log($"Workflow final status recorded: {If(isBlockedFinal, "blocked", "complete")}; checkpointWritten={If(finalCheckpointWritten, "true", "false")}", "diag")
            End If

            Dim finalResponseLength As Integer = If(currentResponse, "").Length
            Dim finalResponseOrigin As String =
    If(context.SequencingState IsNot Nothing AndAlso
       Not String.IsNullOrWhiteSpace(context.SequencingState.FinalResponseOrigin),
       context.SequencingState.FinalResponseOrigin,
       "model_provided")

            context.Log(
    $"Returning final response to host: len={finalResponseLength}; origin={finalResponseOrigin}; blocked={If(context.FinalizationBlocked, "true", "false")}; empty={If(context.EmptyMainModelResponse, "true", "false")}",
    "diag")

            ToolingFileLogger.LogStep(
    $"Host final response prepared: len={finalResponseLength}; origin={finalResponseOrigin}; blocked={If(context.FinalizationBlocked, "true", "false")}; empty={If(context.EmptyMainModelResponse, "true", "false")}")

            Dim sessionSucceeded As Boolean = (Not context.EmptyMainModelResponse) AndAlso (Not context.FinalizationBlocked)
            ToolingFileLogger.EndSession(sessionSucceeded, $"Iterations: {iteration}, Tool calls: {context.AllToolResponses.Count}, Success: {successCount}, Failed: {failedCount}")

            Return currentResponse

        Catch ex As Exception
            context.LogError($"Error in tooling loop: {ex.Message}", ex:=ex)
            ShowCustomMessageBox($"Error during tool execution: {ex.Message}")
            ToolingFileLogger.EndSession(False, $"Exception: {ex.Message}", ex:=ex)
            Return ""
        Finally
            If workflowScope IsNot Nothing Then
                workflowScope.Dispose()
                workflowScope = Nothing
            End If

            _activeToolingContext = parentToolingContext
            INI_APICall_ToolInstructions_2 = ""
            INI_APICall_ToolResponses_2 = ""

            If context.LogWindowForm IsNot Nothing AndAlso Not context.LogWindowForm.IsDisposed Then
                Try
                    If UiSyncContext IsNot Nothing AndAlso
                       System.Threading.Thread.CurrentThread.ManagedThreadId <> UiThreadId Then
                        ' Post (async) — MarkComplete itself uses BeginInvoke on the form,
                        ' but the form may be on a different thread than ours. Using the
                        ' captured UI context guarantees we hit a thread with a pump.
                        UiSyncContext.Post(
                            Sub()
                                Try
                                    context.LogWindowForm.MarkComplete()
                                Catch ex As Exception
                                    ToolingFileLogger.LogWarn("Failed to mark LogWindow complete (UI post).", ex:=ex)
                                End Try
                            End Sub, Nothing)
                    Else
                        context.LogWindowForm.MarkComplete()
                    End If
                Catch ex As Exception
                    ToolingFileLogger.LogWarn("Failed to mark LogWindow complete.", ex:=ex)
                End Try

                context.Log("Session complete - close this window when ready")
                If ToolingFileLogger.IsEnabled Then
                    context.Log($"Log saved to: {ToolingFileLogger.LogFilePath}")
                End If
            End If
        End Try
    End Function


    Private Function ResolveToolingWorkflowId(requestedWorkflowId As String,
                                              subAgentMode As Boolean,
                                              parentContext As ToolExecutionContext) As String
        If Not String.IsNullOrWhiteSpace(requestedWorkflowId) Then
            Return requestedWorkflowId.Trim()
        End If

        If subAgentMode AndAlso parentContext IsNot Nothing AndAlso
           Not String.IsNullOrWhiteSpace(parentContext.WorkflowId) Then
            Return parentContext.WorkflowId
        End If

        If Not String.IsNullOrWhiteSpace(SharedLibrary.Agents.WorkflowContinuity.CurrentWorkflowId) Then
            Return SharedLibrary.Agents.WorkflowContinuity.CurrentWorkflowId
        End If

        Return SharedLibrary.Agents.WorkflowContinuity.CreateWorkflowId()
    End Function

    Private Function BuildRuntimeContextPromptBlock(context As ToolExecutionContext) As String
        If context Is Nothing OrElse String.IsNullOrWhiteSpace(context.WorkflowId) Then
            Return ""
        End If

        Return SharedLibrary.Agents.WorkflowContinuity.BuildPromptContextBlock(
            context.WorkflowId,
            includeRecentWorkflowMemoryStubs:=context.SequencingState IsNot Nothing AndAlso
                                              context.SequencingState.ShouldExposeRecentMemoryStubs)
    End Function

    Private Function ExtractWorkflowSkillName(toolCall As ToolCall,
                                              toolResponse As ToolResponse) As String
        If toolCall Is Nothing OrElse toolResponse Is Nothing OrElse Not toolResponse.Success Then
            Return ""
        End If

        If Not String.IsNullOrWhiteSpace(toolCall.ToolName) AndAlso
           toolCall.ToolName.StartsWith("skill_", StringComparison.OrdinalIgnoreCase) Then
            Return toolCall.ToolName.Substring("skill_".Length)
        End If

        If Not String.Equals(toolCall.ToolName, SharedLibrary.Agents.SkillInvokeTool.ToolName, StringComparison.OrdinalIgnoreCase) Then
            Return ""
        End If

        Try
            Dim obj As JObject = JObject.Parse(If(toolResponse.Response, ""))
            Return If(obj.Value(Of String)("name"), "").Trim()
        Catch
            Return ""
        End Try
    End Function

    Private Sub UpdateWorkflowContinuityAfterToolExecution(context As ToolExecutionContext,
                                                           toolCall As ToolCall,
                                                           toolResponse As ToolResponse)
        If context Is Nothing OrElse toolCall Is Nothing OrElse String.IsNullOrWhiteSpace(context.WorkflowId) Then
            Return
        End If

        Dim rawResponse As String = If(toolResponse?.Response, "")
        Dim resultRef As String = SharedLibrary.Agents.WorkflowContinuity.ExtractStructuredResultReference(rawResponse)
        Dim outputRef As String = SharedLibrary.Agents.WorkflowContinuity.ExtractOutputReference(rawResponse)
        Dim sourceRefs As List(Of String) = SharedLibrary.Agents.WorkflowContinuity.ExtractSourceReferences(rawResponse)

        Dim checkpointWritten As Boolean =
            SharedLibrary.Agents.WorkflowContinuity.NoteToolCallResult(
                context.WorkflowId,
                context.HostKind,
                toolCall.ToolName,
                toolResponse IsNot Nothing AndAlso toolResponse.Success,
                resultRef,
                outputRef,
                sourceRefs,
                Math.Max(context.PrematureTextRetryCount, context.SubAgentEmptyResponseRetryCount))

        Dim skillName As String = ExtractWorkflowSkillName(toolCall, toolResponse)
        Dim skillCheckpointWritten As Boolean = False

        If Not String.IsNullOrWhiteSpace(skillName) Then
            skillCheckpointWritten =
                SharedLibrary.Agents.WorkflowContinuity.NoteSkillLoaded(
                    context.WorkflowId,
                    context.HostKind,
                    skillName)
        End If

        context.RuntimeState = SharedLibrary.Agents.WorkflowContinuity.GetState(context.WorkflowId)

        context.Log(
            $"Runtime state updated: tool={toolCall.ToolName}; success={If(toolResponse IsNot Nothing AndAlso toolResponse.Success, "true", "false")}; checkpointWritten={If(checkpointWritten OrElse skillCheckpointWritten, "true", "false")}; memoryRefWritten={If(Not String.IsNullOrWhiteSpace(resultRef), "true", "false")}; sourceRefWritten={If(sourceRefs IsNot Nothing AndAlso sourceRefs.Count > 0, "true", "false")}", "diag")
    End Sub


    Private Function IsToolAllowedForCurrentContext(toolName As String, context As ToolExecutionContext) As Boolean
        If context Is Nothing OrElse String.IsNullOrWhiteSpace(toolName) Then Return False
        If Not context.EnforceAllowedToolScope Then Return True
        If context.AllowedToolNames Is Nothing Then Return False
        Return context.AllowedToolNames.Contains(toolName.Trim())
    End Function

    Private Function BuildToolNotAllowedResponse(toolCall As ToolCall, context As ToolExecutionContext) As ToolResponse
        Dim message As String = $"Tool '{toolCall.ToolName}' is not allowed for this sub-agent."

        Dim payload As String = JsonConvert.SerializeObject(New With {
        Key .summary = "Tool call was rejected by the sub-agent runtime.",
        Key .result = CType(Nothing, Object),
        Key .resultKind = "error",
        Key .error = New With {
            Key .code = "tool_not_allowed",
            Key .phase = "tool_dispatch",
            Key .message = message
        }
    })

        Dim response As New ToolResponse() With {
        .CallId = toolCall.CallId,
        .ToolName = toolCall.ToolName,
        .Response = payload,
        .Success = False,
        .ErrorMessage = message,
        .ResultKind = "error",
        .ErrorCode = "tool_not_allowed",
        .OriginalCallJson = toolCall.RawJson
    }

        If context IsNot Nothing Then
            context.LogWarn("Blocked tool outside sub-agent scope.",
                        details:=$"host={context.HostKind}; tool={toolCall.ToolName}")
        End If

        Return response
    End Function

    Private Function BuildTurnVisibleToolNameSet(selectedTools As IEnumerable(Of ModelConfig)) As HashSet(Of String)
        Dim result As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        If selectedTools Is Nothing Then
            Return result
        End If

        For Each tool As ModelConfig In selectedTools
            If tool Is Nothing OrElse String.IsNullOrWhiteSpace(tool.ToolName) Then Continue For
            result.Add(tool.ToolName.Trim())
        Next

        Return result
    End Function

    Private Function BuildToolNotExposedThisTurnResponse(toolCall As ToolCall,
                                                         visibleToolNames As IEnumerable(Of String)) As ToolResponse
        Dim visible As New List(Of String)()
        Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        If visibleToolNames IsNot Nothing Then
            For Each name As String In visibleToolNames
                Dim normalized As String = If(name, "").Trim()
                If normalized = "" Then Continue For
                If seen.Add(normalized) Then
                    visible.Add(normalized)
                End If
            Next
        End If

        visible.Sort(StringComparer.OrdinalIgnoreCase)

        If visible.Count > 25 Then
            visible = visible.Take(25).ToList()
        End If

        Dim message As String =
            $"Tool '{toolCall.ToolName}' was not exposed with its full schema at the start of this assistant turn."

        Dim payload As String = JsonConvert.SerializeObject(New With {
            Key .summary = "Tool call was rejected because its schema was not available at the start of the current turn.",
            Key .result = CType(Nothing, Object),
            Key .resultKind = "error",
            Key .visible_tools = visible,
            Key .error = New With {
                Key .code = "tool_not_exposed_in_current_turn",
                Key .phase = "tool_dispatch",
                Key .message = message
            }
        })

        Return New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName,
            .Response = payload,
            .Success = False,
            .ErrorMessage = message,
            .ResultKind = "error",
            .ErrorCode = "tool_not_exposed_in_current_turn",
            .OriginalCallJson = toolCall.RawJson
        }
    End Function

    Private Function BuildInvalidToolArgumentsResponse(toolCall As ToolCall,
                                                       validationError As String) As ToolResponse
        Dim message As String =
            $"Tool '{toolCall.ToolName}' was rejected because its arguments do not match the current schema. " &
            If(validationError, "")

        Dim payload As String = JsonConvert.SerializeObject(New With {
            Key .summary = "Tool call was rejected because its arguments do not match the current schema.",
            Key .result = CType(Nothing, Object),
            Key .resultKind = "error",
            Key .error = New With {
                Key .code = "invalid_tool_arguments",
                Key .phase = "tool_argument_validation",
                Key .message = message
            }
        })

        Return New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName,
            .Response = payload,
            .Success = False,
            .ErrorMessage = message,
            .ResultKind = "error",
            .ErrorCode = "invalid_tool_arguments",
            .OriginalCallJson = toolCall.RawJson
        }
    End Function

    Private Function BuildTurnExposureGuardPrompt(toolName As String,
                                                  Optional preparedForNextTurn As Boolean = False) As String
        Dim normalizedToolName As String = If(toolName, "").Trim()

        If preparedForNextTurn Then
            Return "HOST TOOL EXPOSURE GUARD: The tool '" &
                   normalizedToolName &
                   "' was not exposed with its full schema at the start of the previous assistant turn. " &
                   "It has now been loaded and is exposed in THIS turn. " &
                   "If you still need that tool, call '" & normalizedToolName & "' now through the native tool-calling mechanism. " &
                   "Do NOT call tool_loader for this tool again. " &
                   "Do NOT postpone this tool call to a later turn. " &
                   "Do NOT produce final prose that merely announces a later audit, attestation, or follow-up step."
        End If

        Return "HOST TOOL EXPOSURE GUARD: The tool '" &
               normalizedToolName &
               "' was not exposed with its full schema at the start of the previous assistant turn. " &
               "Do NOT call a newly loaded tool in the same turn in which it becomes available. " &
               "If you need that tool, first load it, then wait for the next turn, and only then call it."
    End Function

    Private Function BuildToolLoaderBarrierGuardPrompt() As String
        Return "HOST TOOL LOADER BARRIER: The previous turn successfully executed tool_loader. " &
               "Newly loaded tool schemas are available only from the next assistant turn onward. " &
               "Do NOT combine tool_loader and a newly loaded tool call in the same turn. " &
               "In this turn, either call one of the now-exposed tools or continue with another already-exposed step."
    End Function

    Private Function BuildInvalidToolArgumentsGuardPrompt(toolName As String,
                                                          validationError As String) As String
        Dim normalizedToolName As String = If(toolName, "").Trim()

        Return "HOST TOOL ARGUMENT VALIDATION: The previous call to '" &
               normalizedToolName &
               "' was rejected because its arguments did not match the current schema. " &
               "Use the exact currently exposed schema. " &
               "Validation error: " & If(validationError, "")
    End Function

    Private Sub ApplyStructuredAgentResult(response As ToolResponse, context As ToolExecutionContext)
        If response Is Nothing Then Return

        Dim errorCode As String = ""
        Dim resultKind As String = ""

        If SharedLibrary.Agents.SubAgentRuntimeHardening.TryGetEnvelopeErrorInfo(response.Response, errorCode, resultKind) Then
            response.ResultKind = If(String.IsNullOrWhiteSpace(resultKind), "error", resultKind)
            response.ErrorCode = errorCode
            response.Success = False

            If String.IsNullOrWhiteSpace(response.ErrorMessage) Then
                response.ErrorMessage = If(String.IsNullOrWhiteSpace(errorCode),
                                       "Agent-layer tool returned an error payload.",
                                       $"Agent-layer tool returned error '{errorCode}'.")
            End If

            If context IsNot Nothing Then
                context.LogWarn("Agent-layer tool returned structured failure.",
                            details:=$"tool={response.ToolName}; resultKind={response.ResultKind}; errorCode={If(response.ErrorCode, "")}")
            End If
            Return
        End If

        If Not String.IsNullOrWhiteSpace(resultKind) Then
            response.ResultKind = resultKind
        End If
    End Sub




    Private Function ResolveBlockedFallbackUserLanguage(context As ToolExecutionContext) As String
        Return If(context?.SequencingState?.UserLanguage, "").Trim()
    End Function


    Private Async Function LocalizeHostMessageIfNeededAsync(message As String,
                                                            userLanguage As String,
                                                            useSecondAPI As Boolean,
                                                            hideSplash As Boolean,
                                                            cancellationToken As System.Threading.CancellationToken) As Task(Of String)
        Dim baseMessage As String =
            SharedLibrary.Agents.ToolCallSequencing.StripTaskStatusBlocksFromUserFacingText(
                If(message, "").Trim())

        Dim targetLanguage As String = If(userLanguage, "").Trim()

        If baseMessage = "" OrElse targetLanguage = "" Then
            Return baseMessage
        End If

        If String.Equals(targetLanguage, "en", StringComparison.OrdinalIgnoreCase) OrElse
           targetLanguage.StartsWith("en-", StringComparison.OrdinalIgnoreCase) Then
            Return baseMessage
        End If

        Dim systemPrompt As String =
            "Translate the provided short end-user office-assistant message into the requested language. " &
            "Do not add explanations, quotes, code fences, markdown, technical detail, or any TASK_STATUS block. " &
            "Return only the final localized message."

        Dim userPrompt As String =
            "<TARGET_LANGUAGE>" & targetLanguage & "</TARGET_LANGUAGE>" & vbCrLf &
            "<MESSAGE>" & baseMessage & "</MESSAGE>"

        Try
            Dim localized As String = Await LLM(
                systemPrompt,
                userPrompt,
                "", "", 0,
                useSecondAPI,
                hideSplash,
                       "",
                     "",
                     True,
                     cancellationToken)

            If String.IsNullOrWhiteSpace(localized) Then
                Return baseMessage
            End If

            Return SharedLibrary.Agents.ToolCallSequencing.StripTaskStatusBlocksFromUserFacingText(localized.Trim())
        Catch
            Return baseMessage
        End Try
    End Function

    Private Async Function BuildLocalizedCreatedDeliverableSuccessMessageAsync(context As ToolExecutionContext,
                                                                            toolResponse As ToolResponse,
                                                                            useSecondAPI As Boolean,
                                                                            hideSplash As Boolean,
                                                                            cancellationToken As System.Threading.CancellationToken) As Task(Of String)
        Dim references As List(Of String) =
            SharedLibrary.Agents.ToolCallSequencing.ExtractCreatedDeliverableReferences(
                If(If(toolResponse Is Nothing, Nothing, toolResponse.Response), ""))

        Dim displayNames As New List(Of String)()

        For Each reference As String In references
            Dim candidate As String = If(reference, "").Trim()
            If candidate = "" Then Continue For

            Try
                Dim fileName As String = Path.GetFileName(candidate)
                If fileName <> "" Then
                    candidate = fileName
                End If
            Catch
            End Try

            Dim alreadyAdded As Boolean = False

            For Each existing As String In displayNames
                If String.Equals(existing, candidate, StringComparison.OrdinalIgnoreCase) Then
                    alreadyAdded = True
                    Exit For
                End If
            Next

            If Not alreadyAdded Then
                displayNames.Add(candidate)
            End If
        Next

        Dim baseMessage As String

        If displayNames.Count = 1 Then
            baseMessage = "Done. Created: " & displayNames(0) & "."
        ElseIf displayNames.Count > 1 Then
            baseMessage = "Done. Created: " & String.Join(", ", displayNames) & "."
        Else
            baseMessage = "Done. The requested output was created."
        End If

        Return Await LocalizeHostMessageIfNeededAsync(
            baseMessage,
            ResolveBlockedFallbackUserLanguage(context),
            useSecondAPI,
            hideSplash, cancellationToken)
    End Function


    Private Async Function BuildTaskSpecificPartialBlockedMessageAsync(context As ToolExecutionContext,
                                                                       errorCode As String,
                                                                       message As String,
                                                                       useSecondAPI As Boolean,
                                                                       hideSplash As Boolean) As Task(Of String)
        If context Is Nothing OrElse context.AllToolResponses Is Nothing Then
            Return ""
        End If

        Dim completedFacts As New List(Of String)()
        Dim unresolvedFacts As New List(Of String)()

        For Each item In context.AllToolResponses.Where(Function(r) r IsNot Nothing AndAlso r.Success)
            Dim excerpt As String = If(item.Response, "").Trim()
            excerpt = excerpt.Replace(vbCr, " ").Replace(vbLf, " ").Trim()

            If excerpt.Length > 240 Then
                excerpt = excerpt.Substring(0, 240) & "..."
            End If

            If excerpt <> "" Then
                completedFacts.Add(excerpt)
            End If

            If completedFacts.Count >= 3 Then
                Exit For
            End If
        Next

        For Each item In context.AllToolResponses.Where(Function(r) r IsNot Nothing AndAlso Not r.Success)
            Dim failureText As String = If(item.ErrorMessage, "").Trim()
            failureText = failureText.Replace(vbCr, " ").Replace(vbLf, " ").Trim()

            If failureText.Length > 180 Then
                failureText = failureText.Substring(0, 180) & "..."
            End If

            If failureText <> "" Then
                unresolvedFacts.Add(failureText)
            End If

            If unresolvedFacts.Count >= 2 Then
                Exit For
            End If
        Next

        If completedFacts.Count = 0 Then
            Return ""
        End If

        Dim targetLanguage As String = ResolveBlockedFallbackUserLanguage(context)
        If targetLanguage = "" Then
            targetLanguage = "en"
        End If

        Dim systemPrompt As String =
            "Write a short end-user office-assistant status message in the requested language. " &
            "Briefly say what was already completed or prepared, and then what could not be completed or may still be incomplete. " &
            "Keep it short, non-technical, task-specific, and easy to understand. " &
            "Do not mention tools, validators, workflows, JSON, retries, memory grounding, checkpoints, or internal state. " &
            "Do not use bullet points. Do not add a TASK_STATUS footer. Do not invent details."

        Dim userPrompt As String =
            "<TARGET_LANGUAGE>" & targetLanguage & "</TARGET_LANGUAGE>" & vbCrLf &
            "<COMPLETED_FACTS>" & vbCrLf &
            String.Join(vbCrLf, completedFacts) & vbCrLf &
            "</COMPLETED_FACTS>" & vbCrLf &
            "<UNRESOLVED_FACTS>" & vbCrLf &
            String.Join(vbCrLf, unresolvedFacts) & vbCrLf &
            "</UNRESOLVED_FACTS>" & vbCrLf &
            "<HOST_BLOCK_REASON>" & If(message, "") & "</HOST_BLOCK_REASON>"

        Try
            Dim responseText As String = Await LLM(
                systemPrompt,
                userPrompt,
                "", "", 0,
                useSecondAPI,
                hideSplash,
                "",
                "",
                True)

            Return If(responseText, "").Trim()
        Catch
            Return ""
        End Try
    End Function

    Private Async Function BuildBlockedToolingResultAsync(context As ToolExecutionContext,
                                                          errorCode As String,
                                                          message As String,
                                                          useSecondAPI As Boolean,
                                                          hideSplash As Boolean, cancellationtoken As System.Threading.CancellationToken) As Task(Of String)
        Dim lastToolName As String = ""
        Dim lastErrorCode As String = ""
        Dim lastErrorMessage As String = ""

        If context IsNot Nothing AndAlso context.SequencingState IsNot Nothing Then
            lastToolName = If(context.SequencingState.LastToolName, "")
            lastErrorCode = If(context.SequencingState.LastErrorCode, "")
            lastErrorMessage = If(context.SequencingState.LastErrorMessage, "")
        End If

        Dim structuredPayload As String = SharedLibrary.Agents.ToolCallSequencing.BuildBlockedResultPayload(
            errorCode,
            "finalization",
            message,
            lastToolName,
            lastErrorCode,
            lastErrorMessage)

        ToolingFileLogger.LogWarn(
            "Structured blocked payload recorded for logs only; user-safe prose will be returned instead.",
            details:=$"host={If(context?.HostKind, "")}; payload={structuredPayload}")

        Dim successCount As Integer = 0
        Dim failedCount As Integer = 0

        If context IsNot Nothing AndAlso context.AllToolResponses IsNot Nothing Then
            successCount = context.AllToolResponses.Where(Function(r) r IsNot Nothing AndAlso r.Success).Count()
            failedCount = context.AllToolResponses.Where(Function(r) r IsNot Nothing AndAlso Not r.Success).Count()
        End If

        Dim partialMessage As String =
            Await BuildTaskSpecificPartialBlockedMessageAsync(
                context,
                errorCode,
                message,
                useSecondAPI,
                hideSplash)

        If Not String.IsNullOrWhiteSpace(partialMessage) Then
            Return partialMessage.Trim() & " " &
                SharedLibrary.Agents.ToolCallSequencing.BuildTaskStatusFooter(
                    "blocked",
                    If(errorCode, "host_generated_blocked"))
        End If

        Dim baseMessage As String =
            SharedLibrary.Agents.ToolCallSequencing.BuildUserSafeBlockedFinalMessage(
                context.SequencingState,
                errorCode,
                message,
                successCount,
                failedCount,
                userLanguage:=ResolveBlockedFallbackUserLanguage(context),
                appendTaskStatusFooter:=True)

        Return Await LocalizeHostMessageIfNeededAsync(
            baseMessage,
            ResolveBlockedFallbackUserLanguage(context),
            useSecondAPI,
            hideSplash, cancellationtoken)
    End Function


    Private Shared Function IsSkipToolErrorHandling(toolConfig As ModelConfig) As Boolean
        Return toolConfig IsNot Nothing AndAlso
           String.Equals(If(toolConfig.ToolErrorHandling, "skip"), "skip", StringComparison.OrdinalIgnoreCase)
    End Function

    Private Function IsStructuredErrorToolResponse(toolResponse As ToolResponse) As Boolean
        If toolResponse Is Nothing OrElse String.IsNullOrWhiteSpace(toolResponse.Response) Then
            Return False
        End If

        Dim errorCode As String = ""
        Dim resultKind As String = ""
        Return SharedLibrary.Agents.SubAgentRuntimeHardening.TryGetEnvelopeErrorInfo(
        toolResponse.Response,
        errorCode,
        resultKind)
    End Function

    Private Function IsSkippedStructuredAgentFailure(toolCall As ToolCall,
                                                 toolResponse As ToolResponse,
                                                 toolConfig As ModelConfig,
                                                 subAgentMode As Boolean) As Boolean
        If subAgentMode Then Return False
        If toolCall Is Nothing OrElse String.IsNullOrWhiteSpace(toolCall.ToolName) Then Return False
        If Not toolCall.ToolName.StartsWith("agent_", StringComparison.OrdinalIgnoreCase) Then Return False
        If toolResponse Is Nothing OrElse toolResponse.Success Then Return False
        If Not IsSkipToolErrorHandling(toolConfig) Then Return False
        Return IsStructuredErrorToolResponse(toolResponse)
    End Function

    Private Function BuildToolFailureReassessmentGuardPrompt(failedToolName As String) As String
        Dim normalizedToolName As String = If(failedToolName, "").Trim()
        Dim failedToolText As String =
            If(normalizedToolName = "",
               "The previous tool step failed. ",
               "The previous tool step using '" & normalizedToolName & "' failed. ")

        Return "HOST TOOL FAILURE RECOVERY: " &
               failedToolText &
               "In THIS turn, first reassess whether that was the right tool for the remaining work. " &
               "If another available tool is more appropriate, use that tool instead. " &
               "If the same tool is still appropriate, retry only with narrower or corrected arguments. " &
               "If enough information is already available, provide the best possible final answer and say clearly if it may be incomplete. " &
               "Return a blocked message only if no reliable answer and no useful recovery step is possible."
    End Function

    Private Function BuildSkippedToolFailureRecoveryGuardPrompt() As String
        Return "HOST TOOL FAILURE GUARD: The previous tool call failed and that structured error is already available in the tool-response history. " &
           "In THIS turn you must take exactly one actionable recovery step: " &
           "(a) write or update failure state for the current item using an available write/state tool; OR " &
           "(b) retry the failed agent/tool with smaller or chunked input; OR " &
           "(c) continue with the next item using the appropriate tool; OR " &
           "(d) produce a valid final blocked response only if all required work is complete or no further tool action is possible. " &
           "Do NOT apologize. Do NOT output ordinary progress prose. Do NOT merely restate the failure."
    End Function




    Private Sub LogToolBatchPlan(context As ToolExecutionContext,
                             plan As SharedLibrary.Agents.ToolCallSequencing.ToolBatchPlan)
        If context Is Nothing OrElse plan Is Nothing Then Return

        context.Log($"Tool-call batch analysis: total={plan.TotalCallCount}; safeBatch={plan.IsFullyBatchSafe}; executed={plan.ExecutedCount}; deferred={plan.DeferredCount}", "diag")
        For Each planned In plan.Calls
            context.Log($"Tool-call classification: index={planned.Index + 1}; tool={planned.ToolName}; class={planned.Classification.ToString().ToLowerInvariant()}; barrier={planned.IsBarrier}; action={If(planned.WillExecute, "execute", "defer")}", "diag")
        Next
    End Sub

    Private Function EnsureVisibleToolLoaded(toolName As String, context As ToolExecutionContext) As ModelConfig
        If context Is Nothing OrElse String.IsNullOrWhiteSpace(toolName) Then
            Return Nothing
        End If

        If context.SelectedTools Is Nothing Then
            context.SelectedTools = New List(Of ModelConfig)()
        End If

        Dim existing = context.SelectedTools.FirstOrDefault(
        Function(t)
            Return t IsNot Nothing AndAlso
                   Not String.IsNullOrWhiteSpace(t.ToolName) AndAlso
                   t.ToolName.Equals(toolName, StringComparison.OrdinalIgnoreCase)
        End Function)

        If existing IsNot Nothing Then
            Return existing
        End If

        If Not IsToolAllowedForCurrentContext(toolName, context) Then
            Return Nothing
        End If

        Dim loaded = context.AllowedToolRegistry.Get(toolName)
        If loaded Is Nothing Then
            Return Nothing
        End If

        context.SelectedTools.Add(loaded)
        Return loaded
    End Function

    Private Function ExpandAllowedToolNamesForRegistry(requestedToolNames As IEnumerable(Of String),
                                                   registry As SharedLibrary.Agents.ToolRegistry) As List(Of String)
        If requestedToolNames Is Nothing Then
            Return Nothing
        End If

        Dim result As New List(Of String)()
        Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        Dim availableNames As List(Of String) =
        If(registry Is Nothing,
           New List(Of String)(),
           registry.ListManifests().
               Where(Function(m) m IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(m.Name)).
               Select(Function(m) m.Name.Trim()).
               Distinct(StringComparer.OrdinalIgnoreCase).
               ToList())

        For Each rawName As String In requestedToolNames
            Dim token As String = If(rawName, "").Trim()
            If token = "" Then Continue For

            Dim expandedNames As IEnumerable(Of String)

            If token = "*" Then
                expandedNames = availableNames

            ElseIf IsSelectedOnlineSourcesAlias(token) Then
                expandedNames =
                availableNames.Where(Function(name) IsSelectedOnlineSourceToolName(name))

            ElseIf ContainsWildcardPattern(token) Then
                expandedNames =
                availableNames.Where(Function(name) WildcardToolNameMatches(token, name))

            Else
                expandedNames = {token}
            End If

            For Each expandedName As String In expandedNames
                Dim normalized As String = If(expandedName, "").Trim()
                If normalized = "" Then Continue For

                If seen.Add(normalized) Then
                    result.Add(normalized)
                End If
            Next
        Next

        Return result
    End Function

    Private Shared Function IsSelectedOnlineSourcesAlias(token As String) As Boolean
        Select Case If(token, "").Trim().ToLowerInvariant()
            Case "selected_online_sources",
             "@selected_online_sources",
             "selected_sources",
             "@selected_sources"
                Return True
            Case Else
                Return False
        End Select
    End Function

    Private Shared Function ContainsWildcardPattern(token As String) As Boolean
        If String.IsNullOrWhiteSpace(token) Then Return False
        Return token.IndexOf("*"c) >= 0 OrElse token.IndexOf("?"c) >= 0
    End Function

    Private Shared Function WildcardToolNameMatches(pattern As String, toolName As String) As Boolean
        If String.IsNullOrWhiteSpace(pattern) OrElse String.IsNullOrWhiteSpace(toolName) Then
            Return False
        End If

        Dim regexPattern As String =
        "^" &
        Regex.Escape(pattern.Trim()).
            Replace("\*", ".*").
            Replace("\?", ".") &
        "$"

        Return Regex.IsMatch(
        toolName.Trim(),
        regexPattern,
        RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)
    End Function

    Private Function IsSelectedOnlineSourceToolName(toolName As String) As Boolean
        If String.IsNullOrWhiteSpace(toolName) Then Return False

        Dim name As String = toolName.Trim()

        If name.StartsWith("skill_", StringComparison.OrdinalIgnoreCase) OrElse
       name.StartsWith("agent_", StringComparison.OrdinalIgnoreCase) Then
            Return False
        End If

        If name.Equals(InternalWebToolName, StringComparison.OrdinalIgnoreCase) OrElse
       name.Equals(InternalDownloadWebFilesToolName, StringComparison.OrdinalIgnoreCase) OrElse
       name.Equals(InternalSearchToolName, StringComparison.OrdinalIgnoreCase) OrElse
       name.Equals(SharedLibrary.Agents.SkillInvokeTool.ToolName, StringComparison.OrdinalIgnoreCase) OrElse
       name.Equals(SharedLibrary.Agents.ToolLoaderTool.LoaderToolName, StringComparison.OrdinalIgnoreCase) Then
            Return False
        End If

        If IsInternalKnowledgeToolName(name) OrElse
       SharedLibrary.Agents.WebGroundingTool.IsWebGroundingTool(name) OrElse
       SharedLibrary.SharedLibrary.M365ToolService.IsM365ToolName(name) OrElse
       SharedLibrary.Agents.MemoryTools.IsMemoryTool(name) OrElse
       SharedLibrary.Agents.TextTools.IsTextTool(name) OrElse
       SharedLibrary.Agents.WorkspaceTools.IsWorkspaceTool(name) OrElse
       SharedLibrary.Agents.WordTools.IsWordTool(name) OrElse
       SharedLibrary.Agents.WordDocTools.IsWordDocTool(name) OrElse
       SharedLibrary.Agents.JsRunTool.IsJsTool(name) Then
            Return False
        End If

        Return True
    End Function


    Private Sub LoadSkillAllowedToolsFromResponse(skillResponse As String, context As ToolExecutionContext)
        If context Is Nothing OrElse String.IsNullOrWhiteSpace(skillResponse) Then
            Return
        End If

        Try
            Dim obj As JObject = JObject.Parse(skillResponse)
            Dim allowedToolsToken As JToken = obj("allowed_tools")

            If allowedToolsToken Is Nothing OrElse allowedToolsToken.Type <> JTokenType.Array Then
                Return
            End If

            Dim requestedToolNames As New List(Of String)()

            For Each item As JToken In DirectCast(allowedToolsToken, JArray)
                Dim toolName As String = item.ToString().Trim()
                If toolName = "" Then Continue For
                requestedToolNames.Add(toolName)
            Next

            Dim expandedToolNames As List(Of String) =
            ExpandAllowedToolNamesForRegistry(
                requestedToolNames,
                context.AllowedToolRegistry)

            For Each toolName As String In expandedToolNames
                EnsureVisibleToolLoaded(toolName, context)
            Next
        Catch
        End Try
    End Sub


    Private Function ExecuteToolLoaderCall(toolCall As ToolCall, context As ToolExecutionContext) As ToolResponse
        Dim response As New ToolResponse() With {
        .CallId = toolCall.CallId,
        .ToolName = toolCall.ToolName
    }

        If context Is Nothing OrElse context.AllowedToolRegistry Is Nothing Then
            response.Success = False
            response.ErrorMessage = "Tool loader is not initialized."
            Return response
        End If

        If context.SelectedTools Is Nothing Then
            context.SelectedTools = New List(Of ModelConfig)()
        End If

        Dim requestedNames = SharedLibrary.Agents.ToolLoaderTool.ExtractRequestedToolNames(toolCall.Arguments)

        If requestedNames.Count = 0 Then
            response.Success = False
            response.ErrorMessage = "No tool names were provided to tool_loader."
            Return response
        End If

        Dim loadedNames As New List(Of String)()
        Dim alreadyLoadedNames As New List(Of String)()
        Dim unavailableNames As New List(Of String)()

        For Each requestedName In requestedNames
            If String.IsNullOrWhiteSpace(requestedName) Then Continue For

            If Not IsToolAllowedForCurrentContext(requestedName, context) Then
                unavailableNames.Add(requestedName)
                Continue For
            End If

            If requestedName.Equals(SharedLibrary.Agents.ToolLoaderTool.LoaderToolName, StringComparison.OrdinalIgnoreCase) Then
                alreadyLoadedNames.Add(requestedName)
                Continue For
            End If

            Dim existing = context.SelectedTools.FirstOrDefault(
            Function(t)
                Return t IsNot Nothing AndAlso
                       Not String.IsNullOrWhiteSpace(t.ToolName) AndAlso
                       t.ToolName.Equals(requestedName, StringComparison.OrdinalIgnoreCase)
            End Function)

            If existing IsNot Nothing Then
                alreadyLoadedNames.Add(existing.ToolName)
                Continue For
            End If

            Dim loaded = EnsureVisibleToolLoaded(requestedName, context)
            If loaded IsNot Nothing Then
                loadedNames.Add(loaded.ToolName)
            Else
                unavailableNames.Add(requestedName)
            End If
        Next

        Dim payload As New JObject(
        New JProperty("loaded", New JArray(loadedNames.ToArray())),
        New JProperty("already_loaded", New JArray(alreadyLoadedNames.ToArray())),
        New JProperty("not_available", New JArray(unavailableNames.ToArray()))
    )

        response.Success = True
        response.Response = payload.ToString(Formatting.None)
        Return response
    End Function

#End Region

#Region "Tooling Helper Functions"



    Private Enum TaskStatusKind
        Missing
        Complete
        ContinueWork
        Blocked
    End Enum

    Private Function ParseTaskStatus(text As String) As TaskStatusKind
        Dim parsed As Agents.TaskStatusFooter = Agents.TaskStatusFooterParser.Parse(text)
        Select Case parsed.Kind
            Case Agents.TaskStatusKind.Complete : Return TaskStatusKind.Complete
            Case Agents.TaskStatusKind.ContinueWork : Return TaskStatusKind.ContinueWork
            Case Agents.TaskStatusKind.Blocked : Return TaskStatusKind.Blocked
            Case Else : Return TaskStatusKind.Missing
        End Select
    End Function


    ''' <summary>
    ''' Word equivalent of the Outlook helper: collects deliverable-capable tool names
    ''' that are still loaded and have NOT yet been invoked in this run. Used by the
    ''' final-turn fallback gate to decide whether 'blocked' is premature.
    ''' </summary>
    Private Function CollectUntriedDeliverableFallbackToolNames(ctx As ToolExecutionContext) As IReadOnlyList(Of String)
        Dim result As New List(Of String)()
        If ctx Is Nothing Then Return result.AsReadOnly()

        Dim deliverableCandidates As New HashSet(Of String)(
            SharedLibrary.Agents.HostToolRegistration.GetDeliverableCapableToolNames(
                SharedLibrary.Agents.ToolingHostKind.Word),
            StringComparer.OrdinalIgnoreCase)

        Dim available As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        If ctx.SelectedTools IsNot Nothing Then
            For Each tool As ModelConfig In ctx.SelectedTools
                If tool Is Nothing OrElse String.IsNullOrWhiteSpace(tool.ToolName) Then Continue For
                available.Add(tool.ToolName.Trim())
            Next
        End If

        If ctx.AuthoritativeToolRegistrySnapshot IsNot Nothing Then
            For Each manifest In ctx.AuthoritativeToolRegistrySnapshot.ListManifests()
                If manifest Is Nothing OrElse String.IsNullOrWhiteSpace(manifest.Name) Then Continue For
                available.Add(manifest.Name.Trim())
            Next
        End If

        Dim used As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        If ctx.AllToolResponses IsNot Nothing Then
            For Each r As ToolResponse In ctx.AllToolResponses
                If r Is Nothing OrElse String.IsNullOrWhiteSpace(r.ToolName) Then Continue For
                used.Add(r.ToolName.Trim())
            Next
        End If

        For Each candidate In deliverableCandidates.OrderBy(Function(n) n, StringComparer.OrdinalIgnoreCase)
            If available.Contains(candidate) AndAlso Not used.Contains(candidate) Then
                result.Add(candidate)
            End If
        Next

        Return result.AsReadOnly()
    End Function

    Private Function StripTaskStatus(text As String) As String
        Return Agents.TaskStatusFooterParser.Strip(text)
    End Function


    Private Function ShouldRetryAfterPrematureTextResponse(context As ToolExecutionContext, lastResponse As String) As Boolean
        If context Is Nothing Then Return False
        If context.PrematureTextRetryCount >= ToolExecutionContext.MaxContinuationRetries Then Return False

        Select Case ParseTaskStatus(lastResponse)
            Case TaskStatusKind.Complete, TaskStatusKind.Blocked
                Return False
            Case Else
                ' Missing footer OR status="continue" → force another iteration.
                Return True
        End Select
    End Function

    Private Function BuildPrematureTextContinuationGuardPrompt() As String
        Return "HOST CONTINUATION GUARD: Your previous turn was a text-only response that did NOT end with a valid <TASK_STATUS> footer declaring 'complete' or 'blocked'. " &
               "Therefore the task is treated as unfinished. In THIS turn you must do ONE of: " &
               "(a) invoke the next appropriate tool call now (no footer in tool-call turns), OR " &
               "(b) deliver the actual final, complete result that fully satisfies the user's request and end it with exactly: " &
               "<TASK_STATUS>{""status"":""complete"",""reason"":""...""}</TASK_STATUS>  " &
               "If the task is genuinely impossible despite reasonable tool attempts, end the turn with: " &
               "<TASK_STATUS>{""status"":""blocked"",""reason"":""...""}</TASK_STATUS>  " &
               "If the user's request covers multiple items (every PDF, alle Dokumente, tous les fichiers, etc.), you MUST iterate via tool calls until ALL items are processed before declaring 'complete'. " &
               "If a tool just failed, try a DIFFERENT applicable tool with corrected arguments. To read files inside the connected workspace use workspace_extract_text (any format) or workspace_read (plain text only)."
    End Function


    Private Function EncodeToolToken(value As String) As String
        If String.IsNullOrWhiteSpace(value) Then Return ""

        ' Use a SHA256 hash (hex, lowercase) to produce a fixed-length token
        ' that is always valid for API function names and stays well under
        ' the 128-character name limit imposed by model APIs (e.g. Gemini).
        Using sha = System.Security.Cryptography.SHA256.Create()
            Dim hashBytes = sha.ComputeHash(Encoding.UTF8.GetBytes(value))
            Dim sb As New StringBuilder(hashBytes.Length * 2)
            For Each b In hashBytes
                sb.Append(b.ToString("x2"))
            Next
            Return sb.ToString()   ' 64 hex chars, always
        End Using
    End Function

    Private Function DecodeToolToken(value As String) As String
        ' Hash-based tokens are one-way; decoding is no longer possible.
        ' Callers must use GetKnowledgeStoreForToolName which matches
        ' by recomputing hashes against known stores.
        Return ""
    End Function



    ''' <summary>
    ''' Converts a canonical tool definition JSON string to a model-specific format using a template string.
    ''' </summary>
    ''' <param name="canonicalDefinition">Canonical tool definition JSON (must parse as a JSON object).</param>
    ''' <param name="template">Template used to render the model-specific definition.</param>
    ''' <returns>Rendered tool definition string, or an empty string on error.</returns>
    Public Function ConvertCanonicalToModelFormat(canonicalDefinition As String, template As String) As String
        If String.IsNullOrWhiteSpace(canonicalDefinition) OrElse String.IsNullOrWhiteSpace(template) Then
            ToolingFileLogger.LogWarn(
                "ConvertCanonicalToModelFormat: Empty input.",
                details:=$"canonicalDefinitionEmpty={String.IsNullOrWhiteSpace(canonicalDefinition)}; templateEmpty={String.IsNullOrWhiteSpace(template)}")
            Return ""
        End If

        Try
            Dim jDef As JObject = JObject.Parse(canonicalDefinition)
            Dim name As String = If(jDef("name")?.ToString(), "")
            Dim description As String = If(jDef("description")?.ToString(), "")

            ' Use Formatting.None to produce compact JSON for the parameters object.
            ' JObject.ToString() defaults to Formatting.Indented, which injects literal
            ' newlines and whitespace that bloat the payload and can break model API
            ' templates that expect single-line JSON values.
            Dim parametersToken As JToken = jDef("parameters")
            Dim parameters As String = If(parametersToken IsNot Nothing,
                parametersToken.ToString(Formatting.None), "{}")

            ' JSON-escape name and description before injecting into the template.
            ' JValue.ToString() returns the RAW unescaped string (e.g., embedded " or \
            ' are not escaped). When the template places these inside JSON string
            ' literals like "name":"{name}", unescaped characters produce invalid JSON.
            ' This is especially critical when combining multiple tools — a single
            ' malformed definition breaks the entire tools array and the API rejects
            ' the request, causing LLM() to return an empty string.
            Dim result As String = template
            result = result.Replace("{name}", EscapeJsonString(name))
            result = result.Replace("{description}", EscapeJsonString(description))
            result = result.Replace("{parameters}", parameters)

            Return result
        Catch ex As Exception
            ToolingFileLogger.LogError("ConvertCanonicalToModelFormat error.", details:=$"canonicalDefinition='{canonicalDefinition}'", ex:=ex)
            Debug.WriteLine($"ConvertCanonicalToModelFormat error: {ex.Message}")
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Builds the model-specific tool definitions block for the current tooling model by converting each selected tool definition
    ''' and injecting the resulting definition list into the model's <see cref="ModelConfig.APICall_ToolInstructions"/> template.
    ''' </summary>
    ''' <param name="selectedTools">Tools to include, sorted by <see cref="ModelConfig.ToolPriority"/>.</param>
    ''' <param name="toolingModel">The tooling model that defines the instruction template.</param>
    ''' <returns>Tool instructions string passed via <c>INI_APICall_ToolInstructions_2</c>.</returns>
    Public Function BuildToolInstructionsForModel(selectedTools As List(Of ModelConfig), toolingModel As ModelConfig) As String
        If toolingModel Is Nothing Then
            ToolingFileLogger.LogWarn("BuildToolInstructionsForModel: toolingModel is Nothing.")
            Return ""
        End If

        If String.IsNullOrWhiteSpace(toolingModel.APICall_ToolInstructions) Then
            ToolingFileLogger.LogWarn("BuildToolInstructionsForModel: toolingModel.APICall_ToolInstructions is empty.")
            Return ""
        End If

        Dim definitions As New StringBuilder()
        Dim isFirst As Boolean = True

        Dim sortedTools = selectedTools.OrderBy(Function(t) t.ToolPriority).ToList()

        For Each tool In sortedTools
            If String.IsNullOrWhiteSpace(tool.ToolDefinition) Then
                ToolingFileLogger.LogWarn("Tool skipped: no ToolDefinition.", details:=$"ToolName='{tool.ToolName}'")
                Continue For
            End If

            Dim modelSpecificDef = ConvertCanonicalToModelFormat(
                tool.ToolDefinition,
                toolingModel.APICall_ToolInstructions_Template)

            If Not String.IsNullOrWhiteSpace(modelSpecificDef) Then
                If Not isFirst Then definitions.Append(",")
                definitions.Append(modelSpecificDef)
                isFirst = False
            Else
                ToolingFileLogger.LogWarn("Tool definition conversion returned empty.", details:=$"ToolName='{tool.ToolName}'")
            End If
        Next

        Dim result = toolingModel.APICall_ToolInstructions.Replace("{definitions}", definitions.ToString())
        result = result.Replace(LLM_APICall_Placeholder_ToolDefinitions.TrimStart("{"c).TrimEnd("}"c), definitions.ToString())

        Return result
    End Function

    ''' <summary>
    ''' Builds the model-specific tool response payload to inject into the next iteration of the tooling loop.
    ''' </summary>
    ''' <param name="responses">Tool execution outcomes to serialize.</param>
    ''' <param name="toolingModel">Tooling model that defines response templates and container structure.</param>
    ''' <returns>Serialized tool response payload.</returns>
    Public Function BuildToolResponsesForModel(responses As List(Of ToolResponse),
                                           toolingModel As ModelConfig,
                                           Optional compactForSubAgent As Boolean = False) As String
        If toolingModel Is Nothing Then
            ToolingFileLogger.LogWarn("BuildToolResponsesForModel: toolingModel is Nothing.")
            Return ""
        End If

        If String.IsNullOrWhiteSpace(toolingModel.APICall_ToolResponses) Then
            ToolingFileLogger.LogWarn("BuildToolResponsesForModel: toolingModel.APICall_ToolResponses is empty.")
            Return ""
        End If

        Dim responsePartTemplate As String = toolingModel.APICall_ToolResponses_Template
        If String.IsNullOrWhiteSpace(responsePartTemplate) Then
            ToolingFileLogger.LogWarn("BuildToolResponsesForModel: toolingModel.APICall_ToolResponses_Template is empty.")
            Return ""
        End If

        Dim callPartTemplate As String = If(toolingModel.APICall_ToolCallPart_Template, "")
        Dim useCallParts As Boolean = Not String.IsNullOrWhiteSpace(callPartTemplate)

        Dim callParts As New StringBuilder()
        Dim responseParts As New StringBuilder()
        Dim firstCall As Boolean = True
        Dim firstResp As Boolean = True

        For Each resp In responses
            If useCallParts Then
                ' Extract the original arguments from the parsed tool call JSON
                Dim argsJson As String = "{}"
                Try
                    Dim jCall = JObject.Parse(resp.OriginalCallJson)
                    Dim argsToken = jCall("arguments")
                    If argsToken IsNot Nothing Then
                        If argsToken.Type = JTokenType.String Then
                            argsJson = argsToken.ToString()
                        Else
                            argsJson = argsToken.ToString(Formatting.None)
                        End If
                    End If
                Catch
                    argsJson = "{}"
                End Try

                ' Determine if arguments should be escaped (template has quoted placeholder)
                Dim escapedArgsJson As String
                If callPartTemplate.Contains("""{arguments}""") Then
                    escapedArgsJson = EscapeJsonString(argsJson)
                Else
                    escapedArgsJson = argsJson
                End If

                ' Build the call part, also support {call} placeholder for raw call JSON
                Dim callPart As String = callPartTemplate _
                    .Replace("{call_id}", If(resp.CallId, "")) _
                    .Replace("{name}", If(resp.ToolName, "")) _
                    .Replace("{arguments}", escapedArgsJson) _
                    .Replace("{call}", resp.OriginalCallJson)

                If Not firstCall Then callParts.Append(",")
                callParts.Append(callPart)
                firstCall = False
            End If

            ' Build response content
            Dim responseContent As String = BuildToolResponseContentForModel(resp, compactForSubAgent)

            ' Model-agnostic handling:
            ' - If the response placeholder is quoted, emit an escaped string.
            ' - If the template is a Gemini-style functionResponse/function_response payload,
            '   force the inserted response to be a JSON object (arrays/scalars wrapped).
            ' - Otherwise preserve raw valid JSON for providers that accept arrays/scalars.
            Dim finalResponseContent As String
            Dim templateRequiresQuotedString As Boolean = responsePartTemplate.Contains("""{response}""")
            Dim templateLooksLikeGeminiFunctionResponse As Boolean =
                    responsePartTemplate.IndexOf("functionResponse", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                    responsePartTemplate.IndexOf("function_response", StringComparison.OrdinalIgnoreCase) >= 0

            If templateRequiresQuotedString Then
                finalResponseContent = EscapeJsonString(responseContent)
            ElseIf responsePartTemplate.Contains("{response}") Then
                Try
                    Dim parsed As JToken = JToken.Parse(responseContent)

                    If templateLooksLikeGeminiFunctionResponse Then
                        If TypeOf parsed Is JObject Then
                            finalResponseContent = parsed.ToString(Formatting.None)
                        ElseIf TypeOf parsed Is JArray Then
                            finalResponseContent = New JObject(
                                    New JProperty("items", parsed)
                                ).ToString(Formatting.None)
                        Else
                            finalResponseContent = New JObject(
                                    New JProperty("result", parsed)
                                ).ToString(Formatting.None)
                        End If
                    Else
                        finalResponseContent = parsed.ToString(Formatting.None)
                    End If
                Catch
                    finalResponseContent = New JObject(
                            New JProperty("result", responseContent)
                        ).ToString(Formatting.None)
                End Try
            Else
                finalResponseContent = EscapeJsonString(responseContent)
            End If

            Dim respPart As String = responsePartTemplate _
                .Replace("{call_id}", If(resp.CallId, "")) _
                .Replace("{name}", If(resp.ToolName, "")) _
                .Replace("{response}", finalResponseContent)

            If Not firstResp Then responseParts.Append(",")
            responseParts.Append(respPart)
            firstResp = False
        Next

        Dim functionCallsOutput As String = callParts.ToString()
        Dim responsesOutput As String = responseParts.ToString()

        ' Replace placeholders - NO comma manipulation by code
        ' Templates are responsible for their own structure
        Dim result As String = toolingModel.APICall_ToolResponses

        ' Simple replacement - if content exists, replace; if empty, remove placeholder
        result = result.Replace("{functioncalls}", functionCallsOutput)
        result = result.Replace("{responses}", responsesOutput)

        ' Clean up any empty structural remnants (empty arrays, double commas, etc.)
        ' This handles cases where one placeholder was empty
        result = Regex.Replace(result, "\[\s*\]", "[]")           ' Normalize empty arrays
        result = Regex.Replace(result, ",\s*,", ",")              ' Remove double commas
        result = Regex.Replace(result, "\[\s*,", "[")             ' Remove leading comma in array
        result = Regex.Replace(result, ",\s*\]", "]")             ' Remove trailing comma in array

        Return result
    End Function

    ''' <summary>
    ''' Determines whether a string represents a JSON object or array.
    ''' </summary>
    ''' <param name="str">Candidate JSON string.</param>
    ''' <returns>True if valid JSON object/array; otherwise False.</returns>
    Private Function IsValidJson(str As String) As Boolean
        If String.IsNullOrWhiteSpace(str) Then Return False
        str = str.Trim()
        If (str.StartsWith("{") AndAlso str.EndsWith("}")) OrElse
           (str.StartsWith("[") AndAlso str.EndsWith("]")) Then
            Try
                JToken.Parse(str)
                Return True
            Catch
                Return False
            End Try
        End If
        Return False
    End Function

    ''' <summary>
    ''' Escapes a string for safe embedding into a JSON string literal.
    ''' </summary>
    ''' <param name="str">Input string.</param>
    ''' <returns>Escaped string content (without surrounding quotes).</returns>
    Private Function EscapeJsonString(str As String) As String
        If String.IsNullOrEmpty(str) Then Return ""

        Dim sb As New StringBuilder()
        For Each c As Char In str
            Select Case c
                Case """"c : sb.Append("\""")
                Case "\"c : sb.Append("\\")
                Case "/"c : sb.Append("\/")
                Case ChrW(8) : sb.Append("\b")   ' Backspace
                Case ChrW(12) : sb.Append("\f")  ' Form feed
                Case vbLf(0) : sb.Append("\n")
                Case vbCr(0) : sb.Append("\r")
                Case vbTab(0) : sb.Append("\t")
                Case Else
                    If AscW(c) < 32 Then
                        ' Other control characters
                        sb.Append("\u" & AscW(c).ToString("X4"))
                    Else
                        sb.Append(c)
                    End If
            End Select
        Next
        Return sb.ToString()
    End Function

    ''' <summary>
    ''' Determines whether a response contains tool calls by applying a detection regex pattern.
    ''' If <paramref name="detectionPattern"/> is empty, the pattern is derived from <c>INI_Response_2</c>.
    ''' </summary>
    ''' <param name="response">LLM response text.</param>
    ''' <param name="detectionPattern">Regex pattern used for detection.</param>
    ''' <returns>True if tool calls are detected; otherwise False.</returns>
    Public Function ContainsToolCalls(response As String, detectionPattern As String) As Boolean
        If String.IsNullOrWhiteSpace(response) Then Return False

        Dim pattern As String = detectionPattern
        If String.IsNullOrWhiteSpace(pattern) Then
            pattern = ExtractToolCallPatternFromResponse(INI_Response_2)
        End If

        If String.IsNullOrWhiteSpace(pattern) Then Return False

        Try
            Return Regex.IsMatch(response, pattern, RegexOptions.Singleline Or RegexOptions.CultureInvariant)
        Catch ex As Exception
            ToolingFileLogger.LogError("Regex match error.", details:=$"pattern='{pattern}'", ex:=ex)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Extracts a regex pattern embedded in a response key using the <c>ToolCallMatching*</c> markers.
    ''' </summary>
    ''' <param name="responseKey">Response configuration key string (e.g., <c>INI_Response_2</c>).</param>
    ''' <returns>Extracted regex pattern, or an empty string if not available/invalid.</returns>
    Private Function ExtractToolCallPatternFromResponse(responseKey As String) As String
        If String.IsNullOrEmpty(responseKey) Then
            Return String.Empty
        End If

        Dim startMarker As String = ToolCallMatchingStart
        Dim startIdx As Integer = responseKey.IndexOf(startMarker, StringComparison.OrdinalIgnoreCase)
        If startIdx < 0 Then Return String.Empty

        Dim endIdx As Integer = responseKey.IndexOf(ToolCallMatchingEnd, startIdx, StringComparison.OrdinalIgnoreCase)
        Dim triggerLen As Integer = If(endIdx >= 0,
               (endIdx - startIdx + ToolCallMatchingEnd.Length),
               (responseKey.Length - startIdx))

        Dim triggerText As String = responseKey.Substring(startIdx, triggerLen)

        Dim lt As Integer = triggerText.IndexOf("<"c)
        Dim gt As Integer = triggerText.LastIndexOf(">"c)

        Dim detectedPattern As String = String.Empty

        If lt >= 0 AndAlso gt > lt Then
            detectedPattern = triggerText.Substring(lt + 1, gt - lt - 1).Trim()
        Else
            Dim colonIdx As Integer = triggerText.IndexOf(ToolCallMatchingMiddle, StringComparison.OrdinalIgnoreCase)
            If colonIdx >= 0 Then
                Dim raw As String = triggerText.Substring(colonIdx + ToolCallMatchingMiddle.Length)
                Dim paren As Integer = raw.LastIndexOf(ToolCallMatchingEnd, StringComparison.OrdinalIgnoreCase)
                If paren >= 0 Then raw = raw.Substring(0, paren)
                detectedPattern = raw.Trim()
            End If
        End If

        If Not String.IsNullOrWhiteSpace(detectedPattern) Then
            Try
                Dim rx As New Regex(detectedPattern)
            Catch ex As ArgumentException
                ToolingFileLogger.LogError("Invalid regex pattern.", details:=$"pattern='{detectedPattern}'", ex:=ex)
                Return String.Empty
            End Try
        End If

        Return detectedPattern
    End Function

    ''' <summary>
    ''' Extracts tool calls from a JSON response according to a JSON "extraction map".
    ''' </summary>
    ''' <param name="response">Response text expected to parse as JSON.</param>
    ''' <param name="extractionMap">JSON map specifying paths for tool call array/id/name/arguments.</param>
    ''' <returns>List of extracted tool calls (may be empty).</returns>
    Public Function ExtractToolCalls(response As String, extractionMap As String) As List(Of ToolCall)
        Dim calls As New List(Of ToolCall)()

        If String.IsNullOrWhiteSpace(response) OrElse String.IsNullOrWhiteSpace(extractionMap) Then
            ToolingFileLogger.LogWarn(
                "ExtractToolCalls: Missing response or extractionMap.",
                details:=$"responseEmpty={String.IsNullOrWhiteSpace(response)}; extractionMapEmpty={String.IsNullOrWhiteSpace(extractionMap)}")
            Return calls
        End If

        Try
            Dim jResponse As JToken = JToken.Parse(response)
            Dim jMap As JObject = JObject.Parse(extractionMap)

            Dim arrayPath = If(jMap("array_path")?.ToString(), "")
            Dim callIdPath = If(jMap("call_id_path")?.ToString(), "id")
            Dim namePath = If(jMap("name_path")?.ToString(), "name")
            Dim argsPath = If(jMap("arguments_path")?.ToString(), "arguments")

            Dim toolCallTokens As IEnumerable(Of JToken)

            If Not String.IsNullOrWhiteSpace(arrayPath) Then
                toolCallTokens = jResponse.SelectTokens(arrayPath).ToList()
            Else
                toolCallTokens = {jResponse}
            End If

            For Each tcToken In toolCallTokens
                Try
                    Dim tc As New ToolCall() With {
                        .CallId = If(tcToken.SelectToken(callIdPath)?.ToString(), Guid.NewGuid().ToString("N")),
                        .ToolName = If(tcToken.SelectToken(namePath)?.ToString(), ""),
                        .RawJson = tcToken.ToString()
                    }

                    Dim argsToken = tcToken.SelectToken(argsPath)
                    If argsToken IsNot Nothing Then
                        If argsToken.Type = JTokenType.String Then
                            Try
                                tc.Arguments = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(argsToken.ToString())
                            Catch ex As Exception
                                tc.Arguments = New Dictionary(Of String, Object) From {{"raw", argsToken.ToString()}}
                                ToolingFileLogger.LogWarn(
                                    "Arguments JSON string could not be deserialized; stored in 'raw'.",
                                    details:=$"ToolName='{tc.ToolName}'; CallId={tc.CallId}",
                                    ex:=ex)
                            End Try
                        Else
                            tc.Arguments = argsToken.ToObject(Of Dictionary(Of String, Object))()
                        End If
                    End If

                    If Not String.IsNullOrWhiteSpace(tc.ToolName) Then
                        calls.Add(tc)
                    Else
                        ToolingFileLogger.LogWarn("ExtractToolCalls: Skipped tool call with empty ToolName.", details:=$"Raw={tc.RawJson}")
                    End If
                Catch ex As Exception
                    ToolingFileLogger.LogError("Error parsing individual tool call.", ex:=ex)
                    Debug.WriteLine($"Error parsing individual tool call: {ex.Message}")
                End Try
            Next

        Catch ex As Exception
            ToolingFileLogger.LogError("ExtractToolCalls error.", details:=$"extractionMap='{extractionMap}'", ex:=ex)
            Debug.WriteLine($"ExtractToolCalls error: {ex.Message}")
        End Try

        Return calls
    End Function

    ''' <summary>
    ''' Builds the tool instructions prompt appended to the tooling session's system prompt.
    ''' </summary>
    ''' <param name="selectedTools">Tools to include, sorted by <see cref="ModelConfig.ToolPriority"/>.</param>
    ''' <returns>System prompt fragment describing tooling usage and available tools.</returns>
    Public Function BuildToolInstructionsPrompt(
        selectedTools As List(Of ModelConfig),
        Optional finalResponseContract As SharedLibrary.Agents.ToolingFinalResponseContract = SharedLibrary.Agents.ToolingFinalResponseContract.UserFacingTaskStatus) As String

        Dim sb As New StringBuilder()

        MaxToolIterations = INI_ToolingMaximumIterations

        If SharedLibrary.Agents.ToolingFinalResponseContractHelpers.RequiresTaskStatusFooter(finalResponseContract) Then
            sb.AppendLine(InterpolateAtRuntime(SP_Add_Tooling))
            sb.AppendLine()
            sb.AppendLine(SharedLibrary.Agents.ToolingOrchestrator.TaskStatusFooterInstruction)
        Else
            sb.AppendLine("ACTIVE-TOOLING CONTRACT: During this active tooling session, every response must be exactly one of: (1) the next required tool call and nothing else; or (2) the final raw text payload requested elsewhere in the prompt and nothing else.")
            sb.AppendLine("Do NOT append any <TASK_STATUS> footer unless the caller explicitly requested one.")
            sb.AppendLine("Do NOT convert the final output into user-facing prose unless the caller explicitly requested prose.")
            sb.AppendLine("When tool work is complete, return only the final raw text payload in the exact caller-defined format.")
        End If

        sb.AppendLine(SharedLibrary.Agents.ToolCallSequencing.DependentBatchingInstruction)

        Dim workflowAddendum As String = BuildToolWorkflowInstructionAddendum(selectedTools)
        If Not String.IsNullOrWhiteSpace(workflowAddendum) Then
            sb.AppendLine()
            sb.AppendLine(workflowAddendum)
        End If

        sb.AppendLine()
        sb.AppendLine("Available tools:")

        If selectedTools Is Nothing Then
            Return sb.ToString()
        End If

        Dim sortedTools = selectedTools.OrderBy(Function(t) t.ToolPriority).ToList()

        For Each tool In sortedTools
            If Not String.IsNullOrWhiteSpace(tool.ToolInstructionsPrompt) Then
                sb.AppendLine()
                sb.AppendLine($"- {tool.ToolInstructionsPrompt}")
            End If
        Next

        Return sb.ToString()
    End Function

    Private Function BuildSubAgentLoaderManifests(context As ToolExecutionContext) As List(Of SharedLibrary.Agents.ToolManifest)
        Dim result As New List(Of SharedLibrary.Agents.ToolManifest)()

        If context Is Nothing OrElse context.AllowedToolRegistry Is Nothing Then
            Return result
        End If

        Dim selectedNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        If context.SelectedTools IsNot Nothing Then
            For Each tool In context.SelectedTools
                If tool Is Nothing OrElse String.IsNullOrWhiteSpace(tool.ToolName) Then Continue For
                selectedNames.Add(tool.ToolName.Trim())
            Next
        End If

        For Each manifest In context.AllowedToolRegistry.ListManifests()
            If manifest Is Nothing OrElse String.IsNullOrWhiteSpace(manifest.Name) Then Continue For

            Dim toolName As String = manifest.Name.Trim()

            If toolName.Equals(SharedLibrary.Agents.ToolLoaderTool.LoaderToolName, StringComparison.OrdinalIgnoreCase) Then Continue For
            If toolName.StartsWith("agent_", StringComparison.OrdinalIgnoreCase) Then Continue For
            If selectedNames.Contains(toolName) Then Continue For

            result.Add(manifest)
        Next

        Return result
    End Function

    Private Function BuildToolInstructionsPromptForSession(
        selectedTools As List(Of ModelConfig),
        subAgentMode As Boolean,
        Optional finalResponseContract As SharedLibrary.Agents.ToolingFinalResponseContract = SharedLibrary.Agents.ToolingFinalResponseContract.UserFacingTaskStatus) As String

        If Not subAgentMode Then
            Return BuildToolInstructionsPrompt(selectedTools, finalResponseContract)
        End If

        Dim sb As New StringBuilder()

        sb.AppendLine("You are running in a native tool-calling runtime.")
        sb.AppendLine()
        sb.AppendLine("SUB-AGENT RULES:")
        sb.AppendLine("- Use tools only if they are actually needed.")
        sb.AppendLine("- If a tool is declared for this turn, call it ONLY through the model's native function/tool-calling mechanism.")
        sb.AppendLine("- Do NOT write tool calls as plain text, JSON text, markdown, Python, JavaScript, SDK examples, print(...), tool(...), tool.run(...), or function wrappers.")
        sb.AppendLine("- Do NOT call any agent_* tool.")
        sb.AppendLine("- Never call tool_loader for agent_* tools.")
        sb.AppendLine("- Use tool_loader only when a needed tool is not yet declared in this turn.")
        sb.AppendLine("- Do NOT append any <TASK_STATUS> footer.")
        sb.AppendLine("- When the work is finished, return exactly one JSON object with keys ""summary"" and ""result"".")
        sb.AppendLine(SharedLibrary.Agents.ToolCallSequencing.DependentBatchingInstruction)
        sb.AppendLine()
        sb.AppendLine("Available tools:")

        If selectedTools IsNot Nothing Then
            Dim sortedTools = selectedTools.OrderBy(Function(t) t.ToolPriority).ToList()

            For Each tool In sortedTools
                If Not String.IsNullOrWhiteSpace(tool.ToolInstructionsPrompt) Then
                    sb.AppendLine()
                    sb.AppendLine($"- {tool.ToolInstructionsPrompt}")
                End If
            Next
        End If

        Return sb.ToString()
    End Function


    Private Sub AddToolByNameIfPresent(target As List(Of ModelConfig),
                                       source As IEnumerable(Of ModelConfig),
                                       toolName As String)
        If target Is Nothing OrElse source Is Nothing OrElse String.IsNullOrWhiteSpace(toolName) Then
            Return
        End If

        If target.Any(Function(t)
                          Return t IsNot Nothing AndAlso
                                 Not String.IsNullOrWhiteSpace(t.ToolName) AndAlso
                                 t.ToolName.Equals(toolName, StringComparison.OrdinalIgnoreCase)
                      End Function) Then
            Return
        End If

        Dim match = source.FirstOrDefault(
            Function(t)
                Return t IsNot Nothing AndAlso
                       Not String.IsNullOrWhiteSpace(t.ToolName) AndAlso
                       t.ToolName.Equals(toolName, StringComparison.OrdinalIgnoreCase)
            End Function)

        If match IsNot Nothing Then
            target.Add(match)
        End If
    End Sub

    Private Sub AddToolNameIfAvailable(target As List(Of String),
                                       availableTools As IEnumerable(Of ModelConfig),
                                       toolName As String)
        If target Is Nothing OrElse availableTools Is Nothing OrElse String.IsNullOrWhiteSpace(toolName) Then
            Return
        End If

        If target.Any(Function(name) name.Equals(toolName, StringComparison.OrdinalIgnoreCase)) Then
            Return
        End If

        If availableTools.Any(Function(t)
                                  Return t IsNot Nothing AndAlso
                                         Not String.IsNullOrWhiteSpace(t.ToolName) AndAlso
                                         t.ToolName.Equals(toolName, StringComparison.OrdinalIgnoreCase)
                              End Function) Then
            target.Add(toolName)
        End If
    End Sub

    Private Function HasToolName(selectedTools As IEnumerable(Of ModelConfig), toolName As String) As Boolean
        If selectedTools Is Nothing OrElse String.IsNullOrWhiteSpace(toolName) Then
            Return False
        End If

        Return selectedTools.Any(
            Function(t)
                Return t IsNot Nothing AndAlso
                       Not String.IsNullOrWhiteSpace(t.ToolName) AndAlso
                       t.ToolName.Equals(toolName, StringComparison.OrdinalIgnoreCase)
            End Function)
    End Function



    Private Function BuildToolArgumentsSignature(arguments As Dictionary(Of String, Object)) As String
        If arguments Is Nothing OrElse arguments.Count = 0 Then
            Return "{}"
        End If

        Try
            Dim normalized As New JObject()

            For Each key In arguments.Keys.OrderBy(Function(k) k, StringComparer.OrdinalIgnoreCase)
                Dim value = arguments(key)

                If TypeOf value Is JToken Then
                    normalized(key) = DirectCast(value, JToken)
                ElseIf value Is Nothing Then
                    normalized(key) = JValue.CreateNull()
                Else
                    normalized(key) = JToken.FromObject(value)
                End If
            Next

            Return normalized.ToString(Formatting.None)
        Catch
            Try
                Return JsonConvert.SerializeObject(arguments)
            Catch
                Return "{}"
            End Try
        End Try
    End Function

    Private Function BuildNormalizedToolCallSignature(toolCall As ToolCall) As String
        If toolCall Is Nothing Then Return ""

        Dim toolName As String = If(toolCall.ToolName, "").Trim()
        If toolName = "" Then Return ""

        ' Append-style operations are legitimately repeatable (each call mutates document
        ' state), so they must never collapse into a duplicate-replay. A unique salt keeps
        ' their signature distinct without affecting idempotent ops like replace/insert_*.
        If IsRepeatableAppendCall(toolCall) Then
            Return toolName.ToLowerInvariant() & "|append|" & Guid.NewGuid().ToString("N")
        End If

        Return toolName.ToLowerInvariant() & "|" & BuildToolArgumentsSignature(toolCall.Arguments)
    End Function

    Private Function IsRepeatableAppendCall(toolCall As ToolCall) As Boolean
        If toolCall Is Nothing OrElse toolCall.Arguments Is Nothing Then Return False

        Dim opValue As Object = Nothing
        If Not toolCall.Arguments.TryGetValue("op", opValue) OrElse opValue Is Nothing Then
            Return False
        End If

        Return System.Convert.ToString(opValue, Globalization.CultureInfo.InvariantCulture).
            Trim().Equals("append", StringComparison.OrdinalIgnoreCase)
    End Function

    Private Function IsDuplicateSuccessGuardedTool(toolName As String) As Boolean
        If String.IsNullOrWhiteSpace(toolName) Then
            Return False
        End If

        Dim normalizedToolName As String = toolName.Trim()

        For Each deliverableToolName As String In SharedLibrary.Agents.HostToolRegistration.GetDeliverableCapableToolNames(
            SharedLibrary.Agents.ToolingHostKind.Word)

            If normalizedToolName.Equals(deliverableToolName, StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If
        Next

        Select Case normalizedToolName.ToLowerInvariant()
            Case "complete_word_tables",
                 "process_word_document",
                 "word_write",
                 "word_markup"
                Return True
        End Select

        Return False
    End Function

    Private Function CloneToolResponseForDuplicateReplay(source As ToolResponse,
                                                         toolCall As ToolCall) As ToolResponse
        Return New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName,
            .Response = source.Response,
            .Success = source.Success,
            .ErrorMessage = source.ErrorMessage,
            .Timestamp = DateTime.Now,
            .OriginalCallJson = toolCall.RawJson,
            .ResultKind = source.ResultKind,
            .ErrorCode = source.ErrorCode,
            .ModelReplayContent = source.ModelReplayContent,
            .ModelReplaySummary = source.ModelReplaySummary,
            .WasCompactedForModelReplay = source.WasCompactedForModelReplay,
            .NormalizedCallSignature = source.NormalizedCallSignature,
            .WasDuplicateReplay = True
        }
    End Function

    Private Function TryBuildDuplicateSuccessfulToolReplay(toolCall As ToolCall,
                                                           normalizedCallSignature As String,
                                                           context As ToolExecutionContext,
                                                           ByRef replayResponse As ToolResponse) As Boolean
        replayResponse = Nothing

        If toolCall Is Nothing OrElse context Is Nothing Then
            Return False
        End If

        If String.IsNullOrWhiteSpace(normalizedCallSignature) Then
            Return False
        End If

        If Not IsDuplicateSuccessGuardedTool(toolCall.ToolName) Then
            Return False
        End If

        If context.AllToolResponses Is Nothing OrElse context.AllToolResponses.Count = 0 Then
            Return False
        End If

        For i As Integer = context.AllToolResponses.Count - 1 To 0 Step -1
            Dim prior As ToolResponse = context.AllToolResponses(i)

            If prior Is Nothing OrElse Not prior.Success Then Continue For
            If prior.WasDuplicateReplay Then Continue For

            If String.Equals(prior.NormalizedCallSignature, normalizedCallSignature, StringComparison.Ordinal) Then
                replayResponse = CloneToolResponseForDuplicateReplay(prior, toolCall)
                Return True
            End If
        Next

        Return False
    End Function

    Private Function BuildDuplicateSuccessfulToolCallGuardPrompt(toolName As String) As String
        Dim normalizedToolName As String = If(toolName, "").Trim()
        Dim toolLabel As String =
            If(normalizedToolName = "",
               "The previous tool call",
               "The tool '" & normalizedToolName & "'")

        Return "HOST DUPLICATE TOOL CALL GUARD: " &
               toolLabel &
               " already completed successfully earlier in this same run with the same arguments. " &
               "Do NOT call it again. Reuse that result. " &
               "If the task is now finished, provide the final answer. " &
               "Otherwise choose the next distinct tool step."
    End Function


    Private Function BuildExecutedToolSignature(toolCall As ToolCall, toolResponse As ToolResponse) As String
        Dim toolName As String = If(toolCall?.ToolName, "")
        Dim argsSig As String = BuildToolArgumentsSignature(toolCall?.Arguments)
        Dim successSig As String = If(toolResponse IsNot Nothing AndAlso toolResponse.Success, "ok", "err")
        Dim responseSig As String = If(toolResponse?.Response, "")
        Dim errorSig As String = If(toolResponse?.ErrorMessage, "")

        If responseSig.Length > 500 Then
            responseSig = responseSig.Substring(0, 500)
        End If

        If errorSig.Length > 300 Then
            errorSig = errorSig.Substring(0, 300)
        End If

        Return toolName & "|" & argsSig & "|" & successSig & "|" & responseSig & "|" & errorSig
    End Function

    Private Function IsEmptyJsRunResult(rawResponse As String) As Boolean
        Dim raw As String = If(rawResponse, "").Trim()

        If raw = "" OrElse raw = "{}" OrElse raw = "[]" OrElse raw.Equals("null", StringComparison.OrdinalIgnoreCase) Then
            Return True
        End If

        Try
            Dim token As JToken = JToken.Parse(raw)

            If TypeOf token Is JObject Then
                Dim obj = DirectCast(token, JObject)

                If Not obj.Properties().Any() Then
                    Return True
                End If

                Dim errorToken = obj("error")
                If errorToken IsNot Nothing AndAlso errorToken.Type <> JTokenType.Null AndAlso errorToken.ToString().Trim() <> "" Then
                    Return False
                End If

                Dim okToken = obj("ok")
                Dim resultToken = obj("result")

                Dim okValue As Boolean = False
                If okToken IsNot Nothing Then
                    If okToken.Type = JTokenType.Boolean Then
                        okValue = okToken.Value(Of Boolean)()
                    Else
                        Boolean.TryParse(okToken.ToString(), okValue)
                    End If
                End If

                If okValue Then
                    If resultToken Is Nothing OrElse resultToken.Type = JTokenType.Null Then
                        Return True
                    End If

                    If resultToken.Type = JTokenType.String AndAlso resultToken.ToString().Trim() = "" Then
                        Return True
                    End If
                End If
            End If
        Catch
        End Try

        Return False
    End Function

    Private Function RegisterToolFailureLoopState(toolCall As ToolCall, toolResponse As ToolResponse, context As ToolExecutionContext) As Boolean
        If toolCall Is Nothing OrElse toolResponse Is Nothing OrElse context Is Nothing Then
            Return False
        End If

        If toolResponse.Success Then
            context.ConsecutiveFailedToolName = ""
            context.ConsecutiveFailedToolCount = 0
            Return False
        End If

        Dim failedName As String = If(toolCall.ToolName, "").Trim()

        If failedName.Equals(context.ConsecutiveFailedToolName, StringComparison.OrdinalIgnoreCase) Then
            context.ConsecutiveFailedToolCount += 1
        Else
            context.ConsecutiveFailedToolName = failedName
            context.ConsecutiveFailedToolCount = 1
        End If

        Return context.ConsecutiveFailedToolCount >= context.ConsecutiveToolFailureAbortThreshold
    End Function



    Private Shared Function GetWordSettingString(name As String) As String
        Try
            Dim value = My.Settings(name)
            Return If(value, "").ToString()
        Catch
            Return ""
        End Try
    End Function

    Private Shared Function GetWordSettingBoolean(name As String, defaultValue As Boolean) As Boolean
        Try
            Dim value = My.Settings(name)
            If value Is Nothing Then Return defaultValue
            Return CBool(value)
        Catch
            Return defaultValue
        End Try
    End Function

    Private Shared Sub SetWordSettingValue(name As String, value As Object)
        Try
            My.Settings(name) = value
        Catch
        End Try
    End Sub

    Friend Shared Function SplitPersistedToolNames(raw As String) As List(Of String)
        If String.IsNullOrWhiteSpace(raw) Then
            Return New List(Of String)()
        End If

        Return raw.Split("|"c).
            Select(Function(s) s.Trim()).
            Where(Function(s) s.Length > 0).
            Distinct(StringComparer.OrdinalIgnoreCase).
            ToList()
    End Function

    Private Shared Function JoinPersistedToolNames(names As IEnumerable(Of String)) As String
        If names Is Nothing Then Return ""
        Return String.Join("|", names.Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                                      Select(Function(s) s.Trim()).
                                      Distinct(StringComparer.OrdinalIgnoreCase))
    End Function


    Private Shared Function DeduplicateToolsByName(tools As IEnumerable(Of ModelConfig)) As List(Of ModelConfig)
        Dim result As New List(Of ModelConfig)()
        Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        If tools Is Nothing Then Return result

        For Each tool In tools
            If tool Is Nothing OrElse String.IsNullOrWhiteSpace(tool.ToolName) Then Continue For
            If seen.Add(tool.ToolName.Trim()) Then
                result.Add(tool)
            End If
        Next

        Return result
    End Function


    Private Function HasSuccessfulToolResponses(context As ToolExecutionContext) As Boolean
        If context Is Nothing OrElse context.AllToolResponses Is Nothing Then
            Return False
        End If

        For Each tr In context.AllToolResponses
            If tr IsNot Nothing AndAlso tr.Success Then
                Return True
            End If
        Next

        Return False
    End Function

    Private Function BuildInvalidTurnSignature(responseText As String, invalidReason As String) As String
        Dim raw As String = If(responseText, "")
        Dim reason As String = If(invalidReason, "").Trim().ToLowerInvariant()
        Dim hashText As String

        Using sha = System.Security.Cryptography.SHA256.Create()
            Dim bytes = sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(raw))
            hashText = BitConverter.ToString(bytes).Replace("-", "").ToLowerInvariant()
        End Using

        Return reason & "|" & hashText
    End Function

    Private Async Function RunForcedNoToolFinalizationAfterRepeatedInvalidTurnAsync(
        context As ToolExecutionContext,
        enhancedSysPrompt As String,
        fullUserPrompt As String,
        otherPrompt As String,
        fileObject As Object,
        useSecondAPI As Boolean,
        hideSplash As Boolean) As Task(Of String)

        INI_APICall_ToolInstructions_2 = ""

        Dim forcedSysPrompt As String =
            enhancedSysPrompt & Environment.NewLine & Environment.NewLine &
            "IMPORTANT: A successful tool step already completed, but the model then repeated an invalid non-final response. " &
            "Do NOT call any more tools in this turn. " &
            "Do NOT repeat the previous invalid response format. " &
            "Do NOT output internal protocol content, intermediate control text, raw structured data, or non-final status text. " &
            "Based only on the tool results already gathered, produce a concise user-facing answer now and append exactly one valid TASK_STATUS footer with status 'complete' or 'blocked'."

        context.LogWarn(
            "Escalating to forced no-tool finalization after repeated identical invalid turn.",
            details:=$"invalidTurnReason={If(context.ForceNoToolFinalizationReason, "")}; host={context.HostKind}")

        ToolingFileLogger.LogStep("Forcing final LLM call without tools after repeated invalid turn")
        ToolingFileLogger.LogPreMainLlmCallSnapshot()

        Try
            Dim forcedFinalResponse As String = Await LLM(
                forcedSysPrompt,
                fullUserPrompt,
                "", "", 0,
                useSecondAPI,
                hideSplash,
                otherPrompt,
                fileObject,
                True)

            If Not String.IsNullOrWhiteSpace(forcedFinalResponse) Then
                ToolingFileLogger.LogRawResponseStub("Main LLM() - Forced Final After Invalid Turn", forcedFinalResponse)
            Else
                ToolingFileLogger.LogWarn(
                    "Empty forced-final response after repeated invalid turn.",
                    details:=$"host={context.HostKind}")
            End If

            Return If(forcedFinalResponse, "")
        Catch ex As Exception
            context.LogError($"Error during forced no-tool finalization: {ex.Message}", ex:=ex)
            Return ""
        End Try
    End Function


    Private Function TryPrepareToolForNextTurnAfterExposureMiss(toolName As String,
                                                            context As ToolExecutionContext) As Boolean
        If context Is Nothing OrElse String.IsNullOrWhiteSpace(toolName) Then
            Return False
        End If

        If Not IsToolAllowedForCurrentContext(toolName, context) Then
            Return False
        End If

        Dim loaded As ModelConfig = EnsureVisibleToolLoaded(toolName.Trim(), context)
        Return loaded IsNot Nothing
    End Function


#End Region


End Class
