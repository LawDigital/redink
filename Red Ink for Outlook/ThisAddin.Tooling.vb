' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddin.Tooling.vb
' Purpose: Core model-agnostic tooling loop for LLM tool/function calling.
'          Orchestrates tool selection, call detection/extraction, execution, and response injection.
'
' Architecture:
'  - Tooling Execution Loop (ExecuteToolingLoop):
'      - Builds system prompt augmentation via BuildToolInstructionsPrompt.
'      - Injects model-specific tool definitions into INI_APICall_ToolInstructions_2.
'      - Iteratively calls LLM(...) until max iterations or no tool calls detected.
'      - On each iteration:
'          - ContainsToolCalls(): Regex-based tool call detection.
'          - ExtractToolCalls(): JSON parsing + extraction map for tool call tokens.
'          - ExecuteToolCall(): Per-tool execution (internal tools, external models, knowledge stores).
'          - BuildToolResponsesForModel(): Packages ToolResponse objects for next LLM iteration.
'          - Assigns response payload to INI_APICall_ToolResponses_2.
'      - Handles duplicate execution detection, consecutive failures, and early abort conditions.
'  - Tool Execution Modes:
'      - Internal Web Tool: retrieve_web_content (URL-based content fetching via WebView2).
'      - Internal M365 Tools: Bridge via ExecuteInternalM365Tool to SharedLibrary.M365ToolService.
'      - Internal Knowledge Tools: Execute via knowledge store APIs with trigger building.
'      - External Tools: Apply ModelConfig, prepare API payload, force JSON response, restore settings.
'  - Tool Selection & Persistence:
'      - LoadToolingServices(): Loads tools from INI_SpecialServicePath.
'      - Adds internal web retrieval tool via GetInternalWebTool.
'      - Persists selections through My.Settings.SelectedToolNames.
'      - Restores via LoadPersistedToolSelection.
'  - User Request & Task Tracking:
'      - ResolveLatestUserRequestRaw(): Extracts latest user turn from dialog/prompt.
'      - BuildLatestUserRequestMetadataBlock(): Authoritative [CURRENT_USER_REQUEST] block.
'      - ResolveOptionalHostTaskSummary(): Optional task summary for host context.
'      - BuildCompletedFactsPromptBlock(): Injects completed tool responses as facts for grounding.
'  - Diagnostics & Logging:
'      - ToolingFileLogger: Single-file per-run logging when INI_APIDebug is enabled.
'      - Optional LogWindow UI when enabled.
'      - BuildPromptDiagnosticStub(): SHA256-based diagnostic fingerprints for prompts.
'  - State Management (ToolExecutionContext):
'      - Tracks selected tools, iteration counters, all responses, cancellation, logs.
'      - Holds active context during ExecuteToolingLoop (used by ApDashboardLog routing).
'      - Persists SequencingState for memory grounding and tool call sequencing decisions.
'
' Data Classes:
'  - ToolCall: Represents single tool call from LLM with CallId, ToolName, Arguments, RawJson.
'  - ToolResponse: Outcome of tool execution with CallId, ToolName, Response, Success, ErrorMessage, Timestamp.
' =============================================================================


Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.Reflection
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Web.WebView2.Core
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods

''' <summary>
''' Provides tooling support helpers for model-agnostic tool/function calling in LLM interactions.
''' </summary>
Partial Public Class ThisAddIn

    ''' <summary>User-friendly name for the tooling feature.</summary>
    Public Const ToolFriendlyName As String = "Agents"

    ''' <summary>Auto-close delay for tooling log window.</summary>
    Public Shared Property ToolingLog_AutoCloseDefaultSeconds As Integer = 30

    Private Const SubAgentLargeToolResponseThresholdChars As Integer = 30000
    Private Const SubAgentLargeToolResponseExcerptChars As Integer = 8000

    Private Const MaxDownloadedWebFileBytes As Long = 50L * 1024L * 1024L
    Private Const WebToolTimeoutSeconds As Integer = 45

    ''' <summary>Selected tool names for persistence.</summary>
    Public Shared Property SelectedToolNames As List(Of String) = New List(Of String)()



    ''' <summary>
    ''' Suffix appended to internal tool descriptions when shown to the user.
    ''' </summary>
    Private Const InternalToolSuffix As String = ""

    ''' <summary>
    ''' Maximum tool loop iterations used by prompt-building helpers.
    ''' </summary>
    Public MaxToolIterations As Integer = Agents.ToolingConstants.DefaultMaxToolIterations

    ''' <summary>
    ''' Holds the active ToolExecutionContext during an ExecuteToolingLoop run.
    ''' Used by ApDashboardLog to route messages to the Chat Agent's LogWindow.
    ''' </summary>
    Private _activeToolingContext As ToolExecutionContext = Nothing



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
    Public Class ToolResponse

        ''' <summary>Tool call identifier used to correlate call and response objects.</summary>
        Public Property CallId As String

        ''' <summary>Name of the tool that was executed.</summary>
        Public Property ToolName As String

        ''' <summary>Raw response returned by the tool execution.</summary>
        Public Property Response As String

        ''' <summary>True if the tool execution completed successfully; otherwise False.</summary>
        Public Property Success As Boolean

        ''' <summary>Error message populated when <see cref="Success"/> is False.</summary>
        Public Property ErrorMessage As String

        ''' <summary>Timestamp captured at response creation time.</summary>
        Public Property Timestamp As DateTime

        ''' <summary>Original tool call JSON as extracted from the LLM response.</summary>
        Public Property OriginalCallJson As String

        Public Property ResultKind As String
        Public Property ErrorCode As String

        Public Property ModelReplayContent As String
        Public Property ModelReplaySummary As String
        Public Property WasCompactedForModelReplay As Boolean

        Public Property NormalizedCallSignature As String
        Public Property WasDuplicateReplay As Boolean

        ''' <summary>True when a repair-loop advisor determined the failure is terminal (budget exhausted or non-recoverable).</summary>
        Public Property RepairLoopTerminal As Boolean

        ''' <summary>Human-readable reason for <see cref="RepairLoopTerminal"/>, surfaced to abort/finalization handling.</summary>
        Public Property RepairLoopTerminalReason As String

        ''' <summary>
        ''' Initializes a new tool response instance with default success state.
        ''' </summary>
        Public Sub New()
            Timestamp = DateTime.Now
            Success = True
        End Sub
    End Class


#End Region



    Private Function ResolveOptionalHostTaskSummary(latestUserRequestRaw As String,
                                                    userText As String,
                                                    otherPrompt As String,
                                                    fullPromptOverride As String,
                                                    insertDocs As String,
                                                    slideInsert As String,
                                                    bubblesText As String) As String
        Dim summary As String =
            BuildToolSelectionHintText(
                userText,
                fullPromptOverride,
                otherPrompt,
                insertDocs,
                slideInsert,
                bubblesText).Trim()

        If summary = "" Then
            Return ""
        End If

        If String.Equals(summary, If(latestUserRequestRaw, "").Trim(), StringComparison.Ordinal) Then
            Return ""
        End If

        Return summary
    End Function



    Private Function BuildCompletedToolResultSummaryBlock(context As ToolExecutionContext,
                                                         Optional maxItems As Integer = 3) As String
        If context Is Nothing OrElse context.AllToolResponses Is Nothing Then
            Return ""
        End If

        Dim summaries As New List(Of String)()

        For Each resp In context.AllToolResponses
            If resp Is Nothing OrElse Not resp.Success Then Continue For

            Dim summary As String = Regex.Replace(If(BuildToolReplaySummary(resp), ""), "\s+", " ").Trim()
            If summary = "" Then Continue For

            summaries.Add("- " & summary)

            If summaries.Count >= maxItems Then
                Exit For
            End If
        Next

        If summaries.Count = 0 Then
            Return ""
        End If

        Dim sb As New StringBuilder()
        sb.AppendLine("<COMPLETED_TOOL_RESULTS>")
        For Each summary In summaries
            sb.AppendLine(summary)
        Next
        sb.AppendLine("</COMPLETED_TOOL_RESULTS>")
        Return sb.ToString().TrimEnd()
    End Function

    Private Function BuildDeliverableCompletionContinuationBlock(context As ToolExecutionContext) As String
        If context Is Nothing Then
            Return ""
        End If

        If String.IsNullOrWhiteSpace(context.LatestUserRequestRaw) Then
            Return ""
        End If

        Dim completedBlock As String = BuildCompletedToolResultSummaryBlock(context)
        Dim needsDeliverable As Boolean =
            context.SequencingState IsNot Nothing AndAlso
            context.SequencingState.RequestRequiresCreatedDeliverable AndAlso
            Not SharedLibrary.Agents.ToolCallSequencing.HasProducedUserDeliverable(context.SequencingState)

        If Not needsDeliverable AndAlso String.IsNullOrWhiteSpace(completedBlock) Then
            Return ""
        End If

        Dim sb As New StringBuilder()
        sb.AppendLine("[HOST REQUEST CONTINUITY]")
        sb.AppendLine("The original latest user request remains authoritative.")
        sb.AppendLine("<LATEST_USER_REQUEST_RAW>")
        sb.AppendLine(context.LatestUserRequestRaw)
        sb.AppendLine("</LATEST_USER_REQUEST_RAW>")

        If Not String.IsNullOrWhiteSpace(completedBlock) Then
            sb.AppendLine(completedBlock)
        End If

        If needsDeliverable Then
            sb.AppendLine("<REMAINING_REQUESTED_DELIVERABLES>")
            sb.AppendLine("A requested deliverable artifact has not yet been actually produced.")
            If context.SequencingState.LastToolProducesIntermediateData Then
                sb.AppendLine("The latest successful tool result is preparatory or intermediate data only.")
            End If
            If Not String.IsNullOrWhiteSpace(context.SequencingState.RequestDeliverableSummary) Then
                sb.AppendLine("Requested deliverable: " & context.SequencingState.RequestDeliverableSummary)
            End If
            sb.AppendLine("</REMAINING_REQUESTED_DELIVERABLES>")
            sb.AppendLine("Do not finalize yet. Use an appropriate creation, write, export, or save tool before finalizing, or explain briefly why the deliverable cannot be created.")
        Else
            sb.AppendLine("Continue with any remaining requested deliverables. Do not finalize until the full request is complete, or explain briefly why it cannot be completed.")
        End If

        Dim consolidationGuidance As String =
            SharedLibrary.Agents.ToolCallSequencing.BuildConsolidatableToolGuidance(context.SequencingState)
        If Not String.IsNullOrWhiteSpace(consolidationGuidance) Then
            sb.AppendLine(consolidationGuidance)
        End If

        sb.AppendLine("[/HOST REQUEST CONTINUITY]")
        Return sb.ToString().TrimEnd()
    End Function

    Private Function BuildEmptyResponseAfterProgressRecoveryPrompt(context As ToolExecutionContext) As String
        Dim sb As New StringBuilder()
        sb.AppendLine("HOST EMPTY-RESPONSE RECOVERY: The previous model turn was empty after successful partial progress.")

        If context IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(context.LatestUserRequestRaw) Then
            sb.AppendLine("<LATEST_USER_REQUEST_RAW>")
            sb.AppendLine(context.LatestUserRequestRaw)
            sb.AppendLine("</LATEST_USER_REQUEST_RAW>")
        End If

        Dim completedBlock As String = BuildCompletedToolResultSummaryBlock(context)
        If Not String.IsNullOrWhiteSpace(completedBlock) Then
            sb.AppendLine(completedBlock)
        End If

        Dim requiresMissingOutputFileRecovery As Boolean =
            context IsNot Nothing AndAlso
            context.SequencingState IsNot Nothing AndAlso
            context.SequencingState.RequestRequiresCreatedDeliverable AndAlso
            context.SequencingState.LastToolProducesIntermediateData AndAlso
            Not SharedLibrary.Agents.ToolCallSequencing.HasProducedUserDeliverable(context.SequencingState)

        If requiresMissingOutputFileRecovery Then
            sb.AppendLine("Intermediate data is available. The requested output file has not been created yet. Call an appropriate create/save/export tool now. Do not finalise unless the file is created or no suitable tool exists.")
        ElseIf context IsNot Nothing AndAlso
               context.SequencingState IsNot Nothing AndAlso
               context.SequencingState.RequestRequiresCreatedDeliverable AndAlso
               Not SharedLibrary.Agents.ToolCallSequencing.HasProducedUserDeliverable(context.SequencingState) Then

            sb.AppendLine("A requested deliverable artifact has not yet been actually produced.")
            sb.AppendLine("Do not finalize yet. Continue with the next appropriate creation, write, export, or save step, or return a valid blocked answer only if the deliverable cannot be created.")
        Else
            sb.AppendLine("Continue with any remaining requested deliverables, or return a valid final answer only if the full request is complete.")
        End If

        Return sb.ToString().TrimEnd()
    End Function

    Private Async Function ResolveToolingUserLanguageAsync(userText As String,
                                                           otherPrompt As String,
                                                           fullPromptOverride As String,
                                                           useSecondAPI As Boolean,
                                                           hideSplash As Boolean,
                                                           cancellationToken As System.Threading.CancellationToken) As Task(Of String)
        Dim sourceText As String =
            ResolveLatestUserRequestRaw(userText, otherPrompt, fullPromptOverride)

        sourceText = If(sourceText, "").Trim()
        If sourceText = "" Then Return ""

        If sourceText.Length > 4000 Then
            sourceText = sourceText.Substring(0, 4000)
        End If

        Dim detectionSystemPrompt As String =
            "Determine the language in which the assistant must answer the user's latest request. " &
            "Return ONLY valid JSON in the form {""language"":""...""}. " &
            "Use a concrete runtime language value suitable for later localization, preferably a BCP-47 tag when clear. " &
            "Do not add explanations."

        Dim detectionUserPrompt As String =
            "<USER_ENTRY>" & sourceText & "</USER_ENTRY>"

        Try
            Dim raw As String = Await LLM(
                detectionSystemPrompt,
                detectionUserPrompt,
                "", "", 0,
                useSecondAPI,
                hideSplash,
                "",
                "",
                cancellationToken,
                True,
                False)

            If String.IsNullOrWhiteSpace(raw) Then Return ""

            Try
                Dim obj As JObject = JObject.Parse(raw)
                Return If(obj.Value(Of String)("language"), "").Trim()
            Catch
                Return raw.Trim().Trim(""""c)
            End Try
        Catch
            Return ""
        End Try
    End Function



#Region "Execute Tooling"


    ''' <summary>
    ''' Derives a tool-agnostic, user-facing progress line for an upcoming tool call,
    ''' preferring the tool's configured ModelDescription over its raw name.
    ''' </summary>
    Private Function BuildFriendlyProgressText(tc As ToolCall, context As ToolExecutionContext) As String
        Dim toolName As String = If(tc?.ToolName, "").Trim()
        If toolName = "" Then Return "Working..."

        Dim friendly As String = toolName
        Try
            Dim cfg = If(context?.SelectedTools, New List(Of ModelConfig)()).
                FirstOrDefault(Function(t) t IsNot Nothing AndAlso
                                         String.Equals(t.ToolName, toolName, StringComparison.OrdinalIgnoreCase))
            If cfg IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(cfg.ModelDescription) Then
                friendly = cfg.ModelDescription.Trim()
            End If
        Catch
        End Try

        Return $"Working on: {friendly}..."
    End Function


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
    ''' <summary>
    ''' Executes an iterative tool-enabled LLM loop until either no tool calls are detected, the maximum iteration count is reached,
    ''' or the user cancels. Tool call detection/extraction and response injection are controlled by the active tooling model config.
    ''' </summary>
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
        Optional binaryOutputDirectory As String = Nothing,
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
        Optional finalResponseContract As SharedLibrary.Agents.ToolingFinalResponseContract = SharedLibrary.Agents.ToolingFinalResponseContract.UserFacingTaskStatus,
        Optional progressSink As Action(Of String) = Nothing) As Task(Of String)

        ' Check for power transition BEFORE starting (matches RunLlmAsync pattern)
        If System.Threading.Interlocked.CompareExchange(powerChanging, 0, 0) <> 0 Then
            Return "Operation cancelled due to power transition."
        End If

        ToolingFileLogger.StartSession()

        ' Per-turn Chat Agent deliverables reset (prevents prior-turn output files
        ' from being re-presented). Sub-agents inherit parent state and must not
        ' reset. AutoPilot is single-threaded per Q4.
        If Not subAgentMode AndAlso _chatAgentActive Then
            ResetChatAgentDeliverableTrackingForNewTurn()
        End If

        Dim parentToolingContext = _activeToolingContext
        Dim workflowScope As IDisposable = Nothing
        Dim acceptedFinalStatus As String = ""

        Dim fullAllowedTools As List(Of ModelConfig) =
            If(selectedTools, New List(Of ModelConfig)()).
                Where(Function(t) t IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(t.ToolName)).
                GroupBy(Function(t) t.ToolName, StringComparer.OrdinalIgnoreCase).
                Select(Function(g) g.First()).
                ToList()

        selectedTools = fullAllowedTools

        Try
            SharedLibrary.Agents.HostToolRegistration.RegisterResolvedInternalTools(
                SharedLibrary.Agents.ToolingHostKind.Outlook,
                fullAllowedTools)
        Catch ex As Exception
            ToolingFileLogger.LogWarn("Failed to register selected Outlook tooling tools.", ex:=ex)
        End Try

        Try
            SharedLibrary.Agents.SkillInvokeTool.CurrentHostProvider =
                Function()
                    If _apActive Then
                        Return "Outlook AutoPilot"
                    End If

                    If _chatAgentActive Then
                        Return "Outlook Local Chat"
                    End If

                    Return "Outlook"
                End Function
        Catch ex As Exception
            ToolingFileLogger.LogWarn("Failed to set skill host provider for Outlook.", ex:=ex)
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
                   hideSplash,
                   cancellationToken))

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

        Dim authoritativeRegistrySource As SharedLibrary.Agents.ToolRegistry = Nothing

        If subAgentMode AndAlso subAgentAuthoritativeRegistry IsNot Nothing Then
            authoritativeRegistrySource = subAgentAuthoritativeRegistry.Snapshot()
        Else
            authoritativeRegistrySource = SharedLibrary.Agents.ToolRegistryBuilder.FromModelConfigs(fullAllowedTools, "selected")

            ' Only register the skills/agents the user actually selected.
            ' This fixes AutoPilot making unselected skills/agents available anyway.
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
        context.HostKind = "Outlook"
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
            ' The LogWindow must be created on the Outlook UI/STA thread.
            ' ExecuteToolingLoop may run on a thread-pool continuation, so always
            ' marshal creation/show through the captured UI SynchronizationContext.
            Dim createdForm As LogWindow = Nothing
            Dim createError As Exception = Nothing

            Dim createOnUi As Action =
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

        context.ProgressSink = progressSink
        context.ReportProgress("Getting started...")

        context.Log("Starting tooling session...")
        If selectedTools IsNot Nothing Then
            context.Log($"Selected tools: {String.Join(", ", selectedTools.Select(Function(t) t.ToolName))}")
        Else
            context.Log("Selected tools: (none)")
        End If

        ' Calculate effective timeout for LLM calls.
        ' INI_Timeout / INI_Timeout_2 are stored in milliseconds; convert to whole
        ' seconds here because the guard below uses TimeSpan.FromSeconds. A default
        ' of 30 seconds is applied when the configured value is zero/invalid.
        Dim configuredTimeoutMs As Integer = CInt(If(useSecondAPI, INI_Timeout_2, INI_Timeout))
        If configuredTimeoutMs <= 0 Then configuredTimeoutMs = 30000
        Dim effectiveTimeout As Integer = System.Math.Max(1, CInt(System.Math.Ceiling(configuredTimeoutMs / 1000.0)))

        Try
            Await ResolveMemoryGroundingModeAsync(
                context,
                userText,
                otherPrompt,
                fullPromptOverride,
                useSecondAPI,
                hideSplash,
                cancellationToken,
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
            Dim baseSysPrompt As String = sysCommand

            ' B1 progress (major steps only): the model announces significant new phases via the
            ' report_progress tool (a real tool call authored in the dialogue language). Deterministic
            ' host-side progress (A) still surfaces for every tool call, so A always appears even if
            ' the model never calls report_progress.
            baseSysPrompt &= " " & BuildMajorStepProgressToolInstruction()

            If Not String.IsNullOrWhiteSpace(bubblesText) Then
                baseSysPrompt &= " " & SP_Add_BubblesExtract
            End If

            If doTPMarkup Then
                baseSysPrompt &= " " & SP_Add_Revisions
            End If

            If Not String.IsNullOrWhiteSpace(slideDeck) Then
                baseSysPrompt &= " " & SP_Add_Slides
            ElseIf Not noFormatting Then
                If keepFormat Then
                    baseSysPrompt &= " " & SP_Add_KeepHTMLIntact
                Else
                    baseSysPrompt &= " " & SP_Add_KeepInlineIntact
                End If
            End If

            If doMyStyle AndAlso Not String.IsNullOrWhiteSpace(myStyleInsert) Then
                baseSysPrompt &= " " & myStyleInsert
            End If

            If DoChart Then
                baseSysPrompt &= " " & SP_Add_Chart
            End If

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

            ' Language contract — Mandatory rule that the model must produce final
            ' user-facing prose in the detected user language regardless of the
            ' language of guard prompts, host messages, or tool output (P1 fix).
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
            Dim abortDueToToolError As Boolean = False
            Dim abortToolErrorMessage As String = ""
            Dim abortToolName As String = ""
            Dim abortToolParamSummary As String = ""
            Dim abortToolRawCallJson As String = ""
            Dim abortFactsPrompt As String = ""

            Dim abortDueToRepeatedToolLoop As Boolean = False

            ' Tracks the currently exposed tool set so the (potentially large) tool
            ' definitions block is only rebuilt when the exposed tools actually change
            ' (e.g. after tool_loader / lazy-load). Null forces the first build.
            Dim lastExposedToolSignature As String = Nothing

            Dim noSelectedText As Boolean = String.IsNullOrWhiteSpace(userText)
            While iteration < context.MaxIterations AndAlso Not context.IsCancelled

                ' Check for cancellation token at each iteration
                If cancellationToken.IsCancellationRequested Then
                    context.LogWarn("Cancellation requested via token")
                    Exit While
                End If

                ' Check for power transition during loop (matches RunLlmAsync pattern)
                If System.Threading.Interlocked.CompareExchange(powerChanging, 0, 0) <> 0 Then
                    context.LogWarn("Power transition detected, aborting tooling loop")
                    ToolingFileLogger.EndSession(False, "Cancelled due to power transition")
                    Return "Operation cancelled due to power transition."
                End If

                iteration += 1
                context.CurrentIteration = iteration
                context.Log($"--- Iteration {iteration} of {context.MaxIterations} ---")

                context.Log("Calling LLM...", "llm")

                If Not String.IsNullOrWhiteSpace(fullPromptOverride) Then
                    fullUserPrompt = fullPromptOverride
                Else
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

                Dim currentExposedToolSignature As String =
                    If(context.SelectedTools Is Nothing, "",
                       String.Join("|", context.SelectedTools.
                           Where(Function(t) t IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(t.ToolName)).
                           Select(Function(t) t.ToolName.Trim())))

                If Not String.Equals(currentExposedToolSignature, lastExposedToolSignature, StringComparison.Ordinal) Then
                    INI_APICall_ToolInstructions_2 = BuildToolInstructionsForModel(context.SelectedTools, context.ToolingModel)
                    lastExposedToolSignature = currentExposedToolSignature
                End If

                ' Payload-size diagnostic: measures per-iteration model input growth so the
                ' large-result referencing and tool-definition rebuild gating can be verified.
                ToolingFileLogger.LogStep(
                    $"Payload sizes (chars): toolInstructions={If(INI_APICall_ToolInstructions_2, "").Length}; " &
                    $"toolResponses={If(INI_APICall_ToolResponses_2, "").Length}; " &
                    $"combined={If(INI_APICall_ToolInstructions_2, "").Length + If(INI_APICall_ToolResponses_2, "").Length}; " &
                    $"iteration={iteration}")

                Debug.WriteLine(
                    $"[TOOLING-PAYLOAD] iteration={iteration}; " &
                    $"toolInstructions={If(INI_APICall_ToolInstructions_2, "").Length}; " &
                    $"toolResponses={If(INI_APICall_ToolResponses_2, "").Length}; " &
                    $"combined={If(INI_APICall_ToolInstructions_2, "").Length + If(INI_APICall_ToolResponses_2, "").Length}")


                Dim effectiveSysPrompt As String = enhancedSysPrompt
                Dim effectiveUserPrompt As String = fullUserPrompt

                If Not subAgentMode Then
                    Dim runtimeContextBlock As String = BuildRuntimeContextPromptBlock(context)
                    If Not String.IsNullOrWhiteSpace(runtimeContextBlock) Then
                        effectiveSysPrompt &= Environment.NewLine & Environment.NewLine & runtimeContextBlock
                    End If

                    If SharedLibrary.Agents.ToolingFinalResponseContractHelpers.RequiresTaskStatusFooter(context.FinalResponseContract) Then
                        Dim postToolContinuationBlock As String =
                            BuildPostToolContinuationBlock(context)

                        If Not String.IsNullOrWhiteSpace(postToolContinuationBlock) Then
                            effectiveSysPrompt &= Environment.NewLine & Environment.NewLine & postToolContinuationBlock
                            LogLatestUserRequestDiagnostic(context, "continuation")
                        End If
                    End If
                End If


                If Not String.IsNullOrWhiteSpace(context.PendingContinuationGuardPrompt) Then
                    effectiveSysPrompt &= Environment.NewLine & Environment.NewLine & context.PendingContinuationGuardPrompt

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

                    effectiveUserPrompt = fullUserPrompt & Environment.NewLine & guardBlock.ToString()

                    LogLatestUserRequestDiagnostic(context, "repair")
                    context.LogWarn("Applying host-side continuation guard.", visibleToUser:=False)
                    context.PendingContinuationGuardPrompt = ""
                    context.PendingRejectedAssistantTurn = ""
                    context.PendingGuardTitle = ""
                    context.PendingRejectedTurnExplanation = ""
                End If

                LogFinalResponseContractDiagnostics(context, effectiveSysPrompt)

                ' Create linked cancellation with timeout (matches RunLlmAsync pattern)
                Using timeoutCts As New System.Threading.CancellationTokenSource()
                    Dim totalTimeout = effectiveTimeout + 60 ' 60 second buffer
                    timeoutCts.CancelAfter(TimeSpan.FromSeconds(totalTimeout))

                    Using combinedCts As System.Threading.CancellationTokenSource =
                        System.Threading.CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, timeoutCts.Token)

                        Try
                            currentResponse = Await LLM(
                                effectiveSysPrompt,
                                effectiveUserPrompt,
                                "", "", 0,
                                useSecondAPI,
                                hideSplash,
                                otherPrompt,
                                fileObject,
                                combinedCts.Token,
                                True, True, binaryOutputDirectory:=binaryOutputDirectory)

                        Catch ex As OperationCanceledException When timeoutCts.IsCancellationRequested
                            context.LogError($"LLM call timed out after {totalTimeout}s")
                            ToolingFileLogger.EndSession(False, $"Timeout after {totalTimeout}s")
                            Return "Operation timed out. Please try again with a shorter prompt or different model."

                        Catch ex As OperationCanceledException
                            If System.Threading.Interlocked.CompareExchange(powerChanging, 0, 0) <> 0 Then
                                context.LogWarn("Cancelled due to power transition")
                                ToolingFileLogger.EndSession(False, "Cancelled due to power transition")
                                Return "Operation cancelled due to power transition."
                            End If
                            context.LogWarn("Operation was cancelled")
                            ToolingFileLogger.EndSession(False, "Cancelled by user")
                            Return "Operation was canceled by the user."
                        End Try
                    End Using
                End Using

                ToolingFileLogger.LogRawResponseStub("Main LLM()", currentResponse)

                If String.IsNullOrWhiteSpace(currentResponse) Then
                    Dim lastSuccess = GetLastSuccessfulToolResponse(context)

                    If Not subAgentMode AndAlso
                       lastSuccess IsNot Nothing AndAlso
                       context.SequencingState IsNot Nothing AndAlso
                       context.SequencingState.RequestRequiresCreatedDeliverable AndAlso
                       SharedLibrary.Agents.ToolCallSequencing.IsSuccessfulDeliverableResult(
                           If(lastSuccess.Response, "")) Then

                        context.EmptyMainModelResponse = False
                        currentResponse = Await BuildLocalizedCreatedDeliverableSuccessMessageAsync(
                            context,
                            lastSuccess,
                            useSecondAPI,
                            hideSplash,
                            cancellationToken)

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
                                context.ToolingModel,
                                compactForSubAgent:=False)

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
                                hideSplash,
                                cancellationToken)

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
                                context.ToolingModel,
                                compactForSubAgent:=False)

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
                            context.ToolingModel,
                            compactForSubAgent:=False)

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
                        If context.IsCancelled OrElse cancellationToken.IsCancellationRequested Then Exit For

                        ' B1 major-step announcement: surface the model-authored heading and
                        ' acknowledge the call as a successful no-op. Kept out of sequencing,
                        ' exposure, and duplicate-detection so it never counts as real work.
                        If IsReportProgressToolName(tc.ToolName) Then
                            Dim progressNote As String = ""
                            If tc.Arguments IsNot Nothing AndAlso
                               tc.Arguments.ContainsKey("note") AndAlso
                               tc.Arguments("note") IsNot Nothing Then
                                progressNote = tc.Arguments("note").ToString()
                            End If

                            context.ReportMajorProgress(progressNote)

                            Dim progressResponse As New ToolResponse() With {
                                .CallId = tc.CallId,
                                .ToolName = tc.ToolName,
                                .Success = True,
                                .Response = "{""ok"":true}",
                                .OriginalCallJson = tc.RawJson
                            }
                            context.AllToolResponses.Add(progressResponse)

                            context.Log("Progress update reported to the user.", "diag")
                            Continue For
                        End If

                        context.ReportProgress(BuildFriendlyProgressText(tc, context))

                        If Not subAgentMode AndAlso
                               context.SequencingState IsNot Nothing AndAlso
                               context.SequencingState.RequiresParentRecovery Then

                            Dim failedToolName As String = context.SequencingState.LastToolName
                            context.SequencingState.NoteRecoveryByLaterToolCall(tc.ToolName)

                            context.Log($"Recovered from prior skipped tool failure by issuing '{tc.ToolName}'.", "success")
                            ToolingFileLogger.LogStep(
                                    $"Skipped agent failure recovered by later parent tool call. host={context.HostKind}; failedTool={failedToolName}; recoveryTool={tc.ToolName}")
                        End If

                        If System.Threading.Interlocked.CompareExchange(powerChanging, 0, 0) <> 0 Then
                            context.LogWarn("Power transition detected during tool execution")
                            Exit For
                        End If

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

                            If preparedForNextTurn Then
                                ' Expected lazy-load deferral: the tool is allowed and was materialized on
                                ' demand; it becomes callable with its full schema on the next iteration.
                                ' This is normal operation, not a fault, so it is logged informationally.
                                context.Log(
                                    $"Tool '{tc.ToolName}' was lazily loaded on demand; it will be callable with its full schema on the next iteration.",
                                    "diag")
                            Else
                                context.LogWarn(
                                    "Blocked tool because it was not exposed at turn start.",
                                    details:=$"host={context.HostKind}; tool={tc.ToolName}; preparedForNextTurn=false")
                            End If
                            Exit For
                        End If

                        If subAgentMode AndAlso Not IsToolAllowedForCurrentContext(tc.ToolName, context) Then
                            Dim blockedResponse = BuildToolNotAllowedResponse(tc, context)
                            context.AllToolResponses.Add(blockedResponse)
                            Continue For
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
                                Dim enforcePrivacy As Boolean
                                If _apConfig IsNot Nothing Then
                                    enforcePrivacy = _apConfig.EnablePrivacyProtection
                                Else
                                    enforcePrivacy = INI_EnablePrivacyForSearch
                                End If
                                toolConfig = GetInternalSearchTool(enforcePrivacy:=enforcePrivacy)
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

                        Dim toolCallSignature = BuildToolCallSignature(tc)
                        Dim normalizedToolCallSignature As String = BuildNormalizedToolCallSignature(tc)
                        Dim previousFailureCount As Integer = 0

                        If context.FailedToolCallCounts.TryGetValue(toolCallSignature, previousFailureCount) AndAlso
                           previousFailureCount >= context.DuplicateFailureAbortThreshold Then

                            Dim duplicateMsg =
                                        $"Aborting because the same failing tool call was repeated {previousFailureCount} time(s): {tc.ToolName}. " &
                                        "The model should revise its plan instead of retrying the identical call."

                            context.LogError(duplicateMsg, details:=$"CallId={tc.CallId}; RawCall={tc.RawJson}")
                            abortDueToToolError = True
                            abortToolName = tc.ToolName
                            abortToolParamSummary = BuildCondensedParamSummary(tc.Arguments)
                            abortToolRawCallJson = tc.RawJson
                            abortToolErrorMessage = duplicateMsg
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

                        ' A python_execute (or similar) repair-loop advisor may signal that the failure is
                        ' terminal (repair budget exhausted or non-recoverable). Route into the existing
                        ' tool-error abort path so the loop stops instead of guessing further variations.
                        If toolResponse.RepairLoopTerminal Then
                            Dim repairAbortReason As String =
                                If(String.IsNullOrWhiteSpace(toolResponse.RepairLoopTerminalReason),
                                   "The repair loop was stopped because no further automatic recovery is possible.",
                                   toolResponse.RepairLoopTerminalReason)

                            context.LogWarn($"Aborting after terminal repair-loop outcome for '{tc.ToolName}'.",
                                            details:=repairAbortReason)

                            abortDueToToolError = True
                            abortToolName = tc.ToolName
                            abortToolParamSummary = BuildCondensedParamSummary(tc.Arguments)
                            abortToolRawCallJson = tc.RawJson
                            abortToolErrorMessage = repairAbortReason
                            Exit For
                        End If

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

                                    If toolConfig IsNot Nothing AndAlso toolConfig.PrefersSingleInvocation Then
                                        context.SequencingState.NoteConsolidatableToolSuccess(tc.ToolName)
                                    End If
                                End If
                            End If

                            UpdateWorkflowContinuityAfterToolExecution(context, tc, toolResponse)

                            SharedLibrary.Agents.ToolCallSequencing.NoteMemoryGroundingToolResult(
                                context.SequencingState,
                                tc.ToolName,
                                toolResponse.Response,
                                toolResponse.Success)

                            If Not subAgentMode AndAlso
                               tc.ToolName.Equals(SharedLibrary.Agents.MemoryTools.ToolList, StringComparison.OrdinalIgnoreCase) Then
                                Await TryForceRequiredMemoryGetsAsync(context, cancellationToken)
                            End If

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

                            If LoaderResponseExposedNewTools(toolResponse.Response) Then
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

                            ' No new tool schemas were exposed (all requested tools were already
                            ' loaded). Do not consume an extra iteration or stop the batch; instead
                            ' steer the model back toward an already-exposed tool. Capability-driven:
                            ' relies solely on the loader's own "loaded" array, no tool-name heuristics.
                            context.PendingContinuationGuardPrompt = BuildToolLoaderBarrierGuardPrompt()
                            context.PendingGuardTitle = "HOST TOOL LOADER NO-OP"
                            context.PendingRejectedTurnExplanation =
                                "The requested tools were already loaded. Their schemas are already available; call one of the exposed tools instead of loading again."
                            context.PendingRejectedAssistantTurn = ""
                            context.PrematureTextRetryCount = 0

                            context.LogWarn(
                                "Ignoring no-op tool_loader call; all requested tools were already loaded.",
                                details:=$"host={context.HostKind}; tool={tc.ToolName}", visibleToUser:=False)
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

                            If context.FailedToolCallCounts.ContainsKey(toolCallSignature) Then
                                context.FailedToolCallCounts(toolCallSignature) += 1
                            Else
                                context.FailedToolCallCounts(toolCallSignature) = 1
                            End If

                            ' A repair-loop advisor may signal that the failure is terminal (repair budget
                            ' exhausted or non-recoverable). Stop the batch and request a no-tool finalization
                            ' so the model produces a user-facing status instead of guessing further variations.
                            ' This serves both Outlook loops (Local Agent and AutoPilot) that share this path.
                            If toolResponse.RepairLoopTerminal Then
                                Dim repairAbortReason As String =
                                    If(String.IsNullOrWhiteSpace(toolResponse.RepairLoopTerminalReason),
                                       "The repair loop was stopped because no further automatic recovery is possible.",
                                       toolResponse.RepairLoopTerminalReason)

                                context.LogWarn($"Stopping after terminal repair-loop outcome for '{tc.ToolName}'.",
                                                details:=repairAbortReason)

                                context.ForceNoToolFinalizationRequested = True
                                context.ForceNoToolFinalizationReason = repairAbortReason
                                context.PendingContinuationGuardPrompt = BuildToolFailureReassessmentGuardPrompt(tc.ToolName)
                                context.PendingGuardTitle = "HOST TOOL FAILURE RECOVERY"
                                context.PendingRejectedTurnExplanation =
                                    "The previous tool step cannot be automatically recovered. Do not retry it; summarize the outcome for the user."
                                context.PendingRejectedAssistantTurn = ""
                                context.PrematureTextRetryCount = 0
                                stopCurrentBatchAfterTool = True
                                Exit For
                            End If

                            Select Case toolConfig.ToolErrorHandling?.ToLowerInvariant()
                                Case "abort"
                                    context.LogError("Aborting due to tool error (ToolErrorHandling=abort)")

                                    If ShouldShowToolingModalDialogs() Then
                                        ShowCustomMessageBox($"Tool execution failed: {toolResponse.ErrorMessage}")
                                    End If

                                    abortDueToToolError = True
                                    abortToolName = tc.ToolName
                                    abortToolParamSummary = BuildCondensedParamSummary(tc.Arguments)
                                    abortToolRawCallJson = tc.RawJson
                                    abortToolErrorMessage = If(toolResponse.ErrorMessage, "Unknown tool error.")
                                    Exit For

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
                            If context.FailedToolCallCounts.ContainsKey(toolCallSignature) Then
                                context.FailedToolCallCounts.Remove(toolCallSignature)
                            End If

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
                        Dim preparedToolResponsesAfterExposureMiss As String = BuildToolResponsesForModelBudgeted(
                            context.AllToolResponses,
                            context.ToolingModel,
                            compactForSubAgent:=subAgentMode)

                        INI_APICall_ToolResponses_2 = preparedToolResponsesAfterExposureMiss
                        context.Log("Tool was prepared after exposure miss; restarting next iteration with refreshed tool definitions.", "diag")
                        Continue While
                    End If

                    If restartAfterToolLoader Then
                        Dim preparedToolResponsesAfterLoader As String = BuildToolResponsesForModelBudgeted(
                            context.AllToolResponses,
                            context.ToolingModel,
                            compactForSubAgent:=subAgentMode)

                        INI_APICall_ToolResponses_2 = preparedToolResponsesAfterLoader
                        context.Log("tool_loader completed; refreshed tool definitions will be exposed on the next iteration.", "diag")
                        Continue While
                    End If

                    If restartForRequiredMemoryGrounding Then
                        Dim preparedToolResponses As String = BuildToolResponsesForModelBudgeted(
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

                    Dim toolResponses = BuildToolResponsesForModelBudgeted(
    context.AllToolResponses,
    context.ToolingModel,
    compactForSubAgent:=subAgentMode)
                    INI_APICall_ToolResponses_2 = toolResponses
                    context.Log("Tool responses prepared for next iteration", "diag")
                    If abortDueToToolError Then
                        context.LogWarn("Stopping tooling loop after tool error abort")
                        Exit While
                    End If

                Else
                    If subAgentMode Then
                        currentResponse = StripTaskStatus(currentResponse)
                        context.Log("Sub-agent final text response accepted (no tool calls)")
                        Exit While
                    End If

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

                            ' Skill/Agent authoring postcondition: a 'create/modify a skill' task must have
                            ' written under an authorized skill/resource root. A successful workspace_write
                            ' does NOT satisfy this - workspace outputs are temporary and do not install a skill.
                            If SharedLibrary.Agents.SkillAuthoringPostcondition.RequiresSkillRootMutation(context) AndAlso
                               Not SharedLibrary.Agents.SkillAuthoringPostcondition.HasSkillRootMutation(context) Then
                                If System.Convert.ToInt32(context.PrematureTextRetryCount) < ToolExecutionContext.MaxContinuationRetries Then
                                    context.PendingContinuationGuardPrompt = SharedLibrary.Agents.SkillAuthoringPostcondition.GuardPrompt
                                    context.PendingGuardTitle = "Skill not actually created"
                                    context.PendingRejectedTurnExplanation = "Skill authoring task completed without any write under an authorized skill root."
                                    context.PendingRejectedAssistantTurn = currentResponse
                                    context.PrematureTextRetryCount = System.Convert.ToInt32(context.PrematureTextRetryCount) + 1
                                    context.Log("Final-turn rejected at 'complete' branch: skill task wrote only to the workspace.", "warn")
                                    Continue While
                                End If
                                context.LogWarn("Skill-authoring postcondition unmet; repair budget exhausted, accepting candidate.")
                            End If

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
               Not cancellationToken.IsCancellationRequested AndAlso
               Not context.FinalizationBlocked Then

                Dim forcedFinalResponse As String =
                    Await RunForcedNoToolFinalizationAfterRepeatedInvalidTurnAsync(
                        context,
                        enhancedSysPrompt,
                        fullUserPrompt,
                        otherPrompt,
                        fileObject,
                        useSecondAPI,
                        hideSplash,
                        cancellationToken,
                        effectiveTimeout,
                        binaryOutputDirectory)

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

            ' If we hit max iterations and the last response is either a pending tool call or
            ' is not substantive user-facing text (e.g. raw JSON, a structured tool echo, or empty),
            ' force a final text response so the user never receives raw protocol content.
            ' The tool results are already in INI_APICall_ToolResponses_2 from the last iteration.
            Dim maxIterationsNeedsUserFacingFinal As Boolean =
                iteration >= context.MaxIterations AndAlso
                Not context.IsCancelled AndAlso
                Not cancellationToken.IsCancellationRequested AndAlso
                (ContainsToolCalls(currentResponse, context.ToolingModel.ToolCallDetectionPattern) OrElse
                 Not SharedLibrary.Agents.ToolCallSequencing.HasSubstantiveUserFacingText(
                     SharedLibrary.Agents.ToolCallSequencing.StripTaskStatusBlocksFromUserFacingText(
                         StripTaskStatus(If(currentResponse, "")))))

            If maxIterationsNeedsUserFacingFinal Then

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

                Using timeoutCts As New System.Threading.CancellationTokenSource()
                    Dim totalTimeout = effectiveTimeout + 60
                    timeoutCts.CancelAfter(TimeSpan.FromSeconds(totalTimeout))

                    Using combinedCts As System.Threading.CancellationTokenSource =
                        System.Threading.CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, timeoutCts.Token)

                        Try
                            Dim finalResponse As String = Await LLM(
                                finalSysPrompt,
                                fullUserPrompt,
                                "", "", 0,
                                useSecondAPI,
                                hideSplash,
                                otherPrompt,
                                fileObject,
                                combinedCts.Token,
                                True, True, binaryOutputDirectory:=binaryOutputDirectory)

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

                        Catch ex As OperationCanceledException
                            context.LogWarn("Forced final call was cancelled")
                        Catch ex As Exception
                            context.LogError($"Error during forced final call: {ex.Message}", ex:=ex)
                        End Try
                    End Using
                End Using
            End If

            If abortDueToToolError AndAlso
                   Not context.IsCancelled AndAlso
                   Not cancellationToken.IsCancellationRequested Then

                context.Log("Forcing final response after tool error abort...")

                ' Disable further tool calls, but keep accumulated tool responses available.
                INI_APICall_ToolInstructions_2 = ""

                Dim successfulToolFacts As New System.Text.StringBuilder()
                Dim failedToolFacts As New System.Text.StringBuilder()

                For Each tr In context.AllToolResponses
                    If tr Is Nothing Then Continue For

                    If tr.Success Then
                        successfulToolFacts.AppendLine($"- Tool: {tr.ToolName}")
                        If Not String.IsNullOrWhiteSpace(tr.Response) Then
                            successfulToolFacts.AppendLine($"  Result: {BuildResultExcerpt(tr.Response, 160)}")
                        End If
                    Else
                        failedToolFacts.AppendLine($"- Tool: {tr.ToolName}")
                        If Not String.IsNullOrWhiteSpace(tr.ErrorMessage) Then
                            failedToolFacts.AppendLine($"  Error: {tr.ErrorMessage}")
                        End If
                    End If
                Next

                abortFactsPrompt =
                        "<TOOL_ABORT_FACTS>" & Environment.NewLine &
                        "Use ONLY the facts in this block when explaining the failure." & Environment.NewLine &
                        "Do NOT replace the stated failure with another cause." & Environment.NewLine &
                        $"Failed tool: {abortToolName}" & Environment.NewLine &
                        $"Failed tool parameters: {abortToolParamSummary}" & Environment.NewLine &
                        $"Failed tool raw call JSON: {abortToolRawCallJson}" & Environment.NewLine &
                        $"Exact failure message: {abortToolErrorMessage}" & Environment.NewLine &
                        Environment.NewLine &
                        "<COMPLETED_TOOL_STEPS>" & Environment.NewLine &
                        If(successfulToolFacts.Length > 0, successfulToolFacts.ToString().TrimEnd(), "(none)") & Environment.NewLine &
                        "</COMPLETED_TOOL_STEPS>" & Environment.NewLine &
                        Environment.NewLine &
                        "<FAILED_TOOL_STEPS>" & Environment.NewLine &
                        If(failedToolFacts.Length > 0, failedToolFacts.ToString().TrimEnd(), "(none)") & Environment.NewLine &
                        "</FAILED_TOOL_STEPS>" & Environment.NewLine &
                        "</TOOL_ABORT_FACTS>"

                Dim abortFinalSysPrompt As String = enhancedSysPrompt & Environment.NewLine & Environment.NewLine &
                    "IMPORTANT: A tool-assisted run has stopped because a tool call failed. Do NOT call any more tools. " &
                    "Provide a concise, user-friendly status update. " &
                    "You MUST rely only on the explicitly supplied failure facts and completed-step facts. " &
                    "Do NOT infer a different cause. Do NOT rewrite the failure into another tool problem. " &
                    "Treat the exact failure message as authoritative. " &
                    "If the failed tool parameters indicate action='delete' or action='rmdir', describe it as a delete/remove attempt, not as a move. " &
                    "If the exact failure says permission was disabled, state that plainly and do not replace it with a path-validation explanation. " &
                    "Explain clearly: (1) what was completed successfully, (2) what failed, (3) why it failed, and (4) what therefore remains incomplete. " &
                    "Do not mention internal logs, JSON, raw tool protocols, or hidden implementation details."

                Dim abortFinalUserPrompt As String = fullUserPrompt & Environment.NewLine & Environment.NewLine &
                            abortFactsPrompt

                ToolingFileLogger.LogStep("Forcing final LLM call without tools after tool error abort")
                ToolingFileLogger.LogPreMainLlmCallSnapshot()

                Using timeoutCts As New System.Threading.CancellationTokenSource()
                    Dim totalTimeout = effectiveTimeout + 60
                    timeoutCts.CancelAfter(TimeSpan.FromSeconds(totalTimeout))

                    Using combinedCts As System.Threading.CancellationTokenSource =
            System.Threading.CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, timeoutCts.Token)

                        Try
                            Dim finalAbortResponse As String = Await LLM(
                    abortFinalSysPrompt,
                    abortFinalUserPrompt,
                    "", "", 0,
                    useSecondAPI,
                    hideSplash,
                    otherPrompt,
                    fileObject,
                    combinedCts.Token,
                    True, True, binaryOutputDirectory:=binaryOutputDirectory)

                            If Not String.IsNullOrWhiteSpace(finalAbortResponse) Then
                                currentResponse = finalAbortResponse
                                context.Log($"Final abort summary received ({currentResponse.Length} chars)")
                                ToolingFileLogger.LogRawResponseStub("Main LLM() - Tool Error Final", currentResponse)
                            Else
                                context.LogWarn("Empty response from final abort-summary LLM call")
                                currentResponse =
                        If(String.IsNullOrWhiteSpace(abortToolErrorMessage),
                           "The tool-assisted run stopped before completion.",
                           $"The tool-assisted run stopped before completion. Failed tool: {abortToolName}. Reason: {abortToolErrorMessage}")
                            End If

                        Catch ex As OperationCanceledException
                            context.LogWarn("Final abort-summary call was cancelled")
                            currentResponse =
                    If(String.IsNullOrWhiteSpace(abortToolErrorMessage),
                       "The tool-assisted run stopped before completion.",
                       $"The tool-assisted run stopped before completion. Failed tool: {abortToolName}. Reason: {abortToolErrorMessage}")

                        Catch ex As Exception
                            context.LogError($"Error during final abort-summary call: {ex.Message}", ex:=ex)
                            currentResponse =
                    If(String.IsNullOrWhiteSpace(abortToolErrorMessage),
                       "The tool-assisted run stopped before completion.",
                       $"The tool-assisted run stopped before completion. Failed tool: {abortToolName}. Reason: {abortToolErrorMessage}")
                        End Try
                    End Using
                End Using
            End If

            If context.IsCancelled OrElse cancellationToken.IsCancellationRequested Then
                context.LogWarn("Session cancelled")
                ToolingFileLogger.EndSession(False, "Cancelled")
                Return If(cancellationToken.IsCancellationRequested, "Operation was canceled by the user.", "")
            End If

            If iteration >= context.MaxIterations Then
                context.LogWarn($"Maximum iterations ({context.MaxIterations}) reached")
                If ShouldShowToolingModalDialogs() Then
                    ShowCustomMessageBox($"Maximum tool iterations ({context.MaxIterations}) reached. The response may be incomplete.")
                End If
                ToolingFileLogger.LogWarn("Maximum iterations reached.", details:=$"MaxIterations={context.MaxIterations}")

                ' Guarantee user-understandable output: if the forced-final call did not yield
                ' substantive user-facing prose, summarize what was achieved as a blocked result
                ' instead of returning raw tool/JSON content to the user.
                If Not context.FinalizationBlocked AndAlso
                   Not SharedLibrary.Agents.ToolCallSequencing.HasSubstantiveUserFacingText(
                       SharedLibrary.Agents.ToolCallSequencing.StripTaskStatusBlocksFromUserFacingText(
                           StripTaskStatus(If(currentResponse, "")))) Then

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
            ' A lazy-load deferral (tool_not_exposed_in_current_turn) is expected behaviour, not a
            ' failure: the tool was allowed and loaded on demand, and the model re-issues the call
            ' with the real schema on the next turn. Count it separately so it is neither reported
            ' as a failure nor as a successful tool execution.
            Dim deferredCount As Integer = context.AllToolResponses.
                Where(Function(r) Not r.Success AndAlso
                                  String.Equals(r.ErrorCode,
                                                SharedLibrary.Agents.ToolCallSequencing.ToolNotExposedInCurrentTurnCode,
                                                StringComparison.OrdinalIgnoreCase)).Count()
            Dim successCount As Integer = context.AllToolResponses.Where(Function(r) r.Success).Count()
            Dim failedCount As Integer = context.AllToolResponses.
                Where(Function(r) Not r.Success AndAlso
                                  Not String.Equals(r.ErrorCode,
                                                    SharedLibrary.Agents.ToolCallSequencing.ToolNotExposedInCurrentTurnCode,
                                                    StringComparison.OrdinalIgnoreCase)).Count()
            context.Log($"Successful: {successCount}", If(failedCount = 0, "success", "step"))
            context.Log($"Failed: {failedCount}", If(failedCount = 0, "step", "warn"))
            If deferredCount > 0 Then
                context.Log($"Deferred (lazy-loaded, re-issued next turn): {deferredCount}", "diag")
            End If

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
                    abortDueToToolError OrElse
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
    $"Returning final response to host: len={finalResponseLength}; origin={finalResponseOrigin}; blocked={If(context.FinalizationBlocked, "true", "false")}; empty={If(context.EmptyMainModelResponse, "true", "false")}; abortedDueToToolError={If(abortDueToToolError, "true", "false")}",
    "diag")

            ToolingFileLogger.LogStep(
    $"Host final response prepared: len={finalResponseLength}; origin={finalResponseOrigin}; blocked={If(context.FinalizationBlocked, "true", "false")}; empty={If(context.EmptyMainModelResponse, "true", "false")}; abortedDueToToolError={If(abortDueToToolError, "true", "false")}")

            ' A final turn accepted with status "blocked" is an unfinished/blocked outcome
            ' and must not be reported as a successful session, even when it carries
            ' substantive prose (which leaves context.FinalizationBlocked unset).
            Dim sessionSucceeded As Boolean = (Not abortDueToToolError) AndAlso (Not context.EmptyMainModelResponse) AndAlso (Not context.FinalizationBlocked) AndAlso Not String.Equals(acceptedFinalStatus, "blocked", StringComparison.OrdinalIgnoreCase)
            Dim sessionSummary As String =
    $"Iterations: {iteration}, Tool calls: {context.AllToolResponses.Count}, Success: {successCount}, Failed: {failedCount}" &
    If(deferredCount > 0, $", Deferred: {deferredCount}", "")

            If abortDueToToolError AndAlso Not String.IsNullOrWhiteSpace(abortToolErrorMessage) Then
                sessionSummary &= $", Aborted due to tool error: {abortToolErrorMessage}"
            End If

            ToolingFileLogger.EndSession(sessionSucceeded, sessionSummary)
            Return currentResponse

        Catch ex As OperationCanceledException
            context.LogWarn("Operation cancelled")
            ToolingFileLogger.EndSession(False, "Cancelled")
            Return "Operation was canceled by the user."

        Catch ex As Exception
            context.LogError($"Error in tooling loop: {ex.Message}", ex:=ex)
            If ShouldShowToolingModalDialogs() Then
                ShowCustomMessageBox($"Error during tool execution: {ex.Message}")
            End If

            ToolingFileLogger.EndSession(False, $"Exception: {ex.Message}", ex:=ex)
            Return $"Error during tool execution: {ex.Message}"
        Finally

            If workflowScope IsNot Nothing Then
                workflowScope.Dispose()
                workflowScope = Nothing
            End If

            ' Release stored large tool-result bodies for this run. Sub-agents share the
            ' parent's WorkflowId, so only the top-level run may clear the store.
            If Not subAgentMode Then
                Try
                    SharedLibrary.Agents.ToolResultStore.ClearWorkflow(context.WorkflowId)
                    SharedLibrary.Agents.SkillAuthoringPostcondition.Clear(context.WorkflowId)
                    ToolingFileLogger.LogStep($"ClearWorkflow: tool result store cleared for workflow '{context.WorkflowId}'.")
                Catch ex As Exception
                    ToolingFileLogger.LogWarn("Failed to clear tool result store.", ex:=ex)
                End Try
            End If

            _activeToolingContext = parentToolingContext
            INI_APICall_ToolInstructions_2 = ""
            INI_APICall_ToolResponses_2 = ""

            If context.LogWindowForm IsNot Nothing AndAlso Not context.LogWindowForm.IsDisposed Then
                Try
                    If UiSyncContext IsNot Nothing AndAlso
                       System.Threading.Thread.CurrentThread.ManagedThreadId <> UiThreadId Then
                        UiSyncContext.Post(
                            Sub()
                                Try
                                    If context.LogWindowForm IsNot Nothing AndAlso
                                       Not context.LogWindowForm.IsDisposed Then
                                        context.LogWindowForm.MarkComplete()
                                    End If
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

            Try
                _lastCompletedToolResponses = New List(Of ToolResponse)()
                If context IsNot Nothing AndAlso context.AllToolResponses IsNot Nothing Then
                    Debug.WriteLine("[AISearch] Persisting completed tool responses: " & context.AllToolResponses.Count.ToString())

                    For Each r In context.AllToolResponses
                        _lastCompletedToolResponses.Add(New ToolResponse() With {
                    .CallId = r.CallId,
                    .ToolName = r.ToolName,
                    .Response = r.Response,
                    .Success = r.Success,
                    .ErrorMessage = r.ErrorMessage,
                    .Timestamp = r.Timestamp,
                    .OriginalCallJson = r.OriginalCallJson
                })
                    Next
                Else
                    Debug.WriteLine("[AISearch] Persisting completed tool responses: context or response list is Nothing")
                End If
            Catch ex As Exception
                Debug.WriteLine("[AISearch] Persisting completed tool responses failed: " & ex.Message)
            End Try

            ' Keep the restored parent context alive for repeated sub-agent calls.
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



    ''' <summary>
    ''' Returns the set of deliverable-capable tool names that are still loaded in
    ''' <paramref name="ctx"/> and have NOT yet been invoked (successfully or otherwise)
    ''' in the current run. Used by the deliverable-fallback gate to decide whether
    ''' 'blocked' is premature.
    ''' </summary>
    Private Function CollectUntriedDeliverableFallbackToolNames(ctx As ToolExecutionContext) As IReadOnlyList(Of String)
        Dim result As New List(Of String)()
        If ctx Is Nothing Then Return result.AsReadOnly()

        Dim deliverableCandidates As New HashSet(Of String)(
            SharedLibrary.Agents.HostToolRegistration.GetDeliverableCapableToolNames(
                SharedLibrary.Agents.ToolingHostKind.Outlook),
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
                cancellationToken,
                True,
                False)

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
            hideSplash,
            cancellationToken)
    End Function



    Private Async Function BuildTaskSpecificPartialBlockedMessageAsync(context As ToolExecutionContext,
                                                                       errorCode As String,
                                                                       message As String,
                                                                       useSecondAPI As Boolean,
                                                                       hideSplash As Boolean,
                                                                       cancellationToken As System.Threading.CancellationToken) As Task(Of String)
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
                cancellationToken,
                True,
                False)

            Return If(responseText, "").Trim()
        Catch
            Return ""
        End Try
    End Function

    Private Async Function BuildBlockedToolingResultAsync(context As ToolExecutionContext,
                                                          errorCode As String,
                                                          message As String,
                                                          useSecondAPI As Boolean,
                                                          hideSplash As Boolean,
                                                          cancellationToken As System.Threading.CancellationToken) As Task(Of String)
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
                hideSplash,
                cancellationToken)

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
            hideSplash,
            cancellationToken)
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

    ''' <summary>
    ''' Returns True when a tool_loader response actually exposed at least one new tool
    ''' (its "loaded" array is non-empty). Used to avoid consuming an extra orchestration
    ''' iteration on no-op loader calls where every requested tool was already loaded.
    ''' Capability-driven: inspects only the loader's own structured output.
    ''' </summary>
    Private Function LoaderResponseExposedNewTools(loaderResponse As String) As Boolean
        If String.IsNullOrWhiteSpace(loaderResponse) Then
            Return False
        End If

        Try
            Dim obj As JObject = JObject.Parse(loaderResponse)
            Dim loadedToken As JToken = obj("loaded")
            Return loadedToken IsNot Nothing AndAlso
                   loadedToken.Type = JTokenType.Array AndAlso
                   DirectCast(loadedToken, JArray).Count > 0
        Catch
            ' On parse failure, fall back to the conservative (barrier) behavior by
            ' treating it as a real load so the caller preserves existing semantics.
            Return True
        End Try
    End Function

#End Region




    Private Function ShouldShowToolingModalDialogs() As Boolean
        Return Not _chatAgentActive AndAlso Not _apActive
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


    Private Function BuildSkippedToolFailureRecoveryGuardPrompt() As String
        Return "HOST TOOL FAILURE GUARD: The previous tool call failed and that structured error is already available in the tool-response history. " &
           "In THIS turn you must take exactly one actionable recovery step: " &
           "(a) write or update failure state for the current item using an available write/state tool; OR " &
           "(b) retry the failed agent/tool with smaller or chunked input; OR " &
           "(c) continue with the next item using the appropriate tool; OR " &
           "(d) produce a valid final blocked response only if all required work is complete or no further tool action is possible. " &
           "Do NOT apologize. Do NOT output ordinary progress prose. Do NOT merely restate the failure."
    End Function



    Private Function BuildSubAgentEmptyResponseRecoveryPrompt(context As ToolExecutionContext) As String
        Dim lastSuccess = GetLastSuccessfulToolResponse(context)
        Dim summary As String = BuildToolReplaySummary(lastSuccess)

        Dim sb As New System.Text.StringBuilder()
        sb.Append("SUB-AGENT EMPTY-RESPONSE RECOVERY: The previous model turn was empty after a successful tool call. ")
        sb.Append("Do not repeat a large raw tool response. ")
        sb.Append("In THIS turn you must either return the required final JSON object, or call one smaller follow-up tool. ")
        sb.Append("If more source text is needed, request a smaller window using max_chars and start_char/offset.")

        If Not String.IsNullOrWhiteSpace(summary) Then
            sb.AppendLine()
            sb.Append("Last successful tool result: ")
            sb.Append(summary)
        End If

        Return sb.ToString()
    End Function






    Private Function IsPreparatoryToolName(toolName As String) As Boolean
        If String.IsNullOrWhiteSpace(toolName) Then
            Return False
        End If

        Select Case toolName.Trim().ToLowerInvariant()
            Case "tool_loader",
                 "report_progress",
                 "agent_workspace_get",
                 "agent_workspace_list",
                 "agent_workspace_find_files",
                 "agent_workspace_recent_files",
                 "agent_workspace_file_details",
                 "agent_workspace_inventory_report",
                 "workspace_get",
                 "workspace_inventory"
                Return True
            Case Else
                Return False
        End Select
    End Function

    Private Function HasAnyNonPreparatoryToolsAvailable(context As ToolExecutionContext) As Boolean
        If context Is Nothing Then
            Return False
        End If

        Try
            If context.AllowedToolRegistry IsNot Nothing Then
                Dim manifests = context.AllowedToolRegistry.ListManifests()

                If manifests IsNot Nothing AndAlso
                   manifests.Any(Function(m)
                                     Return m IsNot Nothing AndAlso
                                            Not String.IsNullOrWhiteSpace(m.Name) AndAlso
                                            Not IsPreparatoryToolName(m.Name)
                                 End Function) Then
                    Return True
                End If
            End If
        Catch
        End Try

        If context.SelectedTools Is Nothing Then
            Return False
        End If

        Return context.SelectedTools.Any(
            Function(t)
                Return t IsNot Nothing AndAlso
                       Not String.IsNullOrWhiteSpace(t.ToolName) AndAlso
                       Not IsPreparatoryToolName(t.ToolName)
            End Function)
    End Function


    Private Enum TaskStatusKind
        Missing
        Complete
        ContinueWork
        Blocked
    End Enum

    Private Function ParseTaskStatus(text As String) As TaskStatusKind
        ' Strict JSON parser is centralized in SharedLibrary.Agents.TaskStatusFooterParser.
        ' We map the shared TaskStatusKind to the local enum to keep all existing call
        ' sites and Select Case statements in this file working unchanged.
        Dim parsed As Agents.TaskStatusFooter =
            Agents.TaskStatusFooterParser.Parse(text)

        Select Case parsed.Kind
            Case Agents.TaskStatusKind.Complete
                Return TaskStatusKind.Complete
            Case Agents.TaskStatusKind.ContinueWork
                Return TaskStatusKind.ContinueWork
            Case Agents.TaskStatusKind.Blocked
                Return TaskStatusKind.Blocked
            Case Else
                ' Treat Invalid (malformed JSON, duplicated footer, unknown status value)
                ' the same as Missing so existing call sites uniformly retry/repair.
                Return TaskStatusKind.Missing
        End Select
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
               "If a tool just failed, try a DIFFERENT applicable tool with corrected arguments. For workspace files use workspace_get, workspace_inventory, workspace_read, workspace_read_many, workspace_extract_text, workspace_extract_text_many, workspace_search, and workspace_write."
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
            sb.AppendLine(Agents.ToolingOrchestrator.TaskStatusFooterInstruction)
        Else
            sb.AppendLine("ACTIVE-TOOLING CONTRACT: During this active tooling session, every response must be exactly one of: (1) the next required tool call and nothing else; or (2) the final raw text payload requested elsewhere in the prompt and nothing else.")
            sb.AppendLine("Do NOT append any <TASK_STATUS> footer unless the caller explicitly requested one.")
            sb.AppendLine("Do NOT convert the final output into user-facing prose unless the caller explicitly requested prose.")
            sb.AppendLine("When tool work is complete, return only the final raw text payload in the exact caller-defined format.")
        End If

        sb.AppendLine(SharedLibrary.Agents.ToolCallSequencing.DependentBatchingInstruction)

        If selectedTools IsNot Nothing AndAlso selectedTools.Any(
                Function(t) t IsNot Nothing AndAlso
                            (SharedLibrary.Agents.ContextExpandTool.IsContextExpandTool(t.ToolName) OrElse
                             SharedLibrary.Agents.ContextCompactTool.IsContextCompactTool(t.ToolName))) Then
            sb.AppendLine()
            sb.AppendLine(SharedLibrary.Agents.ToolCallSequencing.ContextDrawerInstruction)
        End If

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

    Private Function BuildToolSelectionHintText(userText As String,
                                                fullPromptOverride As String,
                                                otherPrompt As String,
                                                insertDocs As String,
                                                slideInsert As String,
                                                bubblesText As String) As String
        Dim parts As New List(Of String)()

        If Not String.IsNullOrWhiteSpace(fullPromptOverride) Then parts.Add(fullPromptOverride)
        If Not String.IsNullOrWhiteSpace(userText) Then parts.Add(userText)
        If Not String.IsNullOrWhiteSpace(otherPrompt) Then parts.Add(otherPrompt)
        If Not String.IsNullOrWhiteSpace(insertDocs) Then parts.Add(insertDocs)
        If Not String.IsNullOrWhiteSpace(slideInsert) Then parts.Add(slideInsert)
        If Not String.IsNullOrWhiteSpace(bubblesText) Then parts.Add(bubblesText)

        Return String.Join(Environment.NewLine, parts)
    End Function

    Private Function BuildInitialToolExposure(allowedTools As List(Of ModelConfig),
                                              allowedRegistry As SharedLibrary.Agents.ToolRegistry,
                                              promptText As String) As List(Of ModelConfig)
        Dim result As New List(Of ModelConfig)()

        If allowedTools Is Nothing OrElse allowedTools.Count = 0 Then
            Return result
        End If

        ' Always expose the internal report_progress tool so the model can announce major
        ' steps (B1) via a real tool call. Host-derived progress (A) is independent of this.
        Dim progressReportTool As ModelConfig = GetInternalProgressReportTool()
        If progressReportTool IsNot Nothing Then
            result.Add(progressReportTool)
        End If

        Dim deduplicatedTools As List(Of ModelConfig) =
            allowedTools.
                Where(Function(t) t IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(t.ToolName)).
                Where(Function(t) Not IsReportProgressToolName(t.ToolName)).
                GroupBy(Function(t) t.ToolName, StringComparer.OrdinalIgnoreCase).
                Select(Function(g) g.First()).
                ToList()

        If Not SharedLibrary.Agents.ToolLoaderTool.ShouldUseLazyLoading(deduplicatedTools) Then
            result.AddRange(deduplicatedTools)
            Return result
        End If

        If allowedRegistry Is Nothing Then
            result.AddRange(deduplicatedTools)
            Return result
        End If

        Dim loaderManifests As List(Of SharedLibrary.Agents.ToolManifest) =
            allowedRegistry.ListManifests().
                Where(Function(m)
                          Return m IsNot Nothing AndAlso
                                 Not String.IsNullOrWhiteSpace(m.Name) AndAlso
                                 Not m.Name.Equals(SharedLibrary.Agents.ToolLoaderTool.LoaderToolName, StringComparison.OrdinalIgnoreCase)
                      End Function).
                OrderBy(Function(m) m.Name, StringComparer.OrdinalIgnoreCase).
                ToList()

        If loaderManifests.Count = 0 Then
            result.AddRange(deduplicatedTools)
            Return result
        End If

        Dim loader As ModelConfig =
            SharedLibrary.Agents.ToolLoaderTool.Build(loaderManifests)

        If loader Is Nothing Then
            result.AddRange(deduplicatedTools)
            Return result
        End If

        result.Add(loader)
        Return result
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





    Private Sub LogAgentToolCallStatistic(toolName As String)
        Dim surface As String = ""

        If _apActive Then
            surface = "Outlook_AutoPilot"
        ElseIf _chatAgentActive Then
            surface = "Outlook_LocalChatAgent"
        Else
            Return
        End If

        SharedLogger.LogAgentToolCall(_context, _context.RDV, surface, toolName)
    End Sub



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

    ''' <summary>
    ''' Executes a single tool call using an internal tool implementation or an external tool configuration.
    ''' Internal tools: <c>retrieve_web_content</c> and <c>internet_search</c> (when search is enabled).
    ''' </summary>
    Public Async Function ExecuteToolCall(toolCall As ToolCall, toolConfig As ModelConfig, context As ToolExecutionContext, Optional cancellationToken As System.Threading.CancellationToken = Nothing) As Task(Of ToolResponse)

        LogAgentToolCallStatistic(toolCall.ToolName)

        ' ── python_execute: secure sandboxed Python execution ──
        If toolCall.ToolName.Equals(SharedLibrary.Agents.PythonExecuteTool.ToolName, StringComparison.OrdinalIgnoreCase) Then
            Dim pyResponse As ToolResponse = Await ExecutePythonExecuteTool(toolCall, context, cancellationToken)

            If (_apActive OrElse _chatAgentActive) AndAlso _apCurrentToolCallLog IsNot Nothing Then
                Dim pyParamSummary As String = BuildCondensedParamSummary(toolCall.Arguments)
                Dim pyExcerpt As String = BuildResultExcerpt(pyResponse.Response, 80)
                ApDashboardLog($"🔧 External tool: {toolCall.ToolName}{pyParamSummary}", "info")
                If pyResponse.Success Then
                    ApDashboardLog($"   ✓ {pyExcerpt}", "info")
                Else
                    ApDashboardLog($"   ✗ {If(pyResponse.ErrorMessage, pyExcerpt)}", "error")
                End If
                RecordAutoPilotToolCall(
                    toolCall.ToolName,
                    toolCall.ToolName,
                    pyParamSummary,
                    isInternalTool:=True,
                    wasSuccessful:=pyResponse.Success,
                    resultExcerpt:=pyExcerpt,
                    elapsed:=DateTime.Now - pyResponse.Timestamp,
                    urls:=Nothing)
            End If

            Return pyResponse
        End If

        ' ── workspace_extract_text: unified file reader for Local Chat agent (no staging) ──
        If ((_chatAgentActive AndAlso Not _apActive) OrElse _apActive) AndAlso
           toolCall.ToolName.Equals(SharedLibrary.Agents.WorkspaceTools.ToolExtractText, StringComparison.OrdinalIgnoreCase) Then

            Dim resp As New ToolResponse() With {.CallId = toolCall.CallId, .ToolName = toolCall.ToolName}
            Dim relPath As String = If(toolCall.Arguments IsNot Nothing AndAlso toolCall.Arguments.ContainsKey("path"),
                                       System.Convert.ToString(toolCall.Arguments("path")), "")
            Dim maxChars As Integer = 100000
            Try
                If toolCall.Arguments IsNot Nothing AndAlso toolCall.Arguments.ContainsKey("max_chars") Then
                    Integer.TryParse(toolCall.Arguments("max_chars").ToString(), maxChars)
                End If
            Catch
            End Try
            maxChars = Math.Min(Math.Max(maxChars, 1000), 500000)

            If Not IsChatAgentWorkspaceConnected() OrElse Not _chatAgentWorkspace.AllowRead Then
                resp.Success = False
                resp.ErrorMessage = "No readable workspace is connected."
                Return resp
            End If

            Dim fullPath As String = ResolveWorkspacePath(relPath)
            If String.IsNullOrWhiteSpace(fullPath) OrElse Not IO.File.Exists(fullPath) Then
                resp.Success = False
                resp.ErrorMessage = "Workspace file not found: " & If(relPath, "")
                Return resp
            End If

            Dim extracted As String = ""
            Try
                extracted = Await ChatAgentExtractFileText(fullPath).ConfigureAwait(False)
            Catch ex As Exception
                resp.Success = False
                resp.ErrorMessage = "Extraction failed: " & ex.Message
                Return resp
            End Try

            If String.IsNullOrWhiteSpace(extracted) Then
                resp.Success = True
                resp.Response = "(No readable text extracted from '" & relPath & "'.)"
            Else
                If extracted.Length > maxChars Then
                    extracted = extracted.Substring(0, maxChars) & Environment.NewLine & "[Truncated at " & maxChars & " characters.]"
                End If
                resp.Success = True
                resp.Response = extracted
            End If
            Return resp
        End If

        ' ── workspace_read_many: shared UTF-8 text reader for multiple files ──
        If ((_chatAgentActive AndAlso Not _apActive) OrElse _apActive OrElse HasActiveScheduledTaskWorkspace()) AndAlso
           toolCall.ToolName.Equals(SharedLibrary.Agents.WorkspaceTools.ToolReadMany, StringComparison.OrdinalIgnoreCase) Then

            Dim rmResp As New ToolResponse() With {.CallId = toolCall.CallId, .ToolName = toolCall.ToolName}
            If Not IsChatAgentWorkspaceConnected() OrElse Not _chatAgentWorkspace.AllowRead Then
                rmResp.Success = False
                rmResp.ErrorMessage = "No readable workspace is connected."
                Return rmResp
            End If

            rmResp.Response = SharedLibrary.Agents.WorkspaceTools.Execute(toolCall.ToolName, toolCall.Arguments)
            rmResp.Success = True
            Return rmResp
        End If

        ' ── workspace_extract_text_many: extract text from multiple files (Local Chat, AutoPilot, scheduled tasks) ──
        If ((_chatAgentActive AndAlso Not _apActive) OrElse _apActive OrElse HasActiveScheduledTaskWorkspace()) AndAlso
           toolCall.ToolName.Equals(SharedLibrary.Agents.WorkspaceTools.ToolExtractTextMany, StringComparison.OrdinalIgnoreCase) Then

            Dim etmResp As New ToolResponse() With {.CallId = toolCall.CallId, .ToolName = toolCall.ToolName}

            Dim etmMaxFiles As Integer = 20
            Dim etmMaxCharsPerFile As Integer = 100000
            Try
                If toolCall.Arguments IsNot Nothing AndAlso toolCall.Arguments.ContainsKey("max_files") Then
                    Integer.TryParse(toolCall.Arguments("max_files").ToString(), etmMaxFiles)
                End If
                If toolCall.Arguments IsNot Nothing AndAlso toolCall.Arguments.ContainsKey("max_chars_per_file") Then
                    Integer.TryParse(toolCall.Arguments("max_chars_per_file").ToString(), etmMaxCharsPerFile)
                End If
            Catch
            End Try
            etmMaxFiles = Math.Min(Math.Max(etmMaxFiles, 1), 100)
            etmMaxCharsPerFile = Math.Min(Math.Max(etmMaxCharsPerFile, 1000), 500000)

            If Not IsChatAgentWorkspaceConnected() OrElse Not _chatAgentWorkspace.AllowRead Then
                etmResp.Success = False
                etmResp.ErrorMessage = "No readable workspace is connected."
                Return etmResp
            End If

            Dim etmPaths As New List(Of String)()
            If toolCall.Arguments IsNot Nothing AndAlso toolCall.Arguments.ContainsKey("paths") Then
                Dim etmV = toolCall.Arguments("paths")
                If TypeOf etmV Is JArray Then
                    For Each tk In DirectCast(etmV, JArray)
                        Dim s = tk.ToString()
                        If Not String.IsNullOrWhiteSpace(s) Then etmPaths.Add(s)
                    Next
                ElseIf TypeOf etmV Is IEnumerable(Of Object) Then
                    For Each o In DirectCast(etmV, IEnumerable(Of Object))
                        Dim s = System.Convert.ToString(o)
                        If Not String.IsNullOrWhiteSpace(s) Then etmPaths.Add(s)
                    Next
                End If
            End If

            Dim etmRequestedCount As Integer = etmPaths.Count
            Dim etmSelected As List(Of String) = etmPaths.Take(etmMaxFiles).ToList()
            Dim etmItems As New List(Of Object)()

            For Each etmRelPath In etmSelected
                Dim etmFullPath As String = ResolveWorkspacePath(etmRelPath)
                If String.IsNullOrWhiteSpace(etmFullPath) OrElse Not IO.File.Exists(etmFullPath) Then
                    etmItems.Add(New With {Key .path = etmRelPath, Key .error = "not_found", Key .message = "File not found."})
                    Continue For
                End If

                Try
                    Dim etmExtracted As String = Await ChatAgentExtractFileText(etmFullPath).ConfigureAwait(False)
                    Dim etmTruncated As Boolean = False
                    If Not String.IsNullOrWhiteSpace(etmExtracted) AndAlso etmExtracted.Length > etmMaxCharsPerFile Then
                        etmExtracted = etmExtracted.Substring(0, etmMaxCharsPerFile) & vbCrLf & $"[Truncated at {etmMaxCharsPerFile} characters.]"
                        etmTruncated = True
                    End If
                    etmItems.Add(New With {
                        Key .path = etmFullPath,
                        Key .truncated = etmTruncated,
                        Key .text = If(etmExtracted, "")
                    })
                Catch ex As Exception
                    etmItems.Add(New With {Key .path = etmRelPath, Key .error = "extraction_failed", Key .message = ex.Message})
                End Try
            Next

            etmResp.Success = True
            etmResp.Response = JsonConvert.SerializeObject(New With {
                Key .requested_count = etmRequestedCount,
                Key .processed_count = etmSelected.Count,
                Key .items = etmItems
            })
            Return etmResp
        End If

        If ((_chatAgentActive AndAlso Not _apActive) OrElse _apActive OrElse HasActiveScheduledTaskWorkspace()) AndAlso
           SharedLibrary.Agents.WorkspaceTools.IsWorkspaceTool(toolCall.ToolName) Then

            Dim wsResp As New ToolResponse() With {
                .CallId = toolCall.CallId,
                .ToolName = toolCall.ToolName
            }

            Dim hasReadableWorkspace As Boolean = IsChatAgentWorkspaceConnected()
            Dim hasTransientSessionWorkspace As Boolean =
                _chatAgentActive AndAlso Not _apActive AndAlso HasChatAgentSessionWorkspace()
            Dim hasScheduledTaskWorkspace As Boolean = HasActiveScheduledTaskWorkspace()

            If Not hasReadableWorkspace AndAlso Not hasTransientSessionWorkspace AndAlso Not hasScheduledTaskWorkspace Then
                wsResp.Success = False
                wsResp.ErrorMessage = "No readable workspace is connected."
                Return wsResp
            End If

            If hasScheduledTaskWorkspace Then
                ActivateScheduledTaskWorkspace(_apScheduledWorkspaceTaskId)
            ElseIf hasTransientSessionWorkspace Then
                SyncWorkspaceToPathPolicy(includeSessionTempFallback:=True)
            End If

            wsResp.Response = SharedLibrary.Agents.WorkspaceTools.Execute(toolCall.ToolName, toolCall.Arguments)
            wsResp.Success = True

            Try
                Dim wsToken As JToken = JToken.Parse(wsResp.Response)
                If wsToken.Type = JTokenType.Object Then
                    Dim wsObj = DirectCast(wsToken, JObject)
                    Dim errToken = wsObj("error")
                    If errToken IsNot Nothing AndAlso errToken.Type <> JTokenType.Null AndAlso errToken.ToString().Trim() <> "" Then
                        wsResp.Success = False
                        wsResp.ErrorMessage = If(wsObj("message")?.ToString(), errToken.ToString())
                    End If
                End If
            Catch
            End Try

            Return wsResp
        End If

        ' ── Agent workspace tools for Local Chat Agent and AutoPilot temp workspaces ──
        If ((_chatAgentActive AndAlso Not _apActive) OrElse _apActive OrElse HasActiveScheduledTaskWorkspace()) AndAlso
           IsChatAgentWorkspaceTool(toolCall.ToolName) Then

            Dim workspaceResult = Await ExecuteChatAgentWorkspaceTool(toolCall, context, cancellationToken)
            Return workspaceResult
        End If

        ' ── AutoPilot / Chat Agent internal tool routing ──
        If (_apActive OrElse _chatAgentActive) AndAlso IsAutoPilotInternalTool(toolCall.ToolName) Then
            Dim apResult = Await TryExecuteAutoPilotTool(toolCall, context, cancellationToken)
            If apResult IsNot Nothing Then
                ' Record for "Sources used:" footer
                If _apCurrentToolCallLog IsNot Nothing Then
                    Dim elapsed = DateTime.Now - apResult.Timestamp
                    Dim excerpt = BuildResultExcerpt(apResult.Response, 80)
                    RecordAutoPilotToolCall(
                        toolCall.ToolName,
                        If(toolConfig?.ModelDescription, toolCall.ToolName),
                        BuildCondensedParamSummary(toolCall.Arguments),
                        isInternalTool:=True,
                        wasSuccessful:=apResult.Success,
                        resultExcerpt:=excerpt,
                        elapsed:=elapsed,
                        urls:=Nothing)
                End If

                Return apResult
            End If
        End If

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName
        }

        ' Build condensed parameter summary for log window
        Dim paramSummary As String = BuildCondensedParamSummary(toolCall.Arguments)
        Dim visibleToolLabel As String = Regex.Replace(If(toolCall.ToolName, ""), "_+", " ").Trim()

        If visibleToolLabel = "" Then
            visibleToolLabel = "processing step"
        End If

        context.Log("Running " & visibleToolLabel & "...")
        context.Log($"Executing tool: {toolCall.ToolName}{paramSummary}", "diag")


        Try
            ' Check cancellation before execution
            cancellationToken.ThrowIfCancellationRequested()


            If toolCall.ToolName.Equals(SharedLibrary.Agents.ToolLoaderTool.LoaderToolName, StringComparison.OrdinalIgnoreCase) Then
                response = ExecuteToolLoaderCall(toolCall, context)
                ToolingFileLogger.LogRawResponseStub($"Internal tool ({toolCall.ToolName})", response.Response)
                GoTo __AfterDispatch
            End If

            If SharedLibrary.Agents.ToolDescribeTool.IsDescribeTool(toolCall.ToolName) Then
                Dim describeRegistry As SharedLibrary.Agents.ToolRegistry =
                    If(context.AuthoritativeToolRegistrySnapshot,
                       If(context.AuthoritativeToolRegistry, context.AllowedToolRegistry))
                response.Response = SharedLibrary.Agents.ToolDescribeTool.Execute(toolCall.Arguments, describeRegistry)
                response.Success = Not String.IsNullOrWhiteSpace(response.Response)
                ToolingFileLogger.LogRawResponseStub($"Internal tool ({toolCall.ToolName})", response.Response)
                GoTo __AfterDispatch
            End If

            If SharedLibrary.Agents.ContextExpandTool.IsContextExpandTool(toolCall.ToolName) Then
                response.Response = SharedLibrary.Agents.ContextExpandTool.Execute(toolCall.Arguments)
                response.Success = Not String.IsNullOrWhiteSpace(response.Response)
                ToolingFileLogger.LogRawResponseStub($"Internal tool ({toolCall.ToolName})", response.Response)
                GoTo __AfterDispatch
            End If

            If SharedLibrary.Agents.AskUserTool.IsAskUserTool(toolCall.ToolName) Then
                response.Response = SharedLibrary.Agents.AskUserTool.Execute(toolCall.Arguments)
                response.Success = Not String.IsNullOrWhiteSpace(response.Response)
                ToolingFileLogger.LogRawResponseStub($"Internal tool ({toolCall.ToolName})", response.Response)
                GoTo __AfterDispatch
            End If

            If SharedLibrary.Agents.ContextCompactTool.IsContextCompactTool(toolCall.ToolName) Then
                response.Response = SharedLibrary.Agents.ContextCompactTool.Execute(
                    toolCall.Arguments,
                    SharedLibrary.Agents.WorkflowContinuity.CurrentWorkflowId)
                response.Success = Not String.IsNullOrWhiteSpace(response.Response)
                ToolingFileLogger.LogRawResponseStub($"Internal tool ({toolCall.ToolName})", response.Response)
                GoTo __AfterDispatch
            End If

            ' Agent layer (memory_*, skill_use, agent_*) — single-line dispatcher.
            If SharedLibrary.Agents.AgentToolRouter.IsAgentLayerTool(toolCall.ToolName) Then
                Dim __agentJson = Await SharedLibrary.Agents.AgentToolRouter.TryHandleAsync(
        toolCall.ToolName, toolCall.Arguments, CType(Me, SharedLibrary.Agents.ISubAgentHost), cancellationToken, _context).ConfigureAwait(False)

                response.Response = If(__agentJson, "")
                response.Success = Not String.IsNullOrWhiteSpace(response.Response)

                If response.Success Then
                    MarkInPlaceEditAsForcedDeliverable(toolCall.ToolName, response.Response)
                End If

                ApplyStructuredAgentResult(response, context)

                If Not response.Success AndAlso String.IsNullOrWhiteSpace(response.ErrorMessage) Then
                    response.ErrorMessage = "Agent-layer tool returned no usable result."
                ElseIf response.Success AndAlso SharedLibrary.Agents.JsRunTool.IsJsTool(toolCall.ToolName) AndAlso IsEmptyJsRunResult(response.Response) Then
                    response.Success = False
                    response.ResultKind = "error"
                    response.ErrorCode = "agent_empty_result"
                    response.ErrorMessage = "js_run returned no usable result. Ensure the script explicitly returns the computed value."
                    response.Response = "{""summary"":""Sub-agent returned no usable result."",""result"":null,""resultKind"":""error"",""error"":{""code"":""agent_empty_result"",""phase"":""final_output_parse"",""message"":""Sub-agent returned no usable final result.""}}"
                End If

                If String.Equals(toolCall.ToolName, SharedLibrary.Agents.SkillInvokeTool.ToolName, StringComparison.OrdinalIgnoreCase) AndAlso response.Success Then
                    Dim __skillName As String = Nothing
                    If toolCall.Arguments IsNot Nothing AndAlso toolCall.Arguments.ContainsKey("name") Then
                        __skillName = System.Convert.ToString(toolCall.Arguments("name"))
                    End If
                    If Not String.IsNullOrWhiteSpace(__skillName) Then
                        ToolingFileLogger.SetTopLevelSkillName(__skillName)
                    End If
                End If

                ToolingFileLogger.LogSubAgentReturn($"Agent-layer tool ({toolCall.ToolName})", response.Response)
                GoTo __AfterDispatch
            ElseIf toolCall.ToolName.StartsWith("skill_", StringComparison.OrdinalIgnoreCase) Then
                Dim skillArgs As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)

                skillArgs("name") = toolCall.ToolName.Substring("skill_".Length)

                If toolCall.Arguments IsNot Nothing Then
                    For Each kvp In toolCall.Arguments
                        If Not skillArgs.ContainsKey(kvp.Key) Then
                            skillArgs(kvp.Key) = kvp.Value
                        End If
                    Next

                    If Not skillArgs.ContainsKey("input") AndAlso toolCall.Arguments.ContainsKey("instruction") Then
                        skillArgs("input") = toolCall.Arguments("instruction")
                    End If
                End If

                response.Response = SharedLibrary.Agents.SkillInvokeTool.Execute(skillArgs)
                response.Success = Not String.IsNullOrWhiteSpace(response.Response)

                ApplyStructuredAgentResult(response, context)

                If response.Success Then
                    LoadSkillAllowedToolsFromResponse(response.Response, context)
                ElseIf String.IsNullOrWhiteSpace(response.ErrorMessage) Then
                    response.ErrorMessage = "Skill invocation returned no usable result."
                End If

                ToolingFileLogger.SetTopLevelSkillName(toolCall.ToolName.Substring("skill_".Length))
                ToolingFileLogger.LogSubAgentReturn($"Agent-layer skill ({toolCall.ToolName})", response.Response)
                GoTo __AfterDispatch
            End If


            If toolCall.ToolName.Equals(InternalWebToolName, StringComparison.OrdinalIgnoreCase) Then
                response = Await ExecuteInternalWebTool(toolCall, context, cancellationToken)
                ToolingFileLogger.LogRawResponseStub($"Internal tool ({toolCall.ToolName})", response.Response)

            ElseIf toolCall.ToolName.Equals(InternalDownloadWebFilesToolName, StringComparison.OrdinalIgnoreCase) Then
                response = Await ExecuteInternalDownloadWebFilesTool(toolCall, context, cancellationToken)
                ToolingFileLogger.LogRawResponseStub(
                $"Internal tool ({toolCall.ToolName})",
                response.Response)

            ElseIf toolCall.ToolName.Equals(InternalSearchToolName, StringComparison.OrdinalIgnoreCase) Then
                response = Await ExecuteInternalSearchTool(toolCall, context, cancellationToken)
                ToolingFileLogger.LogRawResponseStub($"Internal tool ({toolCall.ToolName})", response.Response)

            ElseIf IsInternalKnowledgeToolName(toolCall.ToolName) Then
                response = Await ExecuteInternalKnowledgeTool(toolCall, context)
                ToolingFileLogger.LogRawResponseStub($"Internal tool ({toolCall.ToolName})", response.Response)

            ElseIf SharedLibrary.SharedLibrary.M365ToolService.IsM365ToolName(toolCall.ToolName) Then
                ' M365 tools are allowed in interactive/local scenarios, but never in AutoPilot.
                ' They may trigger an interactive MSAL sign-in on the user's machine.
                If Not _apActive Then
                    response = Await ExecuteInternalM365Tool(toolCall, context)
                    ToolingFileLogger.LogRawResponseStub($"Internal tool ({toolCall.ToolName})", response.Response)
                Else
                    response.Success = False
                    response.ErrorMessage = "M365 tools cannot run inside AutoPilot because they may require interactive user sign-in."
                    ToolingFileLogger.LogWarn(
                        "M365 tool blocked in AutoPilot.",
                        details:=$"tool={toolCall.ToolName}; _apActive={_apActive}; _chatAgentActive={_chatAgentActive}")
                End If

            Else
                response = Await ExecuteExternalTool(toolCall, toolConfig, context, cancellationToken)
                ToolingFileLogger.LogRawResponseStub($"Tool LLM() ({toolCall.ToolName})", response.Response)
            End If

__AfterDispatch:

            ' Log completion with excerpt
            If response.Success Then
                Dim resultSummary As String = BuildResultExcerpt(response.Response, 80)
                context.Log($"Tool {toolCall.ToolName} completed: {resultSummary}", "success")
                ' Also write the full, unmodified tool response JSON so diagnostics
                ' (e.g. word_search/word_markup results) are not lost to truncation.
                ToolingFileLogger.LogRawResponse($"Tool result ({toolCall.ToolName})", response.Response)
            Else
                context.Log($"Tool {toolCall.ToolName} failed: {response.ErrorMessage}", "error")
            End If

            ' ── AutoPilot / Chat Agent: record tool call for dashboard + sources footer ──
            If (_apActive OrElse _chatAgentActive) AndAlso _apCurrentToolCallLog IsNot Nothing Then
                Dim elapsed = DateTime.Now - response.Timestamp
                Dim excerpt = BuildResultExcerpt(response.Response, 80)

                ' Log to dashboard (ApDashboardLog routes to AutoPilot OR Chat Agent automatically)
                ApDashboardLog($"🔧 External tool: {toolCall.ToolName}{paramSummary}", "info")
                If response.Success Then
                    ApDashboardLog($"   ✓ {excerpt}", "info")
                Else
                    ApDashboardLog($"   ✗ {If(response.ErrorMessage, excerpt)}", "error")
                End If

                ' Record for "Sources used:" footer
                Dim toolUrls As List(Of String) = Nothing
                If toolCall.Arguments IsNot Nothing AndAlso toolCall.Arguments.ContainsKey("urls") Then
                    Try
                        Dim urlsObj = toolCall.Arguments("urls")
                        If TypeOf urlsObj Is JArray Then
                            toolUrls = DirectCast(urlsObj, JArray).Select(Function(t) t.ToString()).ToList()
                        ElseIf TypeOf urlsObj Is IEnumerable(Of Object) Then
                            toolUrls = DirectCast(urlsObj, IEnumerable(Of Object)).Select(Function(o) o.ToString()).ToList()
                        ElseIf TypeOf urlsObj Is String Then
                            toolUrls = New List(Of String) From {CStr(urlsObj)}
                        End If
                    Catch
                    End Try
                End If
                RecordAutoPilotToolCall(
                    toolCall.ToolName,
                    If(toolConfig?.ModelDescription, toolCall.ToolName),
                    paramSummary,
                    isInternalTool:=False,
                    wasSuccessful:=response.Success,
                    resultExcerpt:=excerpt,
                    elapsed:=elapsed,
                    urls:=toolUrls)
            End If

        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation was cancelled"
            context.Log($"Tool {toolCall.ToolName} cancelled")
            ToolingFileLogger.LogWarn($"Tool {toolCall.ToolName} cancelled.")

        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            context.Log($"Tool {toolCall.ToolName} error: {ex.Message}")
            ToolingFileLogger.LogError($"Tool {toolCall.ToolName} execution error.", ex:=ex)
        End Try

        Return response
    End Function





    ''' <summary>
    ''' Loads tooling service configurations from an INI file and returns tool-capable <see cref="ModelConfig"/> entries.
    ''' </summary>
    ''' <param name="iniPath">INI path containing tool model sections.</param>
    ''' <param name="toolsOnly">When True, filters to entries that have tool-specific prompt/definition fields.</param>
    ''' <returns>List of available tool configurations.</returns>
    Public Function LoadToolingServices(iniPath As String, Optional toolsOnly As Boolean = True) As List(Of ModelConfig)
        Dim tools As New List(Of ModelConfig)()

        ' Register host-internal tool names in the executor registry (idempotent).
        ' External (INI) tools are registered below, after they have been confirmed
        ' to carry an APICall template. Anything not registered cannot be advertised.
        Agents.HostToolRegistration.RegisterOutlookInternals()

        iniPath = ExpandEnvironmentVariables(iniPath)

        If String.IsNullOrWhiteSpace(iniPath) OrElse Not File.Exists(iniPath) Then
            Return tools
        End If

        Try
            Dim allModels = LoadAlternativeModels(iniPath, _context, StartWithUpcase(ToolFriendlyName), includeToolOnly:=True, toolsOnly:=toolsOnly)

            For Each mc In allModels
                If mc.Deprecated Then Continue For

                If toolsOnly Then
                    If String.IsNullOrWhiteSpace(mc.ToolInstructionsPrompt) AndAlso
                String.IsNullOrWhiteSpace(mc.ToolDefinition) Then
                        Continue For
                    End If
                End If

                mc.Tool = True
                tools.Add(mc)

                ' Register transport-backed (external) tools that have an APICall
                ' template. A tool with no transport is intentionally NOT registered
                ' here; that prevents it from being advertised and forces the
                ' deliverable-fallback gate to look elsewhere.
                Dim apiTemplate As String =
             If(Not String.IsNullOrWhiteSpace(mc.ToolAPICall), mc.ToolAPICall, mc.APICall)
                If Not String.IsNullOrWhiteSpace(apiTemplate) AndAlso
            Not String.IsNullOrWhiteSpace(mc.ToolName) Then
                    Agents.ToolExecutorRegistry.RegisterExternal(
                 Agents.ToolingHostKind.Outlook, mc.ToolName)
                End If
            Next

        Catch ex As Exception
            Debug.WriteLine($"LoadToolingServices error: {ex.Message}")
            ToolingFileLogger.LogError("LoadToolingServices error.", ex:=ex)
        End Try

        Return tools
    End Function

    ''' <summary>
    ''' Shows the tool selection dialog and persists the selected tool names into <c>My.Settings.SelectedToolNames</c>.
    ''' </summary>
    ''' <param name="availableTools">List of available tool configurations.</param>
    ''' <param name="preselectAll">Unused parameter in this method body (caller passes a value).</param>
    ''' <returns>Selected tools when the dialog result is OK; otherwise Nothing.</returns>
    Public Function ShowToolSelectionDialog(availableTools As List(Of ModelConfig), Optional preselectAll As Boolean = True, Optional FriendlyName As String = "Tools") As List(Of ModelConfig)
        If availableTools Is Nothing OrElse availableTools.Count = 0 Then
            Return New List(Of ModelConfig)()
        End If

        ' Note: do NOT add built-in/internal suffixes here.
        ' MultiModelSelectorForm renders those labels centrally via HostToolRegistration.

        Dim selector As New MultiModelSelectorForm(
        availableTools,
        "",
        $"{AN} - Select {FriendlyName}",
        resetChecked:=False,
        preselectMany:=If(SelectedToolNames, New List(Of String)()),
        $"Select the {Globals.ThisAddIn.ToolFriendlyName.ToLower} you want to make available to the model:")

        selector.AddExtraButton("Skills && Agents…",
            Sub(s, e)
                Using f As New SharedLibrary.Agents.AgentResourcesViewerForm()
                    f.ShowDialog(selector)
                End Using
            End Sub)
        selector.AddExtraButton("Memory…",
            Sub(s, e)
                Using f As New SharedLibrary.Agents.SessionMemoryViewerForm()
                    f.ShowDialog(selector)
                End Using
            End Sub)

        If selector.ShowDialog() = DialogResult.OK Then
            Dim selected = selector.SelectedModels
            SelectedToolNames = selected.Select(Function(t) t.ToolName).ToList()
            Try
                My.Settings.SelectedToolNames = String.Join("|", SelectedToolNames)
                My.Settings.Save()
            Catch ex As Exception
                ToolingFileLogger.LogWarn("Failed to persist SelectedToolNames.", ex:=ex)
            End Try
            Return selected
        Else
            Return Nothing
        End If
    End Function

    ''' <summary>
    ''' Returns all available tools by loading external tools from <c>INI_SpecialServicePath</c>,
    ''' adding the internal web tool, conditionally adding the internal search tool
    ''' (only when <c>INI_ISearch</c> is enabled and <c>INI_ISearch_URL</c> is configured),
    ''' and conditionally adding the internal knowledge store search tool
    ''' (only when a knowledge store path is configured and at least one store is indexed).
    ''' </summary>
    ''' <returns>List of available tools.</returns>
    Public Function GetAvailableTools(Optional includeInteractiveM365Tools As Boolean = False) As List(Of ModelConfig)
        Dim tools As New List(Of ModelConfig)()
        Dim specialServicePath As String = ExpandEnvironmentVariables(INI_SpecialServicePath)

        If Not String.IsNullOrWhiteSpace(specialServicePath) Then
            Dim externalTools = LoadToolingServices(specialServicePath, True)
            tools.AddRange(externalTools)
        End If

        tools.Add(GetInternalWebTool())
        tools.Add(GetInternalDownloadWebFilesTool())

        Dim webGroundingTool =
    SharedLibrary.Agents.WebGroundingTool.Build(
        _context,
        enforcePrivacy:=INI_EnablePrivacyForSearch,
        toolPriority:=997,
        displaySuffix:=InternalToolSuffix)

        If webGroundingTool IsNot Nothing Then
            tools.Add(webGroundingTool)
        End If

        If INI_ISearch AndAlso Not String.IsNullOrWhiteSpace(INI_ISearch_URL) Then
            tools.Add(GetInternalSearchTool(enforcePrivacy:=INI_EnablePrivacyForSearch))
        End If

        tools.AddRange(GetInternalKnowledgeTools())

        ' python_execute: secure sandboxed Python execution.
        ' Only advertised when the executor path is set, the exe is available, and
        ' (when requested) its authenticity has been verified.
        Dim pythonExecuteTool As ModelConfig = Nothing
        If TryConfigureAndBuildPythonExecuteTool(pythonExecuteTool) Then
            tools.Add(pythonExecuteTool)
        End If

        ' M365 tools are interactive-only. Show them for Local Chat Agent source
        ' selection and execution, but never for AutoPilot.
        If (includeInteractiveM365Tools OrElse _chatAgentActive) AndAlso Not _apActive Then
            tools.AddRange(SharedLibrary.SharedLibrary.M365ToolService.GetTools(_context, InternalToolSuffix))
        End If

        ' Agent layer: session memory, skill loader, and discovered skills/agents (lazy registry-backed).
        Try
            SharedLibrary.Agents.AgentResources.Refresh()
            tools.AddRange(SharedLibrary.Agents.MemoryTools.BuildAll())
            tools.AddRange(SharedLibrary.Agents.TextTools.BuildAll())
            tools.AddRange(SharedLibrary.Agents.WordTools.BuildAll())
            tools.AddRange(SharedLibrary.Agents.BrowserTools.BuildAll(_context))

            Dim jsRunTool As ModelConfig = SharedLibrary.Agents.JsRunTool.Build(_context)
            If jsRunTool IsNot Nothing Then
                tools.Add(jsRunTool)
            End If

            tools.Add(SharedLibrary.Agents.SkillInvokeTool.Build())
            tools.Add(SharedLibrary.Agents.ToolDescribeTool.Build())
            tools.Add(SharedLibrary.Agents.ContextExpandTool.Build())
            tools.Add(SharedLibrary.Agents.ContextCompactTool.Build())
            If Not _apActive Then tools.Add(SharedLibrary.Agents.AskUserTool.Build())

            Dim __agentReg As New SharedLibrary.Agents.ToolRegistry()
            SharedLibrary.Agents.ToolRegistryBuilder.AddSkills(__agentReg, SharedLibrary.Agents.AgentResources.Skills)
            SharedLibrary.Agents.ToolRegistryBuilder.AddAgents(__agentReg, SharedLibrary.Agents.AgentResources.Agents)
            tools.AddRange(__agentReg.MaterializeAll())
        Catch ex As Exception
            ToolingFileLogger.LogWarn("Agent layer registration failed.", ex:=ex)
        End Try

        Return tools
    End Function

    ''' <summary>
    ''' Loads persisted tool selection from <c>My.Settings.SelectedToolNames</c> into <c>SelectedToolNames</c>.
    ''' </summary>
    Public Sub LoadPersistedToolSelection()
        Try
            Dim saved = My.Settings.SelectedToolNames
            If Not String.IsNullOrWhiteSpace(saved) Then
                SelectedToolNames = saved.Split("|"c).ToList()
            End If
        Catch ex As Exception
            SelectedToolNames = New List(Of String)()
            ToolingFileLogger.LogWarn("Failed to load persisted tool selection.", ex:=ex)
        End Try
    End Sub

    ''' <summary>
    ''' Selects tools for the current session either by reusing persisted selections or by showing the tool selection dialog.
    ''' </summary>
    ''' <param name="forceDialog">If True, always shows the selection dialog.</param>
    ''' <returns>Selected tool configurations, or Nothing when the dialog is canceled or no tools are available.</returns>
    Public Function SelectToolsForSession(Optional forceDialog As Boolean = False,
                                      Optional FriendlyName As String = ToolFriendlyName,
                                      Optional includeInteractiveM365Tools As Boolean = False) As List(Of ModelConfig)
        Dim availableTools = GetAvailableTools(includeInteractiveM365Tools)

        If availableTools.Count = 0 Then
            ShowCustomMessageBox($"No {FriendlyName.ToLower} are available. Configure 'Tooling' in the Special Services configuration file.")
            Return Nothing
        End If

        LoadPersistedToolSelection()

        If Not forceDialog AndAlso SelectedToolNames IsNot Nothing AndAlso SelectedToolNames.Count > 0 Then
            Dim selectedNameSet As New HashSet(Of String)(SelectedToolNames, StringComparer.OrdinalIgnoreCase)

            Dim selected = availableTools.
        Where(Function(t) Not String.IsNullOrWhiteSpace(t.ToolName) AndAlso selectedNameSet.Contains(t.ToolName)).
        ToList()

            If selected.Count > 0 Then
                Return selected
            End If
        End If

        Dim hasPersistedSelection As Boolean = (SelectedToolNames IsNot Nothing AndAlso SelectedToolNames.Count > 0)

        Return ShowToolSelectionDialog(availableTools, preselectAll:=Not hasPersistedSelection, FriendlyName)
    End Function


    Private Class InteractiveToolSelectionState
        Public Property SelectedMainToolNames As New List(Of String)()
        Public Property SelectedAdvancedToolNames As New List(Of String)()
    End Class

    Private Shared Function DeduplicateToolsByName(tools As IEnumerable(Of ModelConfig)) As List(Of ModelConfig)
        Dim result As New List(Of ModelConfig)()
        Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        If tools Is Nothing Then
            Return result
        End If

        For Each tool In tools
            If tool Is Nothing OrElse String.IsNullOrWhiteSpace(tool.ToolName) Then Continue For
            If seen.Add(tool.ToolName.Trim()) Then
                result.Add(tool)
            End If
        Next

        Return result
    End Function


    Private Function IsLocalChatAdvancedToolName(toolName As String) As Boolean
        If String.IsNullOrWhiteSpace(toolName) Then Return False

        ' web_grounding must remain user-visible in the main selector even though
        ' it is also an AutoPilot internal tool.
        If SharedLibrary.Agents.WebGroundingTool.IsWebGroundingTool(toolName) Then
            Return False
        End If

        If IsChatAgentWorkspaceTool(toolName) Then
            Return True
        End If

        If toolName.StartsWith("skill_", StringComparison.OrdinalIgnoreCase) OrElse
           toolName.StartsWith("agent_", StringComparison.OrdinalIgnoreCase) Then
            Return False
        End If

        If toolName.Equals(InternalWebToolName, StringComparison.OrdinalIgnoreCase) OrElse
           toolName.Equals(InternalSearchToolName, StringComparison.OrdinalIgnoreCase) OrElse
           IsInternalKnowledgeToolName(toolName) OrElse
           SharedLibrary.SharedLibrary.M365ToolService.IsM365ToolName(toolName) Then
            Return False
        End If

        If SharedLibrary.Agents.MemoryTools.IsMemoryTool(toolName) OrElse
           SharedLibrary.Agents.TextTools.IsTextTool(toolName) OrElse
           SharedLibrary.Agents.WorkspaceTools.IsWorkspaceTool(toolName) OrElse
           SharedLibrary.Agents.WordTools.IsWordTool(toolName) OrElse
           SharedLibrary.Agents.WordDocTools.IsWordDocTool(toolName) OrElse
           SharedLibrary.Agents.BrowserTools.IsBrowserTool(toolName) OrElse
           SharedLibrary.Agents.JsRunTool.IsJsTool(toolName) OrElse
           SharedLibrary.Agents.PythonExecuteTool.IsPythonTool(toolName) OrElse
           toolName.Equals(SharedLibrary.Agents.SkillInvokeTool.ToolName, StringComparison.OrdinalIgnoreCase) OrElse
           IsAutoPilotInternalTool(toolName) Then
            Return True
        End If

        Return False
    End Function


    Private Function IsLocalChatMainSelectableTool(tool As ModelConfig) As Boolean
        If tool Is Nothing OrElse String.IsNullOrWhiteSpace(tool.ToolName) Then Return False
        Return Not IsLocalChatAdvancedToolName(tool.ToolName)
    End Function

    Public Function GetLocalChatMainSelectableTools(Optional includeInteractiveM365Tools As Boolean = True) As List(Of ModelConfig)
        Dim availableTools = GetAvailableTools(includeInteractiveM365Tools)
        Return DeduplicateToolsByName(
            availableTools.Where(Function(t) IsLocalChatMainSelectableTool(t)))
    End Function


    Public Function GetAvailableToolsForAutoPilotSelection() As List(Of ModelConfig)
        Dim tools As New List(Of ModelConfig)()
        Dim specialServicePath As String = ExpandEnvironmentVariables(INI_SpecialServicePath)

        If Not String.IsNullOrWhiteSpace(specialServicePath) Then
            tools.AddRange(LoadToolingServices(specialServicePath, True))
        End If

        tools.Add(GetInternalWebTool())
        tools.Add(GetInternalDownloadWebFilesTool())

        If INI_ISearch AndAlso Not String.IsNullOrWhiteSpace(INI_ISearch_URL) Then
            tools.Add(GetInternalSearchTool(enforcePrivacy:=INI_EnablePrivacyForSearch))
        End If

        tools.AddRange(GetInternalKnowledgeTools())

        ' AutoPilot should also let the user explicitly select discovered skills and agents.
        Try
            SharedLibrary.Agents.AgentResources.Refresh()

            Dim agentRegistry As New SharedLibrary.Agents.ToolRegistry()
            SharedLibrary.Agents.ToolRegistryBuilder.AddSkills(agentRegistry, SharedLibrary.Agents.AgentResources.Skills)
            SharedLibrary.Agents.ToolRegistryBuilder.AddAgents(agentRegistry, SharedLibrary.Agents.AgentResources.Agents)

            tools.AddRange(agentRegistry.MaterializeAll())
        Catch ex As Exception
            ToolingFileLogger.LogWarn("AutoPilot skill/agent discovery failed.", ex:=ex)
        End Try

        Return DeduplicateToolsByName(tools)
    End Function

    Private Function NormalizeAutoPilotSelectableExternalTools(tools As IEnumerable(Of ModelConfig)) As List(Of ModelConfig)
        Dim result As New List(Of ModelConfig)()

        Dim allowedByName As Dictionary(Of String, ModelConfig) =
            GetAvailableToolsForAutoPilotSelection().
                Where(Function(t) t IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(t.ToolName)).
                GroupBy(Function(t) t.ToolName.Trim(), StringComparer.OrdinalIgnoreCase).
                ToDictionary(Function(g) g.Key, Function(g) g.First(), StringComparer.OrdinalIgnoreCase)

        If tools Is Nothing Then
            Return result
        End If

        For Each tool In tools
            If tool Is Nothing OrElse String.IsNullOrWhiteSpace(tool.ToolName) Then Continue For

            Dim normalized As ModelConfig = Nothing
            If allowedByName.TryGetValue(tool.ToolName.Trim(), normalized) Then
                result.Add(normalized)
            End If
        Next

        Return DeduplicateToolsByName(result)
    End Function

    Public Function GetLocalChatAdvancedSelectableTools(Optional includeInteractiveM365Tools As Boolean = True) As List(Of ModelConfig)
        Dim tools As New List(Of ModelConfig)()

        tools.AddRange(
            GetAvailableTools(includeInteractiveM365Tools).
                Where(Function(t) t IsNot Nothing AndAlso
                                  Not String.IsNullOrWhiteSpace(t.ToolName) AndAlso
                                  IsLocalChatAdvancedToolName(t.ToolName)))

        If IsChatAgentWorkspaceConnected() Then
            tools.AddRange(GetChatAgentWorkspaceTools())
        End If

        Try
            For Each tool In GetAutoPilotInternalTools()
                If tool Is Nothing OrElse String.IsNullOrWhiteSpace(tool.ToolName) Then Continue For
                If tool.ToolName.Equals(AP_Tool_SummarizeThread, StringComparison.OrdinalIgnoreCase) Then Continue For
                tools.Add(tool)
            Next
        Catch
        End Try

        If INI_WebServerBlock <> 4 Then
            tools.Add(BuildManageScheduledTasksTool())
        End If

        Return DeduplicateToolsByName(tools)
    End Function

    ' ── Opt-out resolution helpers ──────────────────────────────────────────────
    ' Effective selection = all tools available for the current workspace/connection
    ' state MINUS the tools the user explicitly deselected. New tools are therefore on
    ' automatically, and workspace tools turn on as soon as a workspace is connected
    ' unless the user expressly deselected them earlier.

    Private Function GetLocalChatAvailableMainToolsForState(Optional includeInteractiveM365Tools As Boolean = True) As List(Of ModelConfig)
        Return GetLocalChatMainSelectableTools(includeInteractiveM365Tools)
    End Function

    Private Function GetLocalChatAvailableAdvancedToolsForState(Optional includeInteractiveM365Tools As Boolean = True) As List(Of ModelConfig)
        ' GetLocalChatAdvancedSelectableTools already includes workspace tools only when a
        ' workspace is connected, so it is the correct "available for current state" set.
        Return GetLocalChatAdvancedSelectableTools(includeInteractiveM365Tools)
    End Function

    Private Function ResolveLocalChatSelectedMainToolNames(st As InkyState,
                                                          Optional includeInteractiveM365Tools As Boolean = True) As List(Of String)
        Dim deselected = BuildToolNameSet(If(st IsNot Nothing, st.DeselectedMainToolNames, Nothing))
        Return GetLocalChatAvailableMainToolsForState(includeInteractiveM365Tools).
            Where(Function(t) t IsNot Nothing AndAlso
                              Not String.IsNullOrWhiteSpace(t.ToolName) AndAlso
                              Not deselected.Contains(t.ToolName)).
            Select(Function(t) t.ToolName).
            Distinct(StringComparer.OrdinalIgnoreCase).
            ToList()
    End Function

    Private Function ResolveLocalChatSelectedAdvancedToolNames(st As InkyState,
                                                              Optional includeInteractiveM365Tools As Boolean = True) As List(Of String)
        Dim deselected = BuildToolNameSet(If(st IsNot Nothing, st.DeselectedAdvancedToolNames, Nothing))
        Return GetLocalChatAvailableAdvancedToolsForState(includeInteractiveM365Tools).
            Where(Function(t) t IsNot Nothing AndAlso
                              Not String.IsNullOrWhiteSpace(t.ToolName) AndAlso
                              Not deselected.Contains(t.ToolName)).
            Select(Function(t) t.ToolName).
            Distinct(StringComparer.OrdinalIgnoreCase).
            ToList()
    End Function

    Public Function ComputeLocalChatDeselectedMainToolNames(checkedMainToolNames As IEnumerable(Of String),
                                                            Optional includeInteractiveM365Tools As Boolean = True) As List(Of String)
        Dim checkedSet = BuildToolNameSet(checkedMainToolNames)
        Return GetLocalChatAvailableMainToolsForState(includeInteractiveM365Tools).
            Where(Function(t) t IsNot Nothing AndAlso
                              Not String.IsNullOrWhiteSpace(t.ToolName) AndAlso
                              Not checkedSet.Contains(t.ToolName)).
            Select(Function(t) t.ToolName).
            Distinct(StringComparer.OrdinalIgnoreCase).
            ToList()
    End Function

    Public Function ComputeLocalChatDeselectedAdvancedToolNames(checkedAdvancedToolNames As IEnumerable(Of String),
                                                                Optional includeInteractiveM365Tools As Boolean = True) As List(Of String)
        Dim checkedSet = BuildToolNameSet(checkedAdvancedToolNames)
        Return GetLocalChatAvailableAdvancedToolsForState(includeInteractiveM365Tools).
            Where(Function(t) t IsNot Nothing AndAlso
                              Not String.IsNullOrWhiteSpace(t.ToolName) AndAlso
                              Not checkedSet.Contains(t.ToolName)).
            Select(Function(t) t.ToolName).
            Distinct(StringComparer.OrdinalIgnoreCase).
            ToList()
    End Function

    Public Function GetLocalChatEffectiveTools(selectedMainToolNames As IEnumerable(Of String),
                                           selectedAdvancedToolNames As IEnumerable(Of String),
                                           advancedToolsEnabled As Boolean,
                                           Optional includeInteractiveM365Tools As Boolean = True) As List(Of ModelConfig)

        ' Opt-out model: callers pass already-resolved *selected* names. This is a pure
        ' intersection with no implicit "select all advanced" fallback, so deselecting
        ' every advanced tool genuinely yields none.
        Dim result As New List(Of ModelConfig)()
        Dim mainSet = BuildToolNameSet(selectedMainToolNames)
        Dim advancedSet = BuildToolNameSet(selectedAdvancedToolNames)

        For Each tool In GetLocalChatMainSelectableTools(includeInteractiveM365Tools)
            If mainSet.Contains(tool.ToolName) Then
                result.Add(tool)
            End If
        Next

        If advancedToolsEnabled Then
            For Each tool In GetLocalChatAdvancedSelectableTools(includeInteractiveM365Tools)
                If advancedSet.Contains(tool.ToolName) Then
                    result.Add(tool)
                End If
            Next
        End If

        Return DeduplicateToolsByName(result)
    End Function

    Private Function ShowLocalChatAdvancedToolSelectionDialog(selectedAdvancedToolNames As IEnumerable(Of String),
                                                          Optional includeInteractiveM365Tools As Boolean = True) As List(Of String)

        Dim availableTools = GetLocalChatAdvancedSelectableTools(includeInteractiveM365Tools)
        Dim preselected = If(selectedAdvancedToolNames, Enumerable.Empty(Of String)()).
            Where(Function(n) Not String.IsNullOrWhiteSpace(n)).
            Select(Function(n) n.Trim()).
            Distinct(StringComparer.OrdinalIgnoreCase).
            ToList()

        Using selector As New MultiModelSelectorForm(
        availableTools,
        "",
        $"{AN} - Select Advanced Tools",
              resetChecked:=False,
        preselectMany:=preselected,
        instruction:="Select the advanced tools that may be callable in Local Chat. " &
                     "All available tools are on by default; uncheck any you want to disable. " &
                     "Workspace tools appear here only while a workspace is connected.")

            If selector.ShowDialog() = DialogResult.OK Then
                Return selector.SelectedModels.
                    Where(Function(t) t IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(t.ToolName)).
                    Select(Function(t) t.ToolName).
                    Distinct(StringComparer.OrdinalIgnoreCase).
                    ToList()
            End If
        End Using

        Return Nothing
    End Function


    Public Function ShowLocalChatToolSelectionDialog(selectedMainToolNames As IEnumerable(Of String),
                                                     selectedAdvancedToolNames As IEnumerable(Of String),
                                                     ByRef updatedAdvancedToolNames As List(Of String),
                                                     Optional includeInteractiveM365Tools As Boolean = True) As List(Of String)

        Dim availableTools = GetLocalChatMainSelectableTools(includeInteractiveM365Tools)
        Dim workingAdvanced As New List(Of String)(
            If(selectedAdvancedToolNames, Enumerable.Empty(Of String)()).
                Where(Function(n) Not String.IsNullOrWhiteSpace(n)).
                Select(Function(n) n.Trim()).
                Distinct(StringComparer.OrdinalIgnoreCase))

        Using selector As New MultiModelSelectorForm(
            availableTools,
            "",
            $"{AN} - Select {ToolFriendlyName}",
            resetChecked:=False,
            preselectMany:=If(selectedMainToolNames, New List(Of String)()),
            instruction:="Select the agents, sources, skills, and connector-oriented tools you want to make available to the model. " &
                         "Note: all available tools are on by default (including newly added ones); uncheck any you want to disable. " &
                         "Advanced tools are managed separately through the 'Advanced tools…' button.")

            selector.AddExtraButton("Advanced tools…",
                Sub(s, e)
                    Dim advanced = ShowLocalChatAdvancedToolSelectionDialog(
                        workingAdvanced,
                        includeInteractiveM365Tools)

                    If advanced IsNot Nothing Then
                        workingAdvanced = advanced
                    End If
                End Sub)

            selector.AddExtraButton("Skills && Agents…",
                Sub(s, e)
                    Using f As New SharedLibrary.Agents.AgentResourcesViewerForm(_context)
                        f.ShowDialog(selector)
                    End Using
                End Sub)

            selector.AddExtraButton("Memory…",
                Sub(s, e)
                    Using f As New SharedLibrary.Agents.SessionMemoryViewerForm()
                        f.ShowDialog(selector)
                    End Using
                End Sub)

            If selector.ShowDialog() = DialogResult.OK Then
                updatedAdvancedToolNames = workingAdvanced.
                    Distinct(StringComparer.OrdinalIgnoreCase).
                    ToList()

                Return selector.SelectedModels.
                    Where(Function(t) t IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(t.ToolName)).
                    Select(Function(t) t.ToolName).
                    Distinct(StringComparer.OrdinalIgnoreCase).
                    ToList()
            End If
        End Using

        Return Nothing
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
        hideSplash As Boolean,
        cancellationToken As System.Threading.CancellationToken,
        effectiveTimeout As Integer,
        binaryOutputDirectory As String) As Task(Of String)

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

        Using timeoutCts As New System.Threading.CancellationTokenSource()
            Dim totalTimeout = effectiveTimeout + 60
            timeoutCts.CancelAfter(TimeSpan.FromSeconds(totalTimeout))

            Using combinedCts As System.Threading.CancellationTokenSource =
                System.Threading.CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, timeoutCts.Token)

                Try
                    Dim forcedFinalResponse As String = Await LLM(
                        forcedSysPrompt,
                        fullUserPrompt,
                        "", "", 0,
                        useSecondAPI,
                        hideSplash,
                        otherPrompt,
                        fileObject,
                        combinedCts.Token,
                        True, True, binaryOutputDirectory:=binaryOutputDirectory)

                    If Not String.IsNullOrWhiteSpace(forcedFinalResponse) Then
                        ToolingFileLogger.LogRawResponseStub("Main LLM() - Forced Final After Invalid Turn", forcedFinalResponse)
                    Else
                        ToolingFileLogger.LogWarn(
                            "Empty forced-final response after repeated invalid turn.",
                            details:=$"host={context.HostKind}")
                    End If

                    Return If(forcedFinalResponse, "")
                Catch ex As OperationCanceledException
                    context.LogWarn("Forced no-tool finalization call was cancelled")
                    Return ""
                Catch ex As Exception
                    context.LogError($"Error during forced no-tool finalization: {ex.Message}", ex:=ex)
                    Return ""
                End Try
            End Using
        End Using
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

End Class
