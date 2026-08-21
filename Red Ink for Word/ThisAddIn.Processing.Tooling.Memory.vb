' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.Processing.Tooling.Memory.vb
' Purpose: Memory grounding mode resolution and classification for the tooling loop.
'
' Responsibilities:
'  - Classify user intent for memory access requirements via LLM classifier.
'  - Build memory grounding system/user prompts for intent detection.
'  - Validate memory tool availability (tools, workflow entries).
'  - Resolve explicit vs. auto-detected memory grounding modes (None, Optional, Required).
'  - Enforce explicit stored-memory demand when necessary.
'  - Manage memory exposure to the model (recent stubs injection).
'  - Handle sub-agent memory isolation (no memory access in delegated execution).
'
' External Dependencies:
'  - SharedLibrary.Agents.ToolCallSequencing for memory classification logic.
'  - SharedLibrary.Agents.MemoryTools and SessionMemory for tool/memory availability.
'  - LLM() for intent classification.
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


Partial Public Class ThisAddIn


    Private Function BuildMemoryGroundingClassifierInput(latestUserRequestRaw As String,
                                                         hostTaskSummary As String) As String
        If String.IsNullOrWhiteSpace(latestUserRequestRaw) Then
            Return ""
        End If

        Return SharedLibrary.Agents.ToolCallSequencing.BuildMemoryGroundingIntentClassifierUserPrompt(
            latestUserRequestRaw,
            hostTaskSummary)
    End Function


    Private Function HasMemoryGroundingClassifierInputsAvailable(context As ToolExecutionContext) As Boolean
        If context Is Nothing Then
            Return False
        End If

        Dim memoryToolsAvailable As Boolean = False

        If context.AuthoritativeToolRegistrySnapshot IsNot Nothing Then
            memoryToolsAvailable =
                context.AuthoritativeToolRegistrySnapshot.Contains(SharedLibrary.Agents.MemoryTools.ToolList) OrElse
                context.AuthoritativeToolRegistrySnapshot.Contains(SharedLibrary.Agents.MemoryTools.ToolGet)
        End If

        If Not memoryToolsAvailable AndAlso context.SelectedTools IsNot Nothing Then
            memoryToolsAvailable =
                context.SelectedTools.Any(
                    Function(tool)
                        Return tool IsNot Nothing AndAlso
                               SharedLibrary.Agents.MemoryTools.IsMemoryTool(tool.ToolName)
                    End Function)
        End If

        Dim workflowMemoryAvailable As Boolean = False

        Try
            workflowMemoryAvailable =
                SharedLibrary.Agents.SessionMemory.ListMostRecentWorkflowEntries(maxItems:=1).Count > 0
        Catch
            workflowMemoryAvailable = False
        End Try

        Return memoryToolsAvailable OrElse workflowMemoryAvailable
    End Function

    Private Function ResolveMemoryGroundingToolConfig(context As ToolExecutionContext,
                                                      toolName As String) As ModelConfig
        If context Is Nothing OrElse String.IsNullOrWhiteSpace(toolName) Then
            Return Nothing
        End If

        If context.SelectedTools IsNot Nothing Then
            Dim existing = context.SelectedTools.FirstOrDefault(
                Function(tool)
                    Return tool IsNot Nothing AndAlso
                           Not String.IsNullOrWhiteSpace(tool.ToolName) AndAlso
                           tool.ToolName.Equals(toolName, StringComparison.OrdinalIgnoreCase)
                End Function)

            If existing IsNot Nothing Then
                Return existing
            End If
        End If

        If context.AllowedToolRegistry Is Nothing Then
            Return Nothing
        End If

        Return EnsureVisibleToolLoaded(toolName, context)
    End Function

    Private Function BuildSyntheticMemoryGroundingToolCallJson(toolName As String,
                                                               arguments As IDictionary(Of String, Object)) As String
        Dim argsObject As New JObject()

        If arguments IsNot Nothing Then
            For Each kvp In arguments
                If String.IsNullOrWhiteSpace(kvp.Key) Then
                    Continue For
                End If

                If TypeOf kvp.Value Is JToken Then
                    argsObject(kvp.Key) = DirectCast(kvp.Value, JToken)
                ElseIf kvp.Value Is Nothing Then
                    argsObject(kvp.Key) = JValue.CreateNull()
                Else
                    argsObject(kvp.Key) = JToken.FromObject(kvp.Value)
                End If
            Next
        End If

        Return New JObject(
            New JProperty("name", If(toolName, "")),
            New JProperty("arguments", argsObject)).ToString(Formatting.None)
    End Function

    Private Sub RecordHostMemoryGroundingPrimeResponse(context As ToolExecutionContext,
                                                       toolCall As ToolCall,
                                                       toolResponse As ToolResponse)
        If context Is Nothing OrElse toolCall Is Nothing OrElse toolResponse Is Nothing Then
            Return
        End If

        toolResponse.OriginalCallJson = toolCall.RawJson
        context.AllToolResponses.Add(toolResponse)

        SharedLibrary.Agents.ToolCallSequencing.NoteToolExecutionMetadata(
            context.SequencingState,
            toolCall.ToolName,
            toolCall.Arguments,
            toolResponse.Success)

        If toolResponse.Success Then
            SharedLibrary.Agents.ToolCallSequencing.NoteToolResultForRepair(
                context.SequencingState,
                toolCall.ToolName,
                toolResponse.Response,
                toolResponse.ResultKind)

            If context.SequencingState IsNot Nothing Then
                context.SequencingState.NoteSuccessfulProgress()
            End If
        End If

        UpdateWorkflowContinuityAfterToolExecution(context, toolCall, toolResponse)

        SharedLibrary.Agents.ToolCallSequencing.NoteMemoryGroundingToolResult(
            context.SequencingState,
            toolCall.ToolName,
            toolResponse.Response,
            toolResponse.Success)

        context.Log(
            "Memory grounding state updated: " &
            SharedLibrary.Agents.ToolCallSequencing.BuildMemoryGroundingStateSummary(context.SequencingState),
            "diag")
    End Sub

    Private Async Function TryPrimeRecentMemoryStubsAsync(context As ToolExecutionContext,
                                                      cancellationToken As System.Threading.CancellationToken) As Task(Of Boolean)
        If context Is Nothing OrElse context.SequencingState Is Nothing Then
            Return False
        End If

        If Not context.SequencingState.IsRequiredMemoryGroundingEnforced Then
            Return False
        End If

        If Not context.SequencingState.ShouldExposeRecentMemoryStubs Then
            Return False
        End If

        If context.SequencingState.MemoryListCalledThisTurn Then
            Return False
        End If

        Dim listConfig As ModelConfig =
            ResolveMemoryGroundingToolConfig(context, SharedLibrary.Agents.MemoryTools.ToolList)

        If listConfig Is Nothing Then
            context.LogWarn(
                "Host memory priming skipped: memory_list is not available for this session.",
                details:=$"host={context.HostKind}")
            Return False
        End If

        Dim arguments As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)

        Dim toolCall As New ToolCall With {
            .CallId = "host_memory_list_" & If(context.RunId, Guid.NewGuid().ToString("N")),
            .ToolName = SharedLibrary.Agents.MemoryTools.ToolList,
            .Arguments = arguments,
            .RawJson = BuildSyntheticMemoryGroundingToolCallJson(
                SharedLibrary.Agents.MemoryTools.ToolList,
                arguments)
        }

        context.Log("Host memory priming started via memory_list.", "diag")

        Dim toolResponse As ToolResponse =
            Await ExecuteToolCall(toolCall, listConfig, context, cancellationToken)

        RecordHostMemoryGroundingPrimeResponse(context, toolCall, toolResponse)

        context.Log(
            "Host memory priming completed: success=" & If(toolResponse.Success, "true", "false"),
            If(toolResponse.Success, "diag", "warn"))

        Return toolResponse.Success
    End Function

    Private Function FindLatestSuccessfulMemoryListResponse(context As ToolExecutionContext) As String
        If context Is Nothing OrElse context.AllToolResponses Is Nothing Then
            Return ""
        End If

        Dim latest As ToolResponse =
            context.AllToolResponses.
                LastOrDefault(
                    Function(response)
                        Return response IsNot Nothing AndAlso
                               response.Success AndAlso
                               Not String.IsNullOrWhiteSpace(response.ToolName) AndAlso
                               response.ToolName.Equals(SharedLibrary.Agents.MemoryTools.ToolList, StringComparison.OrdinalIgnoreCase)
                    End Function)

        Return If(latest?.Response, "")
    End Function

    Private Async Function TryForceRequiredMemoryGetsAsync(context As ToolExecutionContext,
                                                           cancellationToken As System.Threading.CancellationToken) As Task(Of Boolean)
        If context Is Nothing OrElse context.SequencingState Is Nothing Then
            Return False
        End If

        If Not context.SequencingState.IsRequiredMemoryGroundingEnforced Then
            Return False
        End If

        If Not context.SequencingState.MemoryGetRequiredAfterList Then
            Return False
        End If

        If context.SequencingState.MemoryKeysSuggestedForGet Is Nothing OrElse
           context.SequencingState.MemoryKeysSuggestedForGet.Count = 0 Then
            Return False
        End If

        Dim memoryGetConfig As ModelConfig =
            ResolveMemoryGroundingToolConfig(context, SharedLibrary.Agents.MemoryTools.ToolGet)

        If memoryGetConfig Is Nothing Then
            context.LogWarn("Host memory_get follow-up skipped: memory_get tool is unavailable.",
                            details:=$"host={context.HostKind}")
            Return False
        End If

        Dim memoryListResponse As String = FindLatestSuccessfulMemoryListResponse(context)

        Dim selectedKeys As List(Of String) =
            SharedLibrary.Agents.ToolCallSequencing.SelectDeterministicMemoryKeysForHostFollowUp(
                memoryListResponse,
                context.WorkflowId,
                context.LatestUserRequestRaw)

        If selectedKeys.Count = 0 Then
            selectedKeys = context.SequencingState.MemoryKeysSuggestedForGet.ToList()
        End If

        Dim executedAny As Boolean = False

        For Each key As String In selectedKeys
            If String.IsNullOrWhiteSpace(key) Then
                Continue For
            End If

            If context.SequencingState.MemoryKeysRetrievedThisTurn IsNot Nothing AndAlso
               context.SequencingState.MemoryKeysRetrievedThisTurn.Contains(key, StringComparer.OrdinalIgnoreCase) Then
                Continue For
            End If

            Dim arguments As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase) From {
                {"key", key}
            }

            Dim toolCall As New ToolCall With {
                .CallId = "host_memory_get_" & key.Replace(" ", "_"),
                .ToolName = SharedLibrary.Agents.MemoryTools.ToolGet,
                .Arguments = arguments,
                .RawJson = BuildSyntheticMemoryGroundingToolCallJson(
                    SharedLibrary.Agents.MemoryTools.ToolGet,
                    arguments)
            }

            context.Log($"Host memory_get follow-up started for key '{key}'.", "diag")

            Dim toolResponse As ToolResponse =
                Await ExecuteToolCall(toolCall, memoryGetConfig, context, cancellationToken)

            RecordHostMemoryGroundingPrimeResponse(context, toolCall, toolResponse)

            executedAny = True
        Next

        Return executedAny
    End Function

    Private Sub ApplyBootstrapMemoryGroundingDecision(context As ToolExecutionContext,
                                                       decision As SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingIntentDecision)
        If context Is Nothing OrElse context.SequencingState Is Nothing OrElse decision Is Nothing OrElse Not decision.IsValid Then
            Return
        End If

        If decision.MemoryGroundingMode = SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingMode.Required AndAlso
           Not decision.ExplicitStoredMemoryRequired Then

            context.LogWarn(
                "bootstrapMemoryRequiredModeDowngraded",
                details:=$"host={context.HostKind}; reason=missing_explicit_user_demand_for_stored_memory; classifierReason={If(decision.Reason, "")}",
                visibleToUser:=False)

            decision.MemoryGroundingMode = SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingMode.OptionalMode
        End If

        context.SequencingState.MemoryGroundingMode = decision.MemoryGroundingMode
        context.SequencingState.MemoryGroundingAuthority =
            If(decision.MemoryGroundingMode = SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingMode.None,
               SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingAuthority.None,
               SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingAuthority.Classifier)
        context.SequencingState.ShouldExposeRecentMemoryStubs = decision.ShouldExposeRecentMemoryStubs

        If context.SequencingState.MemoryGroundingMode <> SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingMode.None AndAlso
           context.SequencingState.ShouldExposeRecentMemoryStubs Then

            ResolveMemoryGroundingToolConfig(context, SharedLibrary.Agents.MemoryTools.ToolList)
            ResolveMemoryGroundingToolConfig(context, SharedLibrary.Agents.MemoryTools.ToolGet)

            If context.SequencingState.MemoryGroundingStage = SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingStage.NotStarted Then
                context.SequencingState.MemoryGroundingStage = SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingStage.ListRequired
            End If
        End If

        context.Log(
            "Bootstrap memory grounding applied: " &
            SharedLibrary.Agents.ToolCallSequencing.BuildMemoryGroundingStateSummary(context.SequencingState) &
            "; explicitStoredMemoryRequired=" & If(decision.ExplicitStoredMemoryRequired, "true", "false") &
            "; reason=" & If(decision.Reason, ""),
            "diag")
    End Sub


    Private Async Function ResolveMemoryGroundingModeAsync(context As ToolExecutionContext,
                                                           userText As String,
                                                           otherPrompt As String,
                                                           fullPromptOverride As String,
                                                           useSecondAPI As Boolean,
                                                           hideSplash As Boolean,
                                                           explicitMemoryGroundingMode As SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingMode,
                                                           memoryGroundingModeIsExplicit As Boolean,
                                                           subAgentMode As Boolean) As Task
        If context Is Nothing OrElse context.SequencingState Is Nothing Then
            Return
        End If

        If subAgentMode Then
            context.SequencingState.MemoryGroundingMode = SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingMode.None
            context.SequencingState.MemoryGroundingAuthority = SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingAuthority.None
            context.SequencingState.ShouldExposeRecentMemoryStubs = False
            context.Log("Memory grounding classifier skipped for sub-agent mode.", "diag")
            Return
        End If

        If memoryGroundingModeIsExplicit Then
            context.SequencingState.MemoryGroundingMode = explicitMemoryGroundingMode
            context.SequencingState.MemoryGroundingAuthority =
        If(explicitMemoryGroundingMode = SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingMode.None,
           SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingAuthority.None,
           SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingAuthority.ExplicitOverride)
            context.SequencingState.ShouldExposeRecentMemoryStubs =
        explicitMemoryGroundingMode <> SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingMode.None

            If context.SequencingState.MemoryGroundingMode <> SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingMode.None AndAlso
               context.SequencingState.ShouldExposeRecentMemoryStubs Then

                ResolveMemoryGroundingToolConfig(context, SharedLibrary.Agents.MemoryTools.ToolList)
                ResolveMemoryGroundingToolConfig(context, SharedLibrary.Agents.MemoryTools.ToolGet)

                If context.SequencingState.MemoryGroundingStage = SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingStage.NotStarted Then
                    context.SequencingState.MemoryGroundingStage = SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingStage.ListRequired
                End If
            End If

            context.Log("Memory grounding mode applied from explicit override: " &
                SharedLibrary.Agents.ToolCallSequencing.BuildMemoryGroundingStateSummary(context.SequencingState))
            Return
        End If

        If Not HasMemoryGroundingClassifierInputsAvailable(context) Then
            context.SequencingState.MemoryGroundingMode = SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingMode.None
            context.SequencingState.MemoryGroundingAuthority = SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingAuthority.None
            context.SequencingState.ShouldExposeRecentMemoryStubs = False
            context.Log("Memory grounding classifier skipped: no memory tools and no workflow memory available.", "diag")
            Return
        End If

        Dim classifierInput As String =
            BuildMemoryGroundingClassifierInput(
                context.LatestUserRequestRaw,
                context.HostTaskSummary)

        If String.IsNullOrWhiteSpace(classifierInput) Then
            context.SequencingState.MemoryGroundingMode = SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingMode.None
            context.SequencingState.MemoryGroundingAuthority = SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingAuthority.None
            context.SequencingState.ShouldExposeRecentMemoryStubs = False
            context.Log("Memory grounding classifier skipped: no classifier input was available.", "diag")
            Return
        End If

        context.Log("Memory grounding classifier invoked.", "diag")
        Dim raw As String = ""

        LogLatestUserRequestDiagnostic(context, "classifier")

        Try
            raw = Await LLM(
                SharedLibrary.Agents.ToolCallSequencing.BuildMemoryGroundingIntentClassifierSystemPrompt(),
                classifierInput,
                "", "", 0,
                useSecondAPI,
                hideSplash,
                "",
                "",
                True)
        Catch ex As Exception
            context.LogWarn("Memory grounding classifier failed; defaulting to none.",
                            details:=$"host={context.HostKind}; error={ex.Message}")
        End Try

        context.Log("classifierRawOutput=" & If(raw, ""), "diag")

        Dim classifierNormalizedOutput As String = ""
        Dim classifierParseError As String = ""
        Dim decision =
            SharedLibrary.Agents.ToolCallSequencing.ParseMemoryGroundingIntentClassifierDecision(
                raw,
                classifierNormalizedOutput,
                classifierParseError)

        Dim normalizedOutputForLog As String = If(classifierNormalizedOutput, "")
        If normalizedOutputForLog.Length > 600 Then
            normalizedOutputForLog = normalizedOutputForLog.Substring(0, 600) & "..."
        End If

        context.Log("classifierNormalizedOutput=" & normalizedOutputForLog, "diag")
        context.Log("classifierParseSuccess=" & If(decision.IsValid, "true", "false"), "diag")

        If Not decision.IsValid Then
            context.Log("classifierParseError=" & If(classifierParseError, "invalid_classifier_output"), "diag")
            context.LogWarn(
                "classifierParseFailure",
                details:=$"host={context.HostKind}; classifierParseError={If(classifierParseError, "invalid_classifier_output")}")
        End If

        context.Log(
            "parsedMemoryGroundingMode=" &
            SharedLibrary.Agents.ToolCallSequencing.FormatMemoryGroundingMode(decision.MemoryGroundingMode) &
            "; shouldExposeRecentMemoryStubs=" &
            If(decision.ShouldExposeRecentMemoryStubs, "true", "false") &
            "; explicitStoredMemoryRequired=" &
            If(decision.ExplicitStoredMemoryRequired, "true", "false") &
            "; reason=" & If(decision.Reason, ""), "diag")

        If decision.MemoryGroundingMode = SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingMode.Required AndAlso
           Not decision.ExplicitStoredMemoryRequired Then

            context.LogWarn(
                "classifierRequiredModeDowngraded",
                details:=$"host={context.HostKind}; reason=missing_explicit_user_demand_for_stored_memory; classifierReason={If(decision.Reason, "")}")

            decision.MemoryGroundingMode = SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingMode.OptionalMode
        End If

        context.SequencingState.MemoryGroundingMode = decision.MemoryGroundingMode
        context.SequencingState.MemoryGroundingAuthority =
    If(decision.MemoryGroundingMode = SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingMode.None,
       SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingAuthority.None,
       SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingAuthority.Classifier)
        context.SequencingState.ShouldExposeRecentMemoryStubs = decision.ShouldExposeRecentMemoryStubs

        If context.SequencingState.MemoryGroundingMode <> SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingMode.None AndAlso
           context.SequencingState.ShouldExposeRecentMemoryStubs Then

            ResolveMemoryGroundingToolConfig(context, SharedLibrary.Agents.MemoryTools.ToolList)
            ResolveMemoryGroundingToolConfig(context, SharedLibrary.Agents.MemoryTools.ToolGet)

            If context.SequencingState.MemoryGroundingStage = SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingStage.NotStarted Then
                context.SequencingState.MemoryGroundingStage = SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingStage.ListRequired
            End If
        End If


        context.Log(
            "appliedMemoryGroundingMode=" &
            SharedLibrary.Agents.ToolCallSequencing.FormatMemoryGroundingMode(context.SequencingState.MemoryGroundingMode) &
            "; explicitStoredMemoryRequired=" &
            If(decision.ExplicitStoredMemoryRequired, "true", "false"), "diag")

        context.Log(
    "Memory grounding mode applied: " &
    SharedLibrary.Agents.ToolCallSequencing.BuildMemoryGroundingStateSummary(context.SequencingState),
    "diag")

    End Function

End Class
