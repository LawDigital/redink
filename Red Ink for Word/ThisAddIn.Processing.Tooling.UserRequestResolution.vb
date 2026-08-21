' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.Processing.Tooling.UserRequestResolution.vb
' Purpose: User language detection and request text extraction for tooling sessions.
'
' Responsibilities:
'  - Detect user's preferred language via LLM classification.
'  - Extract latest user turn from dialog/prompt structures.
'  - Parse BCP-47 language tags and localization preferences.
'  - Support fallback language handling.
'
' External Dependencies:
'  - LLM() for language detection classification.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.Linq
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


    Private Async Function ResolveToolingUserLanguageAsync(userText As String,
                                                           otherPrompt As String,
                                                           fullPromptOverride As String,
                                                           useSecondAPI As Boolean,
                                                           hideSplash As Boolean) As Task(Of String)
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

    Private Async Function ResolveToolingBootstrapPreflightAsync(
        context As ToolExecutionContext,
        useSecondAPI As Boolean,
        hideSplash As Boolean,
        cancellationToken As System.Threading.CancellationToken,
        explicitUserLanguage As String,
        explicitMemoryGroundingMode As SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingMode,
        memoryGroundingModeIsExplicit As Boolean,
        subAgentMode As Boolean) As System.Threading.Tasks.Task(Of SharedLibrary.Agents.ToolingBootstrapPreflight.Decision)

        If context Is Nothing OrElse subAgentMode Then
            Return Nothing
        End If

        Dim manifests As IEnumerable(Of SharedLibrary.Agents.ToolManifest) =
            If(context.AllowedToolRegistry Is Nothing,
               Enumerable.Empty(Of SharedLibrary.Agents.ToolManifest)(),
               context.AllowedToolRegistry.ListManifests())

        Dim systemPrompt As String = SharedLibrary.Agents.ToolingBootstrapPreflight.BuildSystemPrompt(manifests)
        Dim userPrompt As String = SharedLibrary.Agents.ToolingBootstrapPreflight.BuildUserPrompt(
            context.LatestUserRequestRaw,
            context.HostTaskSummary)

        context.Log("Bootstrap preflight started: response language, memory grounding, capability routing, and first capability load.")
        ToolingFileLogger.LogStep("[PERF] Bootstrap preflight LLM request started.")
        LogLatestUserRequestDiagnostic(context, "bootstrap")

        Dim sw As System.Diagnostics.Stopwatch = System.Diagnostics.Stopwatch.StartNew()
        Dim raw As String = ""

        Try
            raw = Await LLM(
                systemPrompt,
                userPrompt,
                "", "", 0,
                useSecondAPI,
                hideSplash,
                "",
                "",
                ToolExecution:=False,
                cancellationToken:=cancellationToken,
                EnsureUI:=True)
        Catch ex As System.TimeoutException
            sw.Stop()
            ToolingFileLogger.LogStep($"[PERF] Bootstrap preflight timed out: elapsedMs={sw.ElapsedMilliseconds}.")
            context.LogWarn("Bootstrap preflight timed out while contacting the AI model.", details:=$"host={context.HostKind}; error={ex.Message}")
            Throw
        Catch ex As System.OperationCanceledException
            sw.Stop()
            ToolingFileLogger.LogStep($"[PERF] Bootstrap preflight cancelled: elapsedMs={sw.ElapsedMilliseconds}.")
            Throw
        Catch ex As System.Exception
            sw.Stop()
            ToolingFileLogger.LogStep($"[PERF] Bootstrap preflight failed: elapsedMs={sw.ElapsedMilliseconds}.")
            context.LogWarn("Bootstrap preflight failed; falling back to the existing per-decision paths.",
                            details:=$"host={context.HostKind}; error={ex.Message}")
            Return New SharedLibrary.Agents.ToolingBootstrapPreflight.Decision()
        End Try

        sw.Stop()
        ToolingFileLogger.LogStep($"[PERF] Bootstrap preflight LLM completed: elapsedMs={sw.ElapsedMilliseconds}; responseChars={If(raw, "").Length}.")
        context.Log($"Bootstrap preflight model step completed in {sw.ElapsedMilliseconds} ms.")

        Dim decision As SharedLibrary.Agents.ToolingBootstrapPreflight.Decision =
            SharedLibrary.Agents.ToolingBootstrapPreflight.ParseDecision(raw)

        Dim normalizedForLog As String = If(decision.NormalizedOutput, "")
        If normalizedForLog.Length > 1200 Then normalizedForLog = normalizedForLog.Substring(0, 1200) & "..."
        context.Log("bootstrapNormalizedOutput=" & normalizedForLog, "diag")
        If Not String.IsNullOrWhiteSpace(decision.ParseError) Then
            context.Log("bootstrapParseError=" & decision.ParseError, "diag")
        End If

        If Not String.IsNullOrWhiteSpace(explicitUserLanguage) Then
            context.SequencingState.UserLanguage = explicitUserLanguage.Trim()
            decision.LanguageApplied = True
            context.Log("Bootstrap language skipped because an explicit user language was supplied: " & context.SequencingState.UserLanguage, "diag")
        ElseIf decision.LanguageValid Then
            context.SequencingState.UserLanguage = decision.Language
            decision.LanguageApplied = True
            context.Log("Bootstrap response language applied: " & decision.Language, "diag")
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

            decision.MemoryApplied = True
            context.Log("Bootstrap memory classification skipped because an explicit memory mode was supplied; explicit mode applied.", "diag")
        ElseIf Not HasMemoryGroundingClassifierInputsAvailable(context) Then
            context.SequencingState.MemoryGroundingMode = SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingMode.None
            context.SequencingState.MemoryGroundingAuthority = SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingAuthority.None
            context.SequencingState.ShouldExposeRecentMemoryStubs = False
            decision.MemoryApplied = True
            context.Log("Bootstrap memory classification not required: no memory tools and no workflow memory available.", "diag")
        ElseIf decision.MemoryValid Then
            ApplyBootstrapMemoryGroundingDecision(context, decision.MemoryDecision)
            decision.MemoryApplied = True
        End If

        If context.CapabilityRoutingRequired AndAlso decision.RoutingSyntaxValid Then
            If decision.RouteKind = SharedLibrary.Agents.CapabilityRoutingTool.KindNone Then
                context.CapabilityRoutingResolved = True
                context.CapabilityRoutingKind = SharedLibrary.Agents.CapabilityRoutingTool.KindNone
                context.CapabilityRoutingName = ""
                context.CapabilityRoutingEntered = True
                decision.RouteApplied = True
                context.Log("Bootstrap capability route resolved: kind=none; reason=" & decision.RouteReason, "diag")
                context.Log("Bootstrap routing complete: no specific skill or top-level agent selected; normal tool planning will continue.")
                ToolingFileLogger.LogStep("[ROUTE] bootstrap resolved kind=none; host=" & context.HostKind & "; reason=" & decision.RouteReason)
            Else
                Dim matchingManifest As SharedLibrary.Agents.ToolManifest =
                    context.AllowedToolRegistry.ListManifests().FirstOrDefault(
                        Function(m)
                            Return m IsNot Nothing AndAlso
                                   Not String.IsNullOrWhiteSpace(m.Name) AndAlso
                                   m.Name.Equals(decision.RouteName, StringComparison.OrdinalIgnoreCase) AndAlso
                                   String.Equals(m.Category, decision.RouteKind, StringComparison.OrdinalIgnoreCase)
                        End Function)

                If matchingManifest IsNot Nothing Then
                    Dim loaded As ModelConfig = EnsureVisibleToolLoaded(matchingManifest.Name, context)
                    If loaded IsNot Nothing Then
                        context.CapabilityRoutingResolved = True
                        context.CapabilityRoutingKind = decision.RouteKind
                        context.CapabilityRoutingName = matchingManifest.Name
                        context.CapabilityRoutingEntered = False
                        decision.RouteApplied = True
                        context.PendingContinuationGuardPrompt = BuildCapabilityRoutingGuardPrompt(context)
                        context.PendingGuardTitle = "HOST BOOTSTRAP CAPABILITY ROUTE"
                        context.PendingRejectedTurnExplanation = "The host bootstrap already resolved the capability route; invoke the selected capability first."
                        context.PendingRejectedAssistantTurn = ""
                        context.Log("Bootstrap capability route resolved: kind=" & decision.RouteKind & "; name=" & matchingManifest.Name & "; reason=" & decision.RouteReason, "diag")
                        context.Log("Bootstrap selected workflow capability: " & matchingManifest.Name)
                        ToolingFileLogger.LogStep("[ROUTE] bootstrap resolved kind=" & decision.RouteKind & "; name=" & matchingManifest.Name & "; entered=false; host=" & context.HostKind & "; reason=" & decision.RouteReason)
                        ToolingFileLogger.LogStep("[PERF] Bootstrap first capability loaded: " & matchingManifest.Name)
                    End If
                End If
            End If
        End If

        If decision.RouteApplied AndAlso context.SelectedTools IsNot Nothing Then
            context.SelectedTools.RemoveAll(
                Function(tool)
                    Return tool IsNot Nothing AndAlso
                           SharedLibrary.Agents.CapabilityRoutingTool.IsResolverToolName(tool.ToolName)
                End Function)
            context.Log("Bootstrap capability router removed from the first normal turn because routing is already resolved.", "diag")
        End If

        If context.CapabilityRoutingRequired AndAlso Not decision.RouteApplied Then
            context.LogWarn(
                "Bootstrap capability route was not applied; the existing capability router will resolve it in the first normal model turn.",
                details:=$"host={context.HostKind}; syntaxValid={decision.RoutingSyntaxValid}; routeKind={decision.RouteKind}; routeName={decision.RouteName}; parseError={decision.ParseError}",
                visibleToUser:=False)
        End If

        context.Log("Bootstrap preflight completed: languageApplied=" & decision.LanguageApplied.ToString().ToLowerInvariant() &
                    "; memoryApplied=" & decision.MemoryApplied.ToString().ToLowerInvariant() &
                    "; routeApplied=" & decision.RouteApplied.ToString().ToLowerInvariant() &
                    "; route=" & If(context.CapabilityRoutingName, ""), "diag")

        Return decision
    End Function


End Class
