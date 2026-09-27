' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: WebGroundingTool.vb
' Purpose: Shared host-internal web_grounding tool for Outlook and Word.
'          Exposes one tool definition that covers both standard web grounding
'          and deep research, depending on configured special-task models.
'          Executes using the same prompts and JSON-mode behavior as the
'          original Outlook AutoPilot implementation, while allowing host-
'          specific logging callbacks (e.g. AutoPilot dashboard).
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Threading
Imports System.Threading.Tasks

Namespace Agents

    Public NotInheritable Class WebGroundingTool

        Private Sub New()
        End Sub

        Public Const ToolName As String = "web_grounding"

        Public Shared Function IsWebGroundingTool(name As String) As Boolean
            Return Not String.IsNullOrWhiteSpace(name) AndAlso
                   name.Trim().Equals(ToolName, StringComparison.OrdinalIgnoreCase)
        End Function

        Public Shared Function Build(context As SharedLibrary.SharedContext.ISharedContext,
                                     Optional enforcePrivacy As Boolean = True,
                                     Optional toolPriority As Integer = 997,
                                     Optional displaySuffix As String = "") As SharedLibrary.ModelConfig

            If context Is Nothing Then
                Return Nothing
            End If

            Dim webGroundingAvailable As Boolean = SharedLibrary.SharedMethods.IsWebGroundingAvailable(context)
            Dim deepResearchAvailable As Boolean = SharedLibrary.SharedMethods.IsDeepResearchAvailable(context)

            If Not webGroundingAvailable AndAlso Not deepResearchAvailable Then
                Return Nothing
            End If

            Dim modeEnum As String
            Dim modeDescription As String
            Dim modeInstructions As String

            If webGroundingAvailable AndAlso deepResearchAvailable Then
                modeEnum =
                    """mode"":{""type"":""string"",""enum"":[""standard"",""deep_research""]," &
                    """description"":""'standard' for quick factual lookups, current events, verifying claims, or finding specific information. " &
                    "'deep_research' for complex multi-faceted questions requiring comprehensive analysis, comparison of multiple sources, " &
                    "or in-depth investigation. deep_research returns an extensive report. Default: 'standard'""},"

                modeDescription =
                    "Supports two modes: 'standard' for quick lookups and 'deep_research' for comprehensive investigation that returns an extensive report."

                modeInstructions =
                    "Choose mode='standard' for quick factual queries and mode='deep_research' for complex, multi-faceted questions requiring thorough analysis from multiple sources. deep_research produces a detailed, structured report with multiple sections."

            ElseIf webGroundingAvailable Then
                modeEnum =
                    """mode"":{""type"":""string"",""enum"":[""standard""]," &
                    """description"":""Only 'standard' mode is available. Used for factual lookups and current information. Default: 'standard'""},"

                modeDescription =
                    "Uses the web grounding model for factual lookups and current information."

                modeInstructions =
                    "Only 'standard' mode is available."

            Else
                modeEnum =
                    """mode"":{""type"":""string"",""enum"":[""deep_research""]," &
                    """description"":""Only 'deep_research' mode is available. Used for comprehensive, multi-source investigation. Returns an extensive report. Default: 'deep_research'""},"

                modeDescription =
                    "Uses the deep research model for comprehensive, multi-source investigation that returns an extensive report."

                modeInstructions =
                    "Only 'deep_research' mode is available. It produces a detailed, structured report with multiple sections."
            End If

            Dim privacyConstraint As String = ""
            Dim privacyConstraintDef As String = ""
            Dim privacyConstraintQueryParam As String = ""
            Dim privacyConstraintContextParam As String = ""

            If enforcePrivacy Then
                privacyConstraint =
                    "PRIVACY CONSTRAINT: The query is sent to an external AI model with web access. " &
                    "You MUST NOT include any personal data, confidential information, private names, " &
                    "case details, contract terms, internal identifiers, email addresses, phone numbers, " &
                    "account numbers, or any other non-public information in the query. " &
                    "Only well-known public figures, public institutions, published legislation, " &
                    "and other clearly public information may appear in the query. "

                privacyConstraintDef =
                    "PRIVACY: The query is sent to an external model. MUST NOT contain personal data, confidential information, or non-public details."

                privacyConstraintQueryParam =
                    "MUST NOT contain personal data, confidential information, or any non-public details."

                privacyConstraintContextParam =
                    " Same privacy constraints apply — no personal or confidential data."
            End If

            Return New SharedLibrary.ModelConfig() With {
                .ToolOnly = True,
                .Tool = True,
                .ToolName = ToolName,
                .CapabilityTags = "source_retrieval",
                .ToolPriority = toolPriority,
                .ToolErrorHandling = "skip",
                .ModelDescription = "Web Grounding / Deep Research" & If(displaySuffix, ""),
                .ToolInstructionsPrompt =
                    ToolName & ": Searches the internet using a web-grounding AI model to retrieve current, factual, " &
                    "or verifiable information. The model ALWAYS searches the internet — it never answers from memory alone. " &
                    "Results ALWAYS include citations with source URLs. " &
                    "Use this when the user's question requires up-to-date data, fact-checking, research, or information you are not confident about. " &
                    modeInstructions & " " &
                    "If both web_grounding and internet_search are available, prefer web_grounding for current factual research, verification, and cited live-web answers. " &
                    "The 'query' parameter should contain a clear, focused search question or topic. " &
                    privacyConstraint &
                    "RESPONSE HANDLING — MANDATORY SOURCE CITATION RULES: " &
                    "The tool returns the raw JSON response from the grounding model. " &
                    "You MUST extract and present the COMPLETE answer text from the response. Do NOT summarize or shorten it. " &
                    "For deep_research results, the response contains an extensive report — include it in FULL. " &
                    "CRITICAL — SOURCES ARE MANDATORY: " &
                    "Your reply MUST contain a clearly labeled 'Sources' or 'References' section at the end. " &
                    "You MUST extract ALL citations, sources, and grounding metadata from the JSON " &
                    "(look for keys such as 'citations', 'citationSources', 'citationMetadata', 'groundingMetadata', " &
                    "'groundingSupports', 'groundingChunks', 'webSearchQueries', 'searchEntryPoint', 'sources', 'uri', 'url', 'title'). " &
                    "Every single source URL found in the response MUST appear as a clickable Markdown link in your reply: [Title](URL) or [URL](URL) if no title. " &
                    "If the response contains inline citation markers (e.g. [1], [2]), keep them AND provide the numbered source list. " &
                    "FAILURE TO INCLUDE SOURCES IS A CRITICAL ERROR. If you cannot find any URLs in the response JSON, explicitly state: 'No source URLs were returned by the search engine.' " &
                    "NEVER omit sources that are present in the response. NEVER invent or fabricate URLs that are not in the response. " &
                    "A reply without a Sources section when the tool returned URLs is INVALID and MUST be corrected.",
                .ToolDefinition =
                    "{""name"":""" & ToolName & """," &
                    """description"":""Searches the internet using a web-grounding AI model. The model ALWAYS performs a live internet search " &
                    "and ALWAYS returns citations with source URLs. Returns the full JSON response including the answer text and all grounding sources/citations. " &
                    modeDescription & " " &
                    "IMPORTANT: When processing the result, you MUST include the complete answer AND all source URLs as clickable links in a dedicated 'Sources' section. " &
                    "A reply without sources when the tool returned URLs is invalid. " &
                    "For deep_research mode, include the full extensive report without summarizing. " &
                    If(privacyConstraintDef <> "", privacyConstraintDef, "") & """," &
                    """parameters"":{""type"":""object"",""properties"":{" &
                    """query"":{""type"":""string"",""description"":""The search question or topic. Must be clear and focused." &
                    If(privacyConstraintQueryParam <> "", " " & privacyConstraintQueryParam, "") & """}," &
                    modeEnum &
                    """context"":{""type"":""string"",""description"":""Optional additional context to help the grounding model understand the query better." &
                    If(privacyConstraintContextParam <> "", privacyConstraintContextParam, "") & """}" &
                    "},""required"":[""query""]}}"
            }
        End Function

        Public Shared Async Function ExecuteAsync(context As SharedLibrary.SharedContext.ISharedContext,
                                                  arguments As IDictionary(Of String, Object),
                                                  cancellationToken As CancellationToken,
                                                  Optional logStep As Action(Of String) = Nothing,
                                                  Optional logInfo As Action(Of String) = Nothing,
                                                  Optional logWarn As Action(Of String) = Nothing) As Task(Of String)

            If context Is Nothing Then
                Throw New InvalidOperationException("Shared context is required for web_grounding.")
            End If

            Dim query = GetArgumentString(arguments, "query")
            If String.IsNullOrWhiteSpace(query) Then
                Throw New InvalidOperationException("Missing required parameter: query")
            End If

            Dim mode = If(GetArgumentString(arguments, "mode"), "standard").Trim().ToLowerInvariant()
            Dim additionalContext = GetArgumentString(arguments, "context")

            Dim taskKey As String
            Dim modeLabel As String

            If mode = "deep_research" Then
                taskKey = "DeepResearch"
                modeLabel = "deep research"
            Else
                taskKey = "WebGrounding"
                modeLabel = "web grounding"
            End If

            Dim queryPreview = If(query.Length > 120, query.Substring(0, 120) & "...", query)
            SafeInvoke(logStep, $"🌐 {If(mode = "deep_research", "Deep research", "Web grounding")}: {queryPreview}")

            If String.IsNullOrWhiteSpace(context.INI_AlternateModelPath) Then
                Throw New InvalidOperationException(
                    "No alternate model configuration file is configured. Web grounding requires a configured model.")
            End If

            Dim backupConfig As SharedLibrary.ModelConfig = SharedLibrary.SharedMethods.GetCurrentConfig(context)
            Dim previousResponse2 As String = context.INI_Response_2
            Dim modelSwitched As Boolean = False

            Try
                Try
                    modelSwitched =
                        SharedLibrary.SharedMethods.GetSpecialTaskModel(
                            context,
                            context.INI_AlternateModelPath,
                            taskKey)
                Catch
                End Try

                If Not modelSwitched Then
                    Dim fallbackKey = If(taskKey = "WebGrounding", "DeepResearch", "WebGrounding")
                    Try
                        modelSwitched =
                            SharedLibrary.SharedMethods.GetSpecialTaskModel(
                                context,
                                context.INI_AlternateModelPath,
                                fallbackKey)

                        If modelSwitched Then
                            modeLabel = If(fallbackKey = "DeepResearch", "deep research", "web grounding")
                            SafeInvoke(logWarn, $"{taskKey} model not available, using {fallbackKey} instead")
                        End If
                    Catch
                    End Try
                End If

                If Not modelSwitched Then
                    Throw New InvalidOperationException(
                        $"No {taskKey} or fallback model is configured in the alternate models file. Cannot perform web grounding.")
                End If

                context.INI_Response_2 = "JSON"

                Dim systemPrompt As String
                If mode = "deep_research" Then
                    systemPrompt =
                        "You are a deep research assistant. You MUST use your internet search and web grounding " &
                        "capabilities to thoroughly research the user's query. ALWAYS search the internet — " &
                        "even if you believe you already know the answer. Your knowledge may be outdated or incomplete. " &
                        "Produce a comprehensive, well-structured research report with the following requirements: " &
                        "1. Search broadly across multiple sources to gather diverse perspectives. " &
                        "2. Organize your findings into clear sections with headings. " &
                        "3. Provide detailed analysis, not just summaries — include key facts, figures, dates, and context. " &
                        "4. Compare and contrast information from different sources where relevant. " &
                        "5. Note any conflicting information or areas of uncertainty. " &
                        "6. ALWAYS cite your sources with complete URLs. Every factual claim must have a citation. " &
                        "7. Include a 'Sources' section at the end listing all referenced URLs. " &
                        "8. Aim for thoroughness — a deep research report should be extensive and detailed, " &
                        "not a brief summary. Cover the topic comprehensively. " &
                        "NEVER provide an answer without searching the internet first. " &
                        "NEVER omit citations. If you cannot find sources, explicitly state that."
                Else
                    systemPrompt =
                        "You are a web research assistant. You MUST use your internet search and web grounding " &
                        "capabilities to answer the user's query. ALWAYS search the internet — even if you believe " &
                        "you already know the answer. Your knowledge may be outdated or incomplete. " &
                        "Requirements: " &
                        "1. ALWAYS search the web before answering. " &
                        "2. Provide a clear, accurate answer based on what you find. " &
                        "3. ALWAYS cite your sources with complete URLs for every factual claim. " &
                        "4. Include a 'Sources' section at the end listing all referenced URLs. " &
                        "NEVER provide an answer without searching the internet first. " &
                        "NEVER omit citations. If you cannot find sources, explicitly state that."
                End If

                Dim userPrompt As String = query
                If Not String.IsNullOrWhiteSpace(additionalContext) Then
                    userPrompt = query & vbCrLf & vbCrLf & "Additional context: " & additionalContext
                End If

                Dim llmResult =
                    Await SharedLibrary.SharedMethods.LLM(
                        context,
                        systemPrompt,
                        userPrompt,
                        UseSecondAPI:=True,
                        Hidesplash:=True,
                        cancellationToken:=cancellationToken).ConfigureAwait(False)

                If String.IsNullOrWhiteSpace(llmResult) Then
                    SafeInvoke(logWarn, "Web grounding: empty response")
                    Throw New InvalidOperationException(
                        $"The {modeLabel} model returned an empty response. The query may have been blocked or the model may be unavailable.")
                End If

                Dim resultLen = llmResult.Length
                SafeInvoke(logInfo, $"Web grounding ({modeLabel}): {resultLen:N0} chars returned")

                Return llmResult

            Finally
                context.INI_Response_2 = previousResponse2
                If backupConfig IsNot Nothing Then
                    SharedLibrary.SharedMethods.RestoreDefaults(context, backupConfig)
                End If
            End Try
        End Function

        Private Shared Sub SafeInvoke(callback As Action(Of String), message As String)
            If callback Is Nothing Then Return
            Try
                callback(message)
            Catch
            End Try
        End Sub

        Private Shared Function GetArgumentString(arguments As IDictionary(Of String, Object),
                                                  key As String) As String

            If arguments Is Nothing OrElse String.IsNullOrWhiteSpace(key) Then
                Return ""
            End If

            If Not arguments.ContainsKey(key) OrElse arguments(key) Is Nothing Then
                Return ""
            End If

            Try
                Dim value As Object = arguments(key)

                If TypeOf value Is Newtonsoft.Json.Linq.JValue Then
                    Dim jv = DirectCast(value, Newtonsoft.Json.Linq.JValue)
                    Return If(jv.Value, "").ToString().Trim()
                End If

                Return If(value, "").ToString().Trim()
            Catch
                Return ""
            End Try
        End Function

    End Class

End Namespace