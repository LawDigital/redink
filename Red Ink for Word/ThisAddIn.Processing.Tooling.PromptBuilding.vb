' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.Processing.Tooling.PromptBuilding.vb
' Purpose: Prompt construction, user request metadata, and continuation guards.
'
' Responsibilities:
'  - Build tool selection hint text from user/context sources.
'  - Extract latest user turn from dialog structures.
'  - Resolve authoritative user request text (fullPromptOverride precedence).
'  - Build latest-user-request metadata blocks with context enforcement.
'  - Construct completed-facts summaries for model replay.
'  - Build continuation/recovery guard prompts (empty response, repeated tools, missing deliverables).
'  - Classify deliverable requirements (document creation, exports, outputs).
'  - Build final-turn evaluation context (remaining untried tools, pending work).
'  - Append runtime context and workflow continuity information.
'  - Support language contract enforcement in system prompts.
'
' External Dependencies:
'  - SharedLibrary.Agents.ToolCallSequencing for sequencing state and state summary.
'  - SharedLibrary.Agents.WorkflowContinuity for runtime context blocks.
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

    Private Function BuildPromptDiagnosticStub(text As String,
                                              Optional maxExcerptChars As Integer = 120) As String
        Dim raw As String = If(text, "")
        Dim excerpt As String = Regex.Replace(raw, "\s+", " ").Trim()

        If excerpt.Length > maxExcerptChars Then
            excerpt = excerpt.Substring(0, maxExcerptChars) & "..."
        End If

        Dim hashText As String = ""

        Using sha = System.Security.Cryptography.SHA256.Create()
            Dim bytes = sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(raw))
            hashText = BitConverter.ToString(bytes).Replace("-", "").ToLowerInvariant()

            If hashText.Length > 16 Then
                hashText = hashText.Substring(0, 16)
            End If
        End Using

        Return $"len={raw.Length}; sha256={hashText}; excerpt=""{excerpt}"""
    End Function

    Private Sub LogLatestUserRequestDiagnostic(context As ToolExecutionContext, stage As String)
        If context Is Nothing Then
            Return
        End If

        context.Log(
            "latestUserRequestRaw[" & If(stage, "") & "] " &
            BuildPromptDiagnosticStub(context.LatestUserRequestRaw), "diag")
    End Sub

    Private Function BuildLatestUserRequestMetadataBlock(context As ToolExecutionContext) As String
        If context Is Nothing OrElse String.IsNullOrWhiteSpace(context.LatestUserRequestRaw) Then
            Return ""
        End If

        Dim sb As New System.Text.StringBuilder()

        sb.AppendLine("[CURRENT_USER_REQUEST]")
        sb.AppendLine("LATEST_USER_REQUEST_RAW is authoritative for this run.")
        sb.AppendLine("Do not replace, narrow, or reinterpret it based on prior context, memory stubs, workflow summaries, host status text, or completed subtasks.")
        sb.AppendLine("<LATEST_USER_REQUEST_RAW>")
        sb.AppendLine(context.LatestUserRequestRaw)
        sb.AppendLine("</LATEST_USER_REQUEST_RAW>")

        If Not String.IsNullOrWhiteSpace(context.HostTaskSummary) Then
            sb.AppendLine("<HOST_TASK_SUMMARY>")
            sb.AppendLine(context.HostTaskSummary)
            sb.AppendLine("</HOST_TASK_SUMMARY>")
        End If

        sb.AppendLine("[/CURRENT_USER_REQUEST]")
        Return sb.ToString().TrimEnd()
    End Function

    Private Function BuildPromptWithAuthoritativeLatestUserRequest(context As ToolExecutionContext,
                                                                   promptBody As String) As String
        Dim requestBlock As String = BuildLatestUserRequestMetadataBlock(context)

        If String.IsNullOrWhiteSpace(requestBlock) Then
            Return If(promptBody, "")
        End If

        If String.IsNullOrWhiteSpace(promptBody) Then
            Return requestBlock
        End If

        Return requestBlock & Environment.NewLine & Environment.NewLine & promptBody
    End Function

    Private Function BuildCompletedFactsPromptBlock(context As ToolExecutionContext,
                                                    Optional maxItems As Integer = 3) As String
        If context Is Nothing OrElse context.AllToolResponses Is Nothing OrElse context.AllToolResponses.Count = 0 Then
            Return ""
        End If

        Dim facts As New List(Of String)()

        For Each resp In context.AllToolResponses
            If resp Is Nothing OrElse Not resp.Success Then Continue For

            Dim summary As String = Regex.Replace(If(BuildToolReplaySummary(resp), ""), "\s+", " ").Trim()
            If summary = "" Then Continue For

            facts.Add("- " & summary)

            If facts.Count >= maxItems Then
                Exit For
            End If
        Next

        If facts.Count = 0 Then
            Return ""
        End If

        Dim sb As New System.Text.StringBuilder()
        sb.AppendLine("<COMPLETED_FACTS>")
        For Each fact In facts
            sb.AppendLine(fact)
        Next
        sb.AppendLine("</COMPLETED_FACTS>")
        Return sb.ToString().TrimEnd()
    End Function

    Private Function BuildPostToolContinuationBlock(context As ToolExecutionContext) As String
        Return BuildDeliverableCompletionContinuationBlock(context)
    End Function



    Private Function ExtractLatestUserTurnFromDialog(promptText As String) As String
        Dim raw As String = If(promptText, "").Trim()
        If raw = "" Then
            Return ""
        End If

        If raw.IndexOf("<DIALOG>", StringComparison.OrdinalIgnoreCase) < 0 OrElse
           raw.IndexOf("[USER]", StringComparison.OrdinalIgnoreCase) < 0 Then
            Return raw
        End If

        Dim matches As MatchCollection =
            Regex.Matches(
                raw,
                "\[USER\]\s*(?<body>.*?)(?=\s*\[(?:USER|ASSISTANT|SYSTEM|TOOL)\]|\s*</DIALOG>|\z)",
                RegexOptions.IgnoreCase Or RegexOptions.Singleline Or RegexOptions.CultureInvariant)

        If matches.Count = 0 Then
            Return raw
        End If

        Dim latest As String = matches(matches.Count - 1).Groups("body").Value.Trim()
        latest = Regex.Replace(latest, "\s*</DIALOG>\s*$", "", RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant).Trim()

        If latest = "" Then
            Return raw
        End If

        Return latest
    End Function

    Private Function ResolveLatestUserRequestRaw(userText As String,
                                                 otherPrompt As String,
                                                 fullPromptOverride As String) As String
        If Not String.IsNullOrWhiteSpace(fullPromptOverride) Then
            Return ExtractLatestUserTurnFromDialog(fullPromptOverride)
        End If

        If Not String.IsNullOrWhiteSpace(userText) Then
            Return ExtractLatestUserTurnFromDialog(userText)
        End If

        If Not String.IsNullOrWhiteSpace(otherPrompt) Then
            Return ExtractLatestUserTurnFromDialog(otherPrompt)
        End If

        Return ""
    End Function

    Private Function RequestExplicitlyRequiresCreatedDeliverable(latestUserRequestRaw As String) As Boolean
        Dim raw As String = If(latestUserRequestRaw, "").Trim()
        If raw = "" Then
            Return False
        End If

        Dim normalized As String = raw.ToLowerInvariant()

        If Regex.IsMatch(
            normalized,
            "\b(create_word_document|create_pdf|create_excel_spreadsheet|create_powerpoint|create_code_file|word_to_pdf|word_save_as|workspace_write|text_write|word_write)\b",
            RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant) Then
            Return True
        End If

        Dim strongArtifactNounPattern As String =
            "(?:\b(?:pdf|docx|xlsx|pptx|spreadsheet|workbook|presentation|powerpoint|file|datei|attachment|anhang|workspace|downloadable|ausgabedatei)\b|word[- ]?(?:document|dokument)|output\s+file)"

        Dim strongArtifactVerbPattern As String =
            "\b(?:create|generate|produce|make|write|save|export|attach|store|persist|download|erstellen|generieren|erzeugen|machen|schreiben|speichern|exportieren|anhängen|ablegen|herunterladen)\b"

        If Regex.IsMatch(
            normalized,
            strongArtifactVerbPattern & "[\s\S]{0,80}" & strongArtifactNounPattern,
            RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant) Then
            Return True
        End If

        If Regex.IsMatch(
            normalized,
            strongArtifactNounPattern & "[\s\S]{0,80}" & strongArtifactVerbPattern,
            RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant) Then
            Return True
        End If

        Dim documentArtifactVerbPattern As String =
            "\b(?:create|generate|save|export|attach|store|persist|erstellen|generieren|speichern|exportieren|anhängen|ablegen)\b"

        Dim documentArtifactNounPattern As String =
            "\b(?:document|dokument)\b"

        If Regex.IsMatch(
            normalized,
            documentArtifactVerbPattern & "[\s\S]{0,80}" & documentArtifactNounPattern,
            RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant) Then
            Return True
        End If

        If Regex.IsMatch(
            normalized,
            documentArtifactNounPattern & "[\s\S]{0,80}" & documentArtifactVerbPattern,
            RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant) Then
            Return True
        End If

        If Regex.IsMatch(
            normalized,
            "\b(?:as|als)\s+(?:a\s+|an\s+|eine\s+|einen\s+)?(?:pdf|docx|xlsx|pptx|word[- ]?(?:document|dokument)|spreadsheet|presentation|powerpoint|file|datei)\b",
            RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant) Then
            Return True
        End If

        Return False
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

        Dim routingManifests As List(Of SharedLibrary.Agents.ToolManifest) = Nothing
        If allowedRegistry IsNot Nothing Then
            routingManifests = allowedRegistry.ListManifests().
                Where(Function(m) m IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(m.Name)).
                ToList()
            Dim capabilityRouter As ModelConfig = SharedLibrary.Agents.CapabilityRoutingTool.Build(routingManifests)
            If capabilityRouter IsNot Nothing Then
                result.Add(capabilityRouter)
            End If
        End If

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


    Private Function BuildToolWorkflowInstructionAddendum(selectedTools As List(Of ModelConfig)) As String
        Dim sb As New StringBuilder()

        sb.AppendLine("PERSISTENCE CHECKLIST:")
        sb.AppendLine("- Remain in tool-calling mode until the whole user request is completed. Do not stop after planning, discovering files, staging files, or finishing only the first subtask.")
        sb.AppendLine("- If one tool fails, returns too little information, or only partially advances the task, and another available tool could still help, call the next suitable tool instead of giving up.")
        sb.AppendLine("- If the request applies to a folder, directory, workspace path, or a collection of files, discover or stage the collection first and then continue processing the returned items until the collection has actually been searched or analyzed.")
        sb.AppendLine("- Before giving a final answer, explicitly check whether any requested next step, remaining file, or reasonable fallback tool is still outstanding.")

        If HasToolName(selectedTools, "extract_pdf_text") Then
            sb.AppendLine("- extract_pdf_text is for a single PDF or staged/session file at a time. Never pass a directory or folder path to extract_pdf_text.")
        End If

        If HasToolName(selectedTools, "agent_workspace_find_files") OrElse
           HasToolName(selectedTools, "agent_workspace_stage") OrElse
           HasToolName(selectedTools, "agent_workspace_read") OrElse
           HasToolName(selectedTools, "workspace_inventory") OrElse
           HasToolName(selectedTools, "workspace_read") Then
            sb.AppendLine("- For local/workspace PDF collections, prefer the workspace workflow: find files, stage them if required, then read/search/extract them. Do not stop after file discovery.")
        End If

        If HasToolName(selectedTools, "agent_workspace_read") Then
            sb.AppendLine("- For one workspace-local PDF or Office file, prefer agent_workspace_read over calling extract_pdf_text directly on a workspace path.")
        End If

        If HasToolName(selectedTools, "search_in_attachments") Then
            sb.AppendLine("- When many staged PDFs must be searched for a term, prefer search_in_attachments across the staged set before falling back to repeated one-file extraction calls.")
        End If

        Return sb.ToString().Trim()
    End Function


End Class
