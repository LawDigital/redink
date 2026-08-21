' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ToolingOrchestrator.vb
' Purpose: Host-agnostic decision helpers used by Outlook, Word and Excel
'          tooling loops. Concentrates rules for (a) accepting/rejecting final-turn
'          candidates, (b) language-contract injection, and (c) the canonical
'          TASK_STATUS footer instruction text.
'
' Final Turn Decisions:
'  - Accept: candidate is valid, return to user.
'  - RejectMalformedFooter: no/invalid TASK_STATUS, force strict-format guard.
'  - RejectActionPromiseWithoutInvocation: prose announced future work not done.
'  - RejectDeliverableFallbackRequired: request requires deliverable, untried fallbacks exist.
'  - Hosts call EvaluateFinalTurn(...) from existing loops; minimal integration.
' =============================================================================

Option Explicit On
Option Strict On

Imports System.Text

Namespace Agents

    Public Enum FinalTurnDecision
        ''' <summary>The final-turn candidate is acceptable and may be returned to the user.</summary>
        Accept
        ''' <summary>The candidate is malformed (no/invalid TASK_STATUS). Force another iteration with a strict-format guard.</summary>
        RejectMalformedFooter
        ''' <summary>The prose announced a future action that was not invoked. Force one more iteration with a promise-recovery guard.</summary>
        RejectActionPromiseWithoutInvocation
        ''' <summary>The request requires a created deliverable, none was produced, and at least one untried fallback tool is loaded. Force one more iteration with a fallback guard.</summary>
        RejectDeliverableFallbackRequired
    End Enum

    Public Class FinalTurnEvaluation
        Public Property Decision As FinalTurnDecision
        Public Property GuardPrompt As String
        Public Property GuardTitle As String
        Public Property Reason As String
    End Class

    Public Module ToolingOrchestrator

        ' ---------------------------------------------------------------------
        ' TASK_STATUS footer contract text (replaces the per-host const).
        ' This is the SAME wording as before with one addition: an explicit rule
        ' that a 'blocked' final must not announce future work and must follow
        ' at least one fallback attempt when the user authorized one.
        ' ---------------------------------------------------------------------
        Public ReadOnly TaskStatusFooterInstruction As String =
    "TASK STATUS FOOTER (MANDATORY CONTRACT, MACHINE-READ): " &
    "Whenever you produce a final prose response instead of invoking a tool, you MUST append, " &
    "as the literal last line of that turn, exactly: " &
    "<TASK_STATUS>{""status"":""<value>"",""reason"":""<short>""}</TASK_STATUS>  " &
    "Allowed values for <value>: " &
    "  'complete' = the user's entire request has been FULLY satisfied in THIS turn. " &
    "  'blocked'  = the task cannot be completed despite reasonable tool attempts. " &
    "Rules: " &
    "(1) NEVER include the footer in a turn that contains a tool call. " &
    "(2) NEVER wrap the footer in code fences or quotes; it must be plain text on its own final line. " &
    "(3) NEVER claim 'complete' while still announcing future work. " &
    "(4) During active tooling, final prose MUST end with exactly one valid TASK_STATUS footer whose status is either 'complete' or 'blocked'. " &
    "(5) If the user's request covers multiple items, you may emit 'complete' only after all required items have actually been processed via tool calls and the full final result is ready. " &
    "(6) If the task is not yet complete and more tool work is required or possible, emit the next required tool call instead of final prose. " &
    "(7) If required Memory grounding is active and the final answer relies only on a retrieved subset of listed Memory entries, include ""memoryGroundingScope"":""subset"" inside the TASK_STATUS JSON footer. " &
    "(8) NEVER pair 'blocked' with prose that announces a future fallback action (""I will try to..."", ""instead I'll...""). Either invoke the fallback now via a tool call, or declare blocked WITHOUT announcing further work. " &
    "(9) If the user explicitly authorized a fallback (e.g. ""create a Word file, otherwise a plain text file""), 'blocked' is only valid AFTER all authorized fallbacks have actually been attempted via tool calls. " &
    "(10) The ""reason"" value MUST be a single very short plain phrase, ideally 2-6 words, with no line breaks, and no more than " &
    ToolCallSequencing.TaskStatusReasonMaxChars.ToString() &
    " characters. Good examples: ""answer ready"", ""no safe completion path"", ""all requested edits applied""."


        ''' <summary>
        ''' Central decision function for the final-turn acceptance gate. Hosts call this
        ''' from their existing loop right where they previously called HasValidTerminalTaskStatus.
        ''' If the returned Decision is anything other than Accept, the host must perform
        ''' one more iteration with the supplied GuardPrompt (and clear it after the iteration).
        ''' </summary>
        ''' <param name="finalText">The raw model output for the candidate final turn.</param>
        ''' <param name="requestRequiresCreatedDeliverable">True if the user request requires a created artifact.</param>
        ''' <param name="hasProducedUserDeliverable">True if a deliverable has actually been produced via tools in this run.</param>
        ''' <param name="availableUntriedFallbackToolNames">Tool names that can plausibly satisfy the deliverable and have not yet been invoked in this run.</param>
        Public Function EvaluateFinalTurn(finalText As String,
                                          requestRequiresCreatedDeliverable As Boolean,
                                          hasProducedUserDeliverable As Boolean,
                                          availableUntriedFallbackToolNames As IReadOnlyList(Of String),
                                          Optional hasValidatedFinalDeliverable As Boolean = True) As FinalTurnEvaluation

            Dim result As New FinalTurnEvaluation()

            Dim parsed As TaskStatusFooter = TaskStatusFooterParser.Parse(finalText)

            ' (a) Strict footer check.
            If parsed.Kind = TaskStatusKind.Missing OrElse parsed.Kind = TaskStatusKind.Invalid OrElse parsed.Kind = TaskStatusKind.ContinueWork Then
                result.Decision = FinalTurnDecision.RejectMalformedFooter
                result.GuardTitle = "task_status_footer_malformed"
                result.Reason = "missing_or_invalid_task_status_footer:" & parsed.Kind.ToString() &
                                If(String.IsNullOrEmpty(parsed.InvalidDetail), "", ":" & parsed.InvalidDetail)
                result.GuardPrompt = BuildMalformedFooterGuardPrompt(parsed)
                Return result
            End If

            ' parsed.Kind is now either Complete or Blocked.
            Dim prose As String = TaskStatusFooterParser.ExtractProse(finalText)

            ' (b) Action-promise gate. Applies to BOTH complete and blocked finals:
            '     - 'complete' must not announce future work (contract rule 3).
            '     - 'blocked'  must not pair with a fallback promise   (contract rule 8).
            If ActionPromiseValidator.ContainsActionPromise(prose) Then
                result.Decision = FinalTurnDecision.RejectActionPromiseWithoutInvocation
                result.GuardTitle = "final_prose_announced_unperformed_action"
                result.Reason = "action_promise_without_tool_invocation:" & parsed.Kind.ToString()
                result.GuardPrompt = ActionPromiseValidator.BuildPromiseRecoveryGuardPrompt()
                Return result
            End If

            ' (c) Deliverable-fallback gate. Only applies to BLOCKED finals:
            '     if the request requires a created deliverable, none was produced,
            '     and at least one plausible untried fallback tool is still loaded.
            If parsed.Kind = TaskStatusKind.Blocked _
               AndAlso requestRequiresCreatedDeliverable _
               AndAlso Not hasProducedUserDeliverable _
               AndAlso availableUntriedFallbackToolNames IsNot Nothing _
               AndAlso availableUntriedFallbackToolNames.Count > 0 Then

                result.Decision = FinalTurnDecision.RejectDeliverableFallbackRequired
                result.GuardTitle = "blocked_before_attempting_authorized_fallback"
                result.Reason = "blocked_without_attempting_available_fallback_tools"
                result.GuardPrompt = BuildFallbackRequiredGuardPrompt(availableUntriedFallbackToolNames)
                Return result
            End If

            ' (d) Deliverable-existence gate. Applies to COMPLETE finals:
            '     if the request requires a created deliverable, the host must possess a
            '     validated (on-disk) final artifact. A 'complete' claim without a real,
            '     existing file is rejected and repaired. This is bounded by the host's
            '     continuation-retry budget, so a genuinely un-registerable deliverable
            '     still finalizes after the budget is exhausted (no hard deadlock).
            If parsed.Kind = TaskStatusKind.Complete _
               AndAlso requestRequiresCreatedDeliverable _
               AndAlso Not hasValidatedFinalDeliverable Then

                result.Decision = FinalTurnDecision.RejectDeliverableFallbackRequired
                result.GuardTitle = "complete_without_validated_deliverable"
                result.Reason = "complete_claimed_without_validated_file_deliverable"
                result.GuardPrompt = BuildMissingDeliverableGuardPrompt(availableUntriedFallbackToolNames)
                Return result
            End If

            ' All gates passed.
            result.Decision = FinalTurnDecision.Accept
            Return result
        End Function

        Private Function BuildMissingDeliverableGuardPrompt(untriedFallbackToolNames As IReadOnlyList(Of String)) As String
            Dim sb As New StringBuilder()
            sb.AppendLine("HOST DELIVERABLE GUARD: You reported the task as 'complete', but the user's request requires a created file and NO validated output file exists yet.")
            sb.AppendLine("A path mentioned in prose or tool metadata does NOT count. The file must actually be written to the working location by a tool.")
            sb.AppendLine("Before you may report 'complete', invoke a tool that actually creates or saves the requested file (for example a write/export/save-as tool).")
            If untriedFallbackToolNames IsNot Nothing AndAlso untriedFallbackToolNames.Count > 0 Then
                sb.AppendLine("Candidate tools that can produce the deliverable and have not yet been used this run:")
                For Each toolName As String In untriedFallbackToolNames
                    If String.IsNullOrWhiteSpace(toolName) Then Continue For
                    sb.AppendLine("  - " & toolName.Trim())
                Next
            End If
            sb.AppendLine("If the deliverable genuinely cannot be created, declare 'blocked' with a short reason instead of 'complete', and do not announce future work.")
            Return sb.ToString().TrimEnd()
        End Function

        Private Function BuildMalformedFooterGuardPrompt(parsed As TaskStatusFooter) As String
            Dim sb As New StringBuilder()
            sb.AppendLine("HOST FOOTER GUARD: Your previous final prose response did not end with a valid <TASK_STATUS> footer.")
            If parsed IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(parsed.InvalidDetail) Then
                sb.AppendLine("Detected issue: " & parsed.InvalidDetail)
            End If
            sb.AppendLine("Repeat your final response now and end with EXACTLY ONE line of the form:")
            sb.AppendLine("<TASK_STATUS>{""status"":""complete|blocked"",""reason"":""short phrase""}</TASK_STATUS>")
            sb.AppendLine("The reason must be a single very short plain phrase, ideally 2-6 words, with no line breaks, and no more than " &
                  ToolCallSequencing.TaskStatusReasonMaxChars.ToString() & " characters.")
            sb.AppendLine("Valid examples:")
            sb.AppendLine("<TASK_STATUS>{""status"":""complete"",""reason"":""answer ready""}</TASK_STATUS>")
            sb.AppendLine("<TASK_STATUS>{""status"":""blocked"",""reason"":""no safe completion path""}</TASK_STATUS>")
            sb.AppendLine("Do not wrap the footer in quotes or code fences. Do not emit a footer if you are calling a tool.")
            Return sb.ToString().TrimEnd()
        End Function

        Private Function BuildFallbackRequiredGuardPrompt(untriedFallbackToolNames As IReadOnlyList(Of String)) As String
            Dim sb As New StringBuilder()
            sb.AppendLine("HOST FALLBACK GUARD: You declared 'blocked' but the user's request requires a created deliverable that has not yet been produced, and at least one authorized fallback tool has not yet been attempted.")
            sb.AppendLine("Before declaring 'blocked' you MUST attempt at least one of the following fallback tools via a tool call:")

            Dim cap As Integer = ToolingConstants.MaxFallbackToolsListedInGuard
            Dim count As Integer = 0
            For Each name As String In untriedFallbackToolNames
                If String.IsNullOrWhiteSpace(name) Then Continue For
                sb.AppendLine("  - " & name)
                count += 1
                If count >= cap Then Exit For
            Next

            sb.AppendLine("Call the most appropriate fallback now. If the fallback also fails with a structured error, THEN you may declare 'blocked' — and at that point your prose MUST NOT promise any further attempts.")
            Return sb.ToString().TrimEnd()
        End Function

        ''' <summary>
        ''' Convenience: builds the per-iteration system-prompt language-contract block.
        ''' Hosts append the returned text to their assembled system prompt for every iteration.
        ''' </summary>
        Public Function BuildLanguageContractSystemPromptFragment(userLanguage As String) As String
            Dim baseFragment As String = LanguageContract.BuildSystemPromptFragment(userLanguage)
            Dim designFragment As String =
                "DESIGN PROVENANCE RULE: Never claim that you know, reproduced, matched, followed, or used a named organization's corporate design from your own knowledge or visual familiarity. " &
                "You may claim organization-specific design only when a concrete design source is available in the current authorized context, such as a configured design_name from the central/local design repository, an attached/template Office file, an approved skill/reference template, or explicit user-provided design specifications (fonts, colors, layout rules, logos/assets). " &
                "If the user asks for a named organization's design but no such source is available, say briefly that the specific design is not available and create a neutral professional design instead unless the user asks you to stop. " &
                "Do not infer a corporate design merely from the organization name, prior general knowledge, or a logo alone. When a concrete source is used, describe it accurately as the source (for example: 'using the configured VISCHER design profile' or 'using the attached template'), not as independent knowledge of the brand."

            Dim webToolRoutingFragment As String =
                "WEB TOOL ROUTING RULE: Use web_grounding when relevant public pages first need to be discovered. Use retrieve_web_content for a known mostly-static URL when readable text/ordinary links are sufficient. Do not expose/load browser tools by default. For a task that needs browser behavior, load them lazily via tool_loader. Then prefer browser_open + browser_snapshot for exploration of a specific website, finding links/pages/downloads within that site, menus/navigation, pagination, or JavaScript/dynamically rendered content. Use browser_interact only when an actual browser action is needed, and take a fresh browser_snapshot after every interaction. If retrieve_web_content returns navigation failure, empty links, incomplete client-side content, or otherwise cannot expose the requested site structure, load browser_open, browser_snapshot and browser_interact via tool_loader and switch to browser_open + browser_snapshot before doing another broad web_grounding search. If the current browser snapshot already exposes the needed link/control, continue via the browser rather than restarting web_grounding."

            Dim repositoryFragment As String = DesignRepository.BuildPromptFragment()
            Dim combinedDesignFragment As String = designFragment & System.Environment.NewLine & webToolRoutingFragment
            If Not System.String.IsNullOrWhiteSpace(repositoryFragment) Then
                combinedDesignFragment &= System.Environment.NewLine & repositoryFragment.Trim()
            End If

            If System.String.IsNullOrWhiteSpace(baseFragment) Then Return combinedDesignFragment
            Return baseFragment.TrimEnd() & System.Environment.NewLine & System.Environment.NewLine & combinedDesignFragment
        End Function


        ''' <summary>
        ''' Convenience: post-egress decision for final-response localization when the
        ''' final prose does not match the detected user language.
        ''' </summary>
        Public Function ShouldPostLocalizeFinal(prose As String,
                                                userLanguage As String,
                                                finalStatus As String) As Boolean
            Return LanguageContract.ShouldPostLocalizeFinal(
                prose,
                userLanguage,
                finalStatus,
                ToolingConstants.MaxLocalizableBlockedFinalChars)
        End Function
    End Module

End Namespace
