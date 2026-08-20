' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.SenderToolPolicy.vb
' Purpose: Additive per-sender hard tool policy for Inky AutoPilot.
'          Narrows the already-selected AutoPilot tool set for a specific sender
'          BEFORE the tooling loop runs, fully in code (not under LLM control).
'
' Architecture / How it works:
'  - A plain-text policy file (path configured at startup) maps sender patterns
'    to a rule that limits which tools/skills/agents/sources that sender may use.
'  - The policy is applied as a final, additive filter on top of the existing
'    session tool list. It NEVER touches mail intake filters or the auto-approval
'    whitelist - those layers remain unchanged.
'  - When no policy file is configured (or it is empty/missing), the session tool
'    list is returned unchanged, so current behaviour is fully preserved.
'
' Policy file format (one rule per line):
'     # or ;                  -> comment line (ignored)
'     <pattern> = ALL         -> sender may use every selected tool
'     <pattern> = NONE        -> sender may use nothing (report_inability only)
'     <pattern> = ONLY skill_<name>
'                              -> sender may run that skill plus helpers declared in its allowed-tools
'     <pattern> = <selector>, <selector>, ...
'                              -> combined include/exclude selector list
'     DEFAULT = <rule>        -> fallback for senders that match no pattern
'
'  - A selector is matched purely against the NAME of a tool, skill, agent, or online
'    resource. Selectors may be:
'      - an exact name (e.g. internet_search, skill_intake, agent_research, some_online_source)
'      - a wildcard name pattern using * and ? (e.g. skill_*, swiss-caselaw*, agent_?)
'      - the universal placeholder * (or ALL) matching every name
'  - Prefix any selector with ! or - to EXCLUDE the matching names.
'  - If a selector list contains only exclusions, it behaves like * except those exclusions.
'
'  - Any rule may append a hard, in-code system-prompt instruction for the sender:
'         <pattern> = <rule> || <system prompt instruction>
'    Example:
'         info@lawdigital.com = ONLY skill_produkteanfragen || Only provide an answer
'             during weekdays and process any response through that skill
'    For an ONLY-skill rule without an explicit instruction, a default instruction is
'    generated automatically that confines the sender to that skill and its declared helpers.
'
'  - <pattern> supports * and ? wildcards and is matched against the SMTP address.
'  - First matching line wins (top-to-bottom). DEFAULT is only used if no pattern
'    matched. If there is no match and no DEFAULT, the sender inherits ALL tools.
'  - report_inability is ALWAYS preserved so the model can decline gracefully.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports SharedLibrary.SharedLibrary

Partial Public Class ThisAddIn

    ''' <summary>The kind of restriction a sender policy rule applies.</summary>
    Private Enum AutoPilotSenderPolicyRuleKind
        ''' <summary>Sender may use every tool in the session list (no restriction).</summary>
        AllowAll
        ''' <summary>Sender may use no tools at all (only report_inability remains).</summary>
        BlockAll
        ''' <summary>Sender may use only the explicitly listed tool names.</summary>
        ToolList
        ''' <summary>Sender may run exactly one skill (plus its helpers); everything else is blocked.</summary>
        OnlySkill
    End Enum

    ''' <summary>A single parsed sender policy rule.</summary>
    Private NotInheritable Class AutoPilotSenderPolicyRule
        Public Property Pattern As String
        Public Property IsDefault As Boolean
        Public Property Kind As AutoPilotSenderPolicyRuleKind
        Public Property AllowedToolSelectors As New List(Of String)()
        Public Property DeniedToolSelectors As New List(Of String)()
        Public Property SkillToolName As String = ""

        ''' <summary>
        ''' Optional hard system-prompt instruction injected for this sender (in code, not
        ''' under LLM control). Specified in the policy file after a "||" separator.
        ''' </summary>
        Public Property SystemPromptAddition As String = ""
    End Class

    ''' <summary>Parsed sender policy rules for the active session (Nothing = no policy in effect).</summary>
    Private _apSenderPolicyRules As List(Of AutoPilotSenderPolicyRule) = Nothing

    ''' <summary>
    ''' Loads and parses the sender tool policy from the configured file path.
    ''' Call once at AutoPilot session start. Fully silent and fail-safe.
    ''' </summary>
    ''' <param name="policyPath">Full path to the policy text file (may be empty).</param>
    Friend Sub LoadSenderToolPolicy(policyPath As String)
        _apSenderPolicyRules = Nothing

        Try
            If String.IsNullOrWhiteSpace(policyPath) Then Return
            If Not File.Exists(policyPath) Then
                Debug.WriteLine($"[AutoPilot] Sender tool policy file not found: {policyPath}")
                Return
            End If

            Dim rules As New List(Of AutoPilotSenderPolicyRule)()
            For Each rawLine In File.ReadAllLines(policyPath)
                Dim line = If(rawLine, "").Trim()
                If line.Length = 0 OrElse line.StartsWith("#") OrElse line.StartsWith(";") Then Continue For

                Dim sepIndex = line.IndexOf("="c)
                If sepIndex <= 0 Then Continue For

                Dim patternPart = line.Substring(0, sepIndex).Trim()
                Dim rulePart = line.Substring(sepIndex + 1).Trim()
                If patternPart.Length = 0 OrElse rulePart.Length = 0 Then Continue For

                ' Optional per-sender system-prompt instruction after a "||" separator:
                '     <pattern> = <rule> || <system prompt instruction>
                Dim promptPart As String = ""
                Dim promptSepIndex = rulePart.IndexOf("||", StringComparison.Ordinal)
                If promptSepIndex >= 0 Then
                    promptPart = rulePart.Substring(promptSepIndex + 2).Trim()
                    rulePart = rulePart.Substring(0, promptSepIndex).Trim()
                    If rulePart.Length = 0 Then Continue For
                End If

                Dim rule As New AutoPilotSenderPolicyRule()
                rule.IsDefault = patternPart.Equals("DEFAULT", StringComparison.OrdinalIgnoreCase)
                rule.Pattern = patternPart
                rule.SystemPromptAddition = promptPart

                If rulePart.Equals("ALL", StringComparison.OrdinalIgnoreCase) Then
                    rule.Kind = AutoPilotSenderPolicyRuleKind.AllowAll
                ElseIf rulePart.Equals("NONE", StringComparison.OrdinalIgnoreCase) Then
                    rule.Kind = AutoPilotSenderPolicyRuleKind.BlockAll
                ElseIf rulePart.StartsWith("ONLY ", StringComparison.OrdinalIgnoreCase) Then
                    rule.Kind = AutoPilotSenderPolicyRuleKind.OnlySkill
                    rule.SkillToolName = rulePart.Substring(5).Trim()
                Else
                    rule.Kind = AutoPilotSenderPolicyRuleKind.ToolList

                    For Each namePart In rulePart.Split(","c)
                        Dim selector As String = namePart.Trim()
                        If selector.Length = 0 Then Continue For

                        Dim isDenied As Boolean =
            selector.StartsWith("!", StringComparison.Ordinal) OrElse
            selector.StartsWith("-", StringComparison.Ordinal)

                        If isDenied Then
                            selector = selector.Substring(1).Trim()
                        End If

                        If selector.Length = 0 Then Continue For

                        If isDenied Then
                            rule.DeniedToolSelectors.Add(selector)
                        Else
                            rule.AllowedToolSelectors.Add(selector)
                        End If
                    Next
                End If

                rules.Add(rule)
            Next

            If rules.Count > 0 Then
                _apSenderPolicyRules = rules
                Debug.WriteLine($"[AutoPilot] Sender tool policy loaded: {rules.Count} rule(s).")
            End If
        Catch ex As Exception
            _apSenderPolicyRules = Nothing
            Debug.WriteLine($"[AutoPilot] Failed to load sender tool policy: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Applies the per-sender tool policy to a session tool list and returns the
    ''' (possibly narrowed) list. Returns the original list unchanged when no policy
    ''' is in effect, so existing behaviour is fully preserved.
    ''' </summary>
    ''' <param name="senderEmail">SMTP address of the current sender.</param>
    ''' <param name="sessionTools">The tools selected for the session.</param>
    ''' <summary>
    ''' Returns the policy rule that applies to a sender (first pattern match wins;
    ''' DEFAULT is used only when no specific pattern matched). Nothing = no policy.
    ''' </summary>
    Private Function MatchSenderPolicyRule(senderEmail As String) As AutoPilotSenderPolicyRule
        If _apSenderPolicyRules Is Nothing OrElse _apSenderPolicyRules.Count = 0 Then Return Nothing

        Dim matched As AutoPilotSenderPolicyRule = Nothing
        Dim defaultRule As AutoPilotSenderPolicyRule = Nothing

        For Each rule In _apSenderPolicyRules
            If rule.IsDefault Then
                If defaultRule Is Nothing Then defaultRule = rule
                Continue For
            End If
            If Not String.IsNullOrWhiteSpace(senderEmail) AndAlso WildcardMatch(senderEmail, rule.Pattern) Then
                matched = rule
                Exit For
            End If
        Next

        If matched Is Nothing Then matched = defaultRule
        Return matched
    End Function

    ''' <summary>
    ''' Returns the hard, in-code system-prompt instruction that applies to a sender,
    ''' or an empty string when none applies. For an ONLY-skill rule without an explicit
    ''' instruction, a default instruction is generated that confines the sender to that skill.
    ''' </summary>
    Friend Function ResolveSystemPromptAdditionForSender(senderEmail As String) As String
        Dim matched = MatchSenderPolicyRule(senderEmail)
        If matched Is Nothing Then Return ""

        Dim addition = If(matched.SystemPromptAddition, "").Trim()

        If addition.Length = 0 AndAlso
           matched.Kind = AutoPilotSenderPolicyRuleKind.OnlySkill AndAlso
           Not String.IsNullOrWhiteSpace(matched.SkillToolName) Then
            addition =
                $"For this sender, you must handle the request exclusively by invoking the '{matched.SkillToolName.Trim()}' skill. " &
                "You may use only the helper tools explicitly declared by that skill in its allowed-tools frontmatter, plus host safety/runtime tools that are exposed automatically. " &
                "Do not use any other skill, agent, tool, online source, or undeclared model knowledge outside the loaded skill's instructions and references. " &
                "If the request cannot be fulfilled by that skill and its declared helpers, briefly decline using report_inability."
        End If

        Return addition
    End Function

    Friend Function ResolveToolsForSender(senderEmail As String, sessionTools As List(Of ModelConfig)) As List(Of ModelConfig)
        ' No policy in effect -> unchanged behaviour.
        If _apSenderPolicyRules Is Nothing OrElse _apSenderPolicyRules.Count = 0 Then Return sessionTools
        If sessionTools Is Nothing OrElse sessionTools.Count = 0 Then Return sessionTools

        Dim matched As AutoPilotSenderPolicyRule = MatchSenderPolicyRule(senderEmail)

        ' No matching rule and no DEFAULT -> sender inherits all tools (unchanged).
        If matched Is Nothing Then Return sessionTools

        Select Case matched.Kind
            Case AutoPilotSenderPolicyRuleKind.AllowAll
                Return sessionTools

            Case AutoPilotSenderPolicyRuleKind.BlockAll
                Return KeepAlwaysAllowedToolsOnly(sessionTools)

            Case AutoPilotSenderPolicyRuleKind.ToolList
                Return FilterToNamedTools(sessionTools, matched.AllowedToolSelectors, matched.DeniedToolSelectors)

            Case AutoPilotSenderPolicyRuleKind.OnlySkill
                Return FilterToExclusiveSkill(sessionTools, matched.SkillToolName)

            Case Else
                Return sessionTools
        End Select
    End Function

    ''' <summary>Returns only the always-allowed safety tools (report_inability).</summary>
    Private Function KeepAlwaysAllowedToolsOnly(sessionTools As List(Of ModelConfig)) As List(Of ModelConfig)
        Return sessionTools.
            Where(Function(t) t IsNot Nothing AndAlso IsAlwaysAllowedTool(t.ToolName)).
            ToList()
    End Function

    ''' <summary>
    ''' Returns only tools that match the allowed selectors and do not match the denied selectors,
    ''' plus the always-allowed safety tools. If only denied selectors are provided, the rule behaves
    ''' like ALL except those exclusions.
    ''' </summary>
    Private Function FilterToNamedTools(
    sessionTools As List(Of ModelConfig),
    allowedSelectors As List(Of String),
    deniedSelectors As List(Of String)) As List(Of ModelConfig)

        Dim allowed As List(Of String) =
        If(allowedSelectors, New List(Of String)()).
            Where(Function(n) Not String.IsNullOrWhiteSpace(n)).
            Select(Function(n) n.Trim()).
            ToList()

        Dim denied As List(Of String) =
        If(deniedSelectors, New List(Of String)()).
            Where(Function(n) Not String.IsNullOrWhiteSpace(n)).
            Select(Function(n) n.Trim()).
            ToList()

        Return sessionTools.
        Where(Function(t)
                  If t Is Nothing OrElse String.IsNullOrWhiteSpace(t.ToolName) Then Return False

                  Dim toolName As String = t.ToolName.Trim()

                  If IsAlwaysAllowedTool(toolName) Then Return True

                  Dim isAllowed As Boolean =
                      allowed.Count = 0 OrElse
                      allowed.Any(Function(selector) ToolSelectorMatches(selector, toolName))

                  If Not isAllowed Then Return False

                  Dim isDenied As Boolean =
                      denied.Any(Function(selector) ToolSelectorMatches(selector, toolName))

                  Return Not isDenied
              End Function).
        ToList()
    End Function

    ''' <summary>
    ''' Agnostic selector match: a selector is matched purely against a tool/skill/agent/online-resource
    ''' NAME. Any named entity can be allowed or denied. "*" (or "ALL") matches everything; selectors
    ''' containing * or ? are treated as wildcard name patterns; otherwise an exact (case-insensitive)
    ''' name comparison is used. No entity-type-specific or name-specific heuristics are applied.
    ''' </summary>
    Private Function ToolSelectorMatches(selector As String, toolName As String) As Boolean
        Dim normalizedSelector As String = If(selector, "").Trim()
        Dim normalizedToolName As String = If(toolName, "").Trim()

        If normalizedSelector = "" OrElse normalizedToolName = "" Then Return False

        If normalizedSelector = "*" OrElse
           normalizedSelector.Equals("ALL", StringComparison.OrdinalIgnoreCase) Then
            Return True
        End If

        If ContainsWildcardPattern(normalizedSelector) Then
            Return WildcardToolNameMatches(normalizedSelector, normalizedToolName)
        End If

        Return normalizedToolName.Equals(normalizedSelector, StringComparison.OrdinalIgnoreCase)
    End Function

    ''' <summary>
    ''' Restricts the already-authorized AutoPilot session to exactly one skill plus only
    ''' the helper tools that skill explicitly declares in allowed-tools. This function is
    ''' purely narrowing: it never enables a tool, agent, source, or service that was not
    ''' already present in the session tool set. Other skills are always blocked.
    ''' </summary>
    Private Function FilterToExclusiveSkill(sessionTools As List(Of ModelConfig), skillToolName As String) As List(Of ModelConfig)
        Dim exclusive As String = If(skillToolName, "").Trim()
        Dim declaredHelperNames As System.Collections.Generic.HashSet(Of String) =
            ResolveExclusiveSkillDeclaredHelperNames(exclusive)

        Return sessionTools.
            Where(Function(t)
                      If t Is Nothing OrElse System.String.IsNullOrWhiteSpace(t.ToolName) Then Return False
                      Dim name As String = t.ToolName.Trim()

                      If IsAlwaysAllowedTool(name) Then Return True

                      If name.StartsWith("skill_", System.StringComparison.OrdinalIgnoreCase) Then
                          Return name.Equals(exclusive, System.StringComparison.OrdinalIgnoreCase)
                      End If

                      Return declaredHelperNames.Contains(name)
                  End Function).
            ToList()
    End Function

    Private Function ResolveExclusiveSkillDeclaredHelperNames(exclusiveSkillToolName As String) As System.Collections.Generic.HashSet(Of String)
        Dim result As New System.Collections.Generic.HashSet(Of String)(System.StringComparer.OrdinalIgnoreCase)

        If System.String.IsNullOrWhiteSpace(exclusiveSkillToolName) Then Return result

        Try
            SharedLibrary.Agents.AgentResources.Refresh()

            For Each skill As SharedLibrary.Agents.SkillDescriptor In SharedLibrary.Agents.AgentResources.Skills
                If skill Is Nothing OrElse System.String.IsNullOrWhiteSpace(skill.Name) Then Continue For

                Dim dynamicToolName As String = BuildDynamicSkillToolName(skill.Name)
                If Not dynamicToolName.Equals(exclusiveSkillToolName.Trim(), System.StringComparison.OrdinalIgnoreCase) Then Continue For

                If skill.AllowedTools IsNot Nothing Then
                    For Each rawToolName As String In skill.AllowedTools
                        Dim helperName As String = If(rawToolName, "").Trim()
                        If helperName <> "" Then result.Add(helperName)
                    Next
                End If

                Exit For
            Next
        Catch ex As System.Exception
            System.Diagnostics.Debug.WriteLine($"[AutoPilot] Failed to resolve ONLY-skill helper declarations: {ex.Message}")
        End Try

        Return result
    End Function

    Private Function BuildDynamicSkillToolName(skillName As String) As String
        If System.String.IsNullOrWhiteSpace(skillName) Then Return ""

        Dim suffixBuilder As New System.Text.StringBuilder()
        Dim lastWasUnderscore As Boolean = False

        For Each ch As Char In skillName.Trim()
            If System.Char.IsLetterOrDigit(ch) Then
                suffixBuilder.Append(System.Char.ToLowerInvariant(ch))
                lastWasUnderscore = False
            ElseIf Not lastWasUnderscore Then
                suffixBuilder.Append("_"c)
                lastWasUnderscore = True
            End If
        Next

        Dim suffix As String = suffixBuilder.ToString().Trim("_"c)
        If suffix = "" Then Return ""
        Return "skill_" & suffix
    End Function

    ''' <summary>Determines whether a tool must never be filtered out (safety fallback).</summary>
    Private Function IsAlwaysAllowedTool(toolName As String) As Boolean
        If String.IsNullOrWhiteSpace(toolName) Then Return False
        Return toolName.Trim().Equals(AP_Tool_ReportInability, StringComparison.OrdinalIgnoreCase)
    End Function


End Class
