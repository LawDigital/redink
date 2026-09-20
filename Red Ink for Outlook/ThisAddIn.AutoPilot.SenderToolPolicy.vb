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
'     <pattern> = <rule>, designs=block, design_sets=block
'                              -> optional resource capabilities for internal designs
'     <pattern> = <rule>, disclaimer="plain text"
'                              -> host-appended disclaimer for AutoPilot mail to that address
'
'  - A selector is matched purely against the NAME of a tool, skill, agent, or online
'    resource. Selectors may be:
'      - an exact name (e.g. internet_search, skill_intake, agent_research, some_online_source)
'      - a wildcard name pattern using * and ? (e.g. skill_*, swiss-caselaw*, agent_?)
'      - the universal placeholder * (or ALL) matching every name
'  - Prefix any selector with ! or - to EXCLUDE the matching names.
'  - If a selector list contains only exclusions, it behaves like * except those exclusions.
'
'  - Resource directives are optional and independent of tool selectors:
'      designs=allow|block       controls all internal named designs
'      design_sets=allow|block   controls only design_sets/active.json routing
'    designs=block dominates design_sets and prevents names/lookups from being exposed.
'  - disclaimer=... is plain text appended by the host. Quote values containing semicolons;
'    use \n for a line break and \" for a literal quote.
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
        Public Property AllowDesigns As System.Nullable(Of System.Boolean) = Nothing
        Public Property AllowDesignSets As System.Nullable(Of System.Boolean) = Nothing
        Public Property DisclaimerText As System.String = System.String.Empty

        ''' <summary>
        ''' Optional hard system-prompt instruction injected for this sender (in code, not
        ''' under LLM control). Specified in the policy file after a "||" separator.
        ''' </summary>
        Public Property SystemPromptAddition As String = ""
    End Class

    Friend NotInheritable Class AutoPilotSenderDesignAccess
        Public Property AllowDesigns As System.Boolean = True
        Public Property AllowDesignSets As System.Boolean = True
    End Class

    ''' <summary>Parsed sender policy rules for the active session (Nothing = no policy in effect).</summary>
    Private _apSenderPolicyRules As List(Of AutoPilotSenderPolicyRule) = Nothing

    ''' <summary>
    ''' Loads and parses the sender tool policy from the configured file path.
    ''' Call once at AutoPilot session start. Fully silent and fail-safe.
    ''' </summary>
    ''' <param name="policyPath">Full path to the policy text file (may be empty).</param>
    Friend Sub LoadSenderToolPolicy(policyPath As String)
        Try
            If String.IsNullOrWhiteSpace(policyPath) Then
                _apSenderPolicyRules = Nothing
                Return
            End If
            If Not File.Exists(policyPath) Then
                ' Keep the last successfully loaded rules on a transient path/read failure.
                ' A configured security policy must not silently become unrestricted merely
                ' because the backing file is temporarily unavailable.
                Debug.WriteLine($"[AutoPilot] Sender tool policy file not found; keeping last known policy: {policyPath}")
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
                Dim promptSepIndex As System.Int32 = IndexOfUnquotedSenderRuleToken(rulePart, "||")
                If promptSepIndex >= 0 Then
                    promptPart = rulePart.Substring(promptSepIndex + 2).Trim()
                    rulePart = rulePart.Substring(0, promptSepIndex).Trim()
                    If rulePart.Length = 0 Then Continue For
                End If

                Dim rule As New AutoPilotSenderPolicyRule()
                rule.IsDefault = patternPart.Equals("DEFAULT", StringComparison.OrdinalIgnoreCase)
                rule.Pattern = patternPart
                rule.SystemPromptAddition = promptPart

                ' Preserve historical comma-separated tool selectors while allowing resource/output
                ' directives after either commas or semicolons. Only known key=value directives are
                ' removed from the tool rule; everything else remains a selector.
                Dim ruleTokens As System.String() = SplitSenderRuleComponents(rulePart)
                Dim toolRuleTokens As New System.Collections.Generic.List(Of System.String)()
                For Each ruleToken As System.String In ruleTokens
                    If IsKnownSenderResourceDirective(ruleToken) Then
                        ApplySenderResourceDirective(rule, ruleToken)
                    ElseIf Not System.String.IsNullOrWhiteSpace(ruleToken) Then
                        toolRuleTokens.Add(ruleToken.Trim())
                    End If
                Next
                rulePart = System.String.Join(", ", toolRuleTokens).Trim()
                If rulePart.Length = 0 Then Continue For

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

            ' Publish the newly parsed set atomically only after the file was read successfully.
            ' A valid empty/comment-only file intentionally disables per-sender rules.
            _apSenderPolicyRules = If(rules.Count > 0, rules, Nothing)
            Debug.WriteLine($"[AutoPilot] Sender tool policy loaded: {rules.Count} rule(s).")
        Catch ex As System.Exception
            ' Preserve the last known good policy on transient read/parse failures.
            Debug.WriteLine($"[AutoPilot] Failed to reload sender tool policy; keeping last known policy: {ex.Message}")
        End Try
    End Sub

    Private Shared Function IndexOfUnquotedSenderRuleToken(ByVal value As System.String, ByVal token As System.String) As System.Int32
        If System.String.IsNullOrEmpty(value) OrElse System.String.IsNullOrEmpty(token) Then Return -1

        Dim inQuotes As System.Boolean = False
        Dim escaped As System.Boolean = False

        For index As System.Int32 = 0 To value.Length - token.Length
            Dim current As System.Char = value(index)
            If escaped Then
                escaped = False
                Continue For
            End If
            If current = "\"c AndAlso inQuotes Then
                escaped = True
                Continue For
            End If
            If current = """"c Then
                inQuotes = Not inQuotes
                Continue For
            End If
            If Not inQuotes AndAlso value.Substring(index, token.Length).Equals(token, System.StringComparison.Ordinal) Then Return index
        Next

        Return -1
    End Function

    Private Shared Function SplitSenderRuleComponents(ByVal value As System.String) As System.String()
        Dim result As New System.Collections.Generic.List(Of System.String)()
        Dim current As New System.Text.StringBuilder()
        Dim inQuotes As System.Boolean = False
        Dim escaped As System.Boolean = False

        For Each character As System.Char In If(value, System.String.Empty)
            If escaped Then
                current.Append(character)
                escaped = False
                Continue For
            End If
            If character = "\"c AndAlso inQuotes Then
                current.Append(character)
                escaped = True
                Continue For
            End If
            If character = """"c Then
                current.Append(character)
                inQuotes = Not inQuotes
                Continue For
            End If
            If (character = ";"c OrElse character = ","c) AndAlso Not inQuotes Then
                result.Add(current.ToString().Trim())
                current.Clear()
            Else
                current.Append(character)
            End If
        Next

        result.Add(current.ToString().Trim())
        Return result.ToArray()
    End Function

    Private Shared Function ParseSenderDisclaimerValue(ByVal rawValue As System.String) As System.String
        Dim value As System.String = If(rawValue, System.String.Empty).Trim()
        If value.Length >= 2 AndAlso value(0) = """"c AndAlso value(value.Length - 1) = """"c Then
            value = value.Substring(1, value.Length - 2)
        End If

        Dim result As New System.Text.StringBuilder()
        Dim escaped As System.Boolean = False
        For Each character As System.Char In value
            If escaped Then
                Select Case character
                    Case "n"c : result.Append(System.Environment.NewLine)
                    Case "r"c : result.Append(ControlChars.Cr)
                    Case "t"c : result.Append(ControlChars.Tab)
                    Case """"c : result.Append(""""c)
                    Case "\"c : result.Append("\"c)
                    Case Else
                        result.Append("\"c)
                        result.Append(character)
                End Select
                escaped = False
            ElseIf character = "\"c Then
                escaped = True
            Else
                result.Append(character)
            End If
        Next
        If escaped Then result.Append("\"c)
        Return result.ToString().Trim()
    End Function

    ''' <summary>
    ''' Applies one optional resource-capability directive from a sender rule.
    ''' Unknown keys are ignored for forward compatibility; malformed known design directives fail closed.
    ''' </summary>
    Private Shared Function IsKnownSenderResourceDirective(ByVal rawToken As System.String) As System.Boolean
        If System.String.IsNullOrWhiteSpace(rawToken) Then Return False
        Dim token As System.String = rawToken.Trim()
        Dim separatorIndex As System.Int32 = token.IndexOf("="c)
        If separatorIndex <= 0 Then Return False
        Dim key As System.String = token.Substring(0, separatorIndex).Trim()
        Return key.Equals("designs", System.StringComparison.OrdinalIgnoreCase) OrElse
               key.Equals("design_sets", System.StringComparison.OrdinalIgnoreCase) OrElse
               key.Equals("designsets", System.StringComparison.OrdinalIgnoreCase) OrElse
               key.Equals("disclaimer", System.StringComparison.OrdinalIgnoreCase)
    End Function

    Private Shared Sub ApplySenderResourceDirective(ByVal rule As AutoPilotSenderPolicyRule,
                                                     ByVal rawDirective As System.String)
        If rule Is Nothing OrElse System.String.IsNullOrWhiteSpace(rawDirective) Then Return

        Dim directive As System.String = rawDirective.Trim()
        Dim separatorIndex As System.Int32 = directive.IndexOf("="c)
        Dim key As System.String = If(separatorIndex > 0,
                                      directive.Substring(0, separatorIndex),
                                      directive).Trim().ToLowerInvariant()

        Dim isDesignsDirective As System.Boolean = key.Equals("designs", System.StringComparison.OrdinalIgnoreCase)
        Dim isDesignSetsDirective As System.Boolean =
            key.Equals("design_sets", System.StringComparison.OrdinalIgnoreCase) OrElse
            key.Equals("designsets", System.StringComparison.OrdinalIgnoreCase)
        Dim isDisclaimerDirective As System.Boolean = key.Equals("disclaimer", System.StringComparison.OrdinalIgnoreCase)

        ' Unknown directives remain forward-compatible and do not affect legacy tool selection.
        If Not isDesignsDirective AndAlso Not isDesignSetsDirective AndAlso Not isDisclaimerDirective Then Return

        If isDisclaimerDirective Then
            rule.DisclaimerText = If(separatorIndex > 0, ParseSenderDisclaimerValue(directive.Substring(separatorIndex + 1)), System.String.Empty)
            Return
        End If

        ' A malformed KNOWN resource directive must fail closed. Otherwise a typo such as
        ' "designs=blok" would silently expose the very repository the rule was intended to hide.
        Dim parsedValue As System.Nullable(Of System.Boolean) = Nothing
        If separatorIndex > 0 AndAlso separatorIndex < directive.Length - 1 Then
            Dim value As System.String = directive.Substring(separatorIndex + 1).Trim()
            parsedValue = ParseSenderCapabilityValue(value)
        End If

        Dim effectiveValue As System.Boolean
        If parsedValue.HasValue Then
            effectiveValue = parsedValue.Value
        Else
            effectiveValue = False
            System.Diagnostics.Debug.WriteLine(
                $"[AutoPilot] Invalid sender resource directive '{directive}' was applied fail-closed (blocked).")
        End If

        If isDesignsDirective Then
            rule.AllowDesigns = effectiveValue
        Else
            rule.AllowDesignSets = effectiveValue
        End If
    End Sub

    Private Shared Function ParseSenderCapabilityValue(ByVal rawValue As System.String) As System.Nullable(Of System.Boolean)
        Select Case If(rawValue, System.String.Empty).Trim().ToLowerInvariant()
            Case "allow", "allowed", "true", "yes", "on", "1"
                Return True
            Case "block", "blocked", "deny", "denied", "false", "no", "off", "0"
                Return False
            Case Else
                Return Nothing
        End Select
    End Function

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

    Friend Function ResolveDesignAccessForSender(ByVal senderEmail As System.String) As AutoPilotSenderDesignAccess
        Dim result As New AutoPilotSenderDesignAccess()
        Dim matched As AutoPilotSenderPolicyRule = MatchSenderPolicyRule(senderEmail)
        If matched Is Nothing Then Return result

        If matched.AllowDesigns.HasValue Then result.AllowDesigns = matched.AllowDesigns.Value
        If matched.AllowDesignSets.HasValue Then result.AllowDesignSets = matched.AllowDesignSets.Value
        If Not result.AllowDesigns Then result.AllowDesignSets = False
        Return result
    End Function

    Friend Function ResolveDisclaimerForSender(ByVal senderEmail As System.String) As System.String
        Dim matched As AutoPilotSenderPolicyRule = MatchSenderPolicyRule(senderEmail)
        If matched Is Nothing Then Return System.String.Empty
        Return If(matched.DisclaimerText, System.String.Empty).Trim()
    End Function

    Private Shared Function BuildAutoPilotDisclaimerHtml(ByVal disclaimerText As System.String) As System.String
        Dim text As System.String = If(disclaimerText, System.String.Empty).Trim()
        If text.Length = 0 Then Return System.String.Empty

        Dim encoded As System.String = System.Net.WebUtility.HtmlEncode(text)
        encoded = encoded.Replace(ControlChars.CrLf, "<br/>").Replace(ControlChars.Cr, "<br/>").Replace(ControlChars.Lf, "<br/>")
        Return "<div data-redink-autopilot-disclaimer='true' style='margin-top:16px;padding-top:10px;border-top:1px solid #d0d0d0;font-family:Arial,sans-serif;font-size:9pt;color:#666666;'>" & encoded & "</div>"
    End Function

    ''' <summary>
    ''' Applies the per-sender tool policy to a session tool list and returns the
    ''' possibly narrowed list. Existing behaviour is preserved when no matching
    ''' policy is configured.
    ''' </summary>
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
            SharedLibrary.Agents.AgentResources.EnsureFresh()

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
        Return SharedLibrary.Agents.ToolRegistryBuilder.BuildSkillToolName(skillName)
    End Function

    ''' <summary>Determines whether a tool must never be filtered out (safety fallback).</summary>
    Private Function IsAlwaysAllowedTool(toolName As String) As Boolean
        If String.IsNullOrWhiteSpace(toolName) Then Return False
        Return toolName.Trim().Equals(AP_Tool_ReportInability, StringComparison.OrdinalIgnoreCase)
    End Function


End Class
