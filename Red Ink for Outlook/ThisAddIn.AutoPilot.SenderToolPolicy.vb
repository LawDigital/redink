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
'     # or ;            -> comment line (ignored)
'     <pattern> = ALL                 -> sender may use every selected tool
'     <pattern> = NONE                -> sender may use nothing (report_inability only)
'     <pattern> = tool1, tool2, ...   -> sender may use only the listed tools
'     <pattern> = ONLY skill_<name>   -> sender may only run that one skill
'     DEFAULT = <rule>                -> fallback for senders that match no pattern
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
        Public Property ToolNames As New List(Of String)()
        Public Property SkillToolName As String = ""
    End Class

    ''' <summary>Parsed sender policy rules for the active session (Nothing = no policy in effect).</summary>
    Private _apSenderPolicyRules As List(Of AutoPilotSenderPolicyRule) = Nothing

    ''' <summary>Cached set of internal (built-in) tool names, used for ONLY-skill helper retention.</summary>
    Private _apInternalToolNamesCache As HashSet(Of String) = Nothing

    ''' <summary>
    ''' Loads and parses the sender tool policy from the configured file path.
    ''' Call once at AutoPilot session start. Fully silent and fail-safe.
    ''' </summary>
    ''' <param name="policyPath">Full path to the policy text file (may be empty).</param>
    Friend Sub LoadSenderToolPolicy(policyPath As String)
        _apSenderPolicyRules = Nothing
        _apInternalToolNamesCache = Nothing

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

                Dim rule As New AutoPilotSenderPolicyRule()
                rule.IsDefault = patternPart.Equals("DEFAULT", StringComparison.OrdinalIgnoreCase)
                rule.Pattern = patternPart

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
                        Dim toolName = namePart.Trim()
                        If toolName.Length > 0 Then rule.ToolNames.Add(toolName)
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
    Friend Function ResolveToolsForSender(senderEmail As String, sessionTools As List(Of ModelConfig)) As List(Of ModelConfig)
        ' No policy in effect -> unchanged behaviour.
        If _apSenderPolicyRules Is Nothing OrElse _apSenderPolicyRules.Count = 0 Then Return sessionTools
        If sessionTools Is Nothing OrElse sessionTools.Count = 0 Then Return sessionTools

        ' First match wins; DEFAULT is only used when no specific pattern matched.
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

        ' No matching rule and no DEFAULT -> sender inherits all tools (unchanged).
        If matched Is Nothing Then Return sessionTools

        Select Case matched.Kind
            Case AutoPilotSenderPolicyRuleKind.AllowAll
                Return sessionTools

            Case AutoPilotSenderPolicyRuleKind.BlockAll
                Return KeepAlwaysAllowedToolsOnly(sessionTools)

            Case AutoPilotSenderPolicyRuleKind.ToolList
                Return FilterToNamedTools(sessionTools, matched.ToolNames)

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

    ''' <summary>Returns only tools whose name is in the allowed list, plus the always-allowed safety tools.</summary>
    Private Function FilterToNamedTools(sessionTools As List(Of ModelConfig), allowedNames As List(Of String)) As List(Of ModelConfig)
        Dim allowed As New HashSet(Of String)(
            If(allowedNames, New List(Of String)()).Where(Function(n) Not String.IsNullOrWhiteSpace(n)).Select(Function(n) n.Trim()),
            StringComparer.OrdinalIgnoreCase)

        Return sessionTools.
            Where(Function(t) t IsNot Nothing AndAlso
                              (IsAlwaysAllowedTool(t.ToolName) OrElse allowed.Contains(If(t.ToolName, "").Trim()))).
            ToList()
    End Function

    ''' <summary>
    ''' Restricts the session to a single skill (Option B): keeps the named skill, the
    ''' skill loader, the always-allowed safety tools, and the internal helper tools the
    ''' skill may need, while removing all other skills, all agents, and all external sources.
    ''' </summary>
    Private Function FilterToExclusiveSkill(sessionTools As List(Of ModelConfig), skillToolName As String) As List(Of ModelConfig)
        Dim exclusive = If(skillToolName, "").Trim()
        Dim internalNames = GetInternalToolNames()

        Return sessionTools.
            Where(Function(t)
                      If t Is Nothing OrElse String.IsNullOrWhiteSpace(t.ToolName) Then Return False
                      Dim name = t.ToolName.Trim()

                      ' Always keep the safety tool and the skill loader.
                      If IsAlwaysAllowedTool(name) Then Return True
                      If name.Equals(SharedLibrary.Agents.SkillInvokeTool.ToolName, StringComparison.OrdinalIgnoreCase) Then Return True

                      ' Keep exactly the exclusive skill; block all other skills and all agents.
                      If name.StartsWith("skill_", StringComparison.OrdinalIgnoreCase) Then
                          Return name.Equals(exclusive, StringComparison.OrdinalIgnoreCase)
                      End If
                      If name.StartsWith("agent_", StringComparison.OrdinalIgnoreCase) Then Return False

                      ' Keep internal helper tools (the skill may need them); block external sources.
                      Return internalNames.Contains(name)
                  End Function).
            ToList()
    End Function

    ''' <summary>Determines whether a tool must never be filtered out (safety fallback).</summary>
    Private Function IsAlwaysAllowedTool(toolName As String) As Boolean
        If String.IsNullOrWhiteSpace(toolName) Then Return False
        Return toolName.Trim().Equals(AP_Tool_ReportInability, StringComparison.OrdinalIgnoreCase)
    End Function

    ''' <summary>Builds (and caches) the set of internal/built-in AutoPilot tool names.</summary>
    Private Function GetInternalToolNames() As HashSet(Of String)
        If _apInternalToolNamesCache IsNot Nothing Then Return _apInternalToolNamesCache

        Dim names As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Try
            For Each tool In GetAutoPilotInternalTools()
                If tool IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(tool.ToolName) Then
                    names.Add(tool.ToolName.Trim())
                End If
            Next
        Catch ex As Exception
            Debug.WriteLine($"[AutoPilot] Failed to enumerate internal tools for sender policy: {ex.Message}")
        End Try

        _apInternalToolNamesCache = names
        Return names
    End Function

End Class
