' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SkillInvokeTool.vb
' Purpose: Implements the universal "skill_use" tool. The model calls
'          skill_use(name, input?) and receives the SKILL.md body (loaded lazily)
'          along with an inventory of the skill's scripts/ and references/ dirs.
'
' Architecture:
'  - Lazy-loads skill bodies on first access via AgentResources.FindSkill().
'  - Returns skill instructions + inventory (names + sizes) as JSON.
'  - Model follows those instructions in subsequent turns.
'  - Text and script bodies are NOT auto-loaded; model fetches what it needs.
'  - Binary reference/script assets are discovered by inventory and can be
'    materialized with file_* tools when the loaded skill allows them.
'  - Security: allowed-tools communicated; enforcement by host runner.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Agents

    Public NotInheritable Class SkillInvokeTool

        Private Sub New()
        End Sub

        Public Const ToolName As String = "skill_use"

        ' Host-identity hook. Each host wires this up at startup so a loaded skill
        ' can deterministically know where it runs ("Word" or "Outlook Local Chat")
        ' instead of guessing from the visible tool set. skill_use is not available
        ' under AutoPilot, so AutoPilot never needs to set this.
        Public Shared Property CurrentHostProvider As Func(Of String)

        Public Shared Function Build() As SharedLibrary.ModelConfig
            Dim def =
                "{""name"":""" & ToolName & """," &
                """description"":""Load and apply a Skill (Claude-style SKILL.md). Returns the skill's instructions and an inventory of its scripts/ and references/ files. Read text files with text_read, materialize binary reference or script assets with the appropriate file_* tools when allowed, and execute scripts with js_run. Use this when a relevant skill is offered above and the user's task matches."",""parameters"":{" &
                """type"":""object""," &
                """properties"":{" &
                """name"":{""type"":""string"",""description"":""The skill name (matches the Skill listed above).""}," &
                """input"":{""type"":""string"",""description"":""Optional input or sub-task description for the skill.""}," &
                """expected_artifacts"":{""type"":""array"",""description"":""Exact expected-final-artifact contract when the selected skill declares deliverable-count > 0 in frontmatter. Use opaque logical_deliverable_id/output_slot_id pairs."",""items"":{""type"":""object"",""properties"":{" &
                """logical_deliverable_id"":{""type"":""string""}," &
                """output_slot_id"":{""type"":""string""}}," &
                """required"":[""logical_deliverable_id"",""output_slot_id""]}}}," &
                """required"":[""name""]}}"

            Return New SharedLibrary.ModelConfig() With {
                .ToolName = ToolName,
                .ToolDefinition = def,
                .ToolInstructionsPrompt = ToolName & ": Load a Skill's instructions (lazy). Call this once per skill, then follow its directions in subsequent turns, using text_read for text resources and the appropriate file_* tools for binary reference assets when the skill allows them. If the selected skill declares deliverable-count > 0, expected_artifacts is mandatory and must contain exactly that many opaque logical_deliverable_id/output_slot_id pairs.",
                .ModelDescription = "Skill loader",
                .Tool = True,
                .ToolPriority = 940,
                .ToolErrorHandling = "skip"
            }
        End Function

        ''' <summary>
        ''' Executes the skill_use call. Returns a JSON string suitable for the tool response.
        ''' Caller passes the dictionary from ToolCall.Arguments.
        ''' </summary>
        Public Shared Function Execute(arguments As IDictionary(Of String, Object)) As String
            Try
                Dim name As String = GetStr(arguments, "name")
                Dim input As String = GetStr(arguments, "input")

                If String.IsNullOrWhiteSpace(name) Then
                    name = GetStr(arguments, "tool")
                End If

                If String.IsNullOrWhiteSpace(name) Then
                    name = GetStr(arguments, "skill")
                End If

                If String.IsNullOrWhiteSpace(name) Then
                    Return JsonConvert.SerializeObject(New With {Key .error = "missing_name"})
                End If

                ' Canonical, agnostic skill resolution: try the name exactly as provided
                ' first, then fall back to a single "skill_" prefix strip only if that
                ' actually resolves. This avoids double-stripping when the tool name
                ' (e.g. "skill_<slug>") has already been reduced to the skill's own name,
                ' which may itself legitimately start with "skill_".
                Dim sk = AgentResources.FindSkill(name)

                If sk Is Nothing AndAlso name.StartsWith("skill_", StringComparison.OrdinalIgnoreCase) Then
                    Dim strippedName As String = name.Substring("skill_".Length)
                    If Not String.IsNullOrWhiteSpace(strippedName) Then
                        Dim strippedSkill = AgentResources.FindSkill(strippedName)
                        If strippedSkill IsNot Nothing Then
                            name = strippedName
                            sk = strippedSkill
                        End If
                    End If
                End If

                If sk Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {Key .error = "skill_not_found", Key .name = name})
                End If

                Dim declaredDeliverableCount As System.Int32 =
                    GetDeclaredDeliverableCount(sk)

                Dim deliverableContractFailure As System.String = System.String.Empty
                If Not ValidateDeclaredDeliverableContract(
                    sk,
                    arguments,
                    deliverableContractFailure) Then

                    Dim contractError As New Newtonsoft.Json.Linq.JObject(
                        New Newtonsoft.Json.Linq.JProperty("code", "skill_deliverable_contract_invalid"),
                        New Newtonsoft.Json.Linq.JProperty("phase", "skill_invocation_validation"),
                        New Newtonsoft.Json.Linq.JProperty("message", deliverableContractFailure),
                        New Newtonsoft.Json.Linq.JProperty("retryable", True),
                        New Newtonsoft.Json.Linq.JProperty("required_deliverable_count", declaredDeliverableCount))

                    Return New Newtonsoft.Json.Linq.JObject(
                        New Newtonsoft.Json.Linq.JProperty("status", "failed"),
                        New Newtonsoft.Json.Linq.JProperty("summary", "Skill invocation rejected until its declared final-artifact contract is supplied."),
                        New Newtonsoft.Json.Linq.JProperty("result", Newtonsoft.Json.Linq.JValue.CreateNull()),
                        New Newtonsoft.Json.Linq.JProperty("resultKind", "error"),
                        New Newtonsoft.Json.Linq.JProperty("error", contractError)).
                        ToString(Newtonsoft.Json.Formatting.None)
                End If

                Dim body As String = sk.LoadBody()
                Dim scripts As List(Of Object) = InventoryDir(sk.ScriptsDir)
                Dim references As List(Of Object) = InventoryDir(sk.ReferencesDir)

                Dim result As New JObject()
                result("name") = sk.Name
                result("description") = If(sk.Description, "")
                result("origin") = If(sk.IsLocal, "local", "central")
                result("dir") = sk.DirectoryPath
                result("network_allowed") = sk.Network
                result("allowed_tools") = If(sk.AllowedTools Is Nothing,
                                             New JArray(),
                                             JArray.FromObject(sk.AllowedTools))
                result("declared_deliverable_count") = declaredDeliverableCount
                result("declared_deliverable_required_effects") =
                    New Newtonsoft.Json.Linq.JArray(GetDeclaredDeliverableRequiredEffects(sk))
                result("declared_required_successful_tools") =
                    New Newtonsoft.Json.Linq.JArray(GetDeclaredRequiredSuccessfulTools(sk))
                result("instructions") = body
                result("scripts") = JArray.FromObject(scripts)
                result("references") = JArray.FromObject(references)

                ' Provide a discovery index of all skills and agents with their exact
                ' file paths. Without this, an authoring skill has to guess where a
                ' resource lives, which leads to failed reads and accidental new files.
                result("resource_index") = BuildResourceIndex()

                If Not String.IsNullOrWhiteSpace(input) Then result("input") = input

                Return result.ToString(Formatting.None)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {Key .error = "skill_invoke_failed", Key .message = ex.Message})
            End Try
        End Function

        Private Shared Function BuildResourceIndex() As JObject
            Dim idx As New JObject()

            Dim skillsArr As New JArray()
            Try
                For Each s In AgentResources.Skills
                    If s Is Nothing Then Continue For
                    Dim o As New JObject()
                    o("name") = If(s.Name, "")
                    o("origin") = If(s.IsLocal, "local", "central")
                    o("file") = If(s.FilePath, "")
                    o("dir") = If(s.DirectoryPath, "")
                    skillsArr.Add(o)
                Next
            Catch
            End Try

            Dim agentsArr As New JArray()
            Try
                For Each a In AgentResources.Agents
                    If a Is Nothing Then Continue For
                    Dim o As New JObject()
                    o("name") = If(a.Name, "")
                    o("origin") = If(a.IsLocal, "local", "central")
                    o("file") = If(a.FilePath, "")
                    o("dir") = If(a.DirectoryPath, "")
                    agentsArr.Add(o)
                Next
            Catch
            End Try

            idx("skills") = skillsArr
            idx("agents") = agentsArr

            Dim localRoot As String = If(AgentResources.ConfiguredLocalPath, "")
            Dim centralRoot As String = If(AgentResources.ConfiguredCentralPath, "")
            Dim authorActive As Boolean = SkillAuthorMode.IsActive
            Dim allowCentral As Boolean = authorActive AndAlso SkillAuthorMode.AllowCentralWrites

            Dim currentHost As String = ""
            Try
                Dim hostProvider As Func(Of String) = CurrentHostProvider
                If hostProvider IsNot Nothing Then
                    Dim resolvedHost As String = hostProvider()
                    If resolvedHost IsNot Nothing Then currentHost = resolvedHost
                End If
            Catch
            End Try
            idx("host") = currentHost

            idx("local_root") = localRoot
            idx("central_root") = centralRoot
            idx("author_mode_active") = authorActive
            idx("local_writes_allowed") = authorActive
            idx("central_writes_allowed") = allowCentral

            ' Deterministic target root for NEW resources: local is writable only while author mode is
            ' active; central is only writable when it was additionally and explicitly enabled.
            Dim newResourceRoot As String =
                If(allowCentral AndAlso Not String.IsNullOrWhiteSpace(centralRoot), centralRoot, localRoot)
            idx("new_resource_root") = newResourceRoot

            idx("create_hint") =
                "ALWAYS write NEW resources under new_resource_root using ABSOLUTE paths — never a relative path, " &
                "because a relative path resolves into the temporary workspace, not the resource tree. " &
                "New skill: new_resource_root + '\skills\<name>\SKILL.md'. " &
                "New agent: new_resource_root + '\agents\<name>\AGENT.md' (or new_resource_root + '\agents\<name>.md'). " &
                "Parent folders are created automatically."

            idx("root_choice_hint") =
                If(String.IsNullOrWhiteSpace(centralRoot),
                   "Only a local resource root is configured; create and edit everything under local_root.",
                   If(allowCentral,
                      "Both a local and a central root are configured and central writing is ENABLED. " &
                      "Create NEW shared resources under central_root; create user-private resources under local_root. " &
                      "When unsure, prefer local_root.",
                      "Both a local and a central root are configured but central writing is DISABLED. " &
                      "Create ALL new resources under local_root. Never write under central_root."))

            idx("note") =
                If(authorActive, "", "Author mode is OFF: skills and agents are READ-ONLY; do not attempt to create or edit any resource files. ") &
                "To modify an EXISTING local resource, edit the exact 'file' path shown for it. " &
                "For an EXISTING central resource, edit that central path only when central_writes_allowed is true. " &
                "When central_writes_allowed is false, treat the central entry as a read-only source and create/update the corresponding local override under local_root using the same resource-relative skills/agents path. " &
                "Do not attempt a central write first and do not require the user to request an explicit fork when central writes are unavailable."

            If authorActive Then
                Dim diagPrefix As String = SharedLibrary.SharedMethods.AN9
                Dim diagnosticsDir As String = ""
                If Not String.IsNullOrWhiteSpace(localRoot) Then
                    diagnosticsDir = Path.Combine(localRoot, "diagnostics")
                End If

                idx("diagnostics_hint") =
                    "Diagnostics: previous tooling runs are logged under local_root + '\diagnostics\'. " &
                    "Tooling-loop traces are named '" & diagPrefix & "_Tooling_Log__<stamp>.txt' or, when a top-level " &
                    "skill was recorded, '" & diagPrefix & "_Tooling_Log__<stamp>__<skill>.txt' (where <stamp> is " &
                    "yyyyMMdd-HHmmss and <skill> is the sanitized skill name). Sub-agent/skill return payloads use the " &
                    "matching '" & diagPrefix & "_SubAgent_Returns__<stamp>[__<skill>].txt' names. Do NOT guess an exact " &
                    "filename: use resource_index.diagnostics_files to get the exact available files, then read the most " &
                    "recent relevant '" & diagPrefix & "_Tooling_Log__' file for the run you are diagnosing."

                idx("diagnostics_files") = JArray.FromObject(InventoryDiagnosticsDir(diagnosticsDir, diagPrefix & "_"))
            End If

            Return idx
        End Function

        ''' <summary>
        ''' Returns optional opaque host-verified effects that every final deliverable slot
        ''' of this skill must satisfy. Unknown but syntactically valid identifiers are
        ''' preserved so the orchestration contract is extensible without tool-specific branches.
        ''' </summary>
        Public Shared Function GetDeclaredDeliverableRequiredEffects(
            skill As SkillDescriptor) As System.Collections.Generic.List(Of System.String)

            Dim result As New System.Collections.Generic.List(Of System.String)()
            If skill Is Nothing OrElse skill.Frontmatter Is Nothing Then Return result

            Dim raw As System.String = Nothing
            If Not skill.Frontmatter.TryGetValue("deliverable-required-effects", raw) OrElse
               System.String.IsNullOrWhiteSpace(raw) Then
                Return result
            End If

            Dim seen As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            For Each part As System.String In raw.Split(New System.Char() {","c, ";"c, " "c}, System.StringSplitOptions.RemoveEmptyEntries)
                Dim effect As System.String = If(part, System.String.Empty).Trim().ToLowerInvariant()
                If effect = System.String.Empty Then Continue For
                If Not System.Text.RegularExpressions.Regex.IsMatch(effect, "^[a-z][a-z0-9_.-]{0,63}$") Then Continue For
                If seen.Add(effect) Then result.Add(effect)
            Next

            Return result
        End Function

        Public Shared Function GetDeclaredRequiredSuccessfulTools(
            skill As SkillDescriptor) As System.Collections.Generic.List(Of System.String)

            Dim result As New System.Collections.Generic.List(Of System.String)()
            If skill Is Nothing OrElse skill.Frontmatter Is Nothing Then Return result

            Dim raw As System.String = Nothing
            If Not skill.Frontmatter.TryGetValue("required-successful-tools", raw) OrElse
               System.String.IsNullOrWhiteSpace(raw) Then
                Return result
            End If

            Dim seen As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            Dim normalized As System.String = raw.Trim()
            If normalized.StartsWith("[", System.StringComparison.Ordinal) AndAlso
               normalized.EndsWith("]", System.StringComparison.Ordinal) AndAlso
               normalized.Length >= 2 Then
                normalized = normalized.Substring(1, normalized.Length - 2)
            End If

            For Each part As System.String In normalized.Split(New System.Char() {","c, ";"c, " "c}, System.StringSplitOptions.RemoveEmptyEntries)
                Dim toolName As System.String = If(part, System.String.Empty).Trim().Trim("'"c, """"c)
                If toolName = System.String.Empty Then Continue For
                If Not System.Text.RegularExpressions.Regex.IsMatch(toolName, "^[A-Za-z][A-Za-z0-9_.-]{0,127}$") Then Continue For
                If seen.Add(toolName) Then result.Add(toolName)
            Next

            Return result
        End Function

        Public Shared Function GetDeclaredDeliverableCount(skill As SkillDescriptor) As System.Int32
            If skill Is Nothing OrElse skill.Frontmatter Is Nothing Then Return 0

            Dim raw As System.String = Nothing
            If Not skill.Frontmatter.TryGetValue("deliverable-count", raw) OrElse
               System.String.IsNullOrWhiteSpace(raw) Then

                Return 0
            End If

            Dim parsed As System.Int32 = 0
            If Not System.Int32.TryParse(
                raw.Trim(),
                System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture,
                parsed) Then

                Return 0
            End If

            Return System.Math.Max(0, parsed)
        End Function

        ''' <summary>
        ''' Enforces a skill-declared final-deliverable count at skill invocation time.
        ''' This makes the expected-artifact contract authoritative before any producer runs,
        ''' so legacy staging files cannot later satisfy or widen the requested output set.
        ''' </summary>
        Public Shared Function ValidateDeclaredDeliverableContract(
            skill As SkillDescriptor,
            arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object),
            ByRef failureReason As System.String) As System.Boolean

            failureReason = System.String.Empty

            Dim requiredCount As System.Int32 =
                GetDeclaredDeliverableCount(skill)

            If requiredCount <= 0 Then Return True

            Dim raw As System.Object = Nothing
            If arguments Is Nothing OrElse
               Not arguments.TryGetValue("expected_artifacts", raw) OrElse
               raw Is Nothing Then

                failureReason =
                    "This skill declares deliverable-count=" &
                    requiredCount.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                    ". Invoke it with expected_artifacts containing exactly that many opaque logical_deliverable_id/output_slot_id pairs."
                Return False
            End If

            Dim token As Newtonsoft.Json.Linq.JToken = Nothing
            Try
                token = Newtonsoft.Json.Linq.JToken.FromObject(raw)
            Catch ex As System.Exception
                failureReason = "expected_artifacts could not be parsed as an array."
                Return False
            End Try

            If token Is Nothing OrElse token.Type <> Newtonsoft.Json.Linq.JTokenType.Array Then
                failureReason = "expected_artifacts must be an array."
                Return False
            End If

            Dim array As Newtonsoft.Json.Linq.JArray =
                DirectCast(token, Newtonsoft.Json.Linq.JArray)

            If array.Count <> requiredCount Then
                failureReason =
                    "This skill requires exactly " &
                    requiredCount.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                    " expected final artifact slot(s), but " &
                    array.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                    " were declared."
                Return False
            End If

            Dim seen As New System.Collections.Generic.HashSet(Of System.String)(
                System.StringComparer.Ordinal)

            For Each item As Newtonsoft.Json.Linq.JToken In array
                Dim obj As Newtonsoft.Json.Linq.JObject =
                    TryCast(item, Newtonsoft.Json.Linq.JObject)

                If obj Is Nothing Then
                    failureReason = "Every expected_artifacts item must be an object."
                    Return False
                End If

                Dim logicalId As System.String =
                    If(obj.Value(Of System.String)("logical_deliverable_id"), "").Trim()

                Dim slotId As System.String =
                    If(obj.Value(Of System.String)("output_slot_id"), "").Trim()

                If logicalId = "" OrElse slotId = "" Then
                    failureReason =
                        "Every expected_artifacts item requires non-empty opaque logical_deliverable_id and output_slot_id values."
                    Return False
                End If

                Dim pairKey As System.String =
                    logicalId & System.Char.ConvertFromUtf32(&H1F) & slotId

                If Not seen.Add(pairKey) Then
                    failureReason =
                        "expected_artifacts contains a duplicate logical_deliverable_id/output_slot_id pair."
                    Return False
                End If
            Next

            Return True
        End Function

        ''' <summary>
        ''' Returns True only for a successful host-generated skill payload whose frontmatter
        ''' declared a positive exact final-deliverable count. Hosts use this after registering
        ''' the invocation's expected_artifacts to lock that slot set for the remainder of the run.
        ''' </summary>
        Public Shared Function ResponseDeclaresFixedDeliverableContract(responseText As System.String) As System.Boolean
            If System.String.IsNullOrWhiteSpace(responseText) Then Return False

            Try
                Dim obj As Newtonsoft.Json.Linq.JObject =
                    Newtonsoft.Json.Linq.JObject.Parse(responseText)

                Dim token As Newtonsoft.Json.Linq.JToken =
                    obj("declared_deliverable_count")

                If token Is Nothing Then Return False

                Dim count As System.Int32 = 0
                If Not System.Int32.TryParse(
                    token.ToString(),
                    System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture,
                    count) Then

                    Return False
                End If

                Return count > 0
            Catch ex As System.Exception
                Return False
            End Try
        End Function

        Public Shared Function GetDeclaredRequiredSuccessfulToolsFromResponse(
            responseText As System.String) As System.Collections.Generic.List(Of System.String)

            Dim result As New System.Collections.Generic.List(Of System.String)()
            If System.String.IsNullOrWhiteSpace(responseText) Then Return result

            Try
                Dim obj As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(responseText)
                Dim token As Newtonsoft.Json.Linq.JToken = obj("declared_required_successful_tools")
                If token Is Nothing OrElse token.Type <> Newtonsoft.Json.Linq.JTokenType.Array Then Return result

                Dim seen As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
                For Each item As Newtonsoft.Json.Linq.JToken In DirectCast(token, Newtonsoft.Json.Linq.JArray)
                    If item Is Nothing OrElse item.Type = Newtonsoft.Json.Linq.JTokenType.Null Then Continue For
                    Dim toolName As System.String = item.ToString().Trim()
                    If toolName = System.String.Empty Then Continue For
                    If Not System.Text.RegularExpressions.Regex.IsMatch(toolName, "^[A-Za-z][A-Za-z0-9_.-]{0,127}$") Then Continue For
                    If seen.Add(toolName) Then result.Add(toolName)
                Next
            Catch ex As System.Exception
            End Try

            Return result
        End Function

        Public Shared Function GetDeclaredDeliverableRequiredEffectsFromResponse(
            responseText As System.String) As System.Collections.Generic.List(Of System.String)

            Dim result As New System.Collections.Generic.List(Of System.String)()
            If System.String.IsNullOrWhiteSpace(responseText) Then Return result

            Try
                Dim obj As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(responseText)
                Dim token As Newtonsoft.Json.Linq.JToken = obj("declared_deliverable_required_effects")
                If token Is Nothing OrElse token.Type <> Newtonsoft.Json.Linq.JTokenType.Array Then Return result

                Dim seen As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
                For Each item As Newtonsoft.Json.Linq.JToken In DirectCast(token, Newtonsoft.Json.Linq.JArray)
                    If item Is Nothing OrElse item.Type = Newtonsoft.Json.Linq.JTokenType.Null Then Continue For
                    Dim effect As System.String = item.ToString().Trim().ToLowerInvariant()
                    If effect = System.String.Empty Then Continue For
                    If Not System.Text.RegularExpressions.Regex.IsMatch(effect, "^[a-z][a-z0-9_.-]{0,63}$") Then Continue For
                    If seen.Add(effect) Then result.Add(effect)
                Next
            Catch ex As System.Exception
            End Try

            Return result
        End Function

        Private Shared Function InventoryDir(dir As String) As List(Of Object)
            Dim list As New List(Of Object)
            If String.IsNullOrWhiteSpace(dir) OrElse Not Directory.Exists(dir) Then Return list
            Try
                For Each f In Directory.EnumerateFiles(dir, "*", SearchOption.AllDirectories)
                    Try
                        Dim fi As New FileInfo(f)
                        Dim rel = f.Substring(dir.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                        list.Add(New With {
                            Key .path = rel,
                            Key .size = fi.Length
                        })
                    Catch
                    End Try
                Next
            Catch
            End Try
            Return list
        End Function

        Private Shared Function InventoryDiagnosticsDir(dir As String, prefix As String) As List(Of Object)
            Dim list As New List(Of Object)
            If String.IsNullOrWhiteSpace(dir) OrElse Not Directory.Exists(dir) Then Return list

            Try
                Dim files As New List(Of FileInfo)(
                    New DirectoryInfo(dir).EnumerateFiles(prefix & "*.txt", SearchOption.TopDirectoryOnly))

                files.Sort(Function(a, b) b.LastWriteTimeUtc.CompareTo(a.LastWriteTimeUtc))

                For Each fi In files
                    Try
                        list.Add(New With {
                            Key .path = fi.Name,
                            Key .size = fi.Length,
                            Key .last_write_utc = fi.LastWriteTimeUtc.ToString("yyyy-MM-ddTHH:mm:ssZ")
                        })
                    Catch
                    End Try
                Next
            Catch
            End Try

            Return list
        End Function

        Private Shared Function GetStr(args As IDictionary(Of String, Object), name As String) As String
            If args Is Nothing Then Return ""
            Dim v As Object = Nothing
            If Not args.TryGetValue(name, v) OrElse v Is Nothing Then Return ""
            Return System.Convert.ToString(v)
        End Function

    End Class

End Namespace
