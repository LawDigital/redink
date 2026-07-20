' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: AgentResources.vb
' Purpose: Discovers and parses Claude-style agent resources (Inky.md, Skills,
'          Agents) from two roots: INI_AgentResourcesPath (central) and
'          INI_AgentResourcesPathLocal (user-local). Local entries win over
'          central entries with the same name.
'
' Resource Types:
'  - Inky.md: Project-wide guidance appended to system prompts by InkyPromptBuilder.
'  - SkillDescriptor: Skill resources with YAML frontmatter and optional scripts/.
'  - AgentDescriptor: Agent resources with YAML frontmatter for sub-agent delegation.
'  - YAML parsing: Supports name, description, allowed-tools, model, network, timeout.
'  - Lazy loading: Bodies and scripts are loaded on demand to keep startup fast.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions

Namespace Agents

    ''' <summary>Common base for a parsed markdown resource with YAML frontmatter.</summary>
    Public MustInherit Class AgentResourceBase
        Public Property Name As String
        Public Property Description As String
        Public Property AllowedTools As New List(Of String)
        Public Property Model As String                 ' optional, e.g. "researchmodel" (special-task-model key)
        Public Property Network As Boolean = False      ' opt-in for tools that touch the network (js.run, fetch)
        Public Property TimeoutSeconds As Integer = 0   ' 0 = use default
        Public Property Frontmatter As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        Public Property FilePath As String              ' path to the .md file
        Public Property DirectoryPath As String             ' directory holding the resource
        Public Property IsLocal As Boolean              ' true if found in the local (user) tree

        Private _bodyCache As String
        Public Function LoadBody() As String
            If _bodyCache IsNot Nothing Then Return _bodyCache

            If String.IsNullOrWhiteSpace(FilePath) OrElse Not File.Exists(FilePath) Then
                _bodyCache = String.Empty
                Return _bodyCache
            End If

            Try
                _bodyCache = AgentResources.ReadBody(FilePath)
            Catch
                _bodyCache = String.Empty
            End Try

            Return _bodyCache
        End Function
    End Class

    Public Class SkillDescriptor
        Inherits AgentResourceBase
        ''' <summary>Optional path to the skill's scripts/ directory (may not exist).</summary>
        Public ReadOnly Property ScriptsDir As String
            Get
                Return System.IO.Path.Combine(If(DirectoryPath, ""), "scripts")
            End Get
        End Property
        ''' <summary>Optional path to the skill's references/ directory (may not exist).</summary>
        Public ReadOnly Property ReferencesDir As String
            Get
                Return System.IO.Path.Combine(If(DirectoryPath, ""), "references")
            End Get
        End Property
    End Class

    Public Class AgentDescriptor
        Inherits AgentResourceBase
    End Class

    ''' <summary>
    ''' Static façade for discovering Inky.md, Skills and Agents from central + local roots.
    ''' Call <see cref="Refresh"/> to rescan; results are cached.
    ''' </summary>
    Public NotInheritable Class AgentResources


        Public Shared ReadOnly Property ConfiguredLocalPath As String
            Get
                Return If(_configuredLocalPath, String.Empty)
            End Get
        End Property

        Public Shared ReadOnly Property ConfiguredCentralPath As String
            Get
                Return If(_configuredCentralPath, String.Empty)
            End Get
        End Property

        ''' <summary>Returns the skills/ and agents/ base directories under the LOCAL root.</summary>
        Public Shared Function GetLocalResourceBaseDirectories() As IReadOnlyList(Of String)
            Return BuildResourceBaseDirectories(_configuredLocalPath)
        End Function

        ''' <summary>Returns the skills/ and agents/ base directories under the CENTRAL root.</summary>
        Public Shared Function GetCentralResourceBaseDirectories() As IReadOnlyList(Of String)
            Return BuildResourceBaseDirectories(_configuredCentralPath)
        End Function

        Private Shared Function BuildResourceBaseDirectories(root As String) As IReadOnlyList(Of String)
            Dim list As New List(Of String)()
            If String.IsNullOrWhiteSpace(root) Then Return list
            list.Add(System.IO.Path.Combine(root, "skills"))
            list.Add(System.IO.Path.Combine(root, "agents"))
            Return list
        End Function

        ''' <summary>
        ''' If <paramref name="writtenPath"/> lies under a configured resource root
        ''' (skills/ or agents/, local or central), rescans the in-memory index so an
        ''' edited SKILL.md/AGENT.md (and its cached body) is picked up in the same
        ''' session. Safe to call after every write; it is a no-op for non-resource paths.
        ''' </summary>
        Public Shared Sub RefreshIfResourcePath(writtenPath As String)
            If String.IsNullOrWhiteSpace(writtenPath) Then Return

            Dim full As String
            Try
                full = System.IO.Path.GetFullPath(writtenPath)
            Catch
                Return
            End Try

            Dim baseDirs As New List(Of String)()
            baseDirs.AddRange(GetLocalResourceBaseDirectories())
            baseDirs.AddRange(GetCentralResourceBaseDirectories())

            For Each baseDir In baseDirs
                If String.IsNullOrWhiteSpace(baseDir) Then Continue For
                Dim b As String
                Try
                    b = System.IO.Path.GetFullPath(baseDir).
                        TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar)
                Catch
                    Continue For
                End Try

                If full.StartsWith(b & System.IO.Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) Then
                    Try
                        Refresh()
                    Catch
                    End Try
                    Return
                End If
            Next
        End Sub

        Private Sub New()
        End Sub

        Private Shared ReadOnly _syncRoot As New Object()
        Private Shared _skills As List(Of SkillDescriptor)
        Private Shared _agents As List(Of AgentDescriptor)
        Private Shared _inkyMd As String
        Private Shared _initialized As Boolean

        Public Shared ReadOnly Property Skills As IReadOnlyList(Of SkillDescriptor)
            Get
                EnsureInitialized()
                Return _skills
            End Get
        End Property

        Public Shared ReadOnly Property Agents As IReadOnlyList(Of AgentDescriptor)
            Get
                EnsureInitialized()
                Return _agents
            End Get
        End Property

        ''' <summary>Concatenation of central + local Inky.md (local appended after a divider).</summary>
        Public Shared ReadOnly Property InkyMd As String
            Get
                EnsureInitialized()
                Return If(_inkyMd, String.Empty)
            End Get
        End Property

        Private Shared Function NormalizeResourceLookupKey(value As String) As String
            If String.IsNullOrWhiteSpace(value) Then Return ""

            Dim sb As New StringBuilder()
            Dim lastWasUnderscore As Boolean = False

            For Each ch As Char In value.Trim()
                If Char.IsLetterOrDigit(ch) Then
                    sb.Append(Char.ToLowerInvariant(ch))
                    lastWasUnderscore = False
                ElseIf Not lastWasUnderscore Then
                    sb.Append("_"c)
                    lastWasUnderscore = True
                End If
            Next

            Return sb.ToString().Trim("_"c)
        End Function

        Public Shared Function FindSkill(name As String) As SkillDescriptor
            If String.IsNullOrWhiteSpace(name) Then Return Nothing
            EnsureInitialized()

            Dim trimmedName As String = name.Trim()

            Dim exact = _skills.FirstOrDefault(
                Function(s)
                    Return s IsNot Nothing AndAlso
                           String.Equals(s.Name, trimmedName, StringComparison.OrdinalIgnoreCase)
                End Function)

            If exact IsNot Nothing Then
                Return exact
            End If

            Dim normalizedName As String = NormalizeResourceLookupKey(trimmedName)
            If normalizedName = "" Then Return Nothing

            Return _skills.FirstOrDefault(
                Function(s)
                    Return s IsNot Nothing AndAlso
                           NormalizeResourceLookupKey(s.Name) = normalizedName
                End Function)
        End Function

        Public Shared Function FindAgent(name As String) As AgentDescriptor
            If String.IsNullOrWhiteSpace(name) Then Return Nothing
            EnsureInitialized()

            Dim trimmedName As String = name.Trim()

            Dim exact = _agents.FirstOrDefault(
                Function(a)
                    Return a IsNot Nothing AndAlso
                           String.Equals(a.Name, trimmedName, StringComparison.OrdinalIgnoreCase)
                End Function)

            If exact IsNot Nothing Then
                Return exact
            End If

            Dim normalizedName As String = NormalizeResourceLookupKey(trimmedName)
            If normalizedName = "" Then Return Nothing

            Return _agents.FirstOrDefault(
                Function(a)
                    Return a IsNot Nothing AndAlso
                           NormalizeResourceLookupKey(a.Name) = normalizedName
                End Function)
        End Function

        Private Shared Sub EnsureInitialized()
            If _initialized Then Return
            Refresh()
        End Sub

        ''' <summary>Rescans both roots and refreshes the in-memory index.</summary>
        Private Shared _configuredCentralPath As String
        Private Shared _configuredLocalPath As String

        Public Shared Sub SetPaths(centralPath As String, localPath As String)
            SyncLock _syncRoot
                _configuredCentralPath = SharedLibrary.SharedMethods.ExpandEnvironmentVariables(centralPath)
                _configuredLocalPath = SharedLibrary.SharedMethods.ExpandEnvironmentVariables(localPath)
                _initialized = False

            End SyncLock
        End Sub

        ''' <summary>Rescans both roots and refreshes the in-memory index.</summary>
        Public Shared Sub Refresh()
            SyncLock _syncRoot
                Dim central = _configuredCentralPath
                Dim localPath = _configuredLocalPath

                Dim skillsCentral = ScanSkills(central, isLocal:=False)
                Dim skillsLocal = ScanSkills(localPath, isLocal:=True)
                _skills = MergeByName(skillsCentral, skillsLocal)

                Dim agentsCentral = ScanAgents(central, isLocal:=False)
                Dim agentsLocal = ScanAgents(localPath, isLocal:=True)
                _agents = MergeByName(agentsCentral, agentsLocal)

                _inkyMd = ReadInkyMd(central, localPath)
                _initialized = True

            End SyncLock
        End Sub

        ' ------------------------------------------------------------------ scanning

        Private Shared Function ScanSkills(root As String, isLocal As Boolean) As List(Of SkillDescriptor)
            Dim list As New List(Of SkillDescriptor)
            If String.IsNullOrWhiteSpace(root) Then Return list
            Dim skillsDir = Path.Combine(root, "skills")
            If Not System.IO.Directory.Exists(skillsDir) Then Return list

            For Each subDir In System.IO.Directory.EnumerateDirectories(skillsDir)
                Dim md = FindMarkdownFile(subDir, {"SKILL.md", "skill.md"})
                If md Is Nothing Then Continue For
                Try
                    Dim sk As New SkillDescriptor()
                    PopulateFromMarkdown(sk, md, isLocal)
                    If String.IsNullOrWhiteSpace(sk.Name) Then sk.Name = Path.GetFileName(subDir)
                    sk.DirectoryPath = subDir
                    list.Add(sk)
                Catch
                    ' ignore broken entries
                End Try
            Next
            Return list
        End Function

        Private Shared Function ScanAgents(root As String, isLocal As Boolean) As List(Of AgentDescriptor)
            Dim list As New List(Of AgentDescriptor)
            If String.IsNullOrWhiteSpace(root) Then Return list
            Dim agentsDir = Path.Combine(root, "agents")
            If Not System.IO.Directory.Exists(agentsDir) Then Return list

            ' (a) agents/<name>.md
            For Each f In System.IO.Directory.EnumerateFiles(agentsDir, "*.md", SearchOption.TopDirectoryOnly)
                Try
                    Dim ag As New AgentDescriptor()
                    PopulateFromMarkdown(ag, f, isLocal)
                    If String.IsNullOrWhiteSpace(ag.Name) Then ag.Name = Path.GetFileNameWithoutExtension(f)
                    ag.DirectoryPath = Path.GetDirectoryName(f)
                    list.Add(ag)
                Catch
                End Try
            Next

            ' (b) agents/<name>/AGENT.md
            For Each subDir In System.IO.Directory.EnumerateDirectories(agentsDir)
                Dim md = FindMarkdownFile(subDir, {"AGENT.md", "agent.md"})
                If md Is Nothing Then Continue For
                Try
                    Dim ag As New AgentDescriptor()
                    PopulateFromMarkdown(ag, md, isLocal)
                    If String.IsNullOrWhiteSpace(ag.Name) Then ag.Name = Path.GetFileName(subDir)
                    ag.DirectoryPath = subDir
                    list.Add(ag)
                Catch
                End Try
            Next
            Return list
        End Function

        Private Shared Function ReadInkyMd(central As String, localPath As String) As String
            Dim sb As New StringBuilder()
            Dim c = TryReadInkyMd(central)
            Dim l = TryReadInkyMd(localPath)
            If Not String.IsNullOrEmpty(c) Then
                sb.AppendLine(c.TrimEnd())
            End If
            If Not String.IsNullOrEmpty(l) Then
                If sb.Length > 0 Then
                    sb.AppendLine()
                    sb.AppendLine("<!-- ----- local Inky.md overrides ----- -->")
                    sb.AppendLine()
                End If
                sb.AppendLine(l.TrimEnd())
            End If
            Return sb.ToString()
        End Function

        Private Shared Function TryReadInkyMd(root As String) As String
            If String.IsNullOrWhiteSpace(root) OrElse Not System.IO.Directory.Exists(root) Then Return Nothing
            For Each candidate In {"Inky.md", "INKY.md", "inky.md"}
                Dim p = Path.Combine(root, candidate)
                If File.Exists(p) Then
                    Try
                        Return File.ReadAllText(p, Encoding.UTF8)
                    Catch
                        Return Nothing
                    End Try
                End If
            Next
            Return Nothing
        End Function

        Private Shared Function FindMarkdownFile(dir As String, names As String()) As String
            For Each n In names
                Dim p = Path.Combine(dir, n)
                If File.Exists(p) Then Return p
            Next
            Return Nothing
        End Function

        ' Local entries override central by name (case-insensitive).
        Private Shared Function MergeByName(Of T As AgentResourceBase)(central As List(Of T), localList As List(Of T)) As List(Of T)
            Dim merged As New Dictionary(Of String, T)(StringComparer.OrdinalIgnoreCase)
            For Each c In central
                If Not String.IsNullOrWhiteSpace(c.Name) Then merged(c.Name) = c
            Next
            For Each l In localList
                If Not String.IsNullOrWhiteSpace(l.Name) Then merged(l.Name) = l
            Next
            Return merged.Values.OrderBy(Function(x) x.Name, StringComparer.OrdinalIgnoreCase).ToList()
        End Function

        ' ------------------------------------------------------------------ markdown parsing

        Private Shared ReadOnly _frontmatterRegex As New Regex(
            "^\uFEFF?---\s*\r?\n(?<yaml>.*?)\r?\n---\s*\r?\n?",
            RegexOptions.Singleline Or RegexOptions.CultureInvariant)

        ''' <summary>Reads the body (markdown after the frontmatter, or whole file if none).</summary>
        Friend Shared Function ReadBody(mdPath As String) As String
            Dim text = File.ReadAllText(mdPath, Encoding.UTF8)
            Dim m = _frontmatterRegex.Match(text)
            If m.Success Then Return text.Substring(m.Length)
            Return text
        End Function

        Private Shared Sub PopulateFromMarkdown(target As AgentResourceBase, mdPath As String, isLocal As Boolean)
            target.FilePath = mdPath
            target.IsLocal = isLocal
            target.DirectoryPath = System.IO.Path.GetDirectoryName(mdPath)

            Dim text = File.ReadAllText(mdPath, Encoding.UTF8)
            Dim m = _frontmatterRegex.Match(text)
            If Not m.Success Then Return

            Dim fm = ParseSimpleYaml(m.Groups("yaml").Value)
            target.Frontmatter = fm

            Dim v As String = Nothing
            If fm.TryGetValue("name", v) Then target.Name = v
            If fm.TryGetValue("description", v) Then target.Description = v
            If fm.TryGetValue("model", v) Then target.Model = v
            If fm.TryGetValue("network", v) Then target.Network = ParseBool(v)
            If fm.TryGetValue("timeout", v) Then
                Dim n As Integer
                If Integer.TryParse(v, n) Then target.TimeoutSeconds = n
            End If
            If fm.TryGetValue("allowed-tools", v) Then target.AllowedTools = ParseList(v)
        End Sub

        ''' <summary>
        ''' Minimal YAML parser sufficient for Claude-style frontmatter:
        '''   key: value
        '''   key: [a, b, c]
        '''   key:
        '''     - a
        '''     - b
        ''' Quoted strings ("..." or '...') are unquoted. Unknown structures are kept verbatim.
        ''' </summary>
        Private Shared Function ParseSimpleYaml(yaml As String) As Dictionary(Of String, String)
            Dim result As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            If String.IsNullOrWhiteSpace(yaml) Then Return result

            Dim lines = yaml.Replace(vbCr, "").Split(ChrW(10))
            Dim i As Integer = 0
            While i < lines.Length
                Dim line = lines(i)
                Dim trimmed = line.TrimEnd()
                If String.IsNullOrWhiteSpace(trimmed) OrElse trimmed.TrimStart().StartsWith("#") Then
                    i += 1
                    Continue While
                End If

                Dim colonIdx = trimmed.IndexOf(":"c)
                If colonIdx <= 0 Then
                    i += 1
                    Continue While
                End If

                Dim key = trimmed.Substring(0, colonIdx).Trim()
                Dim valuePart = trimmed.Substring(colonIdx + 1).Trim()

                If valuePart.Length = 0 Then
                    ' Possibly a block list on following indented lines.
                    Dim items As New List(Of String)
                    Dim j = i + 1
                    While j < lines.Length
                        Dim next_ = lines(j)
                        If String.IsNullOrWhiteSpace(next_) Then
                            j += 1
                            Continue While
                        End If
                        Dim ltrim = next_.TrimStart()
                        If ltrim.StartsWith("- ") OrElse ltrim = "-" Then
                            items.Add(Unquote(ltrim.Substring(1).Trim()))
                            j += 1
                        Else
                            Exit While
                        End If
                    End While
                    If items.Count > 0 Then
                        result(key) = String.Join(",", items)
                        i = j
                        Continue While
                    Else
                        result(key) = ""
                        i += 1
                        Continue While
                    End If
                End If

                result(key) = Unquote(valuePart)
                i += 1
            End While
            Return result
        End Function

        Private Shared Function Unquote(s As String) As String
            If s Is Nothing Then Return Nothing
            s = s.Trim()
            If s.Length >= 2 Then
                Dim first = s(0), last = s(s.Length - 1)
                If (first = """"c AndAlso last = """"c) OrElse (first = "'"c AndAlso last = "'"c) Then
                    Return s.Substring(1, s.Length - 2)
                End If
            End If
            Return s
        End Function

        Private Shared Function ParseList(s As String) As List(Of String)
            Dim list As New List(Of String)
            If String.IsNullOrWhiteSpace(s) Then Return list
            Dim t = s.Trim()
            If t.StartsWith("[") AndAlso t.EndsWith("]") Then
                t = t.Substring(1, t.Length - 2)
            End If
            For Each part In t.Split(","c)
                Dim p = Unquote(part.Trim())
                If Not String.IsNullOrWhiteSpace(p) Then list.Add(p)
            Next
            Return list
        End Function

        Private Shared Function ParseBool(s As String) As Boolean
            If String.IsNullOrWhiteSpace(s) Then Return False
            Select Case s.Trim().ToLowerInvariant()
                Case "true", "yes", "on", "1" : Return True
                Case Else : Return False
            End Select
        End Function


    End Class

End Namespace
