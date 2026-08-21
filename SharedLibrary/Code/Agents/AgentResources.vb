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
'  - YAML parsing: Supports name, description, allowed-tools, optional-tools, model, network, timeout.
'  - Lazy loading: Bodies and scripts are loaded on demand to keep startup fast.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json.Linq

Namespace Agents

    ''' <summary>Common base for a parsed markdown resource with YAML frontmatter.</summary>
    Public MustInherit Class AgentResourceBase
        Public Property Name As String
        Public Property Description As String
        Public Property AllowedTools As New List(Of String)
        Public Property OptionalTools As New List(Of String)
        Public Property Model As String                 ' optional, e.g. "researchmodel" (special-task-model key)
        Public Property Network As Boolean = False      ' opt-in for tools that touch the network (js.run, fetch)
        Public Property TimeoutSeconds As Integer = 0   ' 0 = use default
        Public Property Enabled As Boolean = True       ' opt-out: frontmatter "enabled: false" hides it from the model
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
        ''' Ensures the configured CENTRAL resource root and its skills/, agents/ and designs/
        ''' subdirectories exist so that new skills and agents can be created there when
        ''' central writing is explicitly permitted. Mirrors
        ''' <see cref="EnsureLocalResourceDirectories"/>. Best-effort.
        ''' </summary>
        Public Shared Function EnsureCentralResourceDirectories() As Boolean
            Dim root As String
            SyncLock _syncRoot
                root = _configuredCentralPath
            End SyncLock

            If String.IsNullOrWhiteSpace(root) Then Return False

            Dim createdSomething As Boolean = False
            Try
                For Each resourceDir In New String() {
                    root,
                    System.IO.Path.Combine(root, "skills"),
                    System.IO.Path.Combine(root, "agents"),
                    System.IO.Path.Combine(root, DesignRepository.DesignsDirectoryName)}

                    If Not System.IO.Directory.Exists(resourceDir) Then
                        System.IO.Directory.CreateDirectory(resourceDir)
                        createdSomething = True
                    End If
                Next
            Catch
                Return System.IO.Directory.Exists(root)
            End Try

            If createdSomething Then
                Refresh()
            End If

            Return System.IO.Directory.Exists(root)
        End Function

        ''' <summary>
        ''' Ensures the configured LOCAL resource root and its skills/, agents/ and designs/
        ''' subdirectories exist so that new skills, agents, and Inky.md can be created
        ''' even when the .inky tree was never set up. Best-effort: returns True when the
        ''' tree is present after the call, False when no local path is configured or
        ''' creation failed. Rescans the index when it created anything.
        ''' </summary>
        Public Shared Function EnsureLocalResourceDirectories() As Boolean
            Dim root As String
            SyncLock _syncRoot
                root = _configuredLocalPath
            End SyncLock

            If String.IsNullOrWhiteSpace(root) Then Return False

            Dim createdSomething As Boolean = False
            Try
                For Each resourceDir In New String() {
                    root,
                    System.IO.Path.Combine(root, "skills"),
                    System.IO.Path.Combine(root, "agents"),
                    System.IO.Path.Combine(root, DesignRepository.DesignsDirectoryName)}

                    If Not System.IO.Directory.Exists(resourceDir) Then
                        System.IO.Directory.CreateDirectory(resourceDir)
                        createdSomething = True
                    End If
                Next
            Catch
                Return System.IO.Directory.Exists(root)
            End Try

            If createdSomething Then
                Refresh()
            End If

            Return System.IO.Directory.Exists(root)
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
        Private Shared _resourceWatchers As New List(Of System.IO.FileSystemWatcher)()
        Private Shared _refreshGeneration As Long

        ''' <summary>Skills offered to the model — excludes entries with frontmatter "enabled: false".</summary>
        Public Shared ReadOnly Property Skills As IReadOnlyList(Of SkillDescriptor)
            Get
                EnsureInitialized()
                Return _skills.Where(Function(s) s IsNot Nothing AndAlso s.Enabled).ToList()
            End Get
        End Property

        ''' <summary>Agents offered to the model — excludes entries with frontmatter "enabled: false".</summary>
        Public Shared ReadOnly Property Agents As IReadOnlyList(Of AgentDescriptor)
            Get
                EnsureInitialized()
                Return _agents.Where(Function(a) a IsNot Nothing AndAlso a.Enabled).ToList()
            End Get
        End Property

        ''' <summary>All discovered skills including disabled ones — for management/inspection only.</summary>
        Public Shared ReadOnly Property AllSkills As IReadOnlyList(Of SkillDescriptor)
            Get
                EnsureInitialized()
                Return _skills
            End Get
        End Property

        ''' <summary>All discovered agents including disabled ones — for management/inspection only.</summary>
        Public Shared ReadOnly Property AllAgents As IReadOnlyList(Of AgentDescriptor)
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

        ''' <summary>
        ''' Ensures that the in-memory resource index exists without forcing a disk rescan
        ''' when the configured roots have not changed and no watcher has marked them dirty.
        ''' </summary>
        Public Shared Sub EnsureFresh()
            EnsureInitialized()
        End Sub

        Public Shared ReadOnly Property RefreshGeneration As Long
            Get
                SyncLock _syncRoot
                    Return _refreshGeneration
                End SyncLock
            End Get
        End Property

        Private Shared _configuredCentralPath As String
        Private Shared _configuredLocalPath As String

        Private Shared Function NormalizeConfiguredResourcePath(value As String) As String
            Dim expanded As String = SharedLibrary.SharedMethods.ExpandEnvironmentVariables(If(value, ""))
            If String.IsNullOrWhiteSpace(expanded) Then Return ""

            Try
                Return System.IO.Path.GetFullPath(expanded).
                    TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar)
            Catch
                Return expanded.Trim().
                    TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar)
            End Try
        End Function

        Private Shared Sub DisposeResourceWatchersUnsafe()
            For Each watcher As System.IO.FileSystemWatcher In _resourceWatchers
                If watcher Is Nothing Then Continue For
                Try
                    watcher.EnableRaisingEvents = False
                    watcher.Dispose()
                Catch
                End Try
            Next
            _resourceWatchers.Clear()
        End Sub

        Private Shared Sub MarkResourceIndexDirty(sender As Object, e As System.IO.FileSystemEventArgs)
            SyncLock _syncRoot
                _initialized = False
            End SyncLock
        End Sub

        Private Shared Sub MarkResourceIndexDirtyOnRename(sender As Object, e As System.IO.RenamedEventArgs)
            SyncLock _syncRoot
                _initialized = False
            End SyncLock
        End Sub

        Private Shared Sub AddResourceWatcherUnsafe(root As String)
            If String.IsNullOrWhiteSpace(root) OrElse Not System.IO.Directory.Exists(root) Then Return

            Try
                Dim watcher As New System.IO.FileSystemWatcher(root)
                watcher.IncludeSubdirectories = True
                watcher.NotifyFilter =
                    System.IO.NotifyFilters.FileName Or
                    System.IO.NotifyFilters.DirectoryName Or
                    System.IO.NotifyFilters.LastWrite Or
                    System.IO.NotifyFilters.CreationTime

                AddHandler watcher.Changed, AddressOf MarkResourceIndexDirty
                AddHandler watcher.Created, AddressOf MarkResourceIndexDirty
                AddHandler watcher.Deleted, AddressOf MarkResourceIndexDirty
                AddHandler watcher.Renamed, AddressOf MarkResourceIndexDirtyOnRename
                watcher.EnableRaisingEvents = True
                _resourceWatchers.Add(watcher)
            Catch
                ' Watchers are an optimization/freshness aid only. Explicit resource writes
                ' still call RefreshIfResourcePath, and manual refresh remains available.
            End Try
        End Sub

        Private Shared Sub RebuildResourceWatchersUnsafe()
            DisposeResourceWatchersUnsafe()

            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each root As String In New String() {_configuredCentralPath, _configuredLocalPath}
                If String.IsNullOrWhiteSpace(root) OrElse seen.Contains(root) Then Continue For
                seen.Add(root)
                AddResourceWatcherUnsafe(root)
            Next
        End Sub

        Public Shared Sub SetPaths(centralPath As String, localPath As String)
            Dim normalizedCentral As String = NormalizeConfiguredResourcePath(centralPath)
            Dim normalizedLocal As String = NormalizeConfiguredResourcePath(localPath)

            SyncLock _syncRoot
                If String.Equals(_configuredCentralPath, normalizedCentral, StringComparison.OrdinalIgnoreCase) AndAlso
                   String.Equals(_configuredLocalPath, normalizedLocal, StringComparison.OrdinalIgnoreCase) Then
                    Return
                End If

                _configuredCentralPath = normalizedCentral
                _configuredLocalPath = normalizedLocal
                _initialized = False
                RebuildResourceWatchersUnsafe()
            End SyncLock
        End Sub

        ''' <summary>Rescans both roots and refreshes the in-memory index.</summary>
        Public Shared Sub Refresh()
            Dim stopwatch As System.Diagnostics.Stopwatch = System.Diagnostics.Stopwatch.StartNew()

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
                _refreshGeneration += 1
            End SyncLock

            stopwatch.Stop()
            System.Diagnostics.Debug.WriteLine(
                "[PERF] AgentResources.Refresh: " &
                stopwatch.ElapsedMilliseconds.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                " ms; generation=" &
                RefreshGeneration.ToString(System.Globalization.CultureInfo.InvariantCulture))
        End Sub

        ' ------------------------------------------------------------------ scanning

        Private Shared Function ScanSkills(root As String, isLocal As Boolean) As List(Of SkillDescriptor)
            Dim list As New List(Of SkillDescriptor)
            If String.IsNullOrWhiteSpace(root) Then Return list
            Dim skillsDir = Path.Combine(root, "skills")
            If Not System.IO.Directory.Exists(skillsDir) Then Return list

            For Each subDir In System.IO.Directory.EnumerateDirectories(skillsDir)
                Dim md = ResolveResourceMarkdown(subDir, {"SKILL.md", "skill.md"})
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

        ''' <summary>
        ''' Resolves the markdown descriptor inside a resource folder in a directory-name
        ''' agnostic way: first tries the canonical names (e.g. SKILL.md/AGENT.md), then
        ''' a file named after the folder (&lt;dirname&gt;.md), and finally falls back to the
        ''' single .md file when the folder contains exactly one.
        ''' </summary>
        Private Shared Function ResolveResourceMarkdown(dir As String, canonicalNames As String()) As String
            Dim canonical = FindMarkdownFile(dir, canonicalNames)
            If canonical IsNot Nothing Then Return canonical

            Dim named = Path.Combine(dir, Path.GetFileName(dir) & ".md")
            If File.Exists(named) Then Return named

            Try
                Dim mdFiles = System.IO.Directory.EnumerateFiles(dir, "*.md", SearchOption.TopDirectoryOnly).ToList()
                If mdFiles.Count = 1 Then Return mdFiles(0)
            Catch
            End Try

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
            If fm.TryGetValue("enabled", v) Then target.Enabled = ParseBool(v)
            If fm.TryGetValue("timeout", v) Then
                Dim n As Integer
                If Integer.TryParse(v, n) Then target.TimeoutSeconds = n
            End If
            If fm.TryGetValue("allowed-tools", v) Then target.AllowedTools = ParseList(v)
            If fm.TryGetValue("optional-tools", v) Then target.OptionalTools = ParseList(v)
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

    ''' <summary>
    ''' One named document-design profile loaded from AgentResourcesPath[/Local]\designs\designs.json.
    ''' The JSON profile is authoritative. Optional Office template files are resolved only
    ''' relative to the containing designs directory and can therefore never escape the
    ''' configured agent-resource tree.
    ''' </summary>
    Public Class DocumentDesignDescriptor
        Public Property Id As String
        Public Property Name As String
        Public Property Description As String
        Public Property Aliases As New List(Of String)()
        Public Property Enabled As Boolean = True
        Public Property IsLocal As Boolean
        Public Property CatalogPath As String
        Public Property DirectoryPath As String
        Public Property Raw As JObject

        Public ReadOnly Property Word As JObject
            Get
                Return TryCast(If(Raw Is Nothing, Nothing, Raw("word")), JObject)
            End Get
        End Property

        Public ReadOnly Property PowerPoint As JObject
            Get
                Return TryCast(If(Raw Is Nothing, Nothing, Raw("powerpoint")), JObject)
            End Get
        End Property

        Public ReadOnly Property Excel As JObject
            Get
                Return TryCast(If(Raw Is Nothing, Nothing, Raw("excel")), JObject)
            End Get
        End Property

        Public Function GetApplicationConfig(applicationName As String) As JObject
            Select Case NormalizeApplicationName(applicationName)
                Case "word" : Return Word
                Case "powerpoint" : Return PowerPoint
                Case "excel" : Return Excel
                Case Else : Return Nothing
            End Select
        End Function

        Public Function SupportsApplication(applicationName As String) As Boolean
            Return GetApplicationConfig(applicationName) IsNot Nothing
        End Function

        Public Function ResolveRepositoryFile(relativePath As String) As String
            If String.IsNullOrWhiteSpace(relativePath) OrElse String.IsNullOrWhiteSpace(DirectoryPath) Then Return ""
            If System.IO.Path.IsPathRooted(relativePath) Then Return ""

            Try
                Dim root As String = System.IO.Path.GetFullPath(DirectoryPath).
                    TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar)
                Dim candidate As String = System.IO.Path.GetFullPath(System.IO.Path.Combine(root, relativePath.Trim())).
                    TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar)

                If Not candidate.StartsWith(root & System.IO.Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) Then
                    Return ""
                End If

                Return candidate
            Catch ex As System.Exception
                Return ""
            End Try
        End Function

        Private Shared Function NormalizeApplicationName(value As String) As String
            Dim normalized As String = If(value, "").Trim().ToLowerInvariant()
            Select Case normalized
                Case "ppt", "pptx", "power point", "power-point" : Return "powerpoint"
                Case "doc", "docx" : Return "word"
                Case "xls", "xlsx", "xlsm" : Return "excel"
                Case Else : Return normalized
            End Select
        End Function
    End Class

    ''' <summary>
    ''' Shared, read-only resolver for named Office designs. The CENTRAL catalog is loaded
    ''' first and the LOCAL catalog then overrides a central entry with the same design id.
    ''' No model-side filesystem discovery is required: hosts can expose the available names
    ''' through <see cref="BuildPromptFragment"/> and creator tools resolve the same profile
    ''' deterministically by <c>design_name</c>.
    ''' </summary>
    Public NotInheritable Class DesignRepository
        Public Const DesignsDirectoryName As String = "designs"
        Public Const CatalogFileName As String = "designs.json"
        Public Const SupportedSchemaVersion As Integer = 1

        Private Sub New()
        End Sub

        Public Shared Function GetDesigns() As IReadOnlyList(Of DocumentDesignDescriptor)
            Dim merged As New Dictionary(Of String, DocumentDesignDescriptor)(StringComparer.OrdinalIgnoreCase)

            ' Approved standalone Office template carriers placed in the conventional
            ' designs\word, designs\powerpoint, or designs\excel directories are
            ' exposed as implicit design profiles. An explicit designs.json entry with
            ' the same id always wins. Local resources override central resources.
            LoadLooseTemplateCarriersInto(AgentResources.ConfiguredCentralPath, isLocal:=False, merged:=merged)
            LoadCatalogInto(AgentResources.ConfiguredCentralPath, isLocal:=False, merged:=merged)
            LoadLooseTemplateCarriersInto(AgentResources.ConfiguredLocalPath, isLocal:=True, merged:=merged)
            LoadCatalogInto(AgentResources.ConfiguredLocalPath, isLocal:=True, merged:=merged)

            Return merged.Values.
                Where(Function(d) d IsNot Nothing AndAlso d.Enabled).
                OrderBy(Function(d) If(d.Name, d.Id), StringComparer.OrdinalIgnoreCase).
                ToList()
        End Function

        Public Shared Function FindDesign(requestedName As String) As DocumentDesignDescriptor
            If String.IsNullOrWhiteSpace(requestedName) Then Return Nothing
            Dim wanted As String = NormalizeLookupKey(requestedName)
            If wanted = "" Then Return Nothing

            Dim designs As IReadOnlyList(Of DocumentDesignDescriptor) = GetDesigns()

            For Each d As DocumentDesignDescriptor In designs
                If d Is Nothing Then Continue For
                If NormalizeLookupKey(d.Id) = wanted OrElse NormalizeLookupKey(d.Name) = wanted Then Return d
            Next

            For Each d As DocumentDesignDescriptor In designs
                If d Is Nothing OrElse d.Aliases Is Nothing Then Continue For
                For Each aliasName As String In d.Aliases
                    If NormalizeLookupKey(aliasName) = wanted Then Return d
                Next
            Next

            Return Nothing
        End Function

        Public Shared Function BuildPromptFragment(Optional maxDesigns As Integer = 24) As String
            Dim designs As IReadOnlyList(Of DocumentDesignDescriptor) = GetDesigns()
            If designs Is Nothing OrElse designs.Count = 0 Then
                Return "DESIGN REPOSITORY: No named Office design profiles or approved standalone template carriers are currently configured under AgentResourcesPath/AgentResourcesPathLocal\\designs. Named corporate designs must therefore not be claimed from model knowledge; use neutral professional design unless another concrete authorized source is available."
            End If

            Dim items As New List(Of String)()
            For Each d As DocumentDesignDescriptor In designs.Take(Math.Max(1, maxDesigns))
                Dim apps As New List(Of String)()
                If d.Word IsNot Nothing Then apps.Add("Word")
                If d.PowerPoint IsNot Nothing Then apps.Add("PowerPoint")
                If d.Excel IsNot Nothing Then apps.Add("Excel")
                Dim source As String = If(d.IsLocal, "local", "central")
                items.Add($"{d.Name} [design_name={d.Id}; {String.Join("/", apps)}; {source}]")
            Next

            Dim suffix As String = If(designs.Count > items.Count, $" (+{designs.Count - items.Count} more)", "")
            Return "DESIGN REPOSITORY: Configured named designs: " & String.Join("; ", items) & suffix & ". " &
                   "When the user asks for one of these designs, pass its exact design_name to create_word_document, create_powerpoint, or create_excel_spreadsheet. " &
                   "Do not substitute a similarly named organization and do not claim any design not backed by this repository or another concrete authorized source."
        End Function

        Public Shared Function GetCatalogPaths() As IReadOnlyList(Of String)
            Dim result As New List(Of String)()
            AddCatalogPath(result, AgentResources.ConfiguredCentralPath)
            AddCatalogPath(result, AgentResources.ConfiguredLocalPath)
            Return result
        End Function

        Private Shared Sub AddCatalogPath(result As List(Of String), root As String)
            If result Is Nothing OrElse String.IsNullOrWhiteSpace(root) Then Return
            Try
                result.Add(System.IO.Path.Combine(root, DesignsDirectoryName, CatalogFileName))
            Catch ex As System.Exception
            End Try
        End Sub

        Private Shared Sub LoadCatalogInto(root As String,
                                           isLocal As Boolean,
                                           merged As Dictionary(Of String, DocumentDesignDescriptor))
            If merged Is Nothing OrElse String.IsNullOrWhiteSpace(root) Then Return

            Dim designsDir As String
            Dim catalogPath As String
            Try
                designsDir = System.IO.Path.GetFullPath(System.IO.Path.Combine(root, DesignsDirectoryName))
                catalogPath = System.IO.Path.Combine(designsDir, CatalogFileName)
            Catch ex As System.Exception
                Return
            End Try

            If Not System.IO.File.Exists(catalogPath) Then Return

            Try
                Dim rawText As String = System.IO.File.ReadAllText(catalogPath, System.Text.Encoding.UTF8)
                Dim catalog As JObject = JObject.Parse(rawText)
                Dim schemaVersion As Integer = 0
                Dim schemaToken As JToken = catalog("schema_version")
                If schemaToken IsNot Nothing Then Integer.TryParse(schemaToken.ToString(), schemaVersion)
                If schemaVersion <> SupportedSchemaVersion Then Return

                Dim entries As JArray = TryCast(catalog("designs"), JArray)
                If entries Is Nothing Then Return

                For Each obj As JObject In entries.OfType(Of JObject)()
                    Dim enabled As Boolean = True
                    Dim enabledToken As JToken = obj("enabled")
                    If enabledToken IsNot Nothing AndAlso enabledToken.Type = JTokenType.Boolean Then enabled = enabledToken.Value(Of Boolean)()
                    If Not enabled Then Continue For

                    Dim id As String = If(obj.Value(Of String)("id"), "").Trim()
                    Dim name As String = If(obj.Value(Of String)("name"), "").Trim()
                    If id = "" Then id = name
                    If name = "" Then name = id
                    Dim mergeKey As String = NormalizeLookupKey(id)
                    If mergeKey = "" Then Continue For

                    Dim descriptor As New DocumentDesignDescriptor() With {
                        .Id = id,
                        .Name = name,
                        .Description = If(obj.Value(Of String)("description"), ""),
                        .Enabled = enabled,
                        .IsLocal = isLocal,
                        .CatalogPath = catalogPath,
                        .DirectoryPath = designsDir,
                        .Raw = DirectCast(obj.DeepClone(), JObject)
                    }

                    Dim aliases As JArray = TryCast(obj("aliases"), JArray)
                    If aliases IsNot Nothing Then
                        For Each token As JToken In aliases
                            If token Is Nothing OrElse token.Type <> JTokenType.String Then Continue For
                            Dim aliasName As String = token.ToString().Trim()
                            If aliasName <> "" Then descriptor.Aliases.Add(aliasName)
                        Next
                    End If

                    ' Local entries intentionally replace central entries with the same stable id.
                    merged(mergeKey) = descriptor
                Next
            Catch ex As System.Exception
                ' A malformed optional design catalog must not take down the agent-resource system.
                ' Creator tools simply fall back to neutral design and can surface that the design
                ' could not be resolved.
            End Try
        End Sub

        Private Shared Sub LoadLooseTemplateCarriersInto(root As String,
                                                           isLocal As Boolean,
                                                           merged As Dictionary(Of String, DocumentDesignDescriptor))
            If merged Is Nothing OrElse String.IsNullOrWhiteSpace(root) Then Return

            Dim designsDir As String
            Try
                designsDir = System.IO.Path.GetFullPath(
                    System.IO.Path.Combine(root, DesignsDirectoryName)
                )
            Catch ex As System.Exception
                Return
            End Try

            AddLooseTemplateApplication(
                designsDir,
                isLocal,
                "word",
                New String() {".dotx", ".dotm", ".docx"},
                merged
            )

            AddLooseTemplateApplication(
                designsDir,
                isLocal,
                "powerpoint",
                New String() {".potx", ".pptx"},
                merged
            )

            AddLooseTemplateApplication(
                designsDir,
                isLocal,
                "excel",
                New String() {".xltx"},
                merged
            )
        End Sub

        Private Shared Sub AddLooseTemplateApplication(designsDir As String,
                                                       isLocal As Boolean,
                                                       applicationName As String,
                                                       allowedExtensions As IEnumerable(Of String),
                                                       merged As Dictionary(Of String, DocumentDesignDescriptor))
            If String.IsNullOrWhiteSpace(designsDir) OrElse
               String.IsNullOrWhiteSpace(applicationName) OrElse
               allowedExtensions Is Nothing OrElse
               merged Is Nothing Then

                Return
            End If

            Dim applicationDir As String
            Try
                applicationDir = System.IO.Path.Combine(designsDir, applicationName)
            Catch ex As System.Exception
                Return
            End Try

            If Not System.IO.Directory.Exists(applicationDir) Then Return

            Dim allowed As New HashSet(Of String)(
                allowedExtensions,
                StringComparer.OrdinalIgnoreCase
            )

            Dim files As IEnumerable(Of String)
            Try
                files = System.IO.Directory.EnumerateFiles(
                    applicationDir,
                    "*.*",
                    System.IO.SearchOption.TopDirectoryOnly
                )
            Catch ex As System.Exception
                Return
            End Try

            For Each filePath As String In files
                Dim extension As String = System.IO.Path.GetExtension(filePath)
                If Not allowed.Contains(extension) Then Continue For

                Dim stem As String = System.IO.Path.GetFileNameWithoutExtension(filePath)
                Dim mergeKey As String = NormalizeLookupKey(stem)
                If mergeKey = "" Then Continue For

                Dim descriptor As DocumentDesignDescriptor = Nothing
                If merged.ContainsKey(mergeKey) Then
                    Dim existing As DocumentDesignDescriptor = merged(mergeKey)
                    If existing IsNot Nothing AndAlso
                       existing.IsLocal = isLocal AndAlso
                       String.IsNullOrWhiteSpace(existing.CatalogPath) Then

                        descriptor = existing
                    End If
                End If

                If descriptor Is Nothing Then
                    Dim displayName As String = stem.Replace("_"c, " "c).Replace("-"c, " "c).Trim()
                    If displayName = "" Then displayName = stem

                    descriptor = New DocumentDesignDescriptor() With {
                        .Id = stem,
                        .Name = displayName,
                        .Description = "Approved standalone Office template carrier discovered from the design repository.",
                        .Enabled = True,
                        .IsLocal = isLocal,
                        .CatalogPath = "",
                        .DirectoryPath = designsDir,
                        .Raw = New JObject()
                    }
                    descriptor.Aliases.Add(stem)
                    descriptor.Aliases.Add(System.IO.Path.GetFileName(filePath))
                    merged(mergeKey) = descriptor
                End If

                Dim relativePath As String =
                    applicationName & "/" & System.IO.Path.GetFileName(filePath)

                Dim appConfig As New JObject() From {
                    {"template_file", relativePath}
                }

                ' Optional human-readable companion guidance uses the same basename
                ' as the Office carrier. This keeps design authoring user-friendly:
                ' a user can drop Example.potx + Example.md into designs\powerpoint
                ' without writing or editing designs.json.
                Dim guidanceCandidate As String =
                    System.IO.Path.Combine(
                        applicationDir,
                        System.IO.Path.GetFileNameWithoutExtension(filePath) & ".md"
                    )
                If System.IO.File.Exists(guidanceCandidate) Then
                    appConfig("guidance_file") =
                        applicationName & "/" & System.IO.Path.GetFileName(guidanceCandidate)
                End If

                Select Case applicationName.ToLowerInvariant()
                    Case "word"
                        appConfig("use_template_styles") = True
                    Case "powerpoint"
                        appConfig("preserve_template_slides") = False
                End Select

                descriptor.Raw(applicationName) = appConfig
            Next
        End Sub

        Private Shared Function NormalizeLookupKey(value As String) As String
            If String.IsNullOrWhiteSpace(value) Then Return ""
            Dim sb As New System.Text.StringBuilder()
            For Each ch As Char In value.Trim().ToLowerInvariant()
                If Char.IsLetterOrDigit(ch) Then sb.Append(ch)
            Next
            Return sb.ToString()
        End Function
    End Class


End Namespace
