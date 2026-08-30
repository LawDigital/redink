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
'  - Design guidance treats an explicitly user-selected source artifact as the format
'    authority ahead of implicit/default repository designs.
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
            Return ToolRegistryBuilder.ToSafeToolSuffix(value)
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
    ''' One deterministic binding declared by a Word design companion guide. The visible
    ''' placeholder name is deliberately semantic-free: e.g. [[RI:Text]] and [[RI:Body]]
    ''' are equivalent when their guidance rows map them to the same source.
    ''' </summary>
    Public Class WordTemplateSlotDefinition
        Public Property Placeholder As String
        Public Property SlotId As String
        Public Property Source As String
        Public Property Purpose As String
        Public Property ContentMode As String
        Public Property Required As Boolean

        Public ReadOnly Property UsesMarkdownContent As Boolean
            Get
                Return String.Equals(Source, "markdown_content", StringComparison.OrdinalIgnoreCase)
            End Get
        End Property

        Public ReadOnly Property TemplateFieldKey As String
            Get
                Const prefix As String = "template_fields."
                If String.IsNullOrWhiteSpace(Source) OrElse
                   Not Source.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) Then
                    Return ""
                End If
                Return Source.Substring(prefix.Length)
            End Get
        End Property
    End Class

    ''' <summary>
    ''' One native Word paragraph style mapping declared by a design companion guide.
    ''' Semantic keys are renderer-level concepts rather than organization-specific style names.
    ''' </summary>
    Public Class WordTemplateBodyStyleDefinition
        Public Property Semantic As String
        Public Property StyleName As String
    End Class

    ''' <summary>
    ''' Parsed Word-template contract from a same-basename Markdown companion file.
    ''' Ordinary prose in the guide stays human guidance. Machine-readable mappings live under
    ''' "## Word template slots", "## Word body styles", "## Word native styles",
    ''' "## Word rendering rules", and the concise model-facing "## Word authoring guidance".
    ''' Property lists are the authoring format; legacy Markdown tables remain accepted as backward-compatible input.
    ''' </summary>
    Public Class WordTemplateBindingContract
        Public Property GuidancePath As String
        Public Property StylePolicyPath As String
        Public Property Slots As New System.Collections.Generic.List(Of WordTemplateSlotDefinition)()
        Public Property BodyStyles As New System.Collections.Generic.List(Of WordTemplateBodyStyleDefinition)()
        Public Property NativeStyles As New System.Collections.Generic.List(Of WordTemplateBodyStyleDefinition)()
        Public Property HeadingNumberingMode As String = ""
        Public Property AuthoringGuidance As New System.Collections.Generic.List(Of String)()

        Public ReadOnly Property HasSlots As Boolean
            Get
                Return Slots IsNot Nothing AndAlso Slots.Count > 0
            End Get
        End Property

        Public ReadOnly Property HasBodyStyles As Boolean
            Get
                Return BodyStyles IsNot Nothing AndAlso BodyStyles.Count > 0
            End Get
        End Property

        Public ReadOnly Property HasNativeStyles As Boolean
            Get
                Return NativeStyles IsNot Nothing AndAlso NativeStyles.Count > 0
            End Get
        End Property

        Public ReadOnly Property HasAuthoringGuidance As Boolean
            Get
                Return AuthoringGuidance IsNot Nothing AndAlso AuthoringGuidance.Count > 0
            End Get
        End Property

        Public Function BuildNativeParagraphStyleMap() As System.Collections.Generic.Dictionary(Of String, String)
            Dim result As New System.Collections.Generic.Dictionary(Of String, String)(System.StringComparer.OrdinalIgnoreCase)
            If BodyStyles Is Nothing Then Return result

            For Each definition As WordTemplateBodyStyleDefinition In BodyStyles
                If definition Is Nothing OrElse
                   System.String.IsNullOrWhiteSpace(definition.Semantic) OrElse
                   System.String.IsNullOrWhiteSpace(definition.StyleName) Then
                    Continue For
                End If
                result(definition.Semantic) = definition.StyleName
            Next
            Return result
        End Function

        Public Function BuildPromptSummary() As String
            If Not HasSlots AndAlso Not HasBodyStyles AndAlso Not HasNativeStyles AndAlso Not HasAuthoringGuidance AndAlso System.String.IsNullOrWhiteSpace(HeadingNumberingMode) Then Return ""

            Dim bodySlots As New System.Collections.Generic.List(Of String)()
            Dim requiredFields As New System.Collections.Generic.List(Of String)()
            Dim optionalFields As New System.Collections.Generic.List(Of String)()

            If Slots IsNot Nothing Then
                For Each slot As WordTemplateSlotDefinition In Slots
                    If slot Is Nothing Then Continue For
                    Dim purposeSuffix As String = FormatPurposeForPrompt(slot.Purpose)
                    If slot.UsesMarkdownContent Then
                        bodySlots.Add(slot.Placeholder & purposeSuffix)
                        Continue For
                    End If

                    Dim key As String = slot.TemplateFieldKey
                    If key = "" Then Continue For
                    Dim describedKey As String = key & purposeSuffix
                    If slot.Required Then
                        requiredFields.Add(describedKey)
                    Else
                        optionalFields.Add(describedKey)
                    End If
                Next
            End If

            Dim parts As New System.Collections.Generic.List(Of String)()
            If bodySlots.Count > 0 Then
                parts.Add("markdown_content -> " & System.String.Join(", ", bodySlots))
            End If
            If requiredFields.Count > 0 Then
                parts.Add("required template_fields: " & System.String.Join(", ", requiredFields.Distinct(System.StringComparer.OrdinalIgnoreCase)))
            End If
            If optionalFields.Count > 0 Then
                parts.Add("optional template_fields: " & System.String.Join(", ", optionalFields.Distinct(System.StringComparer.OrdinalIgnoreCase)))
            End If
            If HasBodyStyles Then
                Dim styleParts As New System.Collections.Generic.List(Of String)()
                For Each definition As WordTemplateBodyStyleDefinition In BodyStyles
                    If definition Is Nothing Then Continue For
                    styleParts.Add(definition.Semantic & "=" & definition.StyleName)
                Next
                If styleParts.Count > 0 Then parts.Add("native body styles: " & System.String.Join(", ", styleParts))
            End If
            If Not System.String.IsNullOrWhiteSpace(HeadingNumberingMode) AndAlso Not System.String.Equals(HeadingNumberingMode, "preserve", System.StringComparison.OrdinalIgnoreCase) Then
                parts.Add("heading numbering=" & HeadingNumberingMode & " (do not put manual numbering prefixes in Markdown headings)")
            End If
            If HasNativeStyles Then
                Dim nativeParts As New System.Collections.Generic.List(Of String)()
                For Each definition As WordTemplateBodyStyleDefinition In NativeStyles
                    If definition Is Nothing Then Continue For
                    nativeParts.Add(definition.Semantic & "=" & definition.StyleName)
                Next
                If nativeParts.Count > 0 Then parts.Add("available native styles: " & System.String.Join(", ", nativeParts))
            End If
            If HasAuthoringGuidance Then
                parts.Add("authoring guidance: " & System.String.Join(" ", AuthoringGuidance.Where(Function(item As String) Not System.String.IsNullOrWhiteSpace(item))))
            End If
            Return System.String.Join("; ", parts)
        End Function

        Private Shared Function FormatPurposeForPrompt(value As String) As String
            Dim normalized As String = If(value, "").Replace(vbCr, " ").Replace(vbLf, " ").Trim()
            If normalized = "" Then Return ""
            Do While normalized.Contains("  ")
                normalized = normalized.Replace("  ", " ")
            Loop
            Return " (" & normalized & ")"
        End Function
    End Class

    ''' <summary>
    ''' Shared parser/resolver for Word template contracts. This contains no
    ''' organization-specific aliases and no Word-Interop rendering logic.
    ''' </summary>
    Public NotInheritable Class WordTemplateBindingContractParser
        Private Shared ReadOnly PlaceholderRegex As New System.Text.RegularExpressions.Regex(
            "^\[\[RI:([\p{L}\p{N}_.-]{1,64})\]\]$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.CultureInvariant)

        Private Shared ReadOnly BodyStyleSemanticRegex As New System.Text.RegularExpressions.Regex(
            "^(paragraph|heading[1-6]|bullet[1-9]|numbered[1-9]|quote[1-9])$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.CultureInvariant)

        Private Sub New()
        End Sub

        Public Shared Function ResolveGuidancePath(descriptor As DocumentDesignDescriptor,
                                                   applicationConfig As Newtonsoft.Json.Linq.JObject,
                                                   templatePath As String) As String
            If descriptor Is Nothing Then Return ""

            Dim configured As String = ""
            If applicationConfig IsNot Nothing Then
                configured = If(applicationConfig.Value(Of String)("guidance_file"), "").Trim()
            End If

            If configured <> "" Then
                Dim resolved As String = descriptor.ResolveRepositoryFile(configured)
                If resolved <> "" AndAlso System.IO.File.Exists(resolved) Then Return resolved
            End If

            If Not System.String.IsNullOrWhiteSpace(templatePath) Then
                Try
                    Dim candidate As String = System.IO.Path.Combine(
                        System.IO.Path.GetDirectoryName(templatePath),
                        System.IO.Path.GetFileNameWithoutExtension(templatePath) & ".md")
                    If System.IO.File.Exists(candidate) Then Return candidate
                Catch ex As System.Exception
                End Try
            End If

            Return ""
        End Function

        Public Shared Function ResolveStylePolicyPath(descriptor As DocumentDesignDescriptor,
                                                      applicationConfig As Newtonsoft.Json.Linq.JObject) As String
            If descriptor Is Nothing OrElse applicationConfig Is Nothing Then Return ""
            Dim configured As String = If(applicationConfig.Value(Of String)("style_policy_file"), "").Trim()
            If configured = "" Then Return ""
            Dim resolved As String = descriptor.ResolveRepositoryFile(configured)
            If resolved <> "" AndAlso System.IO.File.Exists(resolved) Then Return resolved
            Return ""
        End Function

        Public Shared Function TryLoadForDesign(descriptor As DocumentDesignDescriptor,
                                               applicationConfig As Newtonsoft.Json.Linq.JObject,
                                               templatePath As String,
                                               ByRef contract As WordTemplateBindingContract,
                                               ByRef validationError As String) As Boolean
            contract = Nothing
            validationError = ""

            Dim guidancePath As String = ResolveGuidancePath(descriptor, applicationConfig, templatePath)
            Dim stylePolicyPath As String = ResolveStylePolicyPath(descriptor, applicationConfig)
            Dim configuredStylePolicy As String = If(applicationConfig Is Nothing, "", If(applicationConfig.Value(Of String)("style_policy_file"), "").Trim())
            If configuredStylePolicy <> "" AndAlso stylePolicyPath = "" Then
                validationError = "Configured Word style policy was not found or is outside the design repository: " & configuredStylePolicy
                Return False
            End If
            If guidancePath = "" AndAlso stylePolicyPath = "" Then Return True

            Dim merged As New WordTemplateBindingContract() With {
                .GuidancePath = guidancePath,
                .StylePolicyPath = stylePolicyPath
            }

            If stylePolicyPath <> "" Then
                Dim policyText As String
                Try
                    policyText = System.IO.File.ReadAllText(stylePolicyPath, System.Text.Encoding.UTF8)
                Catch ex As System.Exception
                    validationError = "Word style policy could not be read: " & ex.Message
                    Return False
                End Try

                Dim policyContract As WordTemplateBindingContract = Nothing
                If Not TryParse(policyText, stylePolicyPath, policyContract, validationError) Then Return False
                If policyContract IsNot Nothing Then
                    If policyContract.HasSlots Then
                        validationError = "Shared Word style policy '" & System.IO.Path.GetFileName(stylePolicyPath) & "' must not declare template slots. Put slots in the design-specific companion guidance instead."
                        Return False
                    End If
                    MergeStyleDefinitions(merged.BodyStyles, policyContract.BodyStyles, stylePolicyPath, validationError, allowOverride:=False)
                    If validationError <> "" Then Return False
                    MergeStyleDefinitions(merged.NativeStyles, policyContract.NativeStyles, stylePolicyPath, validationError, allowOverride:=False)
                    If validationError <> "" Then Return False
                    MergeAuthoringGuidance(merged.AuthoringGuidance, policyContract.AuthoringGuidance)
                    If Not System.String.IsNullOrWhiteSpace(policyContract.HeadingNumberingMode) Then merged.HeadingNumberingMode = policyContract.HeadingNumberingMode
                End If
            End If

            If guidancePath <> "" Then
                Dim guidance As String
                Try
                    guidance = System.IO.File.ReadAllText(guidancePath, System.Text.Encoding.UTF8)
                Catch ex As System.Exception
                    validationError = "Word design guidance could not be read: " & ex.Message
                    Return False
                End Try

                Dim guideContract As WordTemplateBindingContract = Nothing
                If Not TryParse(guidance, guidancePath, guideContract, validationError) Then Return False
                If guideContract IsNot Nothing Then
                    merged.Slots.AddRange(guideContract.Slots)
                    MergeStyleDefinitions(merged.BodyStyles, guideContract.BodyStyles, guidancePath, validationError, allowOverride:=True)
                    If validationError <> "" Then Return False
                    MergeStyleDefinitions(merged.NativeStyles, guideContract.NativeStyles, guidancePath, validationError, allowOverride:=True)
                    If validationError <> "" Then Return False
                    MergeAuthoringGuidance(merged.AuthoringGuidance, guideContract.AuthoringGuidance)
                    If Not System.String.IsNullOrWhiteSpace(guideContract.HeadingNumberingMode) Then merged.HeadingNumberingMode = guideContract.HeadingNumberingMode
                End If
            End If

            If merged.HasSlots AndAlso Not System.String.IsNullOrWhiteSpace(templatePath) Then
                If Not TryValidateTemplateCarrierSlots(templatePath, merged, validationError) Then Return False
            End If

            If Not merged.HasSlots AndAlso Not merged.HasBodyStyles AndAlso Not merged.HasNativeStyles AndAlso Not merged.HasAuthoringGuidance Then Return True
            contract = merged
            Return True
        End Function

        Private Shared Sub MergeAuthoringGuidance(target As System.Collections.Generic.List(Of String),
                                                        source As System.Collections.Generic.IEnumerable(Of String))
            If target Is Nothing OrElse source Is Nothing Then Return
            For Each item As String In source
                Dim normalized As String = If(item, "").Replace(vbCr, " ").Replace(vbLf, " ").Trim()
                Do While normalized.Contains("  ")
                    normalized = normalized.Replace("  ", " ")
                Loop
                If normalized = "" Then Continue For
                If Not target.Any(Function(existing As String) System.String.Equals(existing, normalized, System.StringComparison.OrdinalIgnoreCase)) Then
                    target.Add(normalized)
                End If
            Next
        End Sub

        Private Shared Sub MergeStyleDefinitions(target As System.Collections.Generic.List(Of WordTemplateBodyStyleDefinition),
                                                  source As System.Collections.Generic.IEnumerable(Of WordTemplateBodyStyleDefinition),
                                                  sourcePath As String,
                                                  ByRef validationError As String,
                                                  Optional allowOverride As System.Boolean = False)
            If target Is Nothing OrElse source Is Nothing Then Return
            For Each definition As WordTemplateBodyStyleDefinition In source
                If definition Is Nothing Then Continue For
                Dim existing As WordTemplateBodyStyleDefinition = target.FirstOrDefault(
                    Function(candidate As WordTemplateBodyStyleDefinition) candidate IsNot Nothing AndAlso System.String.Equals(candidate.Semantic, definition.Semantic, System.StringComparison.OrdinalIgnoreCase))
                If existing IsNot Nothing Then
                    If allowOverride Then
                        existing.StyleName = definition.StyleName
                        Continue For
                    End If
                    validationError = "Duplicate Word style semantic '" & definition.Semantic & "' while merging '" & System.IO.Path.GetFileName(sourcePath) & "'. Shared policies must not contain duplicate semantics."
                    Return
                End If
                target.Add(New WordTemplateBodyStyleDefinition() With {.Semantic = definition.Semantic, .StyleName = definition.StyleName})
            Next
        End Sub

        Private Shared Function TryValidateTemplateCarrierSlots(templatePath As String,
                                                                contract As WordTemplateBindingContract,
                                                                ByRef validationError As String) As Boolean
            validationError = ""
            If contract Is Nothing OrElse Not contract.HasSlots Then Return True
            If System.String.IsNullOrWhiteSpace(templatePath) Then Return True
            If Not System.IO.File.Exists(templatePath) Then
                validationError = "Word design template file was not found while validating its companion guidance: " & templatePath
                Return False
            End If

            Dim extension As String = System.IO.Path.GetExtension(templatePath).ToLowerInvariant()
            If extension <> ".docx" AndAlso extension <> ".dotx" AndAlso extension <> ".dotm" Then
                Return True
            End If

            Try
                Dim allParagraphs As New System.Collections.Generic.List(Of System.Tuple(Of String, String))()

                Using input As New System.IO.FileStream(templatePath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite)
                    Using archive As New System.IO.Compression.ZipArchive(input, System.IO.Compression.ZipArchiveMode.Read, leaveOpen:=False)
                        For Each entry As System.IO.Compression.ZipArchiveEntry In archive.Entries
                            Dim entryName As String = If(entry.FullName, "").Replace("\", "/")
                            If Not entryName.StartsWith("word/", System.StringComparison.OrdinalIgnoreCase) OrElse
                               Not entryName.EndsWith(".xml", System.StringComparison.OrdinalIgnoreCase) Then
                                Continue For
                            End If

                            Using entryStream As System.IO.Stream = entry.Open()
                                Dim xml As System.Xml.Linq.XDocument = System.Xml.Linq.XDocument.Load(entryStream, System.Xml.Linq.LoadOptions.PreserveWhitespace)
                                For Each paragraph As System.Xml.Linq.XElement In xml.Descendants().Where(Function(element As System.Xml.Linq.XElement) System.String.Equals(element.Name.LocalName, "p", System.StringComparison.Ordinal))
                                    Dim paragraphText As New System.Text.StringBuilder()
                                    For Each textNode As System.Xml.Linq.XElement In paragraph.Descendants().Where(Function(element As System.Xml.Linq.XElement) System.String.Equals(element.Name.LocalName, "t", System.StringComparison.Ordinal))
                                        paragraphText.Append(If(textNode.Value, ""))
                                    Next
                                    If paragraphText.Length > 0 Then
                                        allParagraphs.Add(System.Tuple.Create(entryName, paragraphText.ToString()))
                                    End If
                                Next
                            End Using
                        Next
                    End Using
                End Using

                For Each slot As WordTemplateSlotDefinition In contract.Slots
                    If slot Is Nothing Then Continue For
                    Dim allCount As Integer = 0
                    Dim mainDocumentCount As Integer = 0

                    For Each paragraphInfo As System.Tuple(Of String, String) In allParagraphs
                        Dim occurrenceCount As Integer = CountOrdinalIgnoreCaseOccurrences(paragraphInfo.Item2, slot.Placeholder)
                        If occurrenceCount <= 0 Then Continue For
                        allCount += occurrenceCount
                        If System.String.Equals(paragraphInfo.Item1, "word/document.xml", System.StringComparison.OrdinalIgnoreCase) Then
                            mainDocumentCount += occurrenceCount
                        End If
                    Next

                    If allCount = 0 Then
                        validationError = "Word design package mismatch: placeholder " & slot.Placeholder &
                                          " is declared in '" & System.IO.Path.GetFileName(contract.GuidancePath) &
                                          "' but is absent from template '" & System.IO.Path.GetFileName(templatePath) &
                                          "'. Deploy the template carrier and its companion .md together."
                        Return False
                    End If

                    If System.String.Equals(slot.ContentMode, "markdown", System.StringComparison.OrdinalIgnoreCase) AndAlso
                       (allCount <> 1 OrElse mainDocumentCount <> 1) Then
                        validationError = "Word design package mismatch: Markdown placeholder " & slot.Placeholder &
                                          " must occur exactly once in the main document part of template '" &
                                          System.IO.Path.GetFileName(templatePath) & "'; found " &
                                          allCount.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                                          " occurrence(s) in all Word XML parts and " &
                                          mainDocumentCount.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                                          " in word/document.xml."
                        Return False
                    End If
                Next

                Return True
            Catch ex As System.Exception
                validationError = "Word design template package could not be validated against its companion guidance: " & ex.Message
                Return False
            End Try
        End Function

        Private Shared Function CountOrdinalIgnoreCaseOccurrences(text As String, value As String) As Integer
            If System.String.IsNullOrEmpty(text) OrElse System.String.IsNullOrEmpty(value) Then Return 0
            Dim count As Integer = 0
            Dim startIndex As Integer = 0
            Do
                Dim matchIndex As Integer = text.IndexOf(value, startIndex, System.StringComparison.OrdinalIgnoreCase)
                If matchIndex < 0 Then Exit Do
                count += 1
                startIndex = matchIndex + value.Length
            Loop While startIndex < text.Length
            Return count
        End Function

        Public Shared Function TryParse(guidance As String,
                                       guidancePath As String,
                                       ByRef contract As WordTemplateBindingContract,
                                       ByRef validationError As String) As Boolean
            contract = Nothing
            validationError = ""
            If System.String.IsNullOrWhiteSpace(guidance) Then Return True

            Dim lines() As String = guidance.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Split(ControlChars.Lf)
            Dim parsed As New WordTemplateBindingContract() With {.GuidancePath = guidancePath}

            Dim slotSectionIndex As Integer = FindSectionIndex(lines, "## Word template slots")
            If slotSectionIndex >= 0 Then
                If Not TryParseSlotSection(lines, slotSectionIndex, guidancePath, parsed, validationError) Then Return False
            End If

            Dim bodyStyleSectionIndex As Integer = FindSectionIndex(lines, "## Word body styles")
            If bodyStyleSectionIndex >= 0 Then
                If Not TryParseBodyStyleSection(lines, bodyStyleSectionIndex, guidancePath, parsed, validationError) Then Return False
            End If

            Dim nativeStyleSectionIndex As Integer = FindSectionIndex(lines, "## Word native styles")
            If nativeStyleSectionIndex >= 0 Then
                If Not TryParseNativeStyleSection(lines, nativeStyleSectionIndex, guidancePath, parsed, validationError) Then Return False
            End If

            Dim renderingRulesSectionIndex As Integer = FindSectionIndex(lines, "## Word rendering rules")
            If renderingRulesSectionIndex >= 0 Then
                If Not TryParseRenderingRulesSection(lines, renderingRulesSectionIndex, guidancePath, parsed, validationError) Then Return False
            End If

            Dim authoringGuidanceSectionIndex As Integer = FindSectionIndex(lines, "## Word authoring guidance")
            If authoringGuidanceSectionIndex >= 0 Then
                If Not TryParseAuthoringGuidanceSection(lines, authoringGuidanceSectionIndex, guidancePath, parsed, validationError) Then Return False
            End If

            ' Prose-only companion guides remain valid and preserve the legacy style-carrier path.
            If Not parsed.HasSlots AndAlso Not parsed.HasBodyStyles AndAlso Not parsed.HasNativeStyles AndAlso Not parsed.HasAuthoringGuidance AndAlso System.String.IsNullOrWhiteSpace(parsed.HeadingNumberingMode) Then Return True

            contract = parsed
            Return True
        End Function

        Private Shared Function FindSectionIndex(lines() As String, heading As String) As Integer
            If lines Is Nothing OrElse System.String.IsNullOrWhiteSpace(heading) Then Return -1
            For i As Integer = 0 To lines.Length - 1
                If System.String.Equals(If(lines(i), "").Trim(), heading, System.StringComparison.OrdinalIgnoreCase) Then Return i
            Next
            Return -1
        End Function

        Private Shared Function FindSectionEndIndex(lines() As String, sectionIndex As Integer) As Integer
            If lines Is Nothing Then Return 0
            For i As Integer = System.Math.Max(0, sectionIndex + 1) To lines.Length - 1
                Dim candidate As String = If(lines(i), "").Trim()
                If candidate.StartsWith("#", System.StringComparison.Ordinal) Then Return i
            Next
            Return lines.Length
        End Function

        Private Shared Function TryParseSlotSection(lines() As String,
                                                    sectionIndex As Integer,
                                                    guidancePath As String,
                                                    parsed As WordTemplateBindingContract,
                                                    ByRef validationError As String) As Boolean
            If SlotSectionUsesPropertyList(lines, sectionIndex) Then
                Return TryParseSlotPropertyList(lines, sectionIndex, guidancePath, parsed, validationError)
            End If

            ' Backward compatibility: older companion guides used Markdown tables. New guidance
            ' should use the property-list form because it is substantially easier to edit by hand.
            Return TryParseSlotTable(lines, sectionIndex, guidancePath, parsed, validationError)
        End Function

        Private Shared Function TryParseBodyStyleSection(lines() As String,
                                                         sectionIndex As Integer,
                                                         guidancePath As String,
                                                         parsed As WordTemplateBindingContract,
                                                         ByRef validationError As String) As Boolean
            If BodyStyleSectionUsesPropertyList(lines, sectionIndex) Then
                Return TryParseBodyStylePropertyList(lines, sectionIndex, guidancePath, parsed, validationError)
            End If

            ' Backward compatibility only; newly authored guides should use one-line mappings.
            Return TryParseBodyStyleTable(lines, sectionIndex, guidancePath, parsed, validationError)
        End Function

        Private Shared Function SlotSectionUsesPropertyList(lines() As String, sectionIndex As Integer) As Boolean
            Dim sectionEnd As Integer = FindSectionEndIndex(lines, sectionIndex)
            For i As Integer = sectionIndex + 1 To sectionEnd - 1
                Dim key As String = ""
                Dim value As String = ""
                If Not TrySplitListPropertyLine(If(lines(i), ""), key, value) Then Continue For
                If System.String.Equals(key, "placeholder", System.StringComparison.OrdinalIgnoreCase) OrElse
                   System.String.Equals(key, "slot", System.StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If
            Next
            Return False
        End Function

        Private Shared Function BodyStyleSectionUsesPropertyList(lines() As String, sectionIndex As Integer) As Boolean
            Dim sectionEnd As Integer = FindSectionEndIndex(lines, sectionIndex)
            For i As Integer = sectionIndex + 1 To sectionEnd - 1
                Dim candidate As String = If(lines(i), "").Trim()
                If candidate = "" Then Continue For
                If candidate.StartsWith("|", System.StringComparison.Ordinal) Then Return False

                Dim key As String = ""
                Dim value As String = ""
                If TrySplitListPropertyLine(candidate, key, value) Then Return True
            Next
            Return False
        End Function

        Private Shared Function TryParseSlotPropertyList(lines() As String,
                                                         sectionIndex As Integer,
                                                         guidancePath As String,
                                                         parsed As WordTemplateBindingContract,
                                                         ByRef validationError As String) As Boolean
            Dim sectionEnd As Integer = FindSectionEndIndex(lines, sectionIndex)
            Dim seenPlaceholders As New System.Collections.Generic.HashSet(Of String)(System.StringComparer.OrdinalIgnoreCase)
            Dim current As System.Collections.Generic.Dictionary(Of String, String) = Nothing
            Dim currentLine As Integer = -1

            For i As Integer = sectionIndex + 1 To sectionEnd - 1
                Dim raw As String = If(lines(i), "")
                Dim trimmed As String = raw.Trim()
                If trimmed = "" Then Continue For

                Dim listKey As String = ""
                Dim listValue As String = ""
                If TrySplitListPropertyLine(raw, listKey, listValue) Then
                    If System.String.Equals(listKey, "placeholder", System.StringComparison.OrdinalIgnoreCase) OrElse
                       System.String.Equals(listKey, "slot", System.StringComparison.OrdinalIgnoreCase) Then

                        If current IsNot Nothing Then
                            If Not TryAddSlotPropertyListEntry(current, currentLine, guidancePath, parsed, seenPlaceholders, validationError) Then Return False
                        End If

                        current = New System.Collections.Generic.Dictionary(Of String, String)(System.StringComparer.OrdinalIgnoreCase)
                        current("placeholder") = listValue
                        currentLine = i + 1
                        Continue For
                    End If

                    If current IsNot Nothing Then
                        validationError = $"Invalid Word template slot list entry in '{System.IO.Path.GetFileName(guidancePath)}' at line {i + 1}. Start each slot with '- placeholder: [[RI:SlotName]]' and put source/purpose/required (and optional content) on the following indented lines."
                        Return False
                    End If
                    Continue For
                End If

                If current Is Nothing Then
                    ' Human prose before the first list entry is allowed.
                    Continue For
                End If

                Dim propertyKey As String = ""
                Dim propertyValue As String = ""
                If Not TrySplitPropertyLine(raw, propertyKey, propertyValue) Then
                    validationError = $"Invalid Word template slot property in '{System.IO.Path.GetFileName(guidancePath)}' at line {i + 1}. Use 'source:', 'purpose:', 'required:', or optional 'content:'."
                    Return False
                End If

                Select Case propertyKey.Trim().ToLowerInvariant()
                    Case "source", "purpose", "description", "content", "required"
                        Dim normalizedKey As String = propertyKey.Trim().ToLowerInvariant()
                        If normalizedKey = "description" Then normalizedKey = "purpose"
                        current(normalizedKey) = propertyValue
                    Case Else
                        validationError = $"Unknown Word template slot property '{propertyKey}' in '{System.IO.Path.GetFileName(guidancePath)}' at line {i + 1}. Use source, purpose, required, or optional content."
                        Return False
                End Select
            Next

            If current IsNot Nothing Then
                If Not TryAddSlotPropertyListEntry(current, currentLine, guidancePath, parsed, seenPlaceholders, validationError) Then Return False
            End If

            If parsed.Slots.Count = 0 Then
                validationError = $"The '## Word template slots' section in '{System.IO.Path.GetFileName(guidancePath)}' contains no bindings. Use '- placeholder: [[RI:SlotName]]' entries with source/purpose/required properties; content is optional."
                Return False
            End If
            Return True
        End Function

        Private Shared Function TryAddSlotPropertyListEntry(values As System.Collections.Generic.Dictionary(Of String, String),
                                                            lineNumber As Integer,
                                                            guidancePath As String,
                                                            parsed As WordTemplateBindingContract,
                                                            seenPlaceholders As System.Collections.Generic.HashSet(Of String),
                                                            ByRef validationError As String) As Boolean
            If values Is Nothing Then Return True

            Dim placeholder As String = If(values.ContainsKey("placeholder"), NormalizeCell(values("placeholder")), "")
            Dim source As String = If(values.ContainsKey("source"), NormalizeCell(values("source")), "")
            Dim purpose As String = If(values.ContainsKey("purpose"), NormalizeCell(values("purpose")), "")
            Dim contentMode As String = If(values.ContainsKey("content"), NormalizeCell(values("content")).ToLowerInvariant(), "")
            Dim requiredText As String = If(values.ContainsKey("required"), NormalizeCell(values("required")).ToLowerInvariant(), "")

            If placeholder = "" OrElse source = "" OrElse requiredText = "" Then
                validationError = $"Incomplete Word template slot entry in '{System.IO.Path.GetFileName(guidancePath)}' at line {lineNumber}. Each slot requires placeholder, source, and required. Add purpose for clear model-facing semantics; content is optional and inferred from source when omitted."
                Return False
            End If

            Dim match As System.Text.RegularExpressions.Match = PlaceholderRegex.Match(placeholder)
            If Not match.Success Then
                validationError = $"Invalid Word placeholder '{placeholder}' in '{System.IO.Path.GetFileName(guidancePath)}'. Use the exact speaking format [[RI:SlotName]]."
                Return False
            End If
            If Not seenPlaceholders.Add(placeholder) Then
                validationError = $"Duplicate Word placeholder '{placeholder}' in '{System.IO.Path.GetFileName(guidancePath)}'."
                Return False
            End If

            Dim validSource As Boolean =
                System.String.Equals(source, "markdown_content", System.StringComparison.OrdinalIgnoreCase) OrElse
                System.Text.RegularExpressions.Regex.IsMatch(source, "^template_fields\.[\p{L}\p{N}_.-]{1,64}$", System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.CultureInvariant)
            If Not validSource Then
                validationError = $"Invalid source '{source}' for Word placeholder '{placeholder}'. Use markdown_content or template_fields.<key>."
                Return False
            End If

            If contentMode = "" Then
                If System.String.Equals(source, "markdown_content", System.StringComparison.OrdinalIgnoreCase) Then
                    contentMode = "markdown"
                Else
                    contentMode = "text"
                End If
            End If

            If contentMode <> "text" AndAlso contentMode <> "markdown" Then
                validationError = $"Invalid content mode '{contentMode}' for Word placeholder '{placeholder}'. Use text or markdown."
                Return False
            End If

            Dim isRequired As Boolean
            Select Case requiredText
                Case "yes", "true", "required", "1" : isRequired = True
                Case "no", "false", "optional", "0" : isRequired = False
                Case Else
                    validationError = $"Invalid required value '{requiredText}' for Word placeholder '{placeholder}'. Use yes/no."
                    Return False
            End Select

            parsed.Slots.Add(New WordTemplateSlotDefinition() With {
                .Placeholder = placeholder,
                .SlotId = match.Groups(1).Value,
                .Source = source,
                .Purpose = purpose,
                .ContentMode = contentMode,
                .Required = isRequired
            })
            Return True
        End Function

        Private Shared Function TryParseBodyStylePropertyList(lines() As String,
                                                              sectionIndex As Integer,
                                                              guidancePath As String,
                                                              parsed As WordTemplateBindingContract,
                                                              ByRef validationError As String) As Boolean
            Dim sectionEnd As Integer = FindSectionEndIndex(lines, sectionIndex)
            Dim seenSemantics As New System.Collections.Generic.HashSet(Of String)(System.StringComparer.OrdinalIgnoreCase)

            For i As Integer = sectionIndex + 1 To sectionEnd - 1
                Dim raw As String = If(lines(i), "")
                Dim trimmed As String = raw.Trim()
                If trimmed = "" Then Continue For

                Dim key As String = ""
                Dim value As String = ""
                If Not TrySplitListPropertyLine(raw, key, value) Then
                    ' Human prose before the first mapping is allowed; once mappings start the section
                    ' stays machine-readable and therefore rejects ambiguous free-form content.
                    If parsed.BodyStyles.Count = 0 Then Continue For
                    validationError = $"Invalid Word body-style mapping in '{System.IO.Path.GetFileName(guidancePath)}' at line {i + 1}. Use '- semantic: Exact Word Style Name'."
                    Return False
                End If

                Dim semantic As String = NormalizeBodyStyleSemantic(key)
                Dim styleName As String = NormalizeCell(value)
                If semantic = "" OrElse Not BodyStyleSemanticRegex.IsMatch(semantic) Then
                    validationError = $"Invalid Word body-style semantic '{key}' in '{System.IO.Path.GetFileName(guidancePath)}'. Use paragraph, heading1..heading6, bullet1..bullet9, numbered1..numbered9, or quote1..quote9."
                    Return False
                End If
                If System.String.IsNullOrWhiteSpace(styleName) OrElse styleName.Length > 128 Then
                    validationError = $"Invalid Word style name for semantic '{semantic}' in '{System.IO.Path.GetFileName(guidancePath)}'."
                    Return False
                End If
                If Not seenSemantics.Add(semantic) Then
                    validationError = $"Duplicate Word body-style semantic '{semantic}' in '{System.IO.Path.GetFileName(guidancePath)}'."
                    Return False
                End If

                parsed.BodyStyles.Add(New WordTemplateBodyStyleDefinition() With {
                    .Semantic = semantic,
                    .StyleName = styleName
                })
            Next

            If parsed.BodyStyles.Count = 0 Then
                validationError = $"The '## Word body styles' section in '{System.IO.Path.GetFileName(guidancePath)}' contains no mappings. Use one-line entries such as '- paragraph: Normal'."
                Return False
            End If
            Return True
        End Function


        Private Shared Function TryParseRenderingRulesSection(lines() As String,
                                                               sectionIndex As Integer,
                                                               guidancePath As String,
                                                               parsed As WordTemplateBindingContract,
                                                               ByRef validationError As String) As Boolean
            Dim sectionEnd As Integer = FindSectionEndIndex(lines, sectionIndex)
            For i As Integer = sectionIndex + 1 To sectionEnd - 1
                Dim raw As String = If(lines(i), "")
                If System.String.IsNullOrWhiteSpace(raw) Then Continue For

                ' Only explicit list-property entries are machine-readable. Human prose, examples,
                ' inline code and explanatory sentences inside the section must never become
                ' configuration merely because they contain a colon. This keeps guidance easy to edit
                ' without making the parser fragile.
                Dim key As String = ""
                Dim value As String = ""
                If Not TrySplitListPropertyLine(raw, key, value) Then Continue For

                Select Case NormalizeCell(key).Trim().ToLowerInvariant()
                    Case "heading_numbering"
                        Dim mode As String = NormalizeCell(value).Trim().ToLowerInvariant()
                        If mode <> "native" AndAlso mode <> "preserve" Then
                            validationError = $"Invalid heading_numbering value '{value}' in '{System.IO.Path.GetFileName(guidancePath)}'. Use native or preserve."
                            Return False
                        End If
                        parsed.HeadingNumberingMode = mode
                    Case Else
                        validationError = $"Unknown Word rendering rule '{key}' in '{System.IO.Path.GetFileName(guidancePath)}'. Currently supported: heading_numbering."
                        Return False
                End Select
            Next
            Return True
        End Function

        Private Shared Function TryParseAuthoringGuidanceSection(lines() As String,
                                                                  sectionIndex As Integer,
                                                                  guidancePath As String,
                                                                  parsed As WordTemplateBindingContract,
                                                                  ByRef validationError As String) As Boolean
            Dim sectionEnd As Integer = FindSectionEndIndex(lines, sectionIndex)
            Dim seen As New System.Collections.Generic.HashSet(Of String)(System.StringComparer.OrdinalIgnoreCase)

            For i As Integer = sectionIndex + 1 To sectionEnd - 1
                Dim raw As String = If(lines(i), "")
                Dim trimmed As String = raw.Trim()
                If trimmed = "" Then Continue For
                If Not trimmed.StartsWith("-", System.StringComparison.Ordinal) Then
                    validationError = $"Invalid Word authoring guidance in '{System.IO.Path.GetFileName(guidancePath)}' at line {i + 1}. Use one concise '- ...' rule per line."
                    Return False
                End If

                Dim rule As String = NormalizeCell(trimmed.Substring(1).Trim())
                If System.String.IsNullOrWhiteSpace(rule) OrElse rule.Length > 768 Then
                    validationError = $"Invalid Word authoring guidance rule in '{System.IO.Path.GetFileName(guidancePath)}' at line {i + 1}."
                    Return False
                End If
                If seen.Add(rule) Then parsed.AuthoringGuidance.Add(rule)
            Next

            If parsed.AuthoringGuidance.Count = 0 Then
                validationError = $"The '## Word authoring guidance' section in '{System.IO.Path.GetFileName(guidancePath)}' contains no rules."
                Return False
            End If
            Return True
        End Function

        Private Shared Function TryParseNativeStyleSection(lines() As String,
                                                            sectionIndex As Integer,
                                                            guidancePath As String,
                                                            parsed As WordTemplateBindingContract,
                                                            ByRef validationError As String) As Boolean
            Dim sectionEnd As Integer = FindSectionEndIndex(lines, sectionIndex)
            Dim seenSemantics As New System.Collections.Generic.HashSet(Of String)(System.StringComparer.OrdinalIgnoreCase)

            For i As Integer = sectionIndex + 1 To sectionEnd - 1
                Dim raw As String = If(lines(i), "")
                Dim trimmed As String = raw.Trim()
                If trimmed = "" Then Continue For

                Dim key As String = ""
                Dim value As String = ""
                If Not TrySplitListPropertyLine(raw, key, value) Then
                    If parsed.NativeStyles.Count = 0 Then Continue For
                    validationError = $"Invalid Word native-style mapping in '{System.IO.Path.GetFileName(guidancePath)}' at line {i + 1}. Use '- semantic: Exact Word Style Name'."
                    Return False
                End If

                Dim semantic As String = NormalizeCell(key).Trim().ToLowerInvariant()
                Dim styleName As String = NormalizeCell(value)
                If Not System.Text.RegularExpressions.Regex.IsMatch(semantic, "^[a-z][a-z0-9_-]{0,63}$", System.Text.RegularExpressions.RegexOptions.CultureInvariant) Then
                    validationError = $"Invalid Word native-style semantic '{key}' in '{System.IO.Path.GetFileName(guidancePath)}'."
                    Return False
                End If
                If System.String.IsNullOrWhiteSpace(styleName) OrElse styleName.Length > 128 Then
                    validationError = $"Invalid Word native style name for semantic '{semantic}' in '{System.IO.Path.GetFileName(guidancePath)}'."
                    Return False
                End If
                If Not seenSemantics.Add(semantic) Then
                    validationError = $"Duplicate Word native-style semantic '{semantic}' in '{System.IO.Path.GetFileName(guidancePath)}'."
                    Return False
                End If

                parsed.NativeStyles.Add(New WordTemplateBodyStyleDefinition() With {
                    .Semantic = semantic,
                    .StyleName = styleName
                })
            Next

            If parsed.NativeStyles.Count = 0 Then
                validationError = $"The '## Word native styles' section in '{System.IO.Path.GetFileName(guidancePath)}' contains no mappings."
                Return False
            End If
            Return True
        End Function

        Private Shared Function TrySplitListPropertyLine(rawLine As String,
                                                         ByRef key As String,
                                                         ByRef value As String) As Boolean
            key = ""
            value = ""
            If System.String.IsNullOrWhiteSpace(rawLine) Then Return False

            Dim trimmed As String = rawLine.Trim()
            If Not trimmed.StartsWith("-", System.StringComparison.Ordinal) Then Return False
            trimmed = trimmed.Substring(1).TrimStart()
            Return TrySplitPropertyValue(trimmed, key, value)
        End Function

        Private Shared Function TrySplitPropertyLine(rawLine As String,
                                                     ByRef key As String,
                                                     ByRef value As String) As Boolean
            key = ""
            value = ""
            If System.String.IsNullOrWhiteSpace(rawLine) Then Return False
            Return TrySplitPropertyValue(rawLine.Trim(), key, value)
        End Function

        Private Shared Function TrySplitPropertyValue(rawValue As String,
                                                      ByRef key As String,
                                                      ByRef value As String) As Boolean
            key = ""
            value = ""
            If System.String.IsNullOrWhiteSpace(rawValue) Then Return False

            Dim separatorIndex As Integer = rawValue.IndexOf(":"c)
            If separatorIndex <= 0 Then Return False

            key = NormalizeCell(rawValue.Substring(0, separatorIndex))
            value = NormalizeCell(rawValue.Substring(separatorIndex + 1))
            Return key <> ""
        End Function

        Private Shared Function TryParseSlotTable(lines() As String,
                                                  sectionIndex As Integer,
                                                  guidancePath As String,
                                                  parsed As WordTemplateBindingContract,
                                                  ByRef validationError As String) As Boolean
            Dim headerIndex As Integer = -1
            Dim columnIndexes As New System.Collections.Generic.Dictionary(Of String, Integer)(System.StringComparer.OrdinalIgnoreCase)

            For i As Integer = sectionIndex + 1 To lines.Length - 1
                Dim candidate As String = If(lines(i), "").Trim()
                If candidate.StartsWith("#", System.StringComparison.Ordinal) Then Exit For

                Dim cells As System.Collections.Generic.List(Of String) = SplitMarkdownTableLine(candidate)
                If cells.Count < 4 Then Continue For

                columnIndexes.Clear()
                For c As Integer = 0 To cells.Count - 1
                    Dim name As String = NormalizeCell(cells(c)).ToLowerInvariant()
                    If name = "placeholder" OrElse name = "source" OrElse name = "content" OrElse name = "required" OrElse name = "purpose" OrElse name = "description" Then
                        If name = "description" Then name = "purpose"
                        columnIndexes(name) = c
                    End If
                Next

                If columnIndexes.ContainsKey("placeholder") AndAlso
                   columnIndexes.ContainsKey("source") AndAlso
                   columnIndexes.ContainsKey("content") AndAlso
                   columnIndexes.ContainsKey("required") Then
                    headerIndex = i
                    Exit For
                End If
            Next

            If headerIndex < 0 Then
                validationError = $"The '## Word template slots' section in '{System.IO.Path.GetFileName(guidancePath)}' must contain property-list entries beginning with '- placeholder: [[RI:SlotName]]'. Legacy Markdown tables are still accepted for backward compatibility."
                Return False
            End If

            Dim seenPlaceholders As New System.Collections.Generic.HashSet(Of String)(System.StringComparer.OrdinalIgnoreCase)
            For i As Integer = headerIndex + 1 To lines.Length - 1
                Dim raw As String = If(lines(i), "").Trim()
                If raw = "" Then
                    If parsed.Slots.Count > 0 Then Exit For
                    Continue For
                End If
                If raw.StartsWith("#", System.StringComparison.Ordinal) AndAlso parsed.Slots.Count > 0 Then Exit For
                If Not raw.StartsWith("|", System.StringComparison.Ordinal) Then
                    If parsed.Slots.Count > 0 Then Exit For
                    Continue For
                End If

                Dim cells As System.Collections.Generic.List(Of String) = SplitMarkdownTableLine(raw)
                If cells.Count = 0 OrElse IsMarkdownSeparatorRow(cells) Then Continue For

                Dim maxIndex As Integer = columnIndexes.Values.Max()
                If cells.Count <= maxIndex Then
                    validationError = $"Invalid legacy Word template slot row in '{System.IO.Path.GetFileName(guidancePath)}' at line {i + 1}: expected Placeholder, Source, Content, Required columns."
                    Return False
                End If

                Dim placeholder As String = NormalizeCell(cells(columnIndexes("placeholder")))
                Dim source As String = NormalizeCell(cells(columnIndexes("source")))
                Dim contentMode As String = NormalizeCell(cells(columnIndexes("content"))).ToLowerInvariant()
                Dim requiredText As String = NormalizeCell(cells(columnIndexes("required"))).ToLowerInvariant()
                If placeholder = "" AndAlso source = "" Then Continue For

                Dim values As New System.Collections.Generic.Dictionary(Of String, String)(System.StringComparer.OrdinalIgnoreCase) From {
                    {"placeholder", placeholder},
                    {"source", source},
                    {"content", contentMode},
                    {"required", requiredText}
                }
                If columnIndexes.ContainsKey("purpose") AndAlso cells.Count > columnIndexes("purpose") Then
                    values("purpose") = NormalizeCell(cells(columnIndexes("purpose")))
                End If
                If Not TryAddSlotPropertyListEntry(values, i + 1, guidancePath, parsed, seenPlaceholders, validationError) Then Return False
            Next

            If parsed.Slots.Count = 0 Then
                validationError = $"The legacy Word template slot table in '{System.IO.Path.GetFileName(guidancePath)}' contains no bindings."
                Return False
            End If
            Return True
        End Function

        Private Shared Function TryParseBodyStyleTable(lines() As String,
                                                       sectionIndex As Integer,
                                                       guidancePath As String,
                                                       parsed As WordTemplateBindingContract,
                                                       ByRef validationError As String) As Boolean
            Dim headerIndex As Integer = -1
            Dim semanticColumn As Integer = -1
            Dim styleColumn As Integer = -1

            For i As Integer = sectionIndex + 1 To lines.Length - 1
                Dim candidate As String = If(lines(i), "").Trim()
                If candidate.StartsWith("#", System.StringComparison.Ordinal) Then Exit For

                Dim cells As System.Collections.Generic.List(Of String) = SplitMarkdownTableLine(candidate)
                If cells.Count < 2 Then Continue For

                semanticColumn = -1
                styleColumn = -1
                For c As Integer = 0 To cells.Count - 1
                    Dim name As String = NormalizeCell(cells(c)).ToLowerInvariant()
                    If name = "semantic" Then semanticColumn = c
                    If name = "word style" OrElse name = "style" Then styleColumn = c
                Next
                If semanticColumn >= 0 AndAlso styleColumn >= 0 Then
                    headerIndex = i
                    Exit For
                End If
            Next

            If headerIndex < 0 Then
                validationError = $"The '## Word body styles' section in '{System.IO.Path.GetFileName(guidancePath)}' must contain one-line mappings such as '- paragraph: Normal'. Legacy Markdown tables are still accepted for backward compatibility."
                Return False
            End If

            Dim seenSemantics As New System.Collections.Generic.HashSet(Of String)(System.StringComparer.OrdinalIgnoreCase)
            For i As Integer = headerIndex + 1 To lines.Length - 1
                Dim raw As String = If(lines(i), "").Trim()
                If raw = "" Then
                    If parsed.BodyStyles.Count > 0 Then Exit For
                    Continue For
                End If
                If raw.StartsWith("#", System.StringComparison.Ordinal) AndAlso parsed.BodyStyles.Count > 0 Then Exit For
                If Not raw.StartsWith("|", System.StringComparison.Ordinal) Then
                    If parsed.BodyStyles.Count > 0 Then Exit For
                    Continue For
                End If

                Dim cells As System.Collections.Generic.List(Of String) = SplitMarkdownTableLine(raw)
                If cells.Count = 0 OrElse IsMarkdownSeparatorRow(cells) Then Continue For
                Dim maxIndex As Integer = System.Math.Max(semanticColumn, styleColumn)
                If cells.Count <= maxIndex Then
                    validationError = $"Invalid legacy Word body style row in '{System.IO.Path.GetFileName(guidancePath)}' at line {i + 1}: expected Semantic and Word style columns."
                    Return False
                End If

                Dim semantic As String = NormalizeBodyStyleSemantic(NormalizeCell(cells(semanticColumn)))
                Dim styleName As String = NormalizeCell(cells(styleColumn))
                If semantic = "" AndAlso styleName = "" Then Continue For

                If semantic = "" OrElse Not BodyStyleSemanticRegex.IsMatch(semantic) Then
                    validationError = $"Invalid Word body-style semantic '{NormalizeCell(cells(semanticColumn))}' in '{System.IO.Path.GetFileName(guidancePath)}'. Use paragraph, heading1..heading6, bullet1..bullet9, numbered1..numbered9, or quote1..quote9."
                    Return False
                End If
                If System.String.IsNullOrWhiteSpace(styleName) OrElse styleName.Length > 128 Then
                    validationError = $"Invalid Word style name for semantic '{semantic}' in '{System.IO.Path.GetFileName(guidancePath)}'."
                    Return False
                End If
                If Not seenSemantics.Add(semantic) Then
                    validationError = $"Duplicate Word body-style semantic '{semantic}' in '{System.IO.Path.GetFileName(guidancePath)}'."
                    Return False
                End If

                parsed.BodyStyles.Add(New WordTemplateBodyStyleDefinition() With {
                    .Semantic = semantic,
                    .StyleName = styleName
                })
            Next

            If parsed.BodyStyles.Count = 0 Then
                validationError = $"The legacy Word body style table in '{System.IO.Path.GetFileName(guidancePath)}' contains no mappings."
                Return False
            End If
            Return True
        End Function

        Private Shared Function NormalizeBodyStyleSemantic(value As String) As String
            Dim normalized As String = If(value, "").Trim().ToLowerInvariant().Replace(" ", "").Replace("_", "").Replace(".", "")
            If normalized = "body" OrElse normalized = "normal" OrElse normalized = "bodyparagraph" Then Return "paragraph"
            Return normalized
        End Function

        Private Shared Function SplitMarkdownTableLine(rawLine As String) As System.Collections.Generic.List(Of String)
            Dim result As New System.Collections.Generic.List(Of String)()
            If System.String.IsNullOrWhiteSpace(rawLine) Then Return result
            Dim value As String = rawLine.Trim()
            If value.StartsWith("|", System.StringComparison.Ordinal) Then value = value.Substring(1)
            If value.EndsWith("|", System.StringComparison.Ordinal) Then value = value.Substring(0, value.Length - 1)
            For Each part As String In value.Split("|"c)
                result.Add(part.Trim())
            Next
            Return result
        End Function

        Private Shared Function NormalizeCell(value As String) As String
            Dim result As String = If(value, "").Trim()
            Dim changed As Boolean = True
            Do While changed AndAlso result.Length >= 2
                changed = False
                Dim first As Char = result(0)
                Dim last As Char = result(result.Length - 1)
                Dim isMatchingWrapper As Boolean =
                    (first = "`"c AndAlso last = "`"c) OrElse
                    (first = Microsoft.VisualBasic.ChrW(34) AndAlso last = Microsoft.VisualBasic.ChrW(34)) OrElse
                    (first = "'"c AndAlso last = "'"c) OrElse
                    (first = Microsoft.VisualBasic.ChrW(&H201C) AndAlso last = Microsoft.VisualBasic.ChrW(&H201D)) OrElse
                    (first = Microsoft.VisualBasic.ChrW(&H2018) AndAlso last = Microsoft.VisualBasic.ChrW(&H2019))

                If isMatchingWrapper Then
                    result = result.Substring(1, result.Length - 2).Trim()
                    changed = True
                End If
            Loop
            Return result
        End Function

        Private Shared Function IsMarkdownSeparatorRow(cells As System.Collections.Generic.List(Of String)) As Boolean
            If cells Is Nothing OrElse cells.Count = 0 Then Return False
            For Each cell As String In cells
                Dim normalized As String = NormalizeCell(cell).Replace(":", "").Trim()
                If normalized.Length < 3 OrElse normalized.Any(Function(ch As Char) ch <> "-"c) Then Return False
            Next
            Return True
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
        Public Const DesignSetsDirectoryName As String = "design_sets"
        Public Const ActiveDesignSetFileName As String = "active.json"
        Public Const SupportedSchemaVersion As Integer = 1

        Private NotInheritable Class DesignAccessState
            Public Property AllowDesigns As System.Boolean = True
            Public Property AllowDesignSets As System.Boolean = True
        End Class

        Private NotInheritable Class DesignAccessScope
            Implements System.IDisposable

            Private ReadOnly _previous As DesignAccessState
            Private _disposed As System.Boolean

            Public Sub New(ByVal previous As DesignAccessState)
                _previous = previous
            End Sub

            Public Sub Dispose() Implements System.IDisposable.Dispose
                If _disposed Then Return
                _disposed = True
                _designAccessState.Value = _previous
            End Sub
        End Class

        Private Shared ReadOnly _designAccessState As New System.Threading.AsyncLocal(Of DesignAccessState)()

        Private Sub New()
        End Sub

        Public Shared Function PushAccessScope(ByVal allowDesigns As System.Boolean,
                                               ByVal allowDesignSets As System.Boolean) As System.IDisposable
            Dim previous As DesignAccessState = _designAccessState.Value
            _designAccessState.Value = New DesignAccessState() With {
                .AllowDesigns = allowDesigns,
                .AllowDesignSets = allowDesigns AndAlso allowDesignSets
            }
            Return New DesignAccessScope(previous)
        End Function

        Private Shared Function DesignsAllowed() As System.Boolean
            Dim state As DesignAccessState = _designAccessState.Value
            Return state Is Nothing OrElse state.AllowDesigns
        End Function

        Private Shared Function DesignSetsAllowed() As System.Boolean
            Dim state As DesignAccessState = _designAccessState.Value
            Return state Is Nothing OrElse (state.AllowDesigns AndAlso state.AllowDesignSets)
        End Function

        Public Shared Function GetDesigns() As IReadOnlyList(Of DocumentDesignDescriptor)
            If Not DesignsAllowed() Then Return New System.Collections.Generic.List(Of DocumentDesignDescriptor)()

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

        Public Shared Function FindDefaultDesign(applicationName As String) As DocumentDesignDescriptor
            Dim designs As IReadOnlyList(Of DocumentDesignDescriptor) = GetDesigns()
            Dim normalizedApplication As String = If(applicationName, "").Trim().ToLowerInvariant()
            Dim matches As New System.Collections.Generic.List(Of DocumentDesignDescriptor)()
            For Each d As DocumentDesignDescriptor In designs
                If d Is Nothing Then Continue For
                Dim appConfig As Newtonsoft.Json.Linq.JObject = d.GetApplicationConfig(normalizedApplication)
                If appConfig Is Nothing Then Continue For
                If appConfig.Value(Of Boolean?)("is_default").GetValueOrDefault(False) Then matches.Add(d)
            Next
            If matches.Count = 1 Then Return matches(0)
            Return Nothing
        End Function

        Public Shared Function FindBestWordDesign(documentType As String,
                                                  documentLanguage As String,
                                                  Optional organization As String = "") As DocumentDesignDescriptor
            Dim wantedType As String = NormalizeLookupKey(documentType)
            If wantedType = "" Then Return Nothing

            Dim wantedLanguage As String = NormalizeLanguageKey(documentLanguage)
            Dim wantedOrganization As String = NormalizeLookupKey(organization)
            Dim candidates As New System.Collections.Generic.List(Of DocumentDesignDescriptor)()

            For Each d As DocumentDesignDescriptor In GetDesigns()
                If d Is Nothing OrElse d.Word Is Nothing Then Continue For
                Dim candidateType As String = NormalizeLookupKey(If(d.Word.Value(Of String)("document_type"), ""))
                If candidateType <> wantedType Then Continue For
                If wantedOrganization <> "" Then
                    Dim candidateOrganization As String = NormalizeLookupKey(If(d.Word.Value(Of String)("organization"), ""))
                    If candidateOrganization <> "" AndAlso candidateOrganization <> wantedOrganization Then Continue For
                End If
                candidates.Add(d)
            Next

            If candidates.Count = 0 Then Return Nothing

            If wantedLanguage <> "" Then
                Dim exactLanguage As New System.Collections.Generic.List(Of DocumentDesignDescriptor)()
                Dim anyLanguage As New System.Collections.Generic.List(Of DocumentDesignDescriptor)()
                For Each d As DocumentDesignDescriptor In candidates
                    Dim candidateLanguage As String = NormalizeLanguageKey(If(d.Word.Value(Of String)("language"), ""))
                    If candidateLanguage = wantedLanguage Then
                        exactLanguage.Add(d)
                    ElseIf candidateLanguage = "any" OrElse candidateLanguage = "" Then
                        anyLanguage.Add(d)
                    End If
                Next
                If exactLanguage.Count = 1 Then Return exactLanguage(0)
                If exactLanguage.Count > 1 Then
                    Dim exactDefault As DocumentDesignDescriptor = FindSingleDefaultFor(exactLanguage, wantedType)
                    If exactDefault IsNot Nothing Then Return exactDefault
                    Return Nothing
                End If
                If anyLanguage.Count = 1 Then Return anyLanguage(0)
                If anyLanguage.Count > 1 Then
                    Dim anyDefault As DocumentDesignDescriptor = FindSingleDefaultFor(anyLanguage, wantedType)
                    If anyDefault IsNot Nothing Then Return anyDefault
                    Return Nothing
                End If
            End If

            If candidates.Count = 1 Then Return candidates(0)
            Return FindSingleDefaultFor(candidates, wantedType)
        End Function

        Private Shared Function FindSingleDefaultFor(candidates As IEnumerable(Of DocumentDesignDescriptor),
                                                     wantedType As String) As DocumentDesignDescriptor
            If candidates Is Nothing Then Return Nothing
            Dim matches As New System.Collections.Generic.List(Of DocumentDesignDescriptor)()
            For Each d As DocumentDesignDescriptor In candidates
                If d Is Nothing OrElse d.Word Is Nothing Then Continue For
                Dim defaultFor As String = NormalizeLookupKey(If(d.Word.Value(Of String)("default_for"), ""))
                If defaultFor = wantedType Then matches.Add(d)
            Next
            If matches.Count = 1 Then Return matches(0)
            Return Nothing
        End Function

        Private Shared Function NormalizeLanguageKey(value As String) As String
            Dim normalized As String = If(value, "").Trim().ToLowerInvariant().Replace("_", "-")
            If normalized = "" Then Return ""
            If normalized = "any" Then Return "any"
            Dim separator As Integer = normalized.IndexOf("-"c)
            If separator > 0 Then normalized = normalized.Substring(0, separator)
            Return normalized
        End Function

        Public Shared Function BuildPromptFragment(Optional maxDesigns As Integer = 24) As String
            If Not DesignsAllowed() Then
                Return "DESIGN REPOSITORY: Named internal design profiles are not authorized in the current execution context. Do not list, infer, resolve, request, or apply repository designs or design sets. Use only another concrete authorized format source supplied for this task, otherwise use neutral professional formatting."
            End If

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
                Dim wordContractNote As String = ""
                If d.Word IsNot Nothing Then
                    Dim templateRelative As String = If(d.Word.Value(Of String)("template_file"), "").Trim()
                    Dim templatePath As String = If(templateRelative = "", "", d.ResolveRepositoryFile(templateRelative))
                    Dim contract As WordTemplateBindingContract = Nothing
                    Dim contractError As String = ""
                    If WordTemplateBindingContractParser.TryLoadForDesign(d, d.Word, templatePath, contract, contractError) Then
                        If contract IsNot Nothing AndAlso (contract.HasSlots OrElse contract.HasBodyStyles OrElse contract.HasNativeStyles OrElse contract.HasAuthoringGuidance) Then
                            Dim summary As String = contract.BuildPromptSummary()
                            If summary <> "" Then wordContractNote = "; Word contract: " & summary
                        End If
                    ElseIf contractError <> "" Then
                        wordContractNote = "; Word design package INVALID: " & contractError
                    End If
                End If
                Dim routingNote As String = ""
                If d.Word IsNot Nothing Then
                    Dim documentType As String = If(d.Word.Value(Of String)("document_type"), "").Trim()
                    Dim language As String = If(d.Word.Value(Of String)("language"), "").Trim()
                    Dim organization As String = If(d.Word.Value(Of String)("organization"), "").Trim()
                    Dim isDefault As Boolean = d.Word.Value(Of Boolean?)("is_default").GetValueOrDefault(False)
                    Dim defaultFor As String = If(d.Word.Value(Of String)("default_for"), "").Trim()
                    Dim routeParts As New System.Collections.Generic.List(Of String)()
                    If documentType <> "" Then routeParts.Add("type=" & documentType)
                    If language <> "" Then routeParts.Add("language=" & language)
                    If organization <> "" Then routeParts.Add("organization=" & organization)
                    If defaultFor <> "" Then routeParts.Add("default_for=" & defaultFor)
                    If isDefault Then routeParts.Add("global_default=yes")
                    If routeParts.Count > 0 Then routingNote = "; routing: " & System.String.Join(", ", routeParts)
                End If
                items.Add($"{d.Name} [design_name={d.Id}; {String.Join("/", apps)}; {source}{routingNote}{wordContractNote}]")
            Next

            Dim suffix As String = If(designs.Count > items.Count, $" (+{designs.Count - items.Count} more)", "")
            Return "DESIGN REPOSITORY: Configured named designs: " & String.Join("; ", items) & suffix & ". " &
                   "SOURCE-FORMAT PRIORITY: if the user explicitly identifies an attachment or supplied artifact as the formatting/layout/design/master/style to preserve or copy, that source is the authorized format carrier and takes precedence over any implicit/default repository design. Do not apply a repository default merely because a create_* tool is used; when a creator is still appropriate, pass use_repository_default_design=false. An explicitly user-requested named repository design remains binding when the user actually asks for that design; do not silently mix it with a conflicting source carrier. " &
                   "For Word creation where no user-selected source format carrier applies, an explicitly user-requested named design comes first. Otherwise DOCUMENT TYPE IS THE PRIMARY ROUTING KEY: a requested memo must be routed only among memo designs, a letter only among letter designs, and a generic/blank design must never win merely because its language is broader. Within the matching document type, use language and then organization/default metadata to choose the best variant. Prefer passing document_type and document_language to create_word_document and let the host resolve the configured design; pass design_name directly only when the user explicitly named a particular design. Prefer a matching default_for entry when several same-type/same-language designs fit. If the format is genuinely ambiguous and no applicable default exists, ask the user only when ask_user is available; in unattended AutoPilot do not invent a format and explain which configured fallback/default was used so the user can give a new instruction if needed. If no Word design_name is supplied and repository defaults are enabled, the host applies the single Word entry marked global_default=yes when configured. " &
                   "When the user asks for one of these designs, pass its exact design_name to create_word_document, create_powerpoint, or create_excel_spreadsheet. " &
                   "For a Word design whose entry lists Word slots, put the substantive body in markdown_content and pass the listed template_fields keys exactly; the host binds them to the template deterministically. " &
                   "If the Word contract lists native body styles, use only the declared Markdown heading/list levels; ordinary Markdown paragraphs and declared list/heading semantics are mapped to the template's exact native Word styles by the host. " &
                   "A user-requested named design remains binding after a creator failure: never obtain apparent success by retrying the same artifact creator without its design/template constraint. " &
                   "Do not infer slot meaning from a visible placeholder name, and do not substitute a similarly named organization or claim any design not backed by this repository or another concrete authorized source."
        End Function

        Public Shared Function GetCatalogPaths() As IReadOnlyList(Of String)
            If Not DesignsAllowed() Then Return New System.Collections.Generic.List(Of System.String)()

            Dim result As New List(Of String)()
            AddCatalogPath(result, AgentResources.ConfiguredCentralPath)
            AddCatalogPath(result, AgentResources.ConfiguredLocalPath)
            Return result
        End Function

        Private Shared Function ResolveDesignsDirectory(root As String) As String
            If System.String.IsNullOrWhiteSpace(root) Then Return ""
            Try
                Dim defaultDir As String = System.IO.Path.GetFullPath(System.IO.Path.Combine(root, DesignsDirectoryName))
                If Not DesignSetsAllowed() Then Return defaultDir

                Dim selectorPath As String = System.IO.Path.Combine(root, DesignSetsDirectoryName, ActiveDesignSetFileName)
                If Not System.IO.File.Exists(selectorPath) Then Return defaultDir

                Dim selector As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(System.IO.File.ReadAllText(selectorPath, System.Text.Encoding.UTF8))
                Dim schemaVersion As Integer = 0
                System.Int32.TryParse(If(selector("schema_version"), "").ToString(), schemaVersion)
                If schemaVersion <> SupportedSchemaVersion Then Return defaultDir

                Dim activeSet As String = If(selector.Value(Of String)("active_set"), "").Trim()
                If activeSet = "" Then Return defaultDir
                ' Explicitly supported opt-out: use the neutral .inky/designs fallback and expose no design-set brand.
                If System.String.Equals(activeSet, "none", System.StringComparison.OrdinalIgnoreCase) Then Return defaultDir
                Dim sets As Newtonsoft.Json.Linq.JObject = TryCast(selector("sets"), Newtonsoft.Json.Linq.JObject)
                If sets Is Nothing Then Return defaultDir
                Dim relative As String = If(sets.Value(Of String)(activeSet), "").Trim()
                If relative = "" Then Return defaultDir

                Dim setsRoot As String = System.IO.Path.GetFullPath(System.IO.Path.Combine(root, DesignSetsDirectoryName))
                Dim selectedDir As String = System.IO.Path.GetFullPath(System.IO.Path.Combine(setsRoot, relative))
                Dim prefix As String = setsRoot.TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar) & System.IO.Path.DirectorySeparatorChar
                If Not selectedDir.StartsWith(prefix, System.StringComparison.OrdinalIgnoreCase) Then Return defaultDir
                If Not System.IO.Directory.Exists(selectedDir) Then Return defaultDir
                Return selectedDir
            Catch ex As System.Exception
                Return System.IO.Path.GetFullPath(System.IO.Path.Combine(root, DesignsDirectoryName))
            End Try
        End Function

        Private Shared Sub AddCatalogPath(result As List(Of String), root As String)
            If result Is Nothing OrElse String.IsNullOrWhiteSpace(root) Then Return
            Try
                result.Add(System.IO.Path.Combine(ResolveDesignsDirectory(root), CatalogFileName))
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
                designsDir = ResolveDesignsDirectory(root)
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
                designsDir = ResolveDesignsDirectory(root)
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
                ' a user can drop an Office carrier plus a same-basename .md into the
                ' conventional designs\<application> folder without editing designs.json.
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
