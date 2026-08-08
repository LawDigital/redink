' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: PathPolicy.vb
' Purpose: Centralized path-access policy for agent-layer file tools.
'          Enforces workspace boundaries, skill reference/script access, and
'          symlink-traversal blocking.
'
' Writable Root Precedence:
'  1. Workspace path (if maintained in current session).
'  2. Otherwise user's Desktop.
'
' Read-only Roots (always allowed for read):
'  - Skill scripts/ and references/ directories (any discovered skill).
'  - In "chat author mode", also writable under skill directories.
'  - All paths canonicalized; ".." and symlink traversal blocked.
'  - Default file size limit: 2 MiB (configurable via MaxFileSizeBytes).
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports System.Threading

Namespace Agents

    Public Enum PathAccess
        Read
        Write
    End Enum

    Public NotInheritable Class PathPolicy

        Private Sub New()
        End Sub

        ' --------------------------------------------------------------- workspace

        Private Shared _workspaceRoot As String = Nothing
        Private Shared ReadOnly _restrictToWorkspaceRootOnly As New AsyncLocal(Of Boolean)

        ' User-configured workspace permissions (mirrored from WorkspaceState by the host).
        ' These apply ONLY to paths that resolve under the workspace root; skill and
        ' staging/Desktop roots are governed by their own gates and are unaffected.
        Private Shared _workspaceAllowRead As Boolean = True
        Private Shared _workspaceAllowWrite As Boolean = True
        Private Shared _workspaceAllowMoveCopyRename As Boolean = True
        Private Shared _workspaceAllowDelete As Boolean = False

        ''' <summary>Sets the active workspace root for this process (host call). Pass Nothing to clear.</summary>
        Public Shared Sub SetWorkspaceRoot(rootOrNothing As String)
            If String.IsNullOrWhiteSpace(rootOrNothing) Then
                _workspaceRoot = Nothing
            Else
                Try
                    _workspaceRoot = Path.GetFullPath(rootOrNothing)
                Catch
                    _workspaceRoot = Nothing
                End Try
            End If
        End Sub

        Public Shared ReadOnly Property WorkspaceRoot As String
            Get
                Return _workspaceRoot
            End Get
        End Property

        ''' <summary>
        ''' Mirrors the user-configured workspace permissions into the policy. Hosts call
        ''' this whenever the active workspace state changes. Read/Write are enforced by
        ''' <see cref="Resolve"/>; MoveCopyRename/Delete are exposed for destructive file
        ''' tools to consult (Resolve only distinguishes Read from Write).
        ''' </summary>
        Public Shared Sub SetWorkspacePermissions(allowRead As Boolean,
                                                  allowWrite As Boolean,
                                                  allowMoveCopyRename As Boolean,
                                                  allowDelete As Boolean)
            _workspaceAllowRead = allowRead
            _workspaceAllowWrite = allowWrite
            _workspaceAllowMoveCopyRename = allowMoveCopyRename
            _workspaceAllowDelete = allowDelete
        End Sub

        Public Shared ReadOnly Property WorkspaceAllowMoveCopyRename As Boolean
            Get
                Return _workspaceAllowMoveCopyRename
            End Get
        End Property

        Public Shared ReadOnly Property WorkspaceAllowDelete As Boolean
            Get
                Return _workspaceAllowDelete
            End Get
        End Property

        ''' <summary>
        ''' When enabled, path resolution is limited strictly to the active workspace root.
        ''' Desktop fallback and skill roots are ignored.
        ''' </summary>
        Public Shared Property RestrictToWorkspaceRootOnly As Boolean
            Get
                Return _restrictToWorkspaceRootOnly.Value
            End Get
            Set(value As Boolean)
                _restrictToWorkspaceRootOnly.Value = value
            End Set
        End Property

        Private Shared _strictExtraRoots As String() = Nothing

        ''' <summary>
        ''' Registers additional roots that are also permitted (read and write) while
        ''' <see cref="RestrictToWorkspaceRootOnly"/> is active. Hosts use this to allow,
        ''' for example, both the AutoPilot temp directory and the scheduled-task
        ''' workspace directory during a locked run. Pass Nothing to clear.
        ''' </summary>
        Public Shared Sub SetStrictExtraRoots(rootsOrNothing As IEnumerable(Of String))
            If rootsOrNothing Is Nothing Then
                _strictExtraRoots = Nothing
                Return
            End If

            Dim collected As New List(Of String)()
            For Each r In rootsOrNothing
                If String.IsNullOrWhiteSpace(r) Then Continue For
                Try
                    collected.Add(Path.GetFullPath(r))
                Catch
                End Try
            Next
            _strictExtraRoots = If(collected.Count > 0, collected.ToArray(), Nothing)
        End Sub

        ' --------------------------------------------------------------- session staging root

        ' The active session's staging/temp directory (host-provided). Files produced by
        ' tools into this directory are always readable and writable, independently of the
        ' workspace/Desktop roots, so tool producers and consumers share common ground even
        ' when a workspace is connected. Hosts set this on session start and clear it on end.
        Private Shared _sessionStagingRoot As String = Nothing

        ''' <summary>Sets the active session staging root (host call). Pass Nothing to clear.</summary>
        Public Shared Sub SetSessionStagingRoot(rootOrNothing As String)
            If String.IsNullOrWhiteSpace(rootOrNothing) Then
                _sessionStagingRoot = Nothing
            Else
                Try
                    _sessionStagingRoot = Path.GetFullPath(rootOrNothing)
                Catch
                    _sessionStagingRoot = Nothing
                End Try
            End If
        End Sub

        Public Shared ReadOnly Property SessionStagingRoot As String
            Get
                Return _sessionStagingRoot
            End Get
        End Property

        ' --------------------------------------------------------------- chat-author scope

        Private Shared ReadOnly _chatAuthor As New AsyncLocal(Of Boolean)

        ''' <summary>
        ''' Marks the current async flow as "chat author" — permits writes under skill
        ''' scripts/references for the duration of the returned scope. Use:
        '''     Using PathPolicy.BeginChatAuthorScope() : ... End Using
        ''' </summary>
        Public Shared Function BeginChatAuthorScope() As IDisposable
            Return New ChatAuthorScope()
        End Function

        Private Class ChatAuthorScope
            Implements IDisposable
            Private ReadOnly _previous As Boolean
            Public Sub New()
                _previous = _chatAuthor.Value
                _chatAuthor.Value = True
            End Sub
            Public Sub Dispose() Implements IDisposable.Dispose
                _chatAuthor.Value = _previous
            End Sub
        End Class

        ' --------------------------------------------------------------- size limits

        Public Shared Property MaxFileSizeBytes As Integer = 2 * 1024 * 1024 ' 2 MiB

        ' --------------------------------------------------------------- writable root

        ''' <summary>
        ''' Returns the canonical writable root for new agent-created files.
        ''' Workspace if maintained, otherwise the current user's Desktop.
        ''' </summary>
        Public Shared Function GetDefaultWritableRoot() As String
            If Not String.IsNullOrWhiteSpace(_workspaceRoot) AndAlso Directory.Exists(_workspaceRoot) Then
                Return _workspaceRoot
            End If
            If Not String.IsNullOrWhiteSpace(_sessionStagingRoot) AndAlso Directory.Exists(_sessionStagingRoot) Then
                Return _sessionStagingRoot
            End If
            Return Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        End Function

        ' --------------------------------------------------------------- validation

        ''' <summary>
        ''' Resolves <paramref name="requestedPath"/> against the policy. Returns the fully
        ''' qualified canonical path on success, or throws <see cref="UnauthorizedAccessException"/>
        ''' on denial.
        '''
        ''' Behavior:
        '''  - If <paramref name="requestedPath"/> is relative, it is resolved against the
        '''    default writable root (workspace or Desktop).
        '''  - Read access is allowed under: workspace, Desktop, and any discovered skill
        '''    scripts/ or references/ directory.
        '''  - Write access is allowed under: workspace and Desktop; additionally under a
        '''    skill's scripts/references when chat-author scope is active.
        ''' </summary>
        Public Shared Function Resolve(requestedPath As String, access As PathAccess) As String
            If String.IsNullOrWhiteSpace(requestedPath) Then
                Throw New ArgumentException("Path is empty.", NameOf(requestedPath))
            End If

            Dim wasRelative As Boolean = Not Path.IsPathRooted(requestedPath)

            Dim full As String
            Try
                If Not wasRelative Then
                    full = Path.GetFullPath(requestedPath)
                Else
                    full = Path.GetFullPath(Path.Combine(GetDefaultWritableRoot(), requestedPath))
                End If
            Catch ex As Exception
                Throw New UnauthorizedAccessException("Invalid path: " & ex.Message)
            End Try

            ' Block raw devices and UNC by default.
            If full.StartsWith("\\?\", StringComparison.Ordinal) OrElse full.StartsWith("\\.\", StringComparison.Ordinal) Then
                Throw New UnauthorizedAccessException("Device paths are not allowed.")
            End If

            Dim ws = If(_workspaceRoot, "")

            ' Enforce user-configured workspace permissions for any path that resolves
            ' under the workspace root. Skill and staging/Desktop roots are governed by
            ' their own gates and are intentionally unaffected here.
            If Not String.IsNullOrWhiteSpace(ws) AndAlso IsUnder(full, Path.GetFullPath(ws)) Then
                If access = PathAccess.Read AndAlso Not _workspaceAllowRead Then
                    Throw New UnauthorizedAccessException("Workspace read access is disabled.")
                End If
                If access = PathAccess.Write AndAlso Not _workspaceAllowWrite Then
                    Throw New UnauthorizedAccessException("Workspace write access is disabled.")
                End If
            End If

            If _restrictToWorkspaceRootOnly.Value Then
                ' Collect the strict allow-set: the active workspace root plus any
                ' additional host-registered roots (e.g. the AutoPilot temp directory and
                ' the scheduled-task workspace). Both read and write are confined to these.
                Dim strictRoots As New List(Of String)()
                If Not String.IsNullOrWhiteSpace(ws) AndAlso Directory.Exists(ws) Then
                    strictRoots.Add(Path.GetFullPath(ws))
                End If
                If Not String.IsNullOrWhiteSpace(_sessionStagingRoot) AndAlso Directory.Exists(_sessionStagingRoot) Then
                    strictRoots.Add(Path.GetFullPath(_sessionStagingRoot))
                End If
                Dim extra = _strictExtraRoots
                If extra IsNot Nothing Then
                    For Each r In extra
                        If Not String.IsNullOrWhiteSpace(r) AndAlso Directory.Exists(r) Then
                            strictRoots.Add(Path.GetFullPath(r))
                        End If
                    Next
                End If

                If strictRoots.Count = 0 Then
                    Throw New UnauthorizedAccessException("Workspace-only mode is active but no workspace root is configured.")
                End If

                ' Writes are confined strictly to the allowed roots.
                For Each r In strictRoots
                    If IsUnder(full, r) Then Return full
                Next

                ' Reads are additionally permitted under discovered skill directories
                ' (SKILL.md, scripts/ and references/), which are always read-only sources.
                ' This lets a skill load its bundled reference files even while the
                ' workspace is otherwise locked to the temp/scheduled directory.
                If access = PathAccess.Read Then
                    Dim skillReadResult As String = TryResolveUnderSkillDirectories(full, requestedPath, wasRelative)
                    If skillReadResult IsNot Nothing Then Return skillReadResult
                End If

                Throw New UnauthorizedAccessException("Path is outside the active workspace root.")
            End If

            ' Build allow-set.
            Dim writeRoots As New List(Of String)()
            Dim readRoots As New List(Of String)()

            If Not String.IsNullOrWhiteSpace(ws) AndAlso Directory.Exists(ws) Then
                writeRoots.Add(Path.GetFullPath(ws))
                readRoots.Add(Path.GetFullPath(ws))
            End If
            Dim desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            If Not String.IsNullOrWhiteSpace(desktop) Then
                writeRoots.Add(Path.GetFullPath(desktop))
                readRoots.Add(Path.GetFullPath(desktop))
            End If
            If Not String.IsNullOrWhiteSpace(_sessionStagingRoot) AndAlso Directory.Exists(_sessionStagingRoot) Then
                writeRoots.Add(Path.GetFullPath(_sessionStagingRoot))
                readRoots.Add(Path.GetFullPath(_sessionStagingRoot))
            End If

            ' Skill scripts/references — always readable.
            ' In skill-author mode (or legacy chat-author scope), the ENTIRE skill folder
            ' is writable (so SKILL.md plus scripts/ and references/ can be edited).
            Dim skillReadRoots As New List(Of String)()
            Dim skillFullRoots As New List(Of String)()
            Try
                For Each sk In AgentResources.Skills
                    If sk Is Nothing OrElse String.IsNullOrWhiteSpace(sk.DirectoryPath) Then Continue For
                    Dim skFull As String = Path.GetFullPath(sk.DirectoryPath)
                    skillFullRoots.Add(skFull)
                    Dim sdir As String = Path.Combine(skFull, "scripts")
                    Dim rdir As String = Path.Combine(skFull, "references")
                    If Directory.Exists(sdir) Then skillReadRoots.Add(Path.GetFullPath(sdir))
                    If Directory.Exists(rdir) Then skillReadRoots.Add(Path.GetFullPath(rdir))
                Next
            Catch
            End Try

            ' Existing agent folders (agents/<name>/ or the agents/ base for single-file agents)
            ' are always readable so the author can inspect and revise them.
            Dim agentFullRoots As New List(Of String)()
            Try
                For Each ag In AgentResources.Agents
                    If ag Is Nothing OrElse String.IsNullOrWhiteSpace(ag.DirectoryPath) Then Continue For
                    agentFullRoots.Add(Path.GetFullPath(ag.DirectoryPath))
                Next
            Catch
            End Try

            ' Resource base directories (skills/ and agents/) so brand-new skills and agents
            ' (and their subdirectories) can be created in author mode. Local is always
            ' permitted; central requires the explicit opt-in flag.
            Dim authorBaseRoots As New List(Of String)()
            Try
                For Each baseDir In AgentResources.GetLocalResourceBaseDirectories()
                    If Not String.IsNullOrWhiteSpace(baseDir) Then authorBaseRoots.Add(Path.GetFullPath(baseDir))
                Next
                If SkillAuthorMode.AllowCentralWrites Then
                    For Each baseDir In AgentResources.GetCentralResourceBaseDirectories()
                        If Not String.IsNullOrWhiteSpace(baseDir) Then authorBaseRoots.Add(Path.GetFullPath(baseDir))
                    Next
                End If
            Catch
            End Try

            readRoots.AddRange(skillReadRoots)
            readRoots.AddRange(skillFullRoots)
            readRoots.AddRange(agentFullRoots)
            readRoots.AddRange(authorBaseRoots)

            ' Local diagnostics folder (previous tooling-run logs) is always readable so the
            ' skill-author skill can inspect the last run to diagnose the tooling loop.
            Try
                Dim diagLocalRoot As String = AgentResources.ConfiguredLocalPath
                If Not String.IsNullOrWhiteSpace(diagLocalRoot) Then
                    Dim diagDir As String = Path.GetFullPath(Path.Combine(diagLocalRoot, "diagnostics"))
                    If Directory.Exists(diagDir) Then readRoots.Add(diagDir)
                End If
            Catch
            End Try
            If _chatAuthor.Value OrElse SkillAuthorMode.IsActive Then
                writeRoots.AddRange(skillFullRoots)
                writeRoots.AddRange(agentFullRoots)
                writeRoots.AddRange(authorBaseRoots)
            End If

            Dim roots = If(access = PathAccess.Write, writeRoots, readRoots)
            For Each r In roots
                If IsUnder(full, r) Then Return full
            Next

            ' For read access, allow a relative path to resolve against a skill directory
            ' when it does not exist under the normal roots (e.g. "references/spec.md").
            If access = PathAccess.Read Then
                Dim skillReadResult As String = TryResolveUnderSkillDirectories(full, requestedPath, wasRelative)
                If skillReadResult IsNot Nothing Then Return skillReadResult
            End If

            Throw New UnauthorizedAccessException("Path is outside the allowed roots for " & access.ToString().ToLowerInvariant() & " access.")
        End Function

        ''' <summary>
        ''' Attempts to resolve a read path under any discovered skill directory. Returns
        ''' the canonical path when the (absolute) path already lies under a skill folder,
        ''' or when the original request was relative and combining it with a skill folder
        ''' yields an existing file. Returns Nothing when no skill match is found.
        ''' </summary>
        Private Shared Function TryResolveUnderSkillDirectories(full As String, requestedPath As String, wasRelative As Boolean) As String
            Try
                For Each sk In AgentResources.Skills
                    If sk Is Nothing OrElse String.IsNullOrWhiteSpace(sk.DirectoryPath) Then Continue For
                    Dim skFull As String = Path.GetFullPath(sk.DirectoryPath)

                    ' Absolute path already inside a skill directory.
                    If IsUnder(full, skFull) Then Return full

                    ' Relative request resolved against the skill directory.
                    If wasRelative AndAlso Not String.IsNullOrWhiteSpace(requestedPath) Then
                        Dim candidate As String = Path.GetFullPath(Path.Combine(skFull, requestedPath))
                        If IsUnder(candidate, skFull) AndAlso File.Exists(candidate) Then
                            Return candidate
                        End If
                    End If
                Next
            Catch
            End Try
            Return Nothing
        End Function

        Private Shared Function IsUnder(candidate As String, root As String) As Boolean
            If String.IsNullOrWhiteSpace(candidate) OrElse String.IsNullOrWhiteSpace(root) Then Return False
            Dim a = candidate.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            Dim b = root.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            If String.Equals(a, b, StringComparison.OrdinalIgnoreCase) Then Return True
            Return a.StartsWith(b & Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)
        End Function

        ' --------------------------------------------------------------- writable-name helper

        ''' <summary>
        ''' Returns a non-colliding writable path inside the default writable root for a
        ''' suggested filename. If the suggested filename already exists, " (n)" is appended.
        ''' </summary>
        Public Shared Function NewWritablePath(suggestedFileName As String) As String
            Dim root = GetDefaultWritableRoot()
            Dim safe = SanitizeFileName(If(suggestedFileName, "untitled.txt"))
            Dim candidate = Path.Combine(root, safe)
            If Not File.Exists(candidate) Then Return candidate
            Dim baseName = Path.GetFileNameWithoutExtension(safe)
            Dim ext = Path.GetExtension(safe)
            For i = 2 To 1000
                Dim p = Path.Combine(root, baseName & " (" & i.ToString() & ")" & ext)
                If Not File.Exists(p) Then Return p
            Next
            Return Path.Combine(root, baseName & "_" & Guid.NewGuid().ToString("N").Substring(0, 8) & ext)
        End Function

        Private Shared Function SanitizeFileName(name As String) As String
            If String.IsNullOrWhiteSpace(name) Then Return "untitled.txt"
            Dim invalid = Path.GetInvalidFileNameChars()
            Dim sb As New System.Text.StringBuilder(name.Length)
            For Each c In name
                If Array.IndexOf(invalid, c) >= 0 Then sb.Append("_"c) Else sb.Append(c)
            Next
            Return sb.ToString()
        End Function

    End Class

End Namespace
