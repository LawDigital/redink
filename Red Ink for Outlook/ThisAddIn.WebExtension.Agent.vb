' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved.
' For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.WebExtension.Agent.vb
' Purpose: Chat Agent mode — provides the same document-processing tooling
'          available in AutoPilot to the browser-based local chat. Users can
'          load multiple files, and output files are collected into a Desktop
'          folder after the LLM tooling loop completes.
'
' Architecture:
'   - Shares the AutoPilot internal tool infrastructure via the same shared
'     fields (_apCurrentTempDir, _apCurrentAttachments, _apCurrentMailInfo).
'   - Blocked when AutoPilot is actively running to avoid state collision.
'   - Creates a per-job output folder on Desktop\Inky\yymmdd_hh-mm and copies
'     result files there after the tooling loop finishes.
'   - The summarize_thread tool is excluded (e-mail-specific).
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports Microsoft.VisualBasic.FileIO
Imports Excel = Microsoft.Office.Interop.Excel
Imports Word = Microsoft.Office.Interop.Word
Imports VBFileIO = Microsoft.VisualBasic.FileIO

Partial Public Class ThisAddIn

    ' ═══════════════════════════════════════════════════════════════════════════
    '  CHAT AGENT STATE
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>True while a chat agent tooling job is executing. Enables internal tool routing.</summary>
    Friend _chatAgentActive As Boolean = False

    ''' <summary>Files loaded by the user for the current chat agent session.</summary>
    Private _chatAgentFiles As New List(Of AutoPilotAttachmentInfo)()

    ''' <summary>
    ''' Snapshot of file paths present in _chatAgentTempDir at the START of the
    ''' current turn. Files in this set were already surfaced in a previous turn
    ''' (or are user uploads), so they MUST NOT be presented again. Reset by
    ''' ResetChatAgentDeliverableTrackingForNewTurn() at the start of every
    ''' non-sub-agent tooling run while the chat agent is active.
    ''' </summary>
    Private _chatAgentSurfacedFiles As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

    ''' <summary>
    ''' Paths that a tool edited in place during the current turn and that must be
    ''' delivered even though they pre-existed (and are therefore in
    ''' _chatAgentSurfacedFiles). Populated by tools such as the live Excel workbook
    ''' completion when editing the existing file in place. Reset at the start of
    ''' every non-sub-agent tooling run.
    ''' </summary>
    Private _chatAgentForcedDeliverables As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

    ''' <summary>Temp directory used by the current chat agent session.</summary>
    Private _chatAgentTempDir As String = Nothing

    ''' <summary>Prefix for chat agent temp directories.</summary>
    Private Const CA_TempPrefix As String = AN2 & "_chatagent_"

    Private Const CA_Tool_WorkspaceList As String = "agent_workspace_list"
    Private Const CA_Tool_WorkspaceRead As String = "agent_workspace_read"
    Private Const CA_Tool_WorkspaceWrite As String = "agent_workspace_write"
    Private Const CA_Tool_WorkspaceFileOp As String = "agent_workspace_file_op"
    Private Const CA_Tool_WorkspaceSaveSessionFile As String = "agent_workspace_save_session_file"
    Private Const CA_Tool_WorkspaceSearch As String = "agent_workspace_search"
    Private Const CA_Tool_WorkspaceFindFiles As String = "agent_workspace_find_files"
    Private Const CA_Tool_WorkspaceMoveTo As String = "agent_workspace_move_to"
    Private Const CA_Tool_WorkspaceCopyTo As String = "agent_workspace_copy_to"
    Private Const CA_Tool_WorkspaceRename As String = "agent_workspace_rename"
    Private Const CA_Tool_WorkspaceBulkRename As String = "agent_workspace_bulk_rename"
    Private Const CA_Tool_WorkspaceFileDetails As String = "agent_workspace_file_details"
    Private Const CA_Tool_WorkspaceRecentFiles As String = "agent_workspace_recent_files"
    Private Const CA_Tool_WorkspaceCreateFolderStructure As String = "agent_workspace_create_folder_structure"
    Private Const CA_Tool_WorkspaceTrash As String = "agent_workspace_trash"
    Private Const CA_Tool_WorkspaceInventoryReport As String = "agent_workspace_inventory_report"

    Private _chatAgentWorkspace As ChatAgentWorkspaceState = Nothing
    Private _chatAgentWorkspaceLoaded As Boolean = False

    Private ReadOnly _chatAgentWorkspaceHistory As New List(Of String)()
    Private Const CA_MaxWorkspaceHistoryEntries As Integer = 12
    Private Const CA_WorkspacePresetCount As Integer = 5
    Private Const CA_WorkspacePresetSettingPrefix As String = "Inky_WorkspacePreset"

    Friend Class ChatAgentWorkspaceState
        Public Property RootPath As String = ""
        Public Property PersistUntilRevoked As Boolean = False
        Public Property AllowRead As Boolean = True
        Public Property AllowWrite As Boolean = True
        Public Property AllowMoveCopyRename As Boolean = True
        Public Property AllowDelete As Boolean = SharedMethods.WorkspaceDeleteByDefaultOn
        Public Property SaveDroppedFilesToWorkspace As Boolean = False
        Public Property IncludeHiddenSystem As Boolean = False
    End Class

    Private Class ChatAgentWorkspacePresetInfo
        Public Property Slot As Integer
        Public Property RootPath As String = ""
        Public Property Label As String = ""
    End Class

    Private Function BuildAutoPilotTempWorkspaceState() As ChatAgentWorkspaceState
        If String.IsNullOrWhiteSpace(_apCurrentTempDir) OrElse Not Directory.Exists(_apCurrentTempDir) Then
            Return Nothing
        End If

        Return New ChatAgentWorkspaceState() With {
            .RootPath = Path.GetFullPath(_apCurrentTempDir),
            .PersistUntilRevoked = False,
            .AllowRead = True,
            .AllowWrite = True,
            .AllowMoveCopyRename = True,
            .AllowDelete = True,
            .SaveDroppedFilesToWorkspace = False,
            .IncludeHiddenSystem = True
        }
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  FILE MANAGEMENT
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Adds a file from disk to the chat agent file list. The file is copied to a
    ''' per-session temp directory so tools can operate on it safely.
    ''' The source file is deleted after a successful copy to avoid orphaned uploads.
    ''' </summary>
    Private Function ChatAgentAddFile(sourcePath As String) As AutoPilotAttachmentInfo
        If String.IsNullOrWhiteSpace(sourcePath) OrElse Not File.Exists(sourcePath) Then Return Nothing

        ' Ensure per-session temp dir exists (also reclaims stale dirs and registers the
        ' staging root with PathPolicy, so uploads and tool outputs share the same session).
        EnsureChatAgentTempDir()

        ' The upload handler (inky_upload) saves files as "{32-hex-guid}_{originalName}".
        ' Strip that GUID prefix so tools see the user's original filename.
        Dim fileName = Path.GetFileName(sourcePath)
        If fileName.Length > 33 AndAlso fileName(32) = "_"c Then
            Dim possibleGuid = fileName.Substring(0, 32)
            Dim g As Guid
            If Guid.TryParse(possibleGuid, g) Then
                fileName = fileName.Substring(33)
            End If
        End If

        Dim destPath = Path.Combine(_chatAgentTempDir, fileName)

        ' Prevent collision
        Dim counter = 1
        While File.Exists(destPath)
            Dim baseName = Path.GetFileNameWithoutExtension(fileName)
            Dim ext = Path.GetExtension(fileName)
            destPath = Path.Combine(_chatAgentTempDir, baseName & $"_{counter}{ext}")
            counter += 1
        End While

        File.Copy(sourcePath, destPath, overwrite:=True)

        ChatAgentWorkspaceSaveDroppedFile(destPath)

        ' Delete the source upload file — the copy in the agent temp dir is the working copy
        Try
            File.Delete(sourcePath)
        Catch
        End Try

        Dim fi = New FileInfo(destPath)
        Dim att As New AutoPilotAttachmentInfo() With {
            .OriginalFileName = Path.GetFileName(destPath),
            .TempFilePath = destPath,
            .Extension = fi.Extension.ToLowerInvariant(),
            .SizeBytes = fi.Length,
            .IsOverSizeLimit = False,
            .StatusMessage = "OK",
            .CreatedTime = fi.CreationTimeUtc,
            .LastModifiedTime = fi.LastWriteTimeUtc,
            .OutputFiles = New List(Of String)(),
            .IsToolOutput = False
        }

        _chatAgentFiles.Add(att)
        Return att
    End Function

    ''' <summary>
    ''' Returns a browser-friendly list of loaded agent files.
    ''' </summary>
    Private Function GetAgentFileListForBrowser() As List(Of Object)
        Dim result As New List(Of Object)()

        ' Surface any files produced by tools into the session staging directory as chips,
        ' so users can see and open/download them alongside drag-and-dropped uploads.
        RegisterStagingFilesAsSessionChips()

        For Each att In _chatAgentFiles
            Dim hydrated = EnsureSessionAttachmentAvailable(att)
            If hydrated Is Nothing Then Continue For

            result.Add(New With {
                .name = hydrated.OriginalFileName,
                .size = hydrated.SizeBytes,
                .ext = hydrated.Extension,
                .overLimit = hydrated.IsOverSizeLimit
            })
        Next

        Return result
    End Function

    ''' <summary>
    ''' Clears all chat agent files and removes the temp directory.
    ''' </summary>
    Private Sub ChatAgentClearFiles()
        CleanupChatAgentTempDir()
    End Sub

    ''' <summary>
    ''' Resolves a registered session file by name (uploaded or produced) and returns its
    ''' current temp path, hydrating it from its source if necessary. Returns Nothing when
    ''' the file cannot be located or materialized.
    ''' </summary>
    Private Function ResolveSessionFilePath(fileName As String) As String
        If String.IsNullOrWhiteSpace(fileName) Then Return Nothing

        Dim trimmedName = fileName.Trim()

        ' The browser chips are rendered from _chatAgentFiles (GetAgentFileListForBrowser),
        ' so resolve against that list first. _apCurrentAttachments (used by FindAttachment)
        ' is only populated during an active tooling run and is often Nothing here.
        Dim att As AutoPilotAttachmentInfo = Nothing
        If _chatAgentFiles IsNot Nothing Then
            att = _chatAgentFiles.FirstOrDefault(
                Function(a) a IsNot Nothing AndAlso
                            Not String.IsNullOrWhiteSpace(a.OriginalFileName) AndAlso
                            a.OriginalFileName.Equals(trimmedName, StringComparison.OrdinalIgnoreCase))
        End If

        If att Is Nothing Then
            att = FindAttachment(trimmedName)
        End If

        att = EnsureSessionAttachmentAvailable(att)

        If att Is Nothing OrElse String.IsNullOrWhiteSpace(att.TempFilePath) OrElse Not File.Exists(att.TempFilePath) Then
            Return Nothing
        End If

        Return att.TempFilePath
    End Function

    ''' <summary>
    ''' Opens a registered session file (uploaded or produced) in the OS default application.
    ''' Returns False when the file cannot be located.
    ''' </summary>
    Private Function ChatAgentOpenFileInDefaultApp(fileName As String) As Boolean
        Dim path = ResolveSessionFilePath(fileName)
        If String.IsNullOrWhiteSpace(path) Then Return False

        Try
            Dim psi As New ProcessStartInfo(path) With {.UseShellExecute = True}
            Process.Start(psi)
            Return True
        Catch ex As Exception
            ToolingFileLogger.LogWarn("Failed to open session file in default application.", ex:=ex)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Resolves a registered session file by name (uploaded or produced) and copies it to
    ''' Desktop\Inky\, returning the saved path. Works for any file surfaced as a browser chip.
    ''' </summary>
    Private Function ChatAgentDownloadFileToDesktop(fileName As String) As String
        Dim path = ResolveSessionFilePath(fileName)
        If String.IsNullOrWhiteSpace(path) Then Return Nothing

        Dim desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
        Dim outputDir = System.IO.Path.Combine(desktopPath, "Inky")
        Directory.CreateDirectory(outputDir)

        Dim destName = System.IO.Path.GetFileName(path)
        Dim destPath = System.IO.Path.Combine(outputDir, destName)

        Dim counter = 1
        While File.Exists(destPath)
            destPath = System.IO.Path.Combine(
                outputDir,
                System.IO.Path.GetFileNameWithoutExtension(destName) & $"_{counter}" & System.IO.Path.GetExtension(destName))
            counter += 1
        End While

        File.Copy(path, destPath, overwrite:=False)
        Return destPath
    End Function

    Private Function EnsureSessionAttachmentAvailable(att As AutoPilotAttachmentInfo) As AutoPilotAttachmentInfo
        If att Is Nothing Then Return Nothing

        If Not String.IsNullOrWhiteSpace(att.TempFilePath) AndAlso File.Exists(att.TempFilePath) Then
            Return att
        End If

        If String.IsNullOrWhiteSpace(att.SourcePath) OrElse Not File.Exists(att.SourcePath) Then
            Try
                _chatAgentFiles.Remove(att)
            Catch
            End Try

            If _apCurrentAttachments IsNot Nothing Then
                Try
                    _apCurrentAttachments.Remove(att)
                Catch
                End Try
            End If

            Return Nothing
        End If

        Dim tempDir = EnsureChatAgentTempDir()
        Dim destName = Path.GetFileName(att.SourcePath)
        Dim destPath = Path.Combine(tempDir, destName)

        Dim counter = 1
        While File.Exists(destPath)
            destPath = Path.Combine(
                tempDir,
                Path.GetFileNameWithoutExtension(destName) & $"_{counter}" & Path.GetExtension(destName))
            counter += 1
        End While

        File.Copy(att.SourcePath, destPath, overwrite:=False)

        Dim fi As New FileInfo(destPath)
        att.TempFilePath = destPath
        att.OriginalFileName = Path.GetFileName(destPath)
        att.Extension = fi.Extension.ToLowerInvariant()
        att.SizeBytes = fi.Length
        att.CreatedTime = fi.CreationTimeUtc
        att.LastModifiedTime = fi.LastWriteTimeUtc
        If String.IsNullOrWhiteSpace(att.StatusMessage) Then
            att.StatusMessage = "Rehydrated from workspace source"
        End If

        If _apCurrentAttachments IsNot Nothing AndAlso Not _apCurrentAttachments.Contains(att) Then
            _apCurrentAttachments.Add(att)
        End If

        Return att
    End Function

    ''' <summary>
    ''' Creates a safe snapshot of the current chat-agent tool-call log before teardown clears it.
    ''' </summary>
    Private Function CloneChatAgentToolCallLog() As List(Of AutoPilotToolCallEntry)
        Dim result As New List(Of AutoPilotToolCallEntry)()

        If _apCurrentToolCallLog Is Nothing OrElse _apCurrentToolCallLog.Count = 0 Then
            Return result
        End If

        For Each entry In _apCurrentToolCallLog
            If entry Is Nothing Then Continue For

            result.Add(New AutoPilotToolCallEntry() With {
                .ToolName = entry.ToolName,
                .ToolDisplayName = entry.ToolDisplayName,
                .ParamSummary = entry.ParamSummary,
                .IsInternalTool = entry.IsInternalTool,
                .WasSuccessful = entry.WasSuccessful,
                .ResultExcerpt = entry.ResultExcerpt,
                .Elapsed = entry.Elapsed,
                .Urls = If(entry.Urls IsNot Nothing, New List(Of String)(entry.Urls), Nothing)
            })
        Next

        Return result
    End Function


    ''' <summary>
    ''' Returns True when the chat-agent run ended abnormally and should show
    ''' a partial report instead of only the raw abort/error text.
    ''' </summary>
    Private Function ShouldBuildChatAgentAbortReport(partialOutput As String, wasCancelled As Boolean) As Boolean
        If wasCancelled Then Return True

        Dim normalized As String = If(partialOutput, "").Trim()
        If String.IsNullOrWhiteSpace(normalized) Then Return False

        If normalized.Equals("Operation was canceled by the user.", StringComparison.OrdinalIgnoreCase) Then Return True
        If normalized.Equals("Operation cancelled due to power transition.", StringComparison.OrdinalIgnoreCase) Then Return True
        If normalized.StartsWith("Tool execution failed:", StringComparison.OrdinalIgnoreCase) Then Return True
        If normalized.StartsWith("Aborting because the same failing tool call was repeated", StringComparison.OrdinalIgnoreCase) Then Return True

        Return False
    End Function

    ''' <summary>
    ''' Builds a partial markdown report for an aborted or failed chat-agent run.
    ''' </summary>
    Private Function BuildChatAgentAbortReport(
        toolCallLog As List(Of AutoPilotToolCallEntry),
        outputFiles As List(Of String),
        partialModelOutput As String) As String

        Dim sb As New StringBuilder()
        Dim successfulEntries As New List(Of AutoPilotToolCallEntry)()
        Dim failedEntries As New List(Of AutoPilotToolCallEntry)()

        If toolCallLog IsNot Nothing Then
            For Each entry In toolCallLog
                If entry Is Nothing Then Continue For

                If entry.WasSuccessful Then
                    successfulEntries.Add(entry)
                Else
                    failedEntries.Add(entry)
                End If
            Next
        End If

        Dim normalizedOutput As String = If(partialModelOutput, "").Trim()

        sb.AppendLine("⚠️ The agent run stopped before completion.")
        sb.AppendLine()

        If normalizedOutput.StartsWith("Tool execution failed:", StringComparison.OrdinalIgnoreCase) Then
            sb.AppendLine("### Failure reason")
            sb.AppendLine(normalizedOutput)
            sb.AppendLine()
        ElseIf normalizedOutput.StartsWith("Aborting because the same failing tool call was repeated", StringComparison.OrdinalIgnoreCase) Then
            sb.AppendLine("### Failure reason")
            sb.AppendLine(normalizedOutput)
            sb.AppendLine()
        ElseIf normalizedOutput.Equals("Operation was canceled by the user.", StringComparison.OrdinalIgnoreCase) Then
            sb.AppendLine("### Failure reason")
            sb.AppendLine("The run was canceled by the user.")
            sb.AppendLine()
        ElseIf normalizedOutput.Equals("Operation cancelled due to power transition.", StringComparison.OrdinalIgnoreCase) Then
            sb.AppendLine("### Failure reason")
            sb.AppendLine("The run was interrupted by a power transition.")
            sb.AppendLine()
        End If

        If successfulEntries.Count > 0 Then
            sb.AppendLine("### Completed steps")
            For Each entry In successfulEntries
                Dim toolLabel As String = If(Not String.IsNullOrWhiteSpace(entry.ToolDisplayName), entry.ToolDisplayName, entry.ToolName)
                If String.IsNullOrWhiteSpace(toolLabel) Then toolLabel = "Tool"

                Dim line As String = "- " & toolLabel

                If Not String.IsNullOrWhiteSpace(entry.ParamSummary) Then
                    line &= " — " & entry.ParamSummary.Trim()
                End If

                If entry.Elapsed.TotalSeconds > 0 Then
                    Dim elapsedSeconds As Integer = Math.Max(1, CInt(Math.Round(entry.Elapsed.TotalSeconds, MidpointRounding.AwayFromZero)))
                    line &= $" ({elapsedSeconds}s)"
                End If

                sb.AppendLine(line)

                If Not String.IsNullOrWhiteSpace(entry.ResultExcerpt) Then
                    sb.AppendLine("  - Result: " & entry.ResultExcerpt.Trim())
                End If
            Next
            sb.AppendLine()
        End If

        If failedEntries.Count > 0 Then
            sb.AppendLine("### Failed or unfinished steps")
            For Each entry In failedEntries
                Dim toolLabel As String = If(Not String.IsNullOrWhiteSpace(entry.ToolDisplayName), entry.ToolDisplayName, entry.ToolName)
                If String.IsNullOrWhiteSpace(toolLabel) Then toolLabel = "Tool"

                Dim line As String = "- " & toolLabel

                If Not String.IsNullOrWhiteSpace(entry.ParamSummary) Then
                    line &= " — " & entry.ParamSummary.Trim()
                End If

                sb.AppendLine(line)
            Next
            sb.AppendLine()
        End If

        If outputFiles IsNot Nothing AndAlso outputFiles.Count > 0 Then
            sb.AppendLine("### Output files produced")
            For Each outputFile In outputFiles
                If String.IsNullOrWhiteSpace(outputFile) Then Continue For
                sb.AppendLine("- " & Path.GetFileName(outputFile))
            Next
            sb.AppendLine()
        End If

        Dim canShowPartialResponse As Boolean =
            Not String.IsNullOrWhiteSpace(normalizedOutput) AndAlso
            Not normalizedOutput.Equals("Operation was canceled by the user.", StringComparison.OrdinalIgnoreCase) AndAlso
            Not normalizedOutput.Equals("Operation cancelled due to power transition.", StringComparison.OrdinalIgnoreCase) AndAlso
            Not normalizedOutput.StartsWith("Tool execution failed:", StringComparison.OrdinalIgnoreCase) AndAlso
            Not normalizedOutput.StartsWith("Aborting because the same failing tool call was repeated", StringComparison.OrdinalIgnoreCase)

        If canShowPartialResponse Then
            sb.AppendLine("### Partial response")
            sb.AppendLine(normalizedOutput)
            sb.AppendLine()
        End If

        If successfulEntries.Count = 0 AndAlso failedEntries.Count = 0 AndAlso
           (outputFiles Is Nothing OrElse outputFiles.Count = 0) AndAlso
           String.IsNullOrWhiteSpace(normalizedOutput) Then
            sb.AppendLine("No completed steps were recorded before the run stopped.")
        End If

        Return sb.ToString().Trim()
    End Function



    ' ═══════════════════════════════════════════════════════════════════════════
    '  AGENT CONTEXT SETUP / TEARDOWN
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Sets up the shared AutoPilot fields so that internal tools can operate
    ''' on the chat agent's files. Must be called before ExecuteToolingLoop.
    ''' Returns the combined tool list (internal agent tools + selected external tools).
    ''' </summary>
    Private Function ChatAgentSetupToolContext(Optional selectedTools As List(Of ModelConfig) = Nothing) As List(Of ModelConfig)
        If String.IsNullOrWhiteSpace(_chatAgentTempDir) OrElse Not Directory.Exists(_chatAgentTempDir) Then
            _chatAgentTempDir = Path.Combine(Path.GetTempPath(), CA_TempPrefix & Guid.NewGuid().ToString("N"))
            Directory.CreateDirectory(_chatAgentTempDir)
        End If

        Dim hydratedFiles As New List(Of AutoPilotAttachmentInfo)()
        For Each att In _chatAgentFiles.ToList()
            Dim hydrated = EnsureSessionAttachmentAvailable(att)
            If hydrated IsNot Nothing Then
                hydratedFiles.Add(hydrated)
            End If
        Next
        _chatAgentFiles = hydratedFiles

        _apCurrentTempDir = _chatAgentTempDir
        _apCurrentAttachments = _chatAgentFiles
        _apCurrentToolCallLog = New List(Of AutoPilotToolCallEntry)()
        SharedLibrary.Agents.PathPolicy.SetSessionStagingRoot(_chatAgentTempDir)

        _apCurrentMailInfo = New AutoPilotMailInfo() With {
            .EntryID = "",
            .Subject = "Chat Agent Session",
            .SenderName = Environment.UserName,
            .SenderEmail = "",
            .Body = "",
            .ReceivedTime = DateTime.Now,
            .HasAutoReplyHeader = False,
            .ThreadAIReplyCount = 0,
            .AttachmentCount = _chatAgentFiles.Count,
            .AttachmentNames = _chatAgentFiles.Select(Function(a) a.OriginalFileName).ToList(),
            .FolderPath = "",
            .MessageClass = "",
            .InternetHeaders = ""
        }

        _chatAgentActive = True
        SyncWorkspaceToPathPolicy(includeSessionTempFallback:=True)
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "AutoPilot (Local) invoked")

        Dim tools As New List(Of ModelConfig)()

        If selectedTools IsNot Nothing Then
            tools.AddRange(selectedTools)
        ElseIf _selectedToolsForChat IsNot Nothing Then
            tools.AddRange(_selectedToolsForChat)
        End If

        tools.AddRange(GetChatAgentWorkspaceTools())
        tools.AddRange(SharedLibrary.Agents.WorkspaceTools.BuildAll())

        Return tools.
            Where(Function(t) t IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(t.ToolName)).
            GroupBy(Function(t) t.ToolName, StringComparer.OrdinalIgnoreCase).
            Select(Function(g) g.First()).
            ToList()
    End Function

    ''' <summary>
    ''' Clears the shared AutoPilot fields after an agent job completes.
    ''' Deletes the temp directory so no temp files remain on disk.
    ''' Clears attachment caches for security.
    ''' </summary>
    Private Sub ChatAgentTeardownToolContext()
        _chatAgentActive = False
        SyncWorkspaceToPathPolicy()
        ClearAttachmentCaches()
        _apCurrentTempDir = Nothing
        _apCurrentAttachments = Nothing
        _apCurrentMailInfo = Nothing
        _apCurrentToolCallLog = Nothing
        _apKnowledgeSourceCopies.Clear()
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  OUTPUT FOLDER
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Collects output files from the agent session, filters to only those
    ''' cited by the LLM, copies them to Desktop\Inky\yymmdd_hh-mm\, and
    ''' opens the folder in Explorer.
    ''' </summary>
    Private Function ChatAgentCollectAndCopyOutputs(Optional extraFilePaths As IEnumerable(Of String) = Nothing) As List(Of String)
        Dim copiedFiles As New List(Of String)()
        Dim filesToCopy As New List(Of String)()

        If Not String.IsNullOrWhiteSpace(_chatAgentTempDir) AndAlso Directory.Exists(_chatAgentTempDir) Then
            Dim resultFiles = CollectResultAttachments(_chatAgentTempDir, _chatAgentFiles)
            If resultFiles IsNot Nothing AndAlso resultFiles.Count > 0 Then
                filesToCopy.AddRange(resultFiles)
            End If
        End If

        If extraFilePaths IsNot Nothing Then
            filesToCopy.AddRange(
                extraFilePaths.
                    Where(Function(p) Not String.IsNullOrWhiteSpace(p) AndAlso File.Exists(p)))
        End If

        filesToCopy = filesToCopy.
            Select(Function(p) Path.GetFullPath(p)).
            Distinct(StringComparer.OrdinalIgnoreCase).
            ToList()

        If filesToCopy.Count = 0 Then
            Return copiedFiles
        End If

        Dim desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
        Dim timestamp = DateTime.Now.ToString("yyMMdd_HH-mm")
        Dim outputDir = Path.Combine(desktopPath, "Inky", timestamp)

        Dim counter = 1
        While Directory.Exists(outputDir)
            outputDir = Path.Combine(desktopPath, "Inky", timestamp & $"_{counter}")
            counter += 1
        End While

        Directory.CreateDirectory(outputDir)

        For Each srcPath In filesToCopy
            Try
                Dim destName = Path.GetFileName(srcPath)
                Dim destPath = Path.Combine(outputDir, destName)

                Dim fileCounter = 1
                While File.Exists(destPath)
                    Dim baseName = Path.GetFileNameWithoutExtension(destName)
                    Dim ext = Path.GetExtension(destName)
                    destPath = Path.Combine(outputDir, baseName & $"_{fileCounter}{ext}")
                    fileCounter += 1
                End While

                File.Copy(srcPath, destPath, overwrite:=False)
                copiedFiles.Add(destPath)
            Catch
            End Try
        Next

        If copiedFiles.Count > 0 Then
            Try
                Process.Start("explorer.exe", outputDir)
            Catch
            End Try
        End If

        Return copiedFiles
    End Function

    Private Shared Function ExtractCitedLocalFilePaths(markdown As String) As List(Of String)
        Dim results As New List(Of String)()
        If String.IsNullOrWhiteSpace(markdown) Then Return results

        Dim rx As New System.Text.RegularExpressions.Regex(
            "file:///[^\s\)>\]""']+",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase)

        For Each m As System.Text.RegularExpressions.Match In rx.Matches(markdown)
            Dim raw As String = m.Value.TrimEnd("."c, ","c, ";"c, ":"c)
            Try
                Dim localPath As String = New Uri(raw).LocalPath
                If Not String.IsNullOrWhiteSpace(localPath) AndAlso File.Exists(localPath) Then
                    results.Add(Path.GetFullPath(localPath))
                End If
            Catch
            End Try
        Next

        Return results.Distinct(StringComparer.OrdinalIgnoreCase).ToList()
    End Function

    ''' <summary>
    ''' Called at the start of each non-sub-agent tooling run while the chat agent
    ''' is active. Marks every file currently present in _chatAgentTempDir as
    ''' "already surfaced", and clears the OutputFiles lists on user-uploaded
    ''' attachments so prior-turn outputs are not re-presented. Sub-agents inherit
    ''' parent state and must NOT call this. AutoPilot is single-threaded (Q4).
    ''' </summary>
    Friend Sub ResetChatAgentDeliverableTrackingForNewTurn()
        Try
            _chatAgentSurfacedFiles = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            _chatAgentForcedDeliverables = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            If Not String.IsNullOrWhiteSpace(_chatAgentTempDir) AndAlso Directory.Exists(_chatAgentTempDir) Then
                For Each filePath In Directory.GetFiles(_chatAgentTempDir, "*.*", IO.SearchOption.AllDirectories)
                    _chatAgentSurfacedFiles.Add(Path.GetFullPath(filePath))
                Next
            End If

            ' OutputFiles accumulated on user uploads from previous turns must not
            ' surface again. The uploads themselves stay; only their derived outputs
            ' are forgotten.
            If _chatAgentFiles IsNot Nothing Then
                For Each att In _chatAgentFiles
                    If att.OutputFiles IsNot Nothing Then att.OutputFiles.Clear()
                Next
            End If
        Catch ex As Exception
            ' Never throw from a per-turn reset. Worst case the user sees a duplicate file once.
            ToolingFileLogger.LogWarn("ResetChatAgentDeliverableTrackingForNewTurn failed.", ex:=ex)
        End Try
    End Sub


    ''' <summary>
    ''' When a shared in-place file-editing tool (e.g. word_write/word_markup/word_format/
    ''' word_comment_*) succeeds, its result JSON carries the edited file 'path'. That path
    ''' pre-existed the turn and would be filtered out by the already-surfaced gate in
    ''' CollectResultAttachments. Force its delivery so in-place edits reach Desktop\Inky.
    ''' No-ops for tools/results without a usable in-temp-dir path.
    ''' </summary>
    Private Sub MarkInPlaceEditAsForcedDeliverable(toolName As String, resultJson As String)
        Try
            If String.IsNullOrWhiteSpace(toolName) OrElse String.IsNullOrWhiteSpace(resultJson) Then Return

            ' Only shared Word document-file tools that mutate the file in place.
            Select Case toolName
                Case SharedLibrary.Agents.WordTools.ToolWrite,
                     SharedLibrary.Agents.WordTools.ToolMarkup,
                     SharedLibrary.Agents.WordTools.ToolFormat,
                     SharedLibrary.Agents.WordTools.ToolCommentAdd,
                     SharedLibrary.Agents.WordTools.ToolCommentRemove
                    ' proceed
                Case Else
                    Return
            End Select

            Dim obj As JObject
            Try
                obj = JObject.Parse(resultJson)
            Catch
                Return
            End Try

            Dim editedPath As String = If(obj.Value(Of String)("path"), "").Trim()
            If String.IsNullOrWhiteSpace(editedPath) OrElse Not File.Exists(editedPath) Then Return

            Dim full As String = Path.GetFullPath(editedPath)

            ' Security: only files inside the active session staging/temp dir are eligible.
            Dim stagingDir As String =
                If(_apActive AndAlso Not String.IsNullOrWhiteSpace(_apCurrentTempDir),
                   _apCurrentTempDir, _chatAgentTempDir)
            If String.IsNullOrWhiteSpace(stagingDir) Then Return
            If Not full.StartsWith(Path.GetFullPath(stagingDir), StringComparison.OrdinalIgnoreCase) Then Return

            _chatAgentForcedDeliverables.Add(full)
        Catch ex As Exception
            ToolingFileLogger.LogWarn("MarkInPlaceEditAsForcedDeliverable failed.", ex:=ex)
        End Try
    End Sub

    ''' <summary>
    ''' Deletes the chat agent temp directory (recursively, including subdirectories)
    ''' and resets the file tracking list. Safe to call multiple times.
    ''' </summary>
    Private Sub CleanupChatAgentTempDir()
        _chatAgentFiles.Clear()
        Try
            If Not String.IsNullOrWhiteSpace(_chatAgentTempDir) AndAlso Directory.Exists(_chatAgentTempDir) Then
                Directory.Delete(_chatAgentTempDir, recursive:=True)
            End If
        Catch
        End Try
        _chatAgentTempDir = Nothing
        SharedLibrary.Agents.PathPolicy.SetSessionStagingRoot(Nothing)
    End Sub

    ''' <summary>
    ''' Registers any files currently present in the active session staging directory
    ''' as session attachments so tool consumers can resolve them by name. Runs on every
    ''' FindAttachment lookup; registration is deduped by temp-file path. Registers into
    ''' _apCurrentAttachments only (in Local Agent this is the same list as _chatAgentFiles;
    ''' in AutoPilot it is the mail-attachment list), so no cross-session state is polluted.
    ''' </summary>
    Private Sub RefreshSessionStagingAttachments()
        Try
            If _apCurrentAttachments Is Nothing Then Return

            Dim stagingDir As String =
                If(_apActive AndAlso Not String.IsNullOrWhiteSpace(_apCurrentTempDir),
                   _apCurrentTempDir, _chatAgentTempDir)
            If String.IsNullOrWhiteSpace(stagingDir) OrElse Not Directory.Exists(stagingDir) Then Return

            For Each filePath In Directory.GetFiles(stagingDir, "*.*", IO.SearchOption.AllDirectories)
                Dim full = Path.GetFullPath(filePath)

                Dim already = _apCurrentAttachments.Any(
                    Function(a) a IsNot Nothing AndAlso
                                Not String.IsNullOrWhiteSpace(a.TempFilePath) AndAlso
                                a.TempFilePath.Equals(full, StringComparison.OrdinalIgnoreCase))
                If already Then Continue For

                Dim fi As New FileInfo(full)
                _apCurrentAttachments.Add(New AutoPilotAttachmentInfo() With {
                    .OriginalFileName = Path.GetFileName(full),
                    .Extension = Path.GetExtension(full).ToLowerInvariant(),
                    .TempFilePath = full,
                    .SizeBytes = fi.Length,
                    .IsOverSizeLimit = False,
                    .StatusMessage = "Session staging file",
                    .IsToolOutput = True,
                    .OutputFiles = New List(Of String)()
                })
            Next
        Catch ex As Exception
            ToolingFileLogger.LogWarn("RefreshSessionStagingAttachments failed.", ex:=ex)
        End Try
    End Sub

    ''' <summary>
    ''' Registers files present in the session staging directory into _chatAgentFiles so
    ''' they appear as browser chips. Deduped by temp-file path. Unlike
    ''' RefreshSessionStagingAttachments (which targets the active tool-run attachment list),
    ''' this keeps the browser chip list current even outside an active tool run.
    ''' </summary>
    Private Sub RegisterStagingFilesAsSessionChips()
        Try
            Dim stagingDir As String =
                If(_apActive AndAlso Not String.IsNullOrWhiteSpace(_apCurrentTempDir),
                   _apCurrentTempDir, _chatAgentTempDir)
            If String.IsNullOrWhiteSpace(stagingDir) OrElse Not Directory.Exists(stagingDir) Then Return

            For Each filePath In Directory.GetFiles(stagingDir, "*.*", IO.SearchOption.AllDirectories)
                Dim full = Path.GetFullPath(filePath)
                If _chatAgentFiles.Any(
                    Function(a) a IsNot Nothing AndAlso
                                Not String.IsNullOrWhiteSpace(a.TempFilePath) AndAlso
                                a.TempFilePath.Equals(full, StringComparison.OrdinalIgnoreCase)) Then Continue For

                RegisterSessionFile(full, "Session staging file", isToolOutput:=True)
            Next
        Catch ex As Exception
            ToolingFileLogger.LogWarn("RegisterStagingFilesAsSessionChips failed.", ex:=ex)
        End Try
    End Sub

    ''' <summary>
    ''' Best-effort reclamation of orphaned chat-agent staging directories from prior
    ''' sessions whose cleanup previously failed. Deletes only directories matching this
    ''' host's temp prefix, excluding the current active dirs, and only when older than
    ''' 24 hours (age gate protects concurrently running sessions).
    ''' </summary>
    Private Sub CleanupStaleChatAgentTempDirs()
        Try
            Dim tempRoot = Path.GetTempPath()
            Dim cutoff = DateTime.UtcNow.AddHours(-24)

            For Each dirPath In Directory.GetDirectories(tempRoot, CA_TempPrefix & "*")
                Try
                    Dim full = Path.GetFullPath(dirPath)

                    If Not String.IsNullOrWhiteSpace(_chatAgentTempDir) AndAlso
                       full.Equals(Path.GetFullPath(_chatAgentTempDir), StringComparison.OrdinalIgnoreCase) Then Continue For
                    If Not String.IsNullOrWhiteSpace(_apCurrentTempDir) AndAlso
                       full.Equals(Path.GetFullPath(_apCurrentTempDir), StringComparison.OrdinalIgnoreCase) Then Continue For

                    If Directory.GetLastWriteTimeUtc(full) > cutoff Then Continue For

                    Directory.Delete(full, recursive:=True)
                Catch
                End Try
            Next
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Returns True if the chat agent mode is blocked because AutoPilot is running.
    ''' </summary>
    Private Function IsChatAgentBlocked() As Boolean
        Return _apActive
    End Function


    ' ═══════════════════════════════════════════════════════════════════════════
    '  CHAT AGENT WORKSPACE
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Function GetChatAgentWorkspaceStatePath() As String
        Dim iniPath = GetDefaultINIPath(_context.RDV)
        Dim iniDir = Path.GetDirectoryName(iniPath)
        Return Path.Combine(iniDir, "local_chat_agent_workspace.json")
    End Function

    Private Sub LoadChatAgentWorkspaceIfNeeded()
        If _chatAgentWorkspaceLoaded Then Return
        _chatAgentWorkspaceLoaded = True
        _chatAgentWorkspace = New ChatAgentWorkspaceState()

        Try
            Dim statePath = GetChatAgentWorkspaceStatePath()
            If File.Exists(statePath) Then
                Dim loaded = JsonConvert.DeserializeObject(Of ChatAgentWorkspaceState)(File.ReadAllText(statePath, Encoding.UTF8))
                If loaded IsNot Nothing AndAlso loaded.PersistUntilRevoked Then
                    _chatAgentWorkspace = loaded
                End If
            End If
        Catch
            _chatAgentWorkspace = New ChatAgentWorkspaceState()
        End Try

        SyncWorkspaceToPathPolicy()
    End Sub

    Private Sub SaveChatAgentWorkspaceState()
        LoadChatAgentWorkspaceIfNeeded()

        Try
            Dim statePath = GetChatAgentWorkspaceStatePath()
            If _chatAgentWorkspace IsNot Nothing AndAlso _chatAgentWorkspace.PersistUntilRevoked AndAlso Not String.IsNullOrWhiteSpace(_chatAgentWorkspace.RootPath) Then
                File.WriteAllText(statePath, JsonConvert.SerializeObject(_chatAgentWorkspace, Formatting.Indented), Encoding.UTF8)
            ElseIf File.Exists(statePath) Then
                File.Delete(statePath)
            End If
        Catch
        End Try
    End Sub

    Private Function IsValidWorkspacePresetSlot(slot As Integer) As Boolean
        Return slot >= 1 AndAlso slot <= CA_WorkspacePresetCount
    End Function

    Private Function GetWorkspacePresetPathSettingName(slot As Integer) As String
        Return CA_WorkspacePresetSettingPrefix & slot.ToString(CultureInfo.InvariantCulture) & "Path"
    End Function

    Private Function GetWorkspacePresetLabelSettingName(slot As Integer) As String
        Return CA_WorkspacePresetSettingPrefix & slot.ToString(CultureInfo.InvariantCulture) & "Label"
    End Function

    Private Function TrySetAppSettingValue(settingName As String, value As Object) As Boolean
        Try
            Dim p = GetType(My.MySettings).GetProperty(
                settingName,
                System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.Instance)

            If p Is Nothing OrElse Not p.CanWrite Then Return False

            p.SetValue(My.Settings, value, Nothing)
            Return True
        Catch
            Return False
        End Try
    End Function

    Private Function GetWorkspacePresetEntries() As List(Of ChatAgentWorkspacePresetInfo)
        Dim result As New List(Of ChatAgentWorkspacePresetInfo)()

        For slot As Integer = 1 To CA_WorkspacePresetCount
            Dim rootPath As String = TryGetAppSetting(Of String)(GetWorkspacePresetPathSettingName(slot), "")
            Dim label As String = TryGetAppSetting(Of String)(GetWorkspacePresetLabelSettingName(slot), "")

            If Not String.IsNullOrWhiteSpace(rootPath) Then
                Try
                    rootPath = Path.GetFullPath(rootPath.Trim())
                Catch
                    rootPath = rootPath.Trim()
                End Try
            Else
                rootPath = ""
            End If

            result.Add(New ChatAgentWorkspacePresetInfo() With {
                .Slot = slot,
                .RootPath = rootPath,
                .Label = If(label, "")
            })
        Next

        Return result
    End Function

    Private Function CleanWorkspaceLabelToken(value As String) As String
        Dim cleaned As String = If(value, "").Trim()

        cleaned = cleaned.Replace("_", " ").Replace("-", " ")
        cleaned = System.Text.RegularExpressions.Regex.Replace(cleaned, "[^\p{L}\p{Nd} ]+", " ")
        cleaned = System.Text.RegularExpressions.Regex.Replace(cleaned, "\s+", " ").Trim()

        If cleaned.Length > 24 Then
            cleaned = cleaned.Substring(0, 24).Trim()
        End If

        Return cleaned
    End Function

    Private Function NormalizeWorkspacePresetLabel(label As String, fallbackLabel As String) As String
        Dim cleaned As String = If(label, "").Trim()

        cleaned = cleaned.Replace("`", " ").Replace("""", " ").Replace("'", " ")
        cleaned = CleanWorkspaceLabelToken(cleaned)

        If String.IsNullOrWhiteSpace(cleaned) Then
            cleaned = CleanWorkspaceLabelToken(fallbackLabel)
        End If

        If String.IsNullOrWhiteSpace(cleaned) Then
            cleaned = "Workspace"
        End If

        Dim words = cleaned.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
        If words.Length > 2 Then
            cleaned = words(0) & " " & words(1)
        End If

        If cleaned.Length > 24 Then
            cleaned = cleaned.Substring(0, 24).Trim()
        End If

        Return cleaned
    End Function

    Private Function SplitWorkspacePathSegments(rootPath As String) As List(Of String)
        Dim result As New List(Of String)()

        If String.IsNullOrWhiteSpace(rootPath) Then Return result

        Dim normalized As String = rootPath

        Try
            normalized = Path.GetFullPath(rootPath)
        Catch
        End Try

        normalized = normalized.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)

        For Each rawSegment In normalized.Split(New Char() {Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar}, StringSplitOptions.RemoveEmptyEntries)
            Dim segment = CleanWorkspaceLabelToken(rawSegment)
            If String.IsNullOrWhiteSpace(segment) Then Continue For
            If segment.EndsWith(":", StringComparison.Ordinal) Then Continue For
            result.Add(segment)
        Next

        Return result
    End Function

    Private Function BuildWorkspacePresetFallbackLabels(entries As List(Of ChatAgentWorkspacePresetInfo)) As Dictionary(Of Integer, String)
        Dim labels As New Dictionary(Of Integer, String)()
        Dim tokenMap As New Dictionary(Of Integer, List(Of String))()

        If entries Is Nothing Then Return labels

        For Each entry In entries
            If entry Is Nothing OrElse String.IsNullOrWhiteSpace(entry.RootPath) Then Continue For

            Dim tokens = SplitWorkspacePathSegments(entry.RootPath)
            tokenMap(entry.Slot) = tokens

            Dim leaf As String = If(tokens.Count > 0, tokens(tokens.Count - 1), "Workspace")
            labels(entry.Slot) = NormalizeWorkspacePresetLabel(leaf, "Workspace " & entry.Slot.ToString(CultureInfo.InvariantCulture))
        Next

        Dim duplicateGroups =
            labels.GroupBy(Function(kvp) kvp.Value, StringComparer.OrdinalIgnoreCase).
                   Where(Function(g) g.Count() > 1).
                   ToList()

        For Each group In duplicateGroups
            For Each item In group
                Dim tokens As List(Of String) = Nothing
                If Not tokenMap.TryGetValue(item.Key, tokens) Then Continue For

                Dim leaf As String = If(tokens.Count > 0, tokens(tokens.Count - 1), "Workspace")
                Dim parent As String = If(tokens.Count > 1, tokens(tokens.Count - 2), "")

                Dim candidate As String
                If Not String.IsNullOrWhiteSpace(parent) Then
                    candidate = parent & " " & leaf
                Else
                    candidate = leaf & " " & item.Key.ToString(CultureInfo.InvariantCulture)
                End If

                labels(item.Key) = NormalizeWorkspacePresetLabel(candidate, leaf & " " & item.Key.ToString(CultureInfo.InvariantCulture))
            Next
        Next

        duplicateGroups =
            labels.GroupBy(Function(kvp) kvp.Value, StringComparer.OrdinalIgnoreCase).
                   Where(Function(g) g.Count() > 1).
                   ToList()

        For Each group In duplicateGroups
            For Each item In group
                Dim tokens As List(Of String) = Nothing
                If Not tokenMap.TryGetValue(item.Key, tokens) Then Continue For

                Dim leaf As String = If(tokens.Count > 0, tokens(tokens.Count - 1), "Workspace")
                labels(item.Key) = NormalizeWorkspacePresetLabel(
                    leaf & " " & item.Key.ToString(CultureInfo.InvariantCulture),
                    "Workspace " & item.Key.ToString(CultureInfo.InvariantCulture))
            Next
        Next

        Return labels
    End Function

    Private Async Function RefreshWorkspacePresetLabelsAsync() As Task(Of Boolean)
        Dim allEntries = GetWorkspacePresetEntries()
        Dim populated =
            allEntries.
                Where(Function(entry) entry IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(entry.RootPath)).
                ToList()

        Dim fallbackLabels = BuildWorkspacePresetFallbackLabels(populated)
        Dim resolvedLabels As New Dictionary(Of Integer, String)()

        For Each kvp In fallbackLabels
            resolvedLabels(kvp.Key) = kvp.Value
        Next

        If populated.Count > 0 Then
            Try
                Dim sysPrompt As String =
                    "You create ultra-short UI labels for saved workspace folder paths. Return strict JSON only."

                Dim userPrompt As New StringBuilder()
                userPrompt.AppendLine("Create unique memory-button labels for these workspace paths.")
                userPrompt.AppendLine("Return JSON only in this exact shape:")
                userPrompt.AppendLine("[{""slot"":1,""label"":""Client Docs""}]")
                userPrompt.AppendLine("Rules:")
                userPrompt.AppendLine("- One object per slot shown below.")
                userPrompt.AppendLine("- label must be 1 or 2 words only.")
                userPrompt.AppendLine("- Keep labels distinct across the full list.")
                userPrompt.AppendLine("- Prefer the most distinguishing folder name or two folder names from the path.")
                userPrompt.AppendLine("- No punctuation except spaces.")
                userPrompt.AppendLine("- No explanations, markdown, or prose.")
                userPrompt.AppendLine("Paths:")

                For Each entry In populated
                    userPrompt.AppendLine("- slot " &
                                          entry.Slot.ToString(CultureInfo.InvariantCulture) &
                                          ": " & entry.RootPath)
                Next

                Dim llmOut As String =
                    Await RunLlmAsync(
                        sysPrompt,
                        userPrompt.ToString(),
                        False,
                        False,
                        "",
                        CancellationToken.None).ConfigureAwait(False)

                llmOut = If(llmOut, "").Trim()

                If llmOut.StartsWith("```", StringComparison.Ordinal) Then
                    Dim firstLf = llmOut.IndexOf(vbLf, StringComparison.Ordinal)
                    If firstLf >= 0 Then
                        llmOut = llmOut.Substring(firstLf + 1).Trim()
                    End If

                    Dim fenceEnd = llmOut.LastIndexOf("```", StringComparison.Ordinal)
                    If fenceEnd >= 0 Then
                        llmOut = llmOut.Substring(0, fenceEnd).Trim()
                    End If
                End If

                Dim arr = JArray.Parse(llmOut)

                For Each item As JToken In arr
                    Dim slot As Integer = 0
                    If Not Integer.TryParse(item("slot")?.ToString(), slot) Then Continue For
                    If Not IsValidWorkspacePresetSlot(slot) Then Continue For

                    Dim fallbackLabel As String =
                        If(fallbackLabels.ContainsKey(slot),
                           fallbackLabels(slot),
                           "Workspace " & slot.ToString(CultureInfo.InvariantCulture))

                    Dim label As String = NormalizeWorkspacePresetLabel(item("label")?.ToString(), fallbackLabel)
                    If Not String.IsNullOrWhiteSpace(label) Then
                        resolvedLabels(slot) = label
                    End If
                Next
            Catch
            End Try
        End If

        For Each entry In allEntries
            Dim label As String = ""

            If Not String.IsNullOrWhiteSpace(entry.RootPath) Then
                Dim fallbackLabel As String =
                    If(fallbackLabels.ContainsKey(entry.Slot),
                       fallbackLabels(entry.Slot),
                       "Workspace " & entry.Slot.ToString(CultureInfo.InvariantCulture))

                label =
                    If(resolvedLabels.ContainsKey(entry.Slot),
                       NormalizeWorkspacePresetLabel(resolvedLabels(entry.Slot), fallbackLabel),
                       fallbackLabel)
            End If

            If Not TrySetAppSettingValue(GetWorkspacePresetLabelSettingName(entry.Slot), label) Then
                Return False
            End If
        Next

        Try
            My.Settings.Save()
            Return True
        Catch
            Return False
        End Try
    End Function

    Friend Async Function ChatAgentWorkspaceSaveCurrentPresetAsync(slot As Integer) As Task(Of String)
        If Not IsValidWorkspacePresetSlot(slot) Then
            Return "Invalid workspace memory slot."
        End If

        LoadChatAgentWorkspaceIfNeeded()

        Dim rootPath As String = If(_chatAgentWorkspace?.RootPath, "")
        If String.IsNullOrWhiteSpace(rootPath) OrElse Not Directory.Exists(rootPath) Then
            Return "No workspace is currently connected."
        End If

        rootPath = Path.GetFullPath(rootPath)

        If Not TrySetAppSettingValue(GetWorkspacePresetPathSettingName(slot), rootPath) Then
            Return "Workspace memory settings are not available."
        End If

        If Not TrySetAppSettingValue(GetWorkspacePresetLabelSettingName(slot), "") Then
            Return "Workspace memory settings are not available."
        End If

        Dim saved As Boolean = Await RefreshWorkspacePresetLabelsAsync().ConfigureAwait(False)
        If Not saved Then
            Return "The workspace memory could not be saved."
        End If

        Return Nothing
    End Function

    Friend Async Function ChatAgentWorkspaceClearPresetAsync(slot As Integer) As Task(Of String)
        If Not IsValidWorkspacePresetSlot(slot) Then
            Return "Invalid workspace memory slot."
        End If

        If Not TrySetAppSettingValue(GetWorkspacePresetPathSettingName(slot), "") Then
            Return "Workspace memory settings are not available."
        End If

        If Not TrySetAppSettingValue(GetWorkspacePresetLabelSettingName(slot), "") Then
            Return "Workspace memory settings are not available."
        End If

        Dim saved As Boolean = Await RefreshWorkspacePresetLabelsAsync().ConfigureAwait(False)
        If Not saved Then
            Return "The workspace memory could not be cleared."
        End If

        Return Nothing
    End Function

    Friend Function ChatAgentWorkspaceApplyPreset(slot As Integer) As String
        If Not IsValidWorkspacePresetSlot(slot) Then
            Return "Invalid workspace memory slot."
        End If

        Dim presetPath As String = TryGetAppSetting(Of String)(GetWorkspacePresetPathSettingName(slot), "")
        If String.IsNullOrWhiteSpace(presetPath) Then
            Return "That workspace memory is empty."
        End If

        Try
            presetPath = Path.GetFullPath(presetPath.Trim())
        Catch
            Return "The saved workspace path is invalid."
        End Try

        If Not Directory.Exists(presetPath) Then
            Return "The saved workspace path no longer exists."
        End If

        LoadChatAgentWorkspaceIfNeeded()

        Dim persistUntilRevoked As Boolean = If(_chatAgentWorkspace?.PersistUntilRevoked, False)

        If Not ChatAgentWorkspaceSetRoot(presetPath, persistUntilRevoked) Then
            Return "The saved workspace path could not be activated."
        End If

        Return Nothing
    End Function

    Friend Function GetAgentWorkspacePresetsForBrowser() As List(Of Object)
        Dim result As New List(Of Object)()
        Dim entries = GetWorkspacePresetEntries()
        Dim populated =
            entries.
                Where(Function(entry) entry IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(entry.RootPath)).
                ToList()

        Dim fallbackLabels = BuildWorkspacePresetFallbackLabels(populated)

        Dim currentRoot As String = ""
        If IsChatAgentWorkspaceConnected() Then
            Try
                currentRoot =
                    Path.GetFullPath(_chatAgentWorkspace.RootPath).
                        TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            Catch
                currentRoot = ""
            End Try
        End If

        For Each entry In entries
            Dim assigned As Boolean = Not String.IsNullOrWhiteSpace(entry.RootPath)
            Dim exists As Boolean = assigned AndAlso Directory.Exists(entry.RootPath)

            Dim fallbackLabel As String =
                If(fallbackLabels.ContainsKey(entry.Slot),
                   fallbackLabels(entry.Slot),
                   "Workspace " & entry.Slot.ToString(CultureInfo.InvariantCulture))

            Dim displayLabel As String =
                If(assigned,
                   NormalizeWorkspacePresetLabel(entry.Label, fallbackLabel),
                   "")

            Dim isActive As Boolean = False
            If assigned AndAlso exists AndAlso Not String.IsNullOrWhiteSpace(currentRoot) Then
                Try
                    Dim presetRoot =
                        Path.GetFullPath(entry.RootPath).
                            TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)

                    isActive = presetRoot.Equals(currentRoot, StringComparison.OrdinalIgnoreCase)
                Catch
                    isActive = False
                End Try
            End If

            result.Add(New With {
                .slot = entry.Slot,
                .path = If(assigned, entry.RootPath, ""),
                .label = displayLabel,
                .assigned = assigned,
                .exists = exists,
                .active = isActive
            })
        Next

        Return result
    End Function

    Friend Function GetAgentWorkspaceForBrowser() As Object
        LoadChatAgentWorkspaceIfNeeded()

        Dim root = If(_chatAgentWorkspace?.RootPath, "")
        Dim connected = Not String.IsNullOrWhiteSpace(root) AndAlso Directory.Exists(root)

        Return New With {
            .connected = connected,
            .rootPath = If(connected, root, ""),
            .name = If(connected, Path.GetFileName(root.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)), ""),
            .persistUntilRevoked = If(_chatAgentWorkspace?.PersistUntilRevoked, False),
            .allowRead = If(_chatAgentWorkspace?.AllowRead, True),
            .allowWrite = If(_chatAgentWorkspace?.AllowWrite, True),
            .allowMoveCopyRename = If(_chatAgentWorkspace?.AllowMoveCopyRename, True),
            .allowDelete = If(_chatAgentWorkspace?.AllowDelete, SharedMethods.WorkspaceDeleteByDefaultOn),
            .saveDroppedFilesToWorkspace = If(_chatAgentWorkspace?.SaveDroppedFilesToWorkspace, False),
            .includeHiddenSystem = If(_chatAgentWorkspace?.IncludeHiddenSystem, False),
            .presets = GetAgentWorkspacePresetsForBrowser()
        }
    End Function

    Friend Function ChatAgentWorkspaceSetRoot(rootPath As String, persistUntilRevoked As Boolean) As Boolean
        If String.IsNullOrWhiteSpace(rootPath) OrElse Not Directory.Exists(rootPath) Then Return False

        LoadChatAgentWorkspaceIfNeeded()

        _chatAgentWorkspace.RootPath = Path.GetFullPath(rootPath)
        _chatAgentWorkspace.PersistUntilRevoked = persistUntilRevoked
        If Not _chatAgentWorkspace.AllowRead Then _chatAgentWorkspace.AllowRead = True

        SaveChatAgentWorkspaceState()
        SyncWorkspaceToPathPolicy()
        Return True
    End Function

    Friend Sub ChatAgentWorkspaceSetPermissions(
        persistUntilRevoked As Boolean,
        allowRead As Boolean,
        allowWrite As Boolean,
        allowMoveCopyRename As Boolean,
        allowDelete As Boolean,
        saveDroppedFilesToWorkspace As Boolean,
        includeHiddenSystem As Boolean)

        LoadChatAgentWorkspaceIfNeeded()

        _chatAgentWorkspace.PersistUntilRevoked = persistUntilRevoked
        _chatAgentWorkspace.AllowRead = allowRead
        _chatAgentWorkspace.AllowWrite = allowWrite
        _chatAgentWorkspace.AllowMoveCopyRename = allowMoveCopyRename
        _chatAgentWorkspace.AllowDelete = allowDelete
        _chatAgentWorkspace.SaveDroppedFilesToWorkspace = saveDroppedFilesToWorkspace
        _chatAgentWorkspace.IncludeHiddenSystem = includeHiddenSystem

        SaveChatAgentWorkspaceState()
        SyncWorkspaceToPathPolicy()
    End Sub

    Friend Sub ChatAgentWorkspaceRevoke()
        _chatAgentWorkspace = New ChatAgentWorkspaceState()
        _chatAgentWorkspaceLoaded = True

        Try
            Dim statePath = GetChatAgentWorkspaceStatePath()
            If File.Exists(statePath) Then File.Delete(statePath)
        Catch
        End Try
        SyncWorkspaceToPathPolicy()
    End Sub

    Private Function IsChatAgentWorkspaceConnected() As Boolean
        LoadChatAgentWorkspaceIfNeeded()
        Return _chatAgentWorkspace IsNot Nothing AndAlso
               Not String.IsNullOrWhiteSpace(_chatAgentWorkspace.RootPath) AndAlso
               Directory.Exists(_chatAgentWorkspace.RootPath)
    End Function

    Private Shared Function IsPathWithinOrEqual(candidatePath As String, requiredParent As String) As Boolean
        Dim fullCandidate = Path.GetFullPath(candidatePath).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
        Dim fullParent = Path.GetFullPath(requiredParent).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)

        If fullCandidate.Equals(fullParent, StringComparison.OrdinalIgnoreCase) Then Return True

        fullParent &= Path.DirectorySeparatorChar
        Return fullCandidate.StartsWith(fullParent, StringComparison.OrdinalIgnoreCase)
    End Function

    Private Function ResolveWorkspacePath(relativePath As String, Optional allowRoot As Boolean = False) As String
        If Not IsChatAgentWorkspaceConnected() Then Return Nothing

        Dim rel = If(relativePath, "").Trim()
        If rel = "" Then
            If allowRoot Then Return Path.GetFullPath(_chatAgentWorkspace.RootPath)
            Return Nothing
        End If

        If Path.IsPathRooted(rel) Then Return Nothing
        If rel.Contains(".." & Path.DirectorySeparatorChar) OrElse
           rel.Contains(".." & Path.AltDirectorySeparatorChar) OrElse
           rel.Equals("..", StringComparison.Ordinal) Then
            Return Nothing
        End If

        Dim root = Path.GetFullPath(_chatAgentWorkspace.RootPath)
        Dim candidate = Path.GetFullPath(Path.Combine(root, rel))

        If Not IsPathWithinOrEqual(candidate, root) Then Return Nothing
        If ContainsReparsePointBetween(root, candidate) Then Return Nothing

        Return candidate
    End Function

    Private Function ContainsReparsePointBetween(rootPath As String, candidatePath As String) As Boolean
        Try
            Dim root = Path.GetFullPath(rootPath).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            Dim current = Path.GetFullPath(candidatePath).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)

            ' If the candidate path does not exist yet (e.g. mkdir "New"), walk up to the
            ' nearest existing ancestor inside the workspace and validate from there.
            While Not String.IsNullOrWhiteSpace(current) AndAlso
                  Not Directory.Exists(current) AndAlso
                  Not File.Exists(current)

                Dim parentPath = Path.GetDirectoryName(current)
                If String.IsNullOrWhiteSpace(parentPath) OrElse
                   parentPath.Equals(current, StringComparison.OrdinalIgnoreCase) Then
                    Exit While
                End If

                current = parentPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            End While

            If String.IsNullOrWhiteSpace(current) Then Return True
            If Not IsPathWithinOrEqual(current, root) Then Return True

            ' If the nearest existing thing is a file, inspect its parent directory.
            If File.Exists(current) Then
                current = Path.GetDirectoryName(current)
            End If

            While Not String.IsNullOrWhiteSpace(current) AndAlso IsPathWithinOrEqual(current, root)
                Dim isRoot = current.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar).
                    Equals(root, StringComparison.OrdinalIgnoreCase)

                ' Do not reject the workspace root itself, because synced folders may be backed
                ' by provider-specific attributes. We only block reparse points below the root.
                If Not isRoot Then
                    Dim attr = File.GetAttributes(current)
                    If (attr And FileAttributes.ReparsePoint) = FileAttributes.ReparsePoint Then
                        Return True
                    End If
                End If

                If isRoot Then Exit While

                Dim parent = Directory.GetParent(current)
                If parent Is Nothing Then Exit While
                current = parent.FullName.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            End While
        Catch
            Return True
        End Try

        Return False
    End Function

    Private Function ShouldSkipHiddenSystem(path As String) As Boolean
        If _chatAgentWorkspace Is Nothing OrElse _chatAgentWorkspace.IncludeHiddenSystem Then Return False

        Try
            Dim attr = File.GetAttributes(path)
            Return (attr And FileAttributes.Hidden) = FileAttributes.Hidden OrElse
                   (attr And FileAttributes.System) = FileAttributes.System OrElse
                   (attr And FileAttributes.ReparsePoint) = FileAttributes.ReparsePoint
        Catch
            Return True
        End Try
    End Function

    Private Function ToWorkspaceRelativePath(fullPath As String) As String
        Dim root = Path.GetFullPath(_chatAgentWorkspace.RootPath).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
        Dim full = Path.GetFullPath(fullPath)

        If full.Equals(root, StringComparison.OrdinalIgnoreCase) Then Return ""

        Dim rel = full.Substring(root.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
        Return rel.Replace(Path.DirectorySeparatorChar, "/"c)
    End Function

    Private Function EnsureChatAgentTempDir() As String
        If _apActive AndAlso
           Not String.IsNullOrWhiteSpace(_apCurrentTempDir) AndAlso
           Directory.Exists(_apCurrentTempDir) Then

            Return _apCurrentTempDir
        End If

        If String.IsNullOrWhiteSpace(_chatAgentTempDir) OrElse Not Directory.Exists(_chatAgentTempDir) Then
            CleanupStaleChatAgentTempDirs()
            _chatAgentTempDir = Path.Combine(Path.GetTempPath(), CA_TempPrefix & Guid.NewGuid().ToString("N"))
            Directory.CreateDirectory(_chatAgentTempDir)
        End If

        If Not _apActive Then
            SharedLibrary.Agents.PathPolicy.SetSessionStagingRoot(_chatAgentTempDir)
        End If

        Return _chatAgentTempDir
    End Function

    Private Function RegisterSessionFile(filePath As String,
                                         statusMessage As String,
                                         Optional sourcePath As String = Nothing,
                                         Optional isToolOutput As Boolean = False) As AutoPilotAttachmentInfo
        If String.IsNullOrWhiteSpace(filePath) OrElse Not File.Exists(filePath) Then Return Nothing

        Dim fileName = Path.GetFileName(filePath)
        Dim existing = _chatAgentFiles.FirstOrDefault(
            Function(a)
                Return a IsNot Nothing AndAlso
                       Not String.IsNullOrWhiteSpace(a.TempFilePath) AndAlso
                       a.TempFilePath.Equals(filePath, StringComparison.OrdinalIgnoreCase)
            End Function)

        If existing IsNot Nothing Then
            If String.IsNullOrWhiteSpace(existing.SourcePath) AndAlso Not String.IsNullOrWhiteSpace(sourcePath) Then
                existing.SourcePath = sourcePath
            End If
            existing.StatusMessage = statusMessage
            Return existing
        End If

        Dim fi As New FileInfo(filePath)
        Dim att As New AutoPilotAttachmentInfo() With {
            .OriginalFileName = fileName,
            .TempFilePath = filePath,
            .SourcePath = sourcePath,
            .Extension = fi.Extension.ToLowerInvariant(),
            .SizeBytes = fi.Length,
            .IsOverSizeLimit = False,
            .StatusMessage = statusMessage,
            .CreatedTime = fi.CreationTimeUtc,
            .LastModifiedTime = fi.LastWriteTimeUtc,
            .OutputFiles = New List(Of String)(),
            .IsToolOutput = isToolOutput
        }

        _chatAgentFiles.Add(att)

        If _apCurrentAttachments IsNot Nothing AndAlso
           Not _apCurrentAttachments.Any(
               Function(a)
                   Return a IsNot Nothing AndAlso
                          Not String.IsNullOrWhiteSpace(a.TempFilePath) AndAlso
                          a.TempFilePath.Equals(filePath, StringComparison.OrdinalIgnoreCase)
               End Function) Then
            _apCurrentAttachments.Add(att)
        End If

        Return att
    End Function

    Private Function StageWorkspaceFile(relativePath As String) As AutoPilotAttachmentInfo
        If Not _chatAgentWorkspace.AllowRead Then Return Nothing

        Dim sourcePath = ResolveWorkspacePath(relativePath)
        If sourcePath Is Nothing OrElse Not File.Exists(sourcePath) OrElse ShouldSkipHiddenSystem(sourcePath) Then Return Nothing

        Dim existing = _chatAgentFiles.FirstOrDefault(
            Function(att)
                Return att IsNot Nothing AndAlso
                       Not String.IsNullOrWhiteSpace(att.SourcePath) AndAlso
                       att.SourcePath.Equals(sourcePath, StringComparison.OrdinalIgnoreCase)
            End Function)

        If existing IsNot Nothing Then
            Return EnsureSessionAttachmentAvailable(existing)
        End If

        Dim tempDir = EnsureChatAgentTempDir()
        Dim destName = Path.GetFileName(sourcePath)
        Dim destPath = Path.Combine(tempDir, destName)

        Dim counter = 1
        While File.Exists(destPath)
            destPath = Path.Combine(
                tempDir,
                Path.GetFileNameWithoutExtension(destName) & $"_{counter}" & Path.GetExtension(destName))
            counter += 1
        End While

        File.Copy(sourcePath, destPath, overwrite:=False)
        Return RegisterSessionFile(destPath, "Loaded from agent workspace", sourcePath:=sourcePath)
    End Function

    Private Function CopySessionFileToWorkspace(sessionFileName As String, targetRelativePath As String, overwrite As Boolean) As String
        If Not _chatAgentWorkspace.AllowWrite Then Return "Error: Workspace write permission is disabled."

        Dim att = FindAttachment(sessionFileName)
        att = EnsureSessionAttachmentAvailable(att)

        If att Is Nothing OrElse String.IsNullOrWhiteSpace(att.TempFilePath) OrElse Not File.Exists(att.TempFilePath) Then
            Dim availableNames = GetAllAvailableFileNames().Distinct(StringComparer.OrdinalIgnoreCase).ToList()
            Dim availableText As String

            If availableNames.Count = 0 Then
                availableText = "No session files are currently available. Stage a workspace file or create a new session output first."
            Else
                availableText = "Currently available session files: " & String.Join(", ", availableNames)
            End If

            Return $"Error: Session file '{sessionFileName}' was not found. {availableText}"
        End If

        Dim targetPath = ResolveWorkspacePath(targetRelativePath)
        If targetPath Is Nothing Then
            Return "Error: Invalid target path."
        End If

        Dim targetDir = Path.GetDirectoryName(targetPath)
        If Not Directory.Exists(targetDir) Then Directory.CreateDirectory(targetDir)

        If File.Exists(targetPath) AndAlso Not overwrite Then
            Return $"Error: Target file already exists: {targetRelativePath}"
        End If

        File.Copy(att.TempFilePath, targetPath, overwrite)
        Return $"Saved '{sessionFileName}' to workspace path '{ToWorkspaceRelativePath(targetPath)}'."
    End Function

    Private Sub ChatAgentWorkspaceSaveDroppedFile(sourcePath As String)
        LoadChatAgentWorkspaceIfNeeded()

        If Not IsChatAgentWorkspaceConnected() Then Return
        If Not _chatAgentWorkspace.SaveDroppedFilesToWorkspace Then Return
        If Not _chatAgentWorkspace.AllowWrite Then Return
        If String.IsNullOrWhiteSpace(sourcePath) OrElse Not File.Exists(sourcePath) Then Return

        Try
            Dim destName = Path.GetFileName(sourcePath)
            Dim destPath = Path.Combine(_chatAgentWorkspace.RootPath, destName)
            Dim counter = 1

            While File.Exists(destPath)
                destPath = Path.Combine(
                    _chatAgentWorkspace.RootPath,
                    Path.GetFileNameWithoutExtension(destName) & $"_{counter}" & Path.GetExtension(destName))
                counter += 1
            End While

            File.Copy(sourcePath, destPath, overwrite:=False)
        Catch
        End Try
    End Sub

    Private Function BuildAgentWorkspaceToolsCore() As List(Of ModelConfig)
        Dim tools As New List(Of ModelConfig)()

        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = CA_Tool_WorkspaceRead,
            .ModelDescription = "Agent Workspace: Read File",
            .ToolPriority = 122,
            .ToolErrorHandling = "skip",
            .ToolInstructionsPrompt =
                CA_Tool_WorkspaceRead & ": Read or extract text from one workspace file. " &
                "For one local workspace PDF or Office document, prefer this tool over direct path-based extraction, because it stages the file into the current session and uses the existing attachment extraction stack. " &
                "When the returned payload contains session_file_name, use that exact name with attachment-based tools such as read_attachment, process_word_document, complete_word_tables, excel_list_live_worksheets, excel_read_live_range, and excel_complete_live_workbook.",
            .ToolDefinition =
                "{""name"":""" & CA_Tool_WorkspaceRead & """," &
                """description"":""Reads or extracts text from one workspace file and stages it into the current session when possible. When session_file_name is returned, use that exact name with attachment-based tools such as read_attachment, process_word_document, complete_word_tables, excel_list_live_worksheets, excel_read_live_range, and excel_complete_live_workbook.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """path"":{""type"":""string"",""description"":""Relative workspace file path.""}," &
                """max_chars"":{""type"":""integer"",""description"":""Maximum characters to return. Default 12000, capped.""}" &
                "},""required"":[""path""]}}"
        })


        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = CA_Tool_WorkspaceRead,
            .ModelDescription = "Agent Workspace: Read File",
            .ToolPriority = 122,
            .ToolErrorHandling = "skip",
            .ToolInstructionsPrompt =
                CA_Tool_WorkspaceRead & ": Read or extract text from one workspace file. " &
                "For one local workspace PDF or Office document, prefer this tool over calling extract_pdf_text directly on a workspace path, because it stages the file and uses the existing attachment extraction stack. " &
                "For many files, use agent_workspace_stage first and then continue processing the full staged set.",
            .ToolDefinition =
                "{""name"":""" & CA_Tool_WorkspaceRead & """," &
                """description"":""Reads or extracts text from one workspace file. Prefer this for a single workspace PDF or Office file instead of calling extract_pdf_text directly on a workspace path, because the file is staged and processed through the existing attachment extraction stack.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """path"":{""type"":""string"",""description"":""Relative workspace file path.""}," &
                """max_chars"":{""type"":""integer"",""description"":""Maximum characters to return. Default 12000, capped.""}" &
                "},""required"":[""path""]}}"
        })

        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = CA_Tool_WorkspaceWrite,
            .ModelDescription = "Agent Workspace: Write File",
            .ToolPriority = 123,
            .ToolErrorHandling = "skip",
            .ToolInstructionsPrompt =
                CA_Tool_WorkspaceWrite & ": Creates a text/code file inside the workspace. " &
                "When updating or replacing an existing workspace file, set overwrite=true. " &
                "Use agent_workspace_save_session_file for Office/PDF/session outputs.",
            .ToolDefinition =
                "{""name"":""" & CA_Tool_WorkspaceWrite & """," &
                """description"":""Creates a text file inside the workspace. To replace an existing file, set overwrite=true.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """path"":{""type"":""string"",""description"":""Relative target file path inside the workspace.""}," &
                """content"":{""type"":""string"",""description"":""File content to write.""}," &
                """overwrite"":{""type"":""boolean"",""description"":""Set to true only when the user wants to replace an existing file. Default false.""}" &
                "},""required"":[""path"",""content""]}}"
        })

        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = CA_Tool_WorkspaceSaveSessionFile,
            .ModelDescription = "Agent Workspace: Save Session File",
            .ToolPriority = 124,
            .ToolErrorHandling = "abort",
            .ToolInstructionsPrompt =
                CA_Tool_WorkspaceSaveSessionFile & ": Copies a file produced or loaded in the current agent session into the workspace. " &
                "Use this after document generation/processing tools when the user wants the result written directly to the workspace.",
            .ToolDefinition =
                "{""name"":""" & CA_Tool_WorkspaceSaveSessionFile & """," &
                """description"":""Copies a current session file or tool output into the workspace.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """session_file_name"":{""type"":""string"",""description"":""Filename available in the current session or produced by a previous tool call.""}," &
                """target_path"":{""type"":""string"",""description"":""Relative workspace target path.""}," &
                """overwrite"":{""type"":""boolean"",""description"":""Overwrite existing target. Default false.""}" &
                "},""required"":[""session_file_name"",""target_path""]}}"
        })

        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = CA_Tool_WorkspaceFileOp,
            .ModelDescription = "Agent Workspace: File Operations",
            .ToolPriority = 125,
            .ToolErrorHandling = "abort",
            .ToolInstructionsPrompt =
                CA_Tool_WorkspaceFileOp & ": Performs copy, move, rename, mkdir, rmdir, or delete inside the workspace. " &
                "Use relative paths only. " &
                "For mkdir, provide target_path with the new folder path. " &
                "For copy/move/rename, provide both source_path and target_path. " &
                "For delete and rmdir, provide source_path. " &
                "For batch copy/move, provide source_paths and target_directory. " &
                "Use dry_run=true to preview the operation without changing files. " &
                "Delete/rmdir require explicit user-enabled permission.",
            .ToolDefinition =
                "{""name"":""" & CA_Tool_WorkspaceFileOp & """," &
                """description"":""Performs safe file operations inside the workspace only. Supports batch copy/move and dry-run preview.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """action"":{""type"":""string"",""enum"":[""copy"",""move"",""rename"",""delete"",""mkdir"",""rmdir""],""description"":""Operation to perform.""}," &
                """source_path"":{""type"":""string"",""description"":""Relative source path. Required for single-file copy/move/rename/delete/rmdir.""}," &
                """source_paths"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Relative source paths for batch copy/move operations.""}," &
                """target_path"":{""type"":""string"",""description"":""Relative target path. Required for mkdir, single copy, move, and rename.""}," &
                """target_directory"":{""type"":""string"",""description"":""Relative destination directory for batch copy/move operations.""}," &
                """recursive"":{""type"":""boolean"",""description"":""For rmdir or directory copy operations. Default false.""}," &
                """overwrite"":{""type"":""boolean"",""description"":""Overwrite target where supported. Default false.""}," &
                """dry_run"":{""type"":""boolean"",""description"":""Preview the operation without changing files. Default false.""}" &
                "},""required"":[""action""]}}"
        })

        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = CA_Tool_WorkspaceSearch,
            .ModelDescription = "Agent Workspace: Search",
            .ToolPriority = 126,
            .ToolErrorHandling = "skip",
            .ToolInstructionsPrompt =
                CA_Tool_WorkspaceSearch & ": Searches visible workspace filenames and text-like file content. Use list/stage/read for deeper Office/PDF work.",
            .ToolDefinition =
                "{""name"":""" & CA_Tool_WorkspaceSearch & """," &
                """description"":""Searches workspace filenames and text-like content.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """query"":{""type"":""string"",""description"":""Search text.""}," &
                """path"":{""type"":""string"",""description"":""Relative folder path. Empty means root.""}," &
                """max_results"":{""type"":""integer"",""description"":""Maximum results. Default 50, capped.""}" &
                "},""required"":[""query""]}}"
        })

        AddWorkspaceMoreTools(tools)

        Return tools

    End Function

    Private Function GetChatAgentWorkspaceTools() As List(Of ModelConfig)
        If Not IsChatAgentWorkspaceConnected() Then
            Return New List(Of ModelConfig)()
        End If

        Return BuildAgentWorkspaceToolsCore()
    End Function

    Private Function GetAutoPilotAgentWorkspaceTools() As List(Of ModelConfig)
        Return BuildAgentWorkspaceToolsCore()
    End Function

    Friend Function IsChatAgentWorkspaceTool(toolName As String) As Boolean
        Select Case toolName
            Case CA_Tool_WorkspaceList,
                 CA_Tool_WorkspaceRead,
                 CA_Tool_WorkspaceWrite,
                 CA_Tool_WorkspaceFileOp,
                 CA_Tool_WorkspaceSaveSessionFile,
                 CA_Tool_WorkspaceSearch,
                 CA_Tool_WorkspaceFindFiles,
                 CA_Tool_WorkspaceMoveTo,
                 CA_Tool_WorkspaceCopyTo,
                 CA_Tool_WorkspaceRename,
                 CA_Tool_WorkspaceBulkRename,
                 CA_Tool_WorkspaceFileDetails,
                 CA_Tool_WorkspaceRecentFiles,
                 CA_Tool_WorkspaceCreateFolderStructure,
                 CA_Tool_WorkspaceTrash,
                 CA_Tool_WorkspaceInventoryReport
                Return True
        End Select

        Return False
    End Function

    Friend Async Function ExecuteChatAgentWorkspaceTool(
        toolCall As ToolCall,
        context As ToolExecutionContext,
        Optional cancellationToken As CancellationToken = Nothing) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName,
            .OriginalCallJson = toolCall.RawJson
        }

        Dim paramSummary As String = BuildCondensedParamSummary(toolCall.Arguments)
        Dim toolDisplayName As String = toolCall.ToolName

        Try
            Dim workspaceToolConfig = GetChatAgentWorkspaceTools().
                FirstOrDefault(Function(t) t IsNot Nothing AndAlso
                                         t.ToolName.Equals(toolCall.ToolName, StringComparison.OrdinalIgnoreCase))
            If workspaceToolConfig IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(workspaceToolConfig.ModelDescription) Then
                toolDisplayName = workspaceToolConfig.ModelDescription
            End If
        Catch
        End Try

        Try
            If Not ((_chatAgentActive AndAlso Not _apActive) OrElse _apActive OrElse HasActiveScheduledTaskWorkspace()) Then
                response.Success = False
                response.ErrorMessage = "Workspace tools are available only in Local Chat Agent mode, during AutoPilot processing, or during scheduled-task execution."
            ElseIf Not IsChatAgentWorkspaceConnected() AndAlso
                   toolCall.ToolName.Equals(CA_Tool_WorkspaceSaveSessionFile, StringComparison.OrdinalIgnoreCase) Then
                ' No workspace is connected, so there is nothing to save into. This is not a hard
                ' failure: session outputs are already delivered to the Desktop output folder and
                ' remain available as session files. Return a soft, self-correcting result so the
                ' tooling loop can finalize cleanly instead of aborting (this tool uses abort).
                response.Success = True
                response.Response = ToWorkspaceJson(New With {
                    .saved = False,
                    .reason = "no_workspace_connected",
                    .message = "No workspace is connected, so no file was saved to a workspace. The produced file is already delivered to the Desktop output folder and remains available as a session file; no further action is needed."
                })
            ElseIf Not IsChatAgentWorkspaceConnected() Then
                response.Success = False
                response.ErrorMessage = "No active workspace is available."
            Else
                Select Case toolCall.ToolName
                    Case CA_Tool_WorkspaceList
                        response.Response = ExecuteWorkspaceList(toolCall)

                    Case CA_Tool_WorkspaceRead
                        response.Response = Await ExecuteWorkspaceRead(toolCall, context)

                    Case CA_Tool_WorkspaceWrite
                        response.Response = ExecuteWorkspaceWrite(toolCall)

                    Case CA_Tool_WorkspaceSaveSessionFile
                        Dim sessionName = GetArgString(toolCall.Arguments, "session_file_name")
                        Dim targetPath = GetArgString(toolCall.Arguments, "target_path")
                        Dim overwrite = GetArgBool(toolCall.Arguments, "overwrite", False)
                        response.Response = CopySessionFileToWorkspace(sessionName, targetPath, overwrite)

                    Case CA_Tool_WorkspaceFileOp
                        response.Response = ExecuteWorkspaceFileOp(toolCall)

                    Case CA_Tool_WorkspaceSearch
                        response.Response = ExecuteWorkspaceSearch(toolCall)

                    Case CA_Tool_WorkspaceFindFiles
                        response.Response = ExecuteWorkspaceFindFiles(toolCall)

                    Case CA_Tool_WorkspaceMoveTo
                        response.Response = ExecuteWorkspaceMoveTo(toolCall)

                    Case CA_Tool_WorkspaceCopyTo
                        response.Response = ExecuteWorkspaceCopyTo(toolCall)

                    Case CA_Tool_WorkspaceRename
                        response.Response = ExecuteWorkspaceRename(toolCall)

                    Case CA_Tool_WorkspaceBulkRename
                        response.Response = ExecuteWorkspaceBulkRename(toolCall)

                    Case CA_Tool_WorkspaceFileDetails
                        response.Response = ExecuteWorkspaceFileDetails(toolCall)

                    Case CA_Tool_WorkspaceRecentFiles
                        response.Response = ExecuteWorkspaceRecentFiles(toolCall)

                    Case CA_Tool_WorkspaceCreateFolderStructure
                        response.Response = ExecuteWorkspaceCreateFolderStructure(toolCall)

                    Case CA_Tool_WorkspaceTrash
                        response.Response = ExecuteWorkspaceTrash(toolCall)

                    Case CA_Tool_WorkspaceInventoryReport
                        response.Response = ExecuteWorkspaceInventoryReport(toolCall)

                    Case Else
                        response.Success = False
                        response.ErrorMessage = "Unknown workspace tool."
                End Select

                If response.Success AndAlso
                   response.Response IsNot Nothing AndAlso
                   response.Response.StartsWith("Error:", StringComparison.OrdinalIgnoreCase) Then
                    response.Success = False
                    response.ErrorMessage = response.Response
                End If
            End If
        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation was cancelled"
        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
        End Try

        If response.Success AndAlso Not String.IsNullOrWhiteSpace(response.Response) Then
            RecordWorkspaceHistory($"{toolCall.ToolName}: {response.Response}")
            context.Log($"Workspace tool completed successfully ({response.Response.Length} chars)", "success")
        Else
            context.Log($"Workspace tool failed: {If(response.ErrorMessage, "Unknown error")}", "error")
        End If

        If _apCurrentToolCallLog IsNot Nothing Then
            Dim elapsed = DateTime.Now - response.Timestamp
            Dim excerpt As String = BuildResultExcerpt(
                If(response.Success, response.Response, If(response.ErrorMessage, "")),
                80)

            ApDashboardLog($"🗂 Workspace tool: {toolCall.ToolName}{paramSummary}", If(response.Success, "info", "error"))
            If response.Success Then
                ApDashboardLog($"   ✓ {excerpt}", "info")
            Else
                ApDashboardLog($"   ✗ {If(response.ErrorMessage, excerpt)}", "error")
            End If

            RecordAutoPilotToolCall(
                toolCall.ToolName,
                toolDisplayName,
                paramSummary,
                isInternalTool:=True,
                wasSuccessful:=response.Success,
                resultExcerpt:=excerpt,
                elapsed:=elapsed,
                urls:=Nothing)
        End If

        Return response
    End Function



    Private Function ExecuteWorkspaceList(toolCall As ToolCall) As String
        If Not _chatAgentWorkspace.AllowRead Then Return "Error: Workspace read permission is disabled."

        Dim rel = GetArgString(toolCall.Arguments, "path")
        Dim recursive = GetArgBool(toolCall.Arguments, "recursive", False)
        Dim maxItems = Math.Min(Math.Max(GetArgInt(toolCall.Arguments, "max_items", 200), 1), 1000)
        Dim dir = ResolveWorkspacePath(rel, allowRoot:=True)

        If dir Is Nothing OrElse Not Directory.Exists(dir) Then Return "Error: Folder not found or invalid path."

        Dim sb As New StringBuilder()
        sb.AppendLine($"Workspace listing: {If(String.IsNullOrWhiteSpace(rel), ".", rel)}")

        Dim count = 0
        Dim pending As New Queue(Of String)()
        pending.Enqueue(dir)

        While pending.Count > 0 AndAlso count < maxItems
            Dim current = pending.Dequeue()

            For Each subdir In Directory.GetDirectories(current)
                If ShouldSkipHiddenSystem(subdir) Then Continue For
                sb.AppendLine($"[dir]  {ToWorkspaceRelativePath(subdir)}")
                count += 1
                If count >= maxItems Then Exit For
                If recursive Then pending.Enqueue(subdir)
            Next

            If count >= maxItems Then Exit While

            For Each filePath In Directory.GetFiles(current)
                If ShouldSkipHiddenSystem(filePath) Then Continue For
                Dim fi As New FileInfo(filePath)
                sb.AppendLine($"[file] {ToWorkspaceRelativePath(filePath)} ({fi.Length} bytes, modified {fi.LastWriteTime:yyyy-MM-dd HH:mm})")
                count += 1
                If count >= maxItems Then Exit For
            Next
        End While

        If count >= maxItems Then sb.AppendLine($"Result capped at {maxItems} item(s).")
        Return sb.ToString()
    End Function


    ''' <summary>
    ''' Unified text extractor for any supported file in the Local Chat workspace.
    ''' Mirrors Word's GetFileContentEx pipeline using existing SharedMethods readers.
    ''' Does NOT stage the file into the chat-agent session/temp store.
    ''' </summary>
    Private Async Function ChatAgentExtractFileText(filePath As String) As Task(Of String)
        If String.IsNullOrWhiteSpace(filePath) OrElse Not IO.File.Exists(filePath) Then Return ""
        Dim ext As String = IO.Path.GetExtension(filePath).ToLowerInvariant()

        Try
            Select Case ext
                Case ".txt", ".ini", ".csv", ".tsv", ".log", ".json", ".xml", ".html", ".htm",
                     ".md", ".yaml", ".yml", ".ics", ".vcf",
                     ".vb", ".cs", ".js", ".ts", ".py", ".java", ".cpp", ".c", ".h", ".sql"
                    Return SharedMethods.ReadTextFile(filePath, ReturnErrorInsteadOfEmpty:=False)
                Case ".rtf"
                    Return SharedMethods.ReadRtfAsText(filePath, ReturnErrorInsteadOfEmpty:=False)

                Case ".doc"
                    If INI_AllowLegacyDocFiles Then Return SharedMethods.ReadWordDocument(filePath, ReturnErrorInsteadOfEmpty:=False)
                    Return ""

                Case ".docx", ".docm"
                    Return SharedMethods.ReadDocxSandboxed(filePath)
                Case ".xlsx", ".xlsm"
                    Return SharedMethods.ReadXlsxSandboxed(filePath, silent:=True, askWorksheetSelection:=False)
                Case ".pptx", ".pptm"
                    Return SharedMethods.ReadPptxSandboxed(filePath)
                Case ".eml"
                    Return SharedMethods.ReadEmlSandboxed(filePath)
                Case ".msg"
                    Return ReadMsgAttachmentText(filePath)

                Case ".pdf"
                    Dim r = Await SharedMethods.ReadPdfAsTextEx(filePath, True, DoOCR:=True, AskUser:=False, context:=_context).ConfigureAwait(False)
                    Return If(r?.Content, "")

                Case Else
                    If SharedMethods.IsBinaryMediaExtension(ext) Then
                        Dim taskFlag = SharedMethods.TaskFlagForExtension(ext)
                        If SharedMethods.IsBinaryMediaSupported(_context, ext, taskFlag) Then
                            Return Await SharedMethods.ReadBinaryFileViaLLM(filePath, _context, "", askUser:=False, taskFlag:=taskFlag).ConfigureAwait(False)
                        End If
                    End If
                    Return ""
            End Select
        Catch ex As Exception
            Debug.WriteLine("ChatAgentExtractFileText failed for '" & filePath & "': " & ex.Message)
            Return ""
        End Try
    End Function

    Private Async Function ExecuteWorkspaceRead(toolCall As ToolCall, context As ToolExecutionContext) As Task(Of String)
        If Not _chatAgentWorkspace.AllowRead Then Return "Error: Workspace read permission is disabled."

        Dim rel = GetArgString(toolCall.Arguments, "path")
        Dim maxChars = Math.Min(Math.Max(GetArgInt(toolCall.Arguments, "max_chars", 12000), 1000), 500000)

        Dim startChar As Integer = GetArgInt(toolCall.Arguments, "start_char", -1)
        If startChar < 0 Then
            startChar = GetArgInt(toolCall.Arguments, "offset", 0)
        End If
        startChar = Math.Max(startChar, 0)

        Dim startPage As Integer = GetArgInt(toolCall.Arguments, "start_page", 0)
        Dim endPage As Integer = GetArgInt(toolCall.Arguments, "end_page", 0)

        Dim fullPath = ResolveWorkspacePath(rel)
        If String.IsNullOrWhiteSpace(fullPath) OrElse Not File.Exists(fullPath) Then
            Return "Error: Workspace file not found."
        End If

        Dim stagedAttachment As AutoPilotAttachmentInfo = StageWorkspaceFile(rel)

        Dim readPath As String = fullPath
        If stagedAttachment IsNot Nothing AndAlso
           Not String.IsNullOrWhiteSpace(stagedAttachment.TempFilePath) AndAlso
           File.Exists(stagedAttachment.TempFilePath) Then
            readPath = stagedAttachment.TempFilePath
        End If

        Dim fullText As String = Await ChatAgentExtractFileText(readPath).ConfigureAwait(False)
        fullText = If(fullText, "")

        Dim totalChars As Integer = fullText.Length
        Dim safeStart As Integer = Math.Min(startChar, totalChars)
        Dim remaining As Integer = Math.Max(totalChars - safeStart, 0)
        Dim takeChars As Integer = Math.Min(maxChars, remaining)
        Dim chunk As String = If(takeChars > 0, fullText.Substring(safeStart, takeChars), "")
        Dim truncated As Boolean = (safeStart + takeChars) < totalChars
        Dim nextOffset As Integer = safeStart + takeChars

        Dim payload As New JObject(
        New JProperty("path", If(rel, "")),
        New JProperty("text", chunk),
        New JProperty("excerpt", BuildResultExcerpt(chunk, 800)),
        New JProperty("total_chars", totalChars),
        New JProperty("start_char", safeStart),
        New JProperty("returned_chars", takeChars),
        New JProperty("truncated", truncated),
        New JProperty("continuation", "If more content is needed, call the same tool again with start_char=next_offset and a suitable max_chars value."))

        If stagedAttachment IsNot Nothing Then
            payload("staged") = True
            payload("session_file_name") = If(stagedAttachment.OriginalFileName, "")
        Else
            payload("staged") = False
        End If

        If truncated Then
            payload("next_offset") = nextOffset
        End If

        If startPage > 0 Then
            payload("start_page") = startPage
        End If

        If endPage > 0 Then
            payload("end_page") = endPage
        End If

        Return payload.ToString(Formatting.None)
    End Function

    Private Function ExecuteWorkspaceWrite(toolCall As ToolCall) As String
        If Not _chatAgentWorkspace.AllowWrite Then Return "Error: Workspace write permission is disabled."

        Dim rel = GetArgString(toolCall.Arguments, "path")
        Dim content = GetArgString(toolCall.Arguments, "content")
        Dim overwrite = GetArgBool(toolCall.Arguments, "overwrite", False)
        Dim targetPath = ResolveWorkspacePath(rel)

        If targetPath Is Nothing Then Return "Error: Invalid workspace target path."

        Dim alreadyExists = File.Exists(targetPath)
        If alreadyExists AndAlso Not overwrite Then
            Return "Error: Target file already exists. Re-run agent_workspace_write with overwrite=true to replace it."
        End If

        Dim dir = Path.GetDirectoryName(targetPath)
        If Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)

        File.WriteAllText(targetPath, If(content, ""), Encoding.UTF8)

        If alreadyExists Then
            Return $"Overwritten workspace file '{ToWorkspaceRelativePath(targetPath)}'."
        End If

        Return $"Written workspace file '{ToWorkspaceRelativePath(targetPath)}'."
    End Function

    Private Function ExecuteWorkspaceFileOp(toolCall As ToolCall) As String
        Dim action = If(GetArgString(toolCall.Arguments, "action"), "").Trim().ToLowerInvariant()
        Dim sourceRel = If(GetArgString(toolCall.Arguments, "source_path"), "").Trim()
        Dim sourcePaths = GetArgStringList(toolCall.Arguments, "source_paths")
        Dim targetRel = If(GetArgString(toolCall.Arguments, "target_path"), "").Trim()
        Dim targetDirectoryRel = If(GetArgString(toolCall.Arguments, "target_directory"), "").Trim()
        Dim recursive = GetArgBool(toolCall.Arguments, "recursive", False)
        Dim overwrite = GetArgBool(toolCall.Arguments, "overwrite", False)
        Dim dryRun = GetArgBool(toolCall.Arguments, "dry_run", False)

        If sourcePaths.Count > 0 AndAlso (action = "copy" OrElse action = "move") Then
            If Not _chatAgentWorkspace.AllowMoveCopyRename Then Return "Error: Copy/move/rename permission is disabled."

            Dim targetDirectoryPath = ResolveWorkspacePath(targetDirectoryRel)
            If String.IsNullOrWhiteSpace(targetDirectoryRel) OrElse targetDirectoryPath Is Nothing Then
                Return "Error: Invalid target_directory for batch operation."
            End If

            Dim messages As New List(Of String)()

            For Each srcRel In sourcePaths
                Dim srcPath = ResolveWorkspacePath(srcRel)
                If srcPath Is Nothing Then
                    messages.Add($"✗ {srcRel}: invalid source path")
                    Continue For
                End If

                Dim destPath = Path.Combine(targetDirectoryPath, Path.GetFileName(srcPath))
                If Not IsPathWithinOrEqual(destPath, _chatAgentWorkspace.RootPath) Then
                    messages.Add($"✗ {srcRel}: invalid destination path")
                    Continue For
                End If

                If dryRun Then
                    messages.Add($"• {action}: {srcRel} -> {ToWorkspaceRelativePath(destPath)}")
                    Continue For
                End If

                If Not Directory.Exists(targetDirectoryPath) Then Directory.CreateDirectory(targetDirectoryPath)

                If File.Exists(srcPath) Then
                    If File.Exists(destPath) AndAlso Not overwrite Then
                        messages.Add($"✗ {srcRel}: target already exists")
                        Continue For
                    End If

                    If action = "copy" Then
                        File.Copy(srcPath, destPath, overwrite)
                    Else
                        If File.Exists(destPath) AndAlso overwrite Then File.Delete(destPath)
                        File.Move(srcPath, destPath)
                    End If

                    messages.Add($"✓ {action}: {srcRel} -> {ToWorkspaceRelativePath(destPath)}")
                Else
                    messages.Add($"✗ {srcRel}: source not found")
                End If
            Next

            If dryRun Then
                Return "Dry run:" & vbCrLf & String.Join(vbCrLf, messages)
            End If

            Return String.Join(vbCrLf, messages)
        End If

        Select Case action
            Case "mkdir"
                If Not _chatAgentWorkspace.AllowWrite Then Return "Error: Workspace write permission is disabled."

                Dim mkdirRel = If(Not String.IsNullOrWhiteSpace(targetRel), targetRel, sourceRel)
                Dim mkdirPath = ResolveWorkspacePath(mkdirRel)

                If String.IsNullOrWhiteSpace(mkdirRel) OrElse mkdirPath Is Nothing Then
                    Return "Error: Invalid folder path. For mkdir, provide target_path (preferred) or source_path as a relative workspace folder path."
                End If

                If dryRun Then
                    Return $"Dry run: mkdir '{mkdirRel}'"
                End If

                Directory.CreateDirectory(mkdirPath)
                Return $"Created folder '{ToWorkspaceRelativePath(mkdirPath)}'."

            Case "delete"
                If Not _chatAgentWorkspace.AllowDelete Then Return "Error: Delete permission is disabled."

                Dim deleteRel = If(Not String.IsNullOrWhiteSpace(sourceRel), sourceRel, targetRel)
                Dim source = ResolveWorkspacePath(deleteRel)

                If source Is Nothing OrElse Not File.Exists(source) Then
                    Return "Error: File not found."
                End If

                If dryRun Then
                    Return $"Dry run: delete '{deleteRel}'"
                End If

                File.Delete(source)
                Return $"Deleted file '{ToWorkspaceRelativePath(source)}'."

            Case "rmdir"
                If Not _chatAgentWorkspace.AllowDelete Then Return "Error: Delete permission is disabled."

                Dim rmdirRel = If(Not String.IsNullOrWhiteSpace(sourceRel), sourceRel, targetRel)
                Dim source = ResolveWorkspacePath(rmdirRel)

                If source Is Nothing OrElse Not Directory.Exists(source) Then
                    Return "Error: Folder not found."
                End If

                If dryRun Then
                    Return $"Dry run: rmdir '{rmdirRel}' (recursive={recursive})"
                End If

                Directory.Delete(source, recursive)
                Return $"Removed folder '{ToWorkspaceRelativePath(source)}'."

            Case "copy", "move", "rename"
                If Not _chatAgentWorkspace.AllowMoveCopyRename Then Return "Error: Copy/move/rename permission is disabled."

                Dim source = ResolveWorkspacePath(sourceRel)
                Dim target = ResolveWorkspacePath(targetRel)

                If String.IsNullOrWhiteSpace(sourceRel) OrElse source Is Nothing Then
                    Return "Error: Invalid source path."
                End If

                If String.IsNullOrWhiteSpace(targetRel) OrElse target Is Nothing Then
                    Return "Error: Invalid target path."
                End If

                If dryRun Then
                    Return $"Dry run: {action} '{sourceRel}' -> '{targetRel}'"
                End If

                Dim targetDir = Path.GetDirectoryName(target)
                If Not Directory.Exists(targetDir) Then Directory.CreateDirectory(targetDir)

                If File.Exists(source) Then
                    If File.Exists(target) AndAlso Not overwrite Then
                        Return "Error: Target file already exists."
                    End If

                    If action = "copy" Then
                        File.Copy(source, target, overwrite)
                    Else
                        If File.Exists(target) AndAlso overwrite Then File.Delete(target)
                        File.Move(source, target)
                    End If

                ElseIf Directory.Exists(source) Then
                    If action = "copy" Then
                        CopyDirectorySafe(source, target, recursive, overwrite)
                    Else
                        If Directory.Exists(target) AndAlso Not overwrite Then
                            Return "Error: Target folder already exists."
                        End If
                        Directory.Move(source, target)
                    End If

                Else
                    Return "Error: Source not found."
                End If

                Return $"{action} completed: '{sourceRel}' -> '{targetRel}'."

            Case Else
                Return "Error: Unsupported workspace file operation."
        End Select
    End Function

    Private Sub CopyDirectorySafe(sourceDir As String, targetDir As String, recursive As Boolean, overwrite As Boolean)
        Directory.CreateDirectory(targetDir)

        For Each filePath In Directory.GetFiles(sourceDir)
            If ShouldSkipHiddenSystem(filePath) Then Continue For
            Dim dest = Path.Combine(targetDir, Path.GetFileName(filePath))
            File.Copy(filePath, dest, overwrite)
        Next

        If recursive Then
            For Each subdir In Directory.GetDirectories(sourceDir)
                If ShouldSkipHiddenSystem(subdir) Then Continue For
                CopyDirectorySafe(subdir, Path.Combine(targetDir, Path.GetFileName(subdir)), recursive, overwrite)
            Next
        End If
    End Sub

    Private Function ExecuteWorkspaceSearch(toolCall As ToolCall) As String
        If Not _chatAgentWorkspace.AllowRead Then Return "Error: Workspace read permission is disabled."

        Dim query = If(GetArgString(toolCall.Arguments, "query"), "").Trim()
        Dim rel = GetArgString(toolCall.Arguments, "path")
        Dim maxResults = Math.Min(Math.Max(GetArgInt(toolCall.Arguments, "max_results", 50), 1), 200)
        Dim root = ResolveWorkspacePath(rel, allowRoot:=True)

        If String.IsNullOrWhiteSpace(query) Then Return "Error: Search query is empty."
        If root Is Nothing OrElse Not Directory.Exists(root) Then Return "Error: Search folder not found."

        Dim results As New List(Of String)()
        Dim textExt = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
            ".txt", ".md", ".csv", ".json", ".xml", ".html", ".htm", ".vb", ".cs", ".js", ".ts", ".css", ".sql", ".ps1", ".bat", ".cmd", ".ini", ".config", ".log"
        }

        For Each filePath In Directory.GetFiles(root, "*", System.IO.SearchOption.AllDirectories)
            If results.Count >= maxResults Then Exit For
            If ShouldSkipHiddenSystem(filePath) Then Continue For

            Dim relPath = ToWorkspaceRelativePath(filePath)
            If relPath.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0 Then
                results.Add("[name] " & relPath)
                Continue For
            End If

            If textExt.Contains(Path.GetExtension(filePath)) Then
                Try
                    Dim text = File.ReadAllText(filePath)
                    Dim ix = text.IndexOf(query, StringComparison.OrdinalIgnoreCase)
                    If ix >= 0 Then
                        Dim startIx = Math.Max(0, ix - 80)
                        Dim len = Math.Min(220, text.Length - startIx)
                        Dim excerpt = text.Substring(startIx, len).Replace(vbCr, " ").Replace(vbLf, " ")
                        results.Add("[content] " & relPath & " — " & excerpt)
                    End If
                Catch
                End Try
            End If
        Next

        If results.Count = 0 Then Return "No workspace search results."
        Return String.Join(vbCrLf, results)
    End Function


    Friend Function BuildAgentWorkspacePromptBlock(Optional maxItems As Integer = 25) As String
        If Not IsChatAgentWorkspaceConnected() Then Return ""

        Dim sb As New StringBuilder()
        Dim root = _chatAgentWorkspace.RootPath

        sb.AppendLine("[AGENT WORKSPACE]")
        sb.AppendLine("A local agent workspace is connected. Workspace files are addressable by their workspace-relative path; you read them in place.")
        sb.AppendLine("Preferred workspace workflow:")
        sb.AppendLine("  1. Use agent_workspace_find_files or agent_workspace_list to locate files.")
        sb.AppendLine("  2. Use workspace_extract_text(path) to read ANY supported file (PDF, DOCX, XLSX, PPTX, RTF, DOC, EML, MSG, plain text/markdown/code; images/audio/video when the model supports them). This is the only file-reading tool you should normally need.")
        sb.AppendLine("  3. Use agent_workspace_write for text/code outputs. Pass overwrite=true to replace existing files.")
        sb.AppendLine("Do NOT copy workspace files into the chat session. Do NOT use list_attachments to look for them — that lists session attachments only, not workspace contents.")
        sb.AppendLine("Safety rules:")
        sb.AppendLine("  - For bulk rename, prefer dry_run=true first.")
        sb.AppendLine("  - For batch move/copy, use dry_run=true when ambiguity exists.")
        sb.AppendLine("  - Prefer agent_workspace_trash over permanent delete.")
        sb.AppendLine("  - Use agent_workspace_inventory_report when the user asks for an overview, list, register, handover list, or archive inventory.")
        sb.AppendLine("Do NOT repeat the same failing workspace tool call with identical arguments.")
        sb.AppendLine($"Workspace root: {root}")

        Try
            Dim visibleItems As New List(Of String)()

            For Each subdir In Directory.GetDirectories(root)
                If ShouldSkipHiddenSystem(subdir) Then Continue For
                visibleItems.Add("[dir]  " & ToWorkspaceRelativePath(subdir))
                If visibleItems.Count >= maxItems Then Exit For
            Next

            If visibleItems.Count < maxItems Then
                For Each filePath In Directory.GetFiles(root)
                    If ShouldSkipHiddenSystem(filePath) Then Continue For
                    Dim fi As New FileInfo(filePath)
                    visibleItems.Add("[file] " & ToWorkspaceRelativePath(filePath) & $" ({fi.Length / 1024.0:F0} KB)")
                    If visibleItems.Count >= maxItems Then Exit For
                Next
            End If

            If visibleItems.Count > 0 Then
                sb.AppendLine("Top-level visible workspace items:")
                For Each item In visibleItems
                    sb.AppendLine("  " & item)
                Next
            Else
                sb.AppendLine("Top-level visible workspace items: (none)")
            End If
        Catch
            sb.AppendLine("Top-level visible workspace items: (could not enumerate)")
        End Try

        sb.AppendLine("[/AGENT WORKSPACE]")
        Return sb.ToString().TrimEnd()
    End Function

    Private Sub RecordWorkspaceHistory(entry As String)
        If String.IsNullOrWhiteSpace(entry) Then Return

        _chatAgentWorkspaceHistory.Add($"{DateTime.Now:HH:mm:ss} {entry}")

        While _chatAgentWorkspaceHistory.Count > CA_MaxWorkspaceHistoryEntries
            _chatAgentWorkspaceHistory.RemoveAt(0)
        End While
    End Sub

    Friend Function BuildAgentWorkspaceHistoryPromptBlock() As String
        If _chatAgentWorkspaceHistory.Count = 0 Then Return ""

        Dim sb As New StringBuilder()
        sb.AppendLine("[WORKSPACE OPERATION HISTORY]")
        For Each item In _chatAgentWorkspaceHistory
            sb.AppendLine("  - " & item)
        Next
        sb.AppendLine("[/WORKSPACE OPERATION HISTORY]")
        Return sb.ToString().TrimEnd()
    End Function

    Friend Function BuildAgentSessionFilesPromptBlock() As String
        Dim hydratedFiles As New List(Of AutoPilotAttachmentInfo)()
        For Each att In _chatAgentFiles.ToList()
            Dim hydrated = EnsureSessionAttachmentAvailable(att)
            If hydrated IsNot Nothing Then
                hydratedFiles.Add(hydrated)
            End If
        Next
        _chatAgentFiles = hydratedFiles

        Dim names = GetAllAvailableFileNames().Distinct(StringComparer.OrdinalIgnoreCase).ToList()
        If names.Count = 0 Then Return ""

        Dim sb As New StringBuilder()
        sb.AppendLine("[SESSION FILES]")
        sb.AppendLine("These files currently exist in the active agent session, persist across follow-up chat turns, and may be referenced by session tools or agent_workspace_save_session_file.")
        If Not IsChatAgentWorkspaceConnected() Then
            sb.AppendLine("No persistent agent workspace is connected for this run. If the user asks for a text, Markdown, JSON, or code deliverable, use workspace_write with a relative filename or relative path to create it inside the temporary session workspace. Files created there will be surfaced back to the user as output files after the run.")
        End If
        For Each name In names
            sb.AppendLine("  - " & name)
        Next
        sb.AppendLine("[/SESSION FILES]")
        Return sb.ToString().TrimEnd()
    End Function

    Private Shared Function GetArgStringList(args As Dictionary(Of String, Object), key As String) As List(Of String)
        Dim result As New List(Of String)()
        If args Is Nothing OrElse Not args.ContainsKey(key) OrElse args(key) Is Nothing Then Return result

        Dim value = args(key)

        If TypeOf value Is JArray Then
            For Each item In DirectCast(value, JArray)
                result.Add(item.ToString())
            Next
        ElseIf TypeOf value Is IEnumerable(Of Object) AndAlso Not TypeOf value Is String Then
            For Each item In DirectCast(value, IEnumerable(Of Object))
                result.Add(If(item, "").ToString())
            Next
        ElseIf TypeOf value Is String Then
            result.Add(CStr(value))
        End If

        Return result.Where(Function(s) Not String.IsNullOrWhiteSpace(s)).Select(Function(s) s.Trim()).ToList()
    End Function


    Private Shared Function CreateWorkspaceToolConfig(
    toolName As String,
    modelDescription As String,
    toolPriority As Integer,
    toolErrorHandling As String,
    toolInstructionsPrompt As String,
    toolDefinition As Object) As ModelConfig

        Return New ModelConfig() With {
            .ToolOnly = True,
            .Tool = True,
            .ToolName = toolName,
            .ModelDescription = modelDescription,
            .ToolPriority = toolPriority,
            .ToolErrorHandling = toolErrorHandling,
            .ToolInstructionsPrompt = toolInstructionsPrompt,
            .ToolDefinition = JsonConvert.SerializeObject(toolDefinition, Formatting.None)
        }
    End Function

    Private Sub AddWorkspaceMoreTools(tools As List(Of ModelConfig))
        tools.Add(CreateWorkspaceToolConfig(
            CA_Tool_WorkspaceFindFiles,
            "Agent Workspace: Find Files",
            127,
            "skip",
            CA_Tool_WorkspaceFindFiles & ": Find files inside the granted workspace by filename/path text, extension, size, and modified date. " &
            "Use this first when the user refers to files vaguely or to a folder of PDFs, but treat it as discovery only: after finding the files, continue with staging/reading/searching and do not stop after listing them.",
            New With {
                .name = CA_Tool_WorkspaceFindFiles,
                .description = "Finds files inside the granted local agent workspace by filename, extension, size, and modified date. Use this for discovery only; after locating files, continue with staging or document-processing tools as needed.",
                .parameters = New With {
                    .type = "object",
                    .properties = New Dictionary(Of String, Object) From {
                        {"query", New With {.type = "string", .description = "Optional filename/folder text to match."}},
                        {"path", New With {.type = "string", .description = "Relative workspace folder path. Empty means root."}},
                        {"extensions", New With {.type = "array", .items = New With {.type = "string"}, .description = "Optional file extensions such as pdf, docx, xlsx."}},
                        {"modified_from", New With {.type = "string", .description = "Optional local date/time lower bound."}},
                        {"modified_to", New With {.type = "string", .description = "Optional local date/time upper bound."}},
                        {"min_size_bytes", New With {.type = "integer", .description = "Optional minimum file size in bytes."}},
                        {"max_size_bytes", New With {.type = "integer", .description = "Optional maximum file size in bytes."}},
                        {"recursive", New With {.type = "boolean", .description = "Whether to recurse into subfolders. Default true."}},
                        {"max_results", New With {.type = "integer", .description = "Maximum files to return. Default 50, capped."}}
                    }
                }
            }))

        tools.Add(CreateWorkspaceToolConfig(
            CA_Tool_WorkspaceMoveTo,
            "Agent Workspace: Move To Folder",
            128,
            "abort",
            CA_Tool_WorkspaceMoveTo & ": Move one or more files/folders into another workspace folder. Create the target folder if needed. Prefer dry_run first for batch moves when ambiguity exists.",
            New With {
                .name = CA_Tool_WorkspaceMoveTo,
                .description = "Moves one or more files or folders to another workspace folder.",
                .parameters = New With {
                    .type = "object",
                    .properties = New Dictionary(Of String, Object) From {
                        {"source_path", New With {.type = "string", .description = "Single relative source path."}},
                        {"source_paths", New With {.type = "array", .items = New With {.type = "string"}, .description = "Multiple relative source paths."}},
                        {"target_directory", New With {.type = "string", .description = "Relative destination folder path."}},
                        {"overwrite", New With {.type = "boolean", .description = "Overwrite or replace collisions where possible. Default false."}},
                        {"dry_run", New With {.type = "boolean", .description = "Preview the move without changing files. Default false."}}
                    },
                    .required = New String() {"target_directory"}
                }
            }))

        tools.Add(CreateWorkspaceToolConfig(
            CA_Tool_WorkspaceCopyTo,
            "Agent Workspace: Copy To Folder",
            129,
            "abort",
            CA_Tool_WorkspaceCopyTo & ": Copy one or more files/folders into another workspace folder. Never deletes the source. Show collisions clearly. Create the target folder if needed.",
            New With {
                .name = CA_Tool_WorkspaceCopyTo,
                .description = "Copies one or more files or folders to another workspace folder without deleting the source.",
                .parameters = New With {
                    .type = "object",
                    .properties = New Dictionary(Of String, Object) From {
                        {"source_path", New With {.type = "string", .description = "Single relative source path."}},
                        {"source_paths", New With {.type = "array", .items = New With {.type = "string"}, .description = "Multiple relative source paths."}},
                        {"target_directory", New With {.type = "string", .description = "Relative destination folder path."}},
                        {"overwrite", New With {.type = "boolean", .description = "Overwrite or merge collisions where possible. Default false."}},
                        {"dry_run", New With {.type = "boolean", .description = "Preview the copy without changing files. Default false."}}
                    },
                    .required = New String() {"target_directory"}
                }
            }))

        tools.Add(CreateWorkspaceToolConfig(
            CA_Tool_WorkspaceRename,
            "Agent Workspace: Rename Item",
            130,
            "abort",
            CA_Tool_WorkspaceRename & ": Rename one file or folder in place. new_name must be a leaf name only, and the item stays in the same parent folder unless a move is explicitly requested.",
            New With {
                .name = CA_Tool_WorkspaceRename,
                .description = "Renames one file or folder inside its current parent folder.",
                .parameters = New With {
                    .type = "object",
                    .properties = New Dictionary(Of String, Object) From {
                        {"source_path", New With {.type = "string", .description = "Relative source file or folder path."}},
                        {"new_name", New With {.type = "string", .description = "New leaf name only, not a path."}},
                        {"overwrite", New With {.type = "boolean", .description = "Overwrite existing target if present. Default false."}},
                        {"dry_run", New With {.type = "boolean", .description = "Preview the rename without changing files. Default false."}}
                    },
                    .required = New String() {"source_path", "new_name"}
                }
            }))

        tools.Add(CreateWorkspaceToolConfig(
            CA_Tool_WorkspaceBulkRename,
            "Agent Workspace: Bulk Rename",
            131,
            "abort",
            CA_Tool_WorkspaceBulkRename & ": Rename many files by prefix, suffix, replace, or pattern. Dry run is recommended by default, and the response returns old → new mappings before commit.",
            New With {
                .name = CA_Tool_WorkspaceBulkRename,
                .description = "Bulk-renames files using prefix, suffix, replace, or pattern rules. For pattern mode, placeholders {name}, {ext}, {index}, {date}, and {file_name} are supported.",
                .parameters = New With {
                    .type = "object",
                    .properties = New Dictionary(Of String, Object) From {
                        {"source_paths", New With {.type = "array", .items = New With {.type = "string"}, .description = "Explicit relative file paths to rename."}},
                        {"path", New With {.type = "string", .description = "Relative file or folder path. If a folder is supplied, files in that folder are renamed."}},
                        {"recursive", New With {.type = "boolean", .description = "When path is a folder, recurse into subfolders. Default false."}},
                        {"mode", New With {.type = "string", .enum = New String() {"prefix", "suffix", "replace", "pattern"}, .description = "Rename mode."}},
                        {"find_text", New With {.type = "string", .description = "Used by replace mode."}},
                        {"replace_text", New With {.type = "string", .description = "Used by replace mode."}},
                        {"prefix", New With {.type = "string", .description = "Used by prefix mode."}},
                        {"suffix", New With {.type = "string", .description = "Used by suffix mode."}},
                        {"pattern", New With {.type = "string", .description = "Used by pattern mode. Supports {name}, {ext}, {index}, {date}, {file_name}."}},
                        {"overwrite", New With {.type = "boolean", .description = "Overwrite collisions where supported. Default false."}},
                        {"dry_run", New With {.type = "boolean", .description = "Preview mappings without changing files. Default true."}}
                    },
                    .required = New String() {"mode"}
                }
            }))

        tools.Add(CreateWorkspaceToolConfig(
            CA_Tool_WorkspaceFileDetails,
            "Agent Workspace: File Details",
            132,
            "skip",
            CA_Tool_WorkspaceFileDetails & ": Return detailed metadata for a file or folder, including timestamps, size, attributes, and optionally a SHA-256 hash for files.",
            New With {
                .name = CA_Tool_WorkspaceFileDetails,
                .description = "Returns detailed metadata for a workspace file or folder.",
                .parameters = New With {
                    .type = "object",
                    .properties = New Dictionary(Of String, Object) From {
                        {"path", New With {.type = "string", .description = "Relative file or folder path."}},
                        {"include_hash", New With {.type = "boolean", .description = "Include SHA-256 for files. Default false."}}
                    },
                    .required = New String() {"path"}
                }
            }))

        tools.Add(CreateWorkspaceToolConfig(
            CA_Tool_WorkspaceRecentFiles,
            "Agent Workspace: Recent Files",
            133,
            "skip",
            CA_Tool_WorkspaceRecentFiles & ": List recently changed files to help the user resume work quickly.",
            New With {
                .name = CA_Tool_WorkspaceRecentFiles,
                .description = "Lists recently changed files in the workspace.",
                .parameters = New With {
                    .type = "object",
                    .properties = New Dictionary(Of String, Object) From {
                        {"path", New With {.type = "string", .description = "Relative workspace folder path. Empty means root."}},
                        {"hours", New With {.type = "integer", .description = "Look back this many hours."}},
                        {"days", New With {.type = "integer", .description = "Look back this many days."}},
                        {"max_results", New With {.type = "integer", .description = "Maximum files to return. Default 25, capped."}}
                    }
                }
            }))

        tools.Add(CreateWorkspaceToolConfig(
            CA_Tool_WorkspaceCreateFolderStructure,
            "Agent Workspace: Create Folder Structure",
            134,
            "abort",
            CA_Tool_WorkspaceCreateFolderStructure & ": Create multiple folders in one operation under a base path. Supports dry_run preview.",
            New With {
                .name = CA_Tool_WorkspaceCreateFolderStructure,
                .description = "Creates multiple folders inside a workspace base path.",
                .parameters = New With {
                    .type = "object",
                    .properties = New Dictionary(Of String, Object) From {
                        {"base_path", New With {.type = "string", .description = "Relative base folder path. Empty means workspace root."}},
                        {"folders", New With {.type = "array", .items = New With {.type = "string"}, .description = "Relative folder names or nested folder paths to create."}},
                        {"dry_run", New With {.type = "boolean", .description = "Preview folder creation without changing files. Default false."}}
                    },
                    .required = New String() {"folders"}
                }
            }))

        tools.Add(CreateWorkspaceToolConfig(
            CA_Tool_WorkspaceTrash,
            "Agent Workspace: Move To Recycle Bin",
            135,
            "abort",
            CA_Tool_WorkspaceTrash & ": Move files/folders to the Windows Recycle Bin instead of permanently deleting them. Prefer this over delete for office-user cleanup.",
            New With {
                .name = CA_Tool_WorkspaceTrash,
                .description = "Moves one or more workspace files or folders to the Recycle Bin.",
                .parameters = New With {
                    .type = "object",
                    .properties = New Dictionary(Of String, Object) From {
                        {"source_path", New With {.type = "string", .description = "Single relative source path."}},
                        {"source_paths", New With {.type = "array", .items = New With {.type = "string"}, .description = "Multiple relative source paths."}},
                        {"dry_run", New With {.type = "boolean", .description = "Preview the trash operation without changing files. Default false."}}
                    }
                }
            }))

        tools.Add(CreateWorkspaceToolConfig(
            CA_Tool_WorkspaceInventoryReport,
            "Agent Workspace: Inventory Report",
            136,
            "abort",
            CA_Tool_WorkspaceInventoryReport & ": Create an Excel or Word report of files in a workspace folder. Use this when the user asks for an overview, list, register, or handover inventory.",
            New With {
                .name = CA_Tool_WorkspaceInventoryReport,
                .description = "Creates an Excel or Word inventory report of files in a workspace folder.",
                .parameters = New With {
                    .type = "object",
                    .properties = New Dictionary(Of String, Object) From {
                        {"path", New With {.type = "string", .description = "Relative workspace folder path. Empty means root."}},
                        {"recursive", New With {.type = "boolean", .description = "Whether to recurse into subfolders. Default false."}},
                        {"max_items", New With {.type = "integer", .description = "Maximum files to include. Default 200, capped."}},
                        {"output_format", New With {.type = "string", .enum = New String() {"excel", "word"}, .description = "Report format."}},
                        {"output_filename", New With {.type = "string", .description = "Leaf filename to create in the target folder. If omitted, a timestamped name is generated."}}
                    },
                    .required = New String() {"output_format"}
                }
            }))
    End Sub


    Private Shared Function ToWorkspaceJson(value As Object) As String
        Return JsonConvert.SerializeObject(value, Formatting.Indented)
    End Function

    Private Shared Function GetArgLongValue(args As Dictionary(Of String, Object), key As String, defaultValue As Long) As Long
        Try
            If args Is Nothing OrElse Not args.ContainsKey(key) OrElse args(key) Is Nothing Then Return defaultValue

            Dim value = args(key)

            If TypeOf value Is Long Then Return CLng(value)
            If TypeOf value Is Integer Then Return CLng(CInt(value))
            If TypeOf value Is Short Then Return CLng(CShort(value))
            If TypeOf value Is Double Then Return CLng(CDbl(value))
            If TypeOf value Is Decimal Then Return CLng(CDec(value))
            If TypeOf value Is JValue Then Return System.Convert.ToInt64(DirectCast(value, JValue).Value, CultureInfo.InvariantCulture)

            Return System.Convert.ToInt64(value, CultureInfo.InvariantCulture)
        Catch
            Return defaultValue
        End Try
    End Function

    Private Shared Function NormalizeExtensionFilter(ext As String) As String
        Dim value = If(ext, "").Trim()
        If value = "" Then Return ""
        If Not value.StartsWith(".", StringComparison.Ordinal) Then value = "." & value
        Return value.ToLowerInvariant()
    End Function

    Private Shared Function IsLeafNameOnly(value As String) As Boolean
        Dim text = If(value, "").Trim()
        If text = "" Then Return False
        If text = "." OrElse text = ".." Then Return False
        If text.IndexOf(Path.DirectorySeparatorChar) >= 0 Then Return False
        If text.IndexOf(Path.AltDirectorySeparatorChar) >= 0 Then Return False

        For Each ch In Path.GetInvalidFileNameChars()
            If text.IndexOf(ch) >= 0 Then Return False
        Next

        Return True
    End Function

    Private Shared Function ReplaceIgnoreCase(sourceText As String, findText As String, replaceText As String) As String
        Dim input = If(sourceText, "")
        Dim findValue = If(findText, "")
        Dim replacement = If(replaceText, "")

        If input = "" OrElse findValue = "" Then Return input

        Dim sb As New StringBuilder(input.Length + Math.Max(0, replacement.Length - findValue.Length) * 4)
        Dim startIx = 0
        Dim hitIx = input.IndexOf(findValue, StringComparison.OrdinalIgnoreCase)

        While hitIx >= 0
            sb.Append(input.Substring(startIx, hitIx - startIx))
            sb.Append(replacement)
            startIx = hitIx + findValue.Length
            hitIx = input.IndexOf(findValue, startIx, StringComparison.OrdinalIgnoreCase)
        End While

        sb.Append(input.Substring(startIx))
        Return sb.ToString()
    End Function

    Private Shared Function TryParseToolDateUtc(value As String, ByRef parsedUtc As DateTime?) As Boolean
        parsedUtc = Nothing

        If String.IsNullOrWhiteSpace(value) Then Return True

        Dim dt As DateTime
        Dim styles = DateTimeStyles.AllowWhiteSpaces Or DateTimeStyles.AssumeLocal

        If DateTime.TryParse(value, CultureInfo.CurrentCulture, styles, dt) OrElse
           DateTime.TryParse(value, CultureInfo.InvariantCulture, styles, dt) Then

            If dt.Kind = DateTimeKind.Unspecified Then
                dt = DateTime.SpecifyKind(dt, DateTimeKind.Local)
            End If

            parsedUtc = dt.ToUniversalTime()
            Return True
        End If

        Return False
    End Function

    Private Shared Function GetPathLeafName(fullPath As String) As String
        If Directory.Exists(fullPath) Then
            Return New DirectoryInfo(fullPath).Name
        End If

        Return Path.GetFileName(fullPath)
    End Function

    Private Function GetWorkspaceVisibleFiles(rootPath As String, recursive As Boolean) As List(Of String)
        Dim results As New List(Of String)()

        Try
            Dim optionValue = If(recursive, System.IO.SearchOption.AllDirectories, System.IO.SearchOption.TopDirectoryOnly)

            For Each filePath In Directory.GetFiles(rootPath, "*", optionValue)
                If ShouldSkipHiddenSystem(filePath) Then Continue For
                results.Add(filePath)
            Next
        Catch
        End Try

        Return results
    End Function

    Private Shared Function WorkspacePathMatchesQuery(relativePath As String, query As String) As Boolean
        Dim q = If(query, "").Trim()
        If q = "" Then Return True

        Dim haystack = If(relativePath, "")
        Dim terms = q.Split(New Char() {" "c, ChrW(9), ","c, ";"c}, StringSplitOptions.RemoveEmptyEntries)
        If terms.Length = 0 Then Return True

        For Each term In terms
            If haystack.IndexOf(term, StringComparison.OrdinalIgnoreCase) < 0 Then
                Return False
            End If
        Next

        Return True
    End Function

    Private Function BuildWorkspaceFileRecord(filePath As String) As Dictionary(Of String, Object)
        Dim fi As New FileInfo(filePath)
        Dim rel = ToWorkspaceRelativePath(filePath)
        Dim folder = Path.GetDirectoryName(rel.Replace("/"c, Path.DirectorySeparatorChar))

        Return New Dictionary(Of String, Object) From {
            {"path", rel},
            {"filename", fi.Name},
            {"folder", If(folder, "").Replace(Path.DirectorySeparatorChar, "/"c)},
            {"type", If(fi.Extension, "").TrimStart("."c).ToLowerInvariant()},
            {"extension", fi.Extension.ToLowerInvariant()},
            {"size_bytes", fi.Length},
            {"modified", fi.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")},
            {"created", fi.CreationTime.ToString("yyyy-MM-dd HH:mm:ss")}
        }
    End Function

    Private Shared Function GetUniqueFilePath(pathValue As String) As String
        If String.IsNullOrWhiteSpace(pathValue) Then
            Return pathValue
        End If

        If Not File.Exists(pathValue) AndAlso Not Directory.Exists(pathValue) Then
            Return pathValue
        End If

        Dim directoryName As String = Path.GetDirectoryName(pathValue)
        Dim baseName As String = Path.GetFileNameWithoutExtension(pathValue)
        Dim extension As String = Path.GetExtension(pathValue)

        For i As Integer = 1 To 9999
            Dim candidate As String = Path.Combine(directoryName, $"{baseName} ({i}){extension}")

            If Not File.Exists(candidate) AndAlso Not Directory.Exists(candidate) Then
                Return candidate
            End If
        Next

        Return Path.Combine(directoryName, $"{baseName}-{Guid.NewGuid():N}{extension}")
    End Function

    Private Shared Function RemoveWorkspacePath(pathValue As String) As Boolean
        If File.Exists(pathValue) Then
            File.Delete(pathValue)
            Return True
        End If

        If Directory.Exists(pathValue) Then
            Directory.Delete(pathValue, recursive:=True)
            Return True
        End If

        Return False
    End Function

    Private Shared Function GetSha256Hex(filePath As String) As String
        Using sha = SHA256.Create()
            Using fs = File.OpenRead(filePath)
                Dim hash = sha.ComputeHash(fs)
                Dim sb As New StringBuilder(hash.Length * 2)

                For Each b In hash
                    sb.Append(b.ToString("x2", CultureInfo.InvariantCulture))
                Next

                Return sb.ToString()
            End Using
        End Using
    End Function

    Private Function GetDirectorySizeSafe(dirPath As String) As Long
        Dim total As Long = 0

        Try
            For Each filePath In Directory.GetFiles(dirPath, "*", System.IO.SearchOption.AllDirectories)
                If ShouldSkipHiddenSystem(filePath) Then Continue For
                total += New FileInfo(filePath).Length
            Next
        Catch
        End Try

        Return total
    End Function

    Private Function CopyWorkspaceItem(sourcePath As String, targetPath As String, overwrite As Boolean, ByRef message As String) As Boolean
        message = ""

        If sourcePath.Equals(targetPath, StringComparison.OrdinalIgnoreCase) Then
            message = "Source and target are identical."
            Return False
        End If

        If File.Exists(sourcePath) Then
            Dim targetDir = Path.GetDirectoryName(targetPath)
            If Not Directory.Exists(targetDir) Then Directory.CreateDirectory(targetDir)

            If Directory.Exists(targetPath) Then
                message = "Target path is an existing folder."
                Return False
            End If

            If File.Exists(targetPath) AndAlso Not overwrite Then
                message = "Target file already exists."
                Return False
            End If

            File.Copy(sourcePath, targetPath, overwrite)
            Return True
        End If

        If Directory.Exists(sourcePath) Then
            If IsPathWithinOrEqual(targetPath, sourcePath) Then
                message = "Cannot copy a folder into itself."
                Return False
            End If

            If File.Exists(targetPath) Then
                message = "Target path is an existing file."
                Return False
            End If

            If Directory.Exists(targetPath) AndAlso Not overwrite Then
                message = "Target folder already exists."
                Return False
            End If

            If Not Directory.Exists(targetPath) Then Directory.CreateDirectory(targetPath)
            CopyDirectorySafe(sourcePath, targetPath, recursive:=True, overwrite:=overwrite)
            Return True
        End If

        message = "Source not found."
        Return False
    End Function

    Private Function MoveWorkspaceItem(sourcePath As String, targetPath As String, overwrite As Boolean, ByRef message As String) As Boolean
        message = ""

        If sourcePath.Equals(targetPath, StringComparison.OrdinalIgnoreCase) Then
            message = "Source and target are identical."
            Return False
        End If

        If File.Exists(sourcePath) Then
            Dim targetDir = Path.GetDirectoryName(targetPath)
            If Not Directory.Exists(targetDir) Then Directory.CreateDirectory(targetDir)

            If Directory.Exists(targetPath) Then
                message = "Target path is an existing folder."
                Return False
            End If

            If File.Exists(targetPath) Then
                If Not overwrite Then
                    message = "Target file already exists."
                    Return False
                End If

                File.Delete(targetPath)
            End If

            File.Move(sourcePath, targetPath)
            Return True
        End If

        If Directory.Exists(sourcePath) Then
            If IsPathWithinOrEqual(targetPath, sourcePath) Then
                message = "Cannot move a folder into itself."
                Return False
            End If

            Dim targetParent = Path.GetDirectoryName(targetPath)
            If Not String.IsNullOrWhiteSpace(targetParent) AndAlso Not Directory.Exists(targetParent) Then
                Directory.CreateDirectory(targetParent)
            End If

            If File.Exists(targetPath) OrElse Directory.Exists(targetPath) Then
                If Not overwrite Then
                    message = "Target already exists."
                    Return False
                End If

                RemoveWorkspacePath(targetPath)
            End If

            Directory.Move(sourcePath, targetPath)
            Return True
        End If

        message = "Source not found."
        Return False
    End Function

    Private Function ExecuteWorkspaceFindFiles(toolCall As ToolCall) As String
        If Not _chatAgentWorkspace.AllowRead Then Return "Error: Workspace read permission is disabled."

        Dim rel = GetArgString(toolCall.Arguments, "path")
        Dim root = ResolveWorkspacePath(rel, allowRoot:=True)
        If root Is Nothing OrElse Not Directory.Exists(root) Then Return "Error: Folder not found or invalid path."

        Dim query = GetArgString(toolCall.Arguments, "query")
        Dim recursive = GetArgBool(toolCall.Arguments, "recursive", True)
        Dim maxResults = Math.Min(Math.Max(GetArgInt(toolCall.Arguments, "max_results", 50), 1), 500)
        Dim minSizeBytes = GetArgLongValue(toolCall.Arguments, "min_size_bytes", Long.MinValue)
        Dim maxSizeBytes = GetArgLongValue(toolCall.Arguments, "max_size_bytes", Long.MaxValue)

        Dim modifiedFromUtc As DateTime? = Nothing
        Dim modifiedToUtc As DateTime? = Nothing

        If Not TryParseToolDateUtc(GetArgString(toolCall.Arguments, "modified_from"), modifiedFromUtc) Then
            Return "Error: Invalid modified_from value."
        End If

        If Not TryParseToolDateUtc(GetArgString(toolCall.Arguments, "modified_to"), modifiedToUtc) Then
            Return "Error: Invalid modified_to value."
        End If

        Dim extFilters = GetArgStringList(toolCall.Arguments, "extensions").
            Select(Function(e) NormalizeExtensionFilter(e)).
            Where(Function(e) e <> "").
            Distinct(StringComparer.OrdinalIgnoreCase).
            ToList()

        Dim results As New List(Of Dictionary(Of String, Object))()

        For Each filePath In GetWorkspaceVisibleFiles(root, recursive).
            OrderByDescending(Function(p) File.GetLastWriteTimeUtc(p)).
            ThenBy(Function(p) ToWorkspaceRelativePath(p), StringComparer.OrdinalIgnoreCase)

            Dim fi As New FileInfo(filePath)
            Dim relPath = ToWorkspaceRelativePath(filePath)

            If Not WorkspacePathMatchesQuery(relPath, query) Then Continue For
            If extFilters.Count > 0 AndAlso Not extFilters.Contains(fi.Extension, StringComparer.OrdinalIgnoreCase) Then Continue For
            If fi.Length < minSizeBytes OrElse fi.Length > maxSizeBytes Then Continue For
            If modifiedFromUtc.HasValue AndAlso fi.LastWriteTimeUtc < modifiedFromUtc.Value Then Continue For
            If modifiedToUtc.HasValue AndAlso fi.LastWriteTimeUtc > modifiedToUtc.Value Then Continue For

            results.Add(BuildWorkspaceFileRecord(filePath))
            If results.Count >= maxResults Then Exit For
        Next

        Return ToWorkspaceJson(New With {
            .path = If(String.IsNullOrWhiteSpace(rel), "", rel),
            .query = If(query, ""),
            .count = results.Count,
            .results = results
        })
    End Function

    Private Function ExecuteWorkspaceTransfer(toolCall As ToolCall, action As String) As String
        If Not _chatAgentWorkspace.AllowMoveCopyRename Then Return "Error: Copy/move/rename permission is disabled."

        Dim requested As New List(Of String)()
        Dim singleSource = If(GetArgString(toolCall.Arguments, "source_path"), "").Trim()
        If singleSource <> "" Then requested.Add(singleSource)
        requested.AddRange(GetArgStringList(toolCall.Arguments, "source_paths"))
        requested = requested.Distinct(StringComparer.OrdinalIgnoreCase).ToList()

        If requested.Count = 0 Then Return "Error: Provide source_path or source_paths."

        Dim targetDirectoryRel = If(GetArgString(toolCall.Arguments, "target_directory"), "").Trim()
        If targetDirectoryRel = "" Then Return "Error: target_directory is required."

        Dim targetDirectoryPath = ResolveWorkspacePath(targetDirectoryRel)
        If targetDirectoryPath Is Nothing Then Return "Error: Invalid target_directory."

        Dim overwrite = GetArgBool(toolCall.Arguments, "overwrite", False)
        Dim dryRun = GetArgBool(toolCall.Arguments, "dry_run", False)

        Dim results As New List(Of Dictionary(Of String, Object))()

        For Each srcRel In requested
            Dim row As New Dictionary(Of String, Object) From {
                {"source_path", srcRel},
                {"dry_run", dryRun}
            }

            Dim sourcePath = ResolveWorkspacePath(srcRel)
            If sourcePath Is Nothing Then
                row("status") = "error"
                row("message") = "Invalid source path."
                results.Add(row)
                Continue For
            End If

            If Not File.Exists(sourcePath) AndAlso Not Directory.Exists(sourcePath) Then
                row("status") = "error"
                row("message") = "Source not found."
                results.Add(row)
                Continue For
            End If

            Dim leafName = GetPathLeafName(sourcePath)
            Dim targetPath = Path.Combine(targetDirectoryPath, leafName)

            If Not IsPathWithinOrEqual(targetPath, _chatAgentWorkspace.RootPath) Then
                row("status") = "error"
                row("message") = "Resolved target path escapes the workspace."
                results.Add(row)
                Continue For
            End If

            row("target_path") = ToWorkspaceRelativePath(targetPath)
            row("item_type") = If(File.Exists(sourcePath), "file", "folder")

            If dryRun Then
                row("status") = "preview"
                row("message") = $"{action} preview"
                results.Add(row)
                Continue For
            End If

            If Not Directory.Exists(targetDirectoryPath) Then Directory.CreateDirectory(targetDirectoryPath)

            Dim opMessage As String = Nothing
            Dim ok As Boolean

            If action = "copy" Then
                ok = CopyWorkspaceItem(sourcePath, targetPath, overwrite, opMessage)
            Else
                ok = MoveWorkspaceItem(sourcePath, targetPath, overwrite, opMessage)
            End If

            row("status") = If(ok, "ok", "error")
            row("message") = If(ok, $"{action} completed", opMessage)
            results.Add(row)
        Next

        Return ToWorkspaceJson(New With {
            .action = action,
            .target_directory = targetDirectoryRel,
            .overwrite = overwrite,
            .dry_run = dryRun,
            .results = results
        })
    End Function

    Private Function ExecuteWorkspaceMoveTo(toolCall As ToolCall) As String
        Return ExecuteWorkspaceTransfer(toolCall, "move")
    End Function

    Private Function ExecuteWorkspaceCopyTo(toolCall As ToolCall) As String
        Return ExecuteWorkspaceTransfer(toolCall, "copy")
    End Function

    Private Function ExecuteWorkspaceRename(toolCall As ToolCall) As String
        If Not _chatAgentWorkspace.AllowMoveCopyRename Then Return "Error: Copy/move/rename permission is disabled."

        Dim sourceRel = If(GetArgString(toolCall.Arguments, "source_path"), "").Trim()
        Dim newName = If(GetArgString(toolCall.Arguments, "new_name"), "").Trim()
        Dim overwrite = GetArgBool(toolCall.Arguments, "overwrite", False)
        Dim dryRun = GetArgBool(toolCall.Arguments, "dry_run", False)

        If sourceRel = "" Then Return "Error: source_path is required."
        If Not IsLeafNameOnly(newName) Then Return "Error: new_name must be a leaf name only."

        Dim sourcePath = ResolveWorkspacePath(sourceRel)
        If sourcePath Is Nothing Then Return "Error: Invalid source path."
        If Not File.Exists(sourcePath) AndAlso Not Directory.Exists(sourcePath) Then Return "Error: Source not found."

        Dim parentDir = Path.GetDirectoryName(sourcePath)
        Dim targetPath = Path.Combine(parentDir, newName)

        Dim result As New Dictionary(Of String, Object) From {
            {"source_path", sourceRel},
            {"target_path", ToWorkspaceRelativePath(targetPath)},
            {"new_name", newName},
            {"dry_run", dryRun},
            {"item_type", If(File.Exists(sourcePath), "file", "folder")}
        }

        If dryRun Then
            result("status") = "preview"
            result("message") = "rename preview"
            Return ToWorkspaceJson(result)
        End If

        Dim opMessage As String = Nothing
        Dim ok = MoveWorkspaceItem(sourcePath, targetPath, overwrite, opMessage)

        result("status") = If(ok, "ok", "error")
        result("message") = If(ok, "rename completed", opMessage)

        Return ToWorkspaceJson(result)
    End Function

    Private Shared Function BuildBulkRenameTargetName(
        originalFileName As String,
        mode As String,
        findText As String,
        replaceText As String,
        prefix As String,
        suffix As String,
        pattern As String,
        index As Integer) As String

        Dim fileName = If(originalFileName, "")
        Dim baseName = Path.GetFileNameWithoutExtension(fileName)
        Dim ext = Path.GetExtension(fileName)

        Select Case If(mode, "").Trim().ToLowerInvariant()
            Case "prefix"
                Return If(prefix, "") & fileName

            Case "suffix"
                Return baseName & If(suffix, "") & ext

            Case "replace"
                Return ReplaceIgnoreCase(baseName, findText, replaceText) & ext

            Case "pattern"
                Dim template = If(pattern, "").Trim()
                If template = "" Then template = "{name}"

                Dim output = template.
                    Replace("{name}", baseName).
                    Replace("{ext}", ext.TrimStart("."c)).
                    Replace("{index}", index.ToString("000", CultureInfo.InvariantCulture)).
                    Replace("{date}", DateTime.Now.ToString("yyyyMMdd", CultureInfo.InvariantCulture)).
                    Replace("{file_name}", fileName)

                If Path.GetExtension(output) = "" AndAlso ext <> "" Then
                    output &= ext
                End If

                Return output

            Case Else
                Return fileName
        End Select
    End Function

    Private Function ExecuteWorkspaceBulkRename(toolCall As ToolCall) As String
        If Not _chatAgentWorkspace.AllowMoveCopyRename Then Return "Error: Copy/move/rename permission is disabled."

        Dim mode = If(GetArgString(toolCall.Arguments, "mode"), "").Trim().ToLowerInvariant()
        If mode = "" Then Return "Error: mode is required."

        Dim overwrite = GetArgBool(toolCall.Arguments, "overwrite", False)
        Dim dryRun = GetArgBool(toolCall.Arguments, "dry_run", True)
        Dim recursive = GetArgBool(toolCall.Arguments, "recursive", False)
        Dim findText = GetArgString(toolCall.Arguments, "find_text")
        Dim replaceText = GetArgString(toolCall.Arguments, "replace_text")
        Dim prefix = GetArgString(toolCall.Arguments, "prefix")
        Dim suffix = GetArgString(toolCall.Arguments, "suffix")
        Dim pattern = GetArgString(toolCall.Arguments, "pattern")
        Dim rel = If(GetArgString(toolCall.Arguments, "path"), "").Trim()

        Dim candidates As New List(Of String)()

        For Each item In GetArgStringList(toolCall.Arguments, "source_paths")
            Dim fullPath = ResolveWorkspacePath(item)
            If fullPath IsNot Nothing AndAlso File.Exists(fullPath) Then
                candidates.Add(fullPath)
            End If
        Next

        If candidates.Count = 0 AndAlso rel <> "" Then
            Dim fullPath = ResolveWorkspacePath(rel)
            If fullPath Is Nothing Then
                Return "Error: Invalid path."
            End If

            If File.Exists(fullPath) Then
                candidates.Add(fullPath)
            ElseIf Directory.Exists(fullPath) Then
                candidates.AddRange(GetWorkspaceVisibleFiles(fullPath, recursive))
            Else
                Return "Error: File or folder not found."
            End If
        End If

        candidates = candidates.
            Distinct(StringComparer.OrdinalIgnoreCase).
            OrderBy(Function(p) ToWorkspaceRelativePath(p), StringComparer.OrdinalIgnoreCase).
            ToList()

        If candidates.Count = 0 Then Return "Error: No files found to rename."

        Dim results As New List(Of Dictionary(Of String, Object))()
        Dim index = 1

        For Each filePath In candidates
            Dim oldName = Path.GetFileName(filePath)
            Dim newName = BuildBulkRenameTargetName(oldName, mode, findText, replaceText, prefix, suffix, pattern, index)
            Dim row As New Dictionary(Of String, Object) From {
                {"source_path", ToWorkspaceRelativePath(filePath)},
                {"old_name", oldName},
                {"new_name", newName},
                {"dry_run", dryRun}
            }

            If Not IsLeafNameOnly(newName) Then
                row("status") = "error"
                row("message") = "Generated name is invalid."
                results.Add(row)
                index += 1
                Continue For
            End If

            Dim targetPath = Path.Combine(Path.GetDirectoryName(filePath), newName)
            row("target_path") = ToWorkspaceRelativePath(targetPath)

            If filePath.Equals(targetPath, StringComparison.OrdinalIgnoreCase) Then
                row("status") = "skipped"
                row("message") = "Name unchanged."
                results.Add(row)
                index += 1
                Continue For
            End If

            If dryRun Then
                row("status") = "preview"
                row("message") = "bulk rename preview"
                results.Add(row)
                index += 1
                Continue For
            End If

            If File.Exists(targetPath) Then
                If Not overwrite Then
                    row("status") = "error"
                    row("message") = "Target file already exists."
                    results.Add(row)
                    index += 1
                    Continue For
                End If

                File.Delete(targetPath)
            End If

            File.Move(filePath, targetPath)
            row("status") = "ok"
            row("message") = "rename completed"
            results.Add(row)
            index += 1
        Next

        Return ToWorkspaceJson(New With {
            .mode = mode,
            .overwrite = overwrite,
            .dry_run = dryRun,
            .results = results
        })
    End Function

    Private Function ExecuteWorkspaceFileDetails(toolCall As ToolCall) As String
        If Not _chatAgentWorkspace.AllowRead Then Return "Error: Workspace read permission is disabled."

        Dim rel = If(GetArgString(toolCall.Arguments, "path"), "").Trim()
        If rel = "" Then Return "Error: path is required."

        Dim includeHash = GetArgBool(toolCall.Arguments, "include_hash", False)
        Dim fullPath = ResolveWorkspacePath(rel, allowRoot:=True)
        If fullPath Is Nothing Then Return "Error: Invalid path."

        If File.Exists(fullPath) Then
            Dim fi As New FileInfo(fullPath)
            Dim attr = File.GetAttributes(fullPath)

            Return ToWorkspaceJson(New Dictionary(Of String, Object) From {
                {"path", ToWorkspaceRelativePath(fullPath)},
                {"item_type", "file"},
                {"size_bytes", fi.Length},
                {"created", fi.CreationTime.ToString("yyyy-MM-dd HH:mm:ss")},
                {"modified", fi.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")},
                {"accessed", fi.LastAccessTime.ToString("yyyy-MM-dd HH:mm:ss")},
                {"extension", fi.Extension.ToLowerInvariant()},
                {"attributes", attr.ToString()},
                {"read_only", (attr And FileAttributes.ReadOnly) = FileAttributes.ReadOnly},
                {"hidden", (attr And FileAttributes.Hidden) = FileAttributes.Hidden},
                {"hash_sha256", If(includeHash, GetSha256Hex(fullPath), Nothing)}
            })
        End If

        If Directory.Exists(fullPath) Then
            Dim di As New DirectoryInfo(fullPath)
            Dim attr = File.GetAttributes(fullPath)

            Return ToWorkspaceJson(New Dictionary(Of String, Object) From {
                {"path", ToWorkspaceRelativePath(fullPath)},
                {"item_type", "folder"},
                {"size_bytes", GetDirectorySizeSafe(fullPath)},
                {"created", di.CreationTime.ToString("yyyy-MM-dd HH:mm:ss")},
                {"modified", di.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")},
                {"accessed", di.LastAccessTime.ToString("yyyy-MM-dd HH:mm:ss")},
                {"extension", ""},
                {"attributes", attr.ToString()},
                {"read_only", (attr And FileAttributes.ReadOnly) = FileAttributes.ReadOnly},
                {"hidden", (attr And FileAttributes.Hidden) = FileAttributes.Hidden},
                {"hash_sha256", Nothing}
            })
        End If

        Return "Error: File or folder not found."
    End Function

    Private Function ExecuteWorkspaceRecentFiles(toolCall As ToolCall) As String
        If Not _chatAgentWorkspace.AllowRead Then Return "Error: Workspace read permission is disabled."

        Dim rel = GetArgString(toolCall.Arguments, "path")
        Dim root = ResolveWorkspacePath(rel, allowRoot:=True)
        If root Is Nothing OrElse Not Directory.Exists(root) Then Return "Error: Folder not found or invalid path."

        Dim hours = Math.Max(GetArgInt(toolCall.Arguments, "hours", 0), 0)
        Dim days = Math.Max(GetArgInt(toolCall.Arguments, "days", 0), 0)
        If hours = 0 AndAlso days = 0 Then days = 1

        Dim maxResults = Math.Min(Math.Max(GetArgInt(toolCall.Arguments, "max_results", 25), 1), 200)
        Dim cutoffUtc = DateTime.UtcNow.AddHours(-(hours + (days * 24.0R)))

        Dim results As New List(Of Dictionary(Of String, Object))()

        For Each filePath In GetWorkspaceVisibleFiles(root, recursive:=True).
            OrderByDescending(Function(p) File.GetLastWriteTimeUtc(p)).
            ThenBy(Function(p) ToWorkspaceRelativePath(p), StringComparer.OrdinalIgnoreCase)

            Dim fi As New FileInfo(filePath)
            If fi.LastWriteTimeUtc < cutoffUtc Then Continue For

            results.Add(BuildWorkspaceFileRecord(filePath))
            If results.Count >= maxResults Then Exit For
        Next

        Return ToWorkspaceJson(New With {
            .path = If(String.IsNullOrWhiteSpace(rel), "", rel),
            .hours = hours,
            .days = days,
            .count = results.Count,
            .results = results
        })
    End Function

    Private Function ExecuteWorkspaceCreateFolderStructure(toolCall As ToolCall) As String
        If Not _chatAgentWorkspace.AllowWrite Then Return "Error: Workspace write permission is disabled."

        Dim baseRel = GetArgString(toolCall.Arguments, "base_path")
        Dim basePath = ResolveWorkspacePath(baseRel, allowRoot:=True)
        If basePath Is Nothing Then Return "Error: Invalid base_path."

        Dim folders = GetArgStringList(toolCall.Arguments, "folders")
        If folders.Count = 0 Then Return "Error: folders is required."

        Dim dryRun = GetArgBool(toolCall.Arguments, "dry_run", False)
        Dim results As New List(Of Dictionary(Of String, Object))()

        For Each folderRelRaw In folders
            Dim folderRel = If(folderRelRaw, "").Trim()

            Dim row As New Dictionary(Of String, Object) From {
                {"folder", folderRel},
                {"dry_run", dryRun}
            }

            If folderRel = "" OrElse Path.IsPathRooted(folderRel) OrElse
               folderRel.Contains(".." & Path.DirectorySeparatorChar) OrElse
               folderRel.Contains(".." & Path.AltDirectorySeparatorChar) OrElse
               folderRel.Equals("..", StringComparison.Ordinal) Then

                row("status") = "error"
                row("message") = "Invalid folder path."
                results.Add(row)
                Continue For
            End If

            Dim combined = Path.GetFullPath(Path.Combine(basePath, folderRel.Replace("/"c, Path.DirectorySeparatorChar)))
            If Not IsPathWithinOrEqual(combined, _chatAgentWorkspace.RootPath) OrElse ContainsReparsePointBetween(_chatAgentWorkspace.RootPath, combined) Then
                row("status") = "error"
                row("message") = "Folder escapes the workspace or crosses a reparse point."
                results.Add(row)
                Continue For
            End If

            row("target_path") = ToWorkspaceRelativePath(combined)

            If dryRun Then
                row("status") = "preview"
                row("message") = "create folder preview"
                results.Add(row)
                Continue For
            End If

            Directory.CreateDirectory(combined)
            row("status") = "ok"
            row("message") = "folder created"
            results.Add(row)
        Next

        Return ToWorkspaceJson(New With {
            .base_path = If(String.IsNullOrWhiteSpace(baseRel), "", baseRel),
            .dry_run = dryRun,
            .results = results
        })
    End Function

    Private Function ExecuteWorkspaceTrash(toolCall As ToolCall) As String
        If Not _chatAgentWorkspace.AllowDelete Then Return "Error: Delete permission is disabled."

        Dim requested As New List(Of String)()
        Dim singleSource = If(GetArgString(toolCall.Arguments, "source_path"), "").Trim()
        If singleSource <> "" Then requested.Add(singleSource)
        requested.AddRange(GetArgStringList(toolCall.Arguments, "source_paths"))
        requested = requested.Distinct(StringComparer.OrdinalIgnoreCase).ToList()

        If requested.Count = 0 Then Return "Error: Provide source_path or source_paths."

        Dim dryRun = GetArgBool(toolCall.Arguments, "dry_run", False)
        Dim results As New List(Of Dictionary(Of String, Object))()

        For Each srcRel In requested
            Dim row As New Dictionary(Of String, Object) From {
                {"source_path", srcRel},
                {"dry_run", dryRun}
            }

            Dim fullPath = ResolveWorkspacePath(srcRel, allowRoot:=True)
            If fullPath Is Nothing Then
                row("status") = "error"
                row("message") = "Invalid source path."
                results.Add(row)
                Continue For
            End If

            If Not File.Exists(fullPath) AndAlso Not Directory.Exists(fullPath) Then
                row("status") = "error"
                row("message") = "Source not found."
                results.Add(row)
                Continue For
            End If

            row("item_type") = If(File.Exists(fullPath), "file", "folder")

            If dryRun Then
                row("status") = "preview"
                row("message") = "trash preview"
                results.Add(row)
                Continue For
            End If

            Try
                If File.Exists(fullPath) Then
                    VBFileIO.FileSystem.DeleteFile(fullPath, VBFileIO.UIOption.OnlyErrorDialogs, VBFileIO.RecycleOption.SendToRecycleBin)
                Else
                    VBFileIO.FileSystem.DeleteDirectory(fullPath, VBFileIO.UIOption.OnlyErrorDialogs, VBFileIO.RecycleOption.SendToRecycleBin)
                End If

                row("status") = "ok"
                row("message") = "moved to Recycle Bin"
            Catch ex As Exception
                row("status") = "error"
                row("message") = ex.Message
            End Try

            results.Add(row)
        Next

        Return ToWorkspaceJson(New With {
            .action = "trash",
            .dry_run = dryRun,
            .results = results
        })
    End Function

    Private Shared Function BuildInventoryHeaders() As String()
        Return New String() {"Filename", "Folder", "Type", "Size", "Modified Date", "Notes/Category"}
    End Function

    Private Function CreateExcelInventoryReport(reportPath As String, sourceRel As String, items As List(Of Dictionary(Of String, Object))) As String
        Dim app As Excel.Application = Nothing
        Dim wb As Excel.Workbook = Nothing
        Dim ws As Excel.Worksheet = Nothing

        Try
            app = New Excel.Application()
            app.Visible = False
            app.DisplayAlerts = False

            wb = app.Workbooks.Add()
            ws = CType(wb.Worksheets(1), Excel.Worksheet)
            ws.Name = "Inventory"

            Dim headers = BuildInventoryHeaders()
            For i = 0 To headers.Length - 1
                ws.Cells(1, i + 1).Value2 = headers(i)
            Next

            Dim row = 2
            For Each item In items
                ws.Cells(row, 1).Value2 = CStr(item("filename"))
                ws.Cells(row, 2).Value2 = CStr(item("folder"))
                ws.Cells(row, 3).Value2 = CStr(item("type"))
                ws.Cells(row, 4).Value2 = CLng(item("size_bytes"))
                ws.Cells(row, 5).Value2 = CStr(item("modified"))
                ws.Cells(row, 6).Value2 = ""
                row += 1
            Next

            ws.Cells(row + 1, 1).Value2 = "Source"
            ws.Cells(row + 1, 2).Value2 = sourceRel
            ws.Cells(row + 2, 1).Value2 = "Generated"
            ws.Cells(row + 2, 2).Value2 = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)

            ws.Range("A1:F1").Font.Bold = True
            ws.Columns("A:F").AutoFit()

            wb.SaveAs(reportPath)
            wb.Close(SaveChanges:=False)
            app.Quit()

            Return Nothing
        Catch ex As Exception
            Try
                If wb IsNot Nothing Then wb.Close(SaveChanges:=False)
            Catch
            End Try

            Try
                If app IsNot Nothing Then app.Quit()
            Catch
            End Try

            Return ex.Message
        Finally
            If ws IsNot Nothing Then
                Try : Marshal.FinalReleaseComObject(ws) : Catch : End Try
            End If
            If wb IsNot Nothing Then
                Try : Marshal.FinalReleaseComObject(wb) : Catch : End Try
            End If
            If app IsNot Nothing Then
                Try : Marshal.FinalReleaseComObject(app) : Catch : End Try
            End If
        End Try
    End Function

    Private Function CreateWordInventoryReport(reportPath As String, sourceRel As String, items As List(Of Dictionary(Of String, Object))) As String
        Dim app As Word.Application = Nothing
        Dim doc As Word.Document = Nothing
        Dim table As Word.Table = Nothing

        Try
            app = New Word.Application()
            app.Visible = False

            doc = app.Documents.Add()

            Dim intro = doc.Range(0, 0)
            intro.Text =
                "Workspace Inventory Report" & vbCrLf &
                "Source: " & sourceRel & vbCrLf &
                "Generated: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) & vbCrLf & vbCrLf

            Dim rowCount = Math.Max(items.Count, 1) + 1
            Dim startRange = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
            table = doc.Tables.Add(startRange, rowCount, 6)
            table.Borders.Enable = 1

            Dim headers = BuildInventoryHeaders()
            For i = 0 To headers.Length - 1
                table.Cell(1, i + 1).Range.Text = headers(i)
            Next

            If items.Count = 0 Then
                table.Cell(2, 1).Range.Text = "(no files found)"
            Else
                Dim row = 2
                For Each item In items
                    table.Cell(row, 1).Range.Text = CStr(item("filename"))
                    table.Cell(row, 2).Range.Text = CStr(item("folder"))
                    table.Cell(row, 3).Range.Text = CStr(item("type"))
                    table.Cell(row, 4).Range.Text = CStr(item("size_bytes"))
                    table.Cell(row, 5).Range.Text = CStr(item("modified"))
                    table.Cell(row, 6).Range.Text = ""
                    row += 1
                Next
            End If

            doc.SaveAs2(reportPath)
            doc.Close(SaveChanges:=False)
            app.Quit()

            Return Nothing
        Catch ex As Exception
            Try
                If doc IsNot Nothing Then doc.Close(SaveChanges:=False)
            Catch
            End Try

            Try
                If app IsNot Nothing Then app.Quit()
            Catch
            End Try

            Return ex.Message
        Finally
            If table IsNot Nothing Then
                Try : Marshal.FinalReleaseComObject(table) : Catch : End Try
            End If
            If doc IsNot Nothing Then
                Try : Marshal.FinalReleaseComObject(doc) : Catch : End Try
            End If
            If app IsNot Nothing Then
                Try : Marshal.FinalReleaseComObject(app) : Catch : End Try
            End If
        End Try
    End Function

    Private Function ExecuteWorkspaceInventoryReport(toolCall As ToolCall) As String
        If Not _chatAgentWorkspace.AllowRead Then Return "Error: Workspace read permission is disabled."
        If Not _chatAgentWorkspace.AllowWrite Then Return "Error: Workspace write permission is disabled."

        Dim rel = GetArgString(toolCall.Arguments, "path")
        Dim root = ResolveWorkspacePath(rel, allowRoot:=True)
        If root Is Nothing OrElse Not Directory.Exists(root) Then Return "Error: Folder not found or invalid path."

        Dim recursive = GetArgBool(toolCall.Arguments, "recursive", False)
        Dim maxItems = Math.Min(Math.Max(GetArgInt(toolCall.Arguments, "max_items", 200), 1), 2000)
        Dim outputFormat = If(GetArgString(toolCall.Arguments, "output_format"), "").Trim().ToLowerInvariant()
        Dim outputFileName = If(GetArgString(toolCall.Arguments, "output_filename"), "").Trim()

        If outputFormat <> "excel" AndAlso outputFormat <> "word" Then
            Return "Error: output_format must be 'excel' or 'word'."
        End If

        If outputFileName = "" Then
            outputFileName = "workspace_inventory_" & DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture)
        End If

        If Not IsLeafNameOnly(outputFileName) OrElse Path.GetFileName(outputFileName) <> outputFileName Then
            Return "Error: output_filename must be a leaf file name only."
        End If

        Dim requiredExt = If(outputFormat = "excel", ".xlsx", ".docx")
        If Path.GetExtension(outputFileName) = "" Then
            outputFileName &= requiredExt
        End If

        Dim items As New List(Of Dictionary(Of String, Object))()

        For Each filePath In GetWorkspaceVisibleFiles(root, recursive).
            OrderBy(Function(p) ToWorkspaceRelativePath(p), StringComparer.OrdinalIgnoreCase)

            items.Add(BuildWorkspaceFileRecord(filePath))
            If items.Count >= maxItems Then Exit For
        Next

        Dim reportPath = GetUniqueFilePath(Path.Combine(root, outputFileName))
        Dim errorMessage As String

        If outputFormat = "excel" Then
            errorMessage = CreateExcelInventoryReport(reportPath, If(String.IsNullOrWhiteSpace(rel), "", rel), items)
        Else
            errorMessage = CreateWordInventoryReport(reportPath, If(String.IsNullOrWhiteSpace(rel), "", rel), items)
        End If

        If Not String.IsNullOrWhiteSpace(errorMessage) Then
            Return "Error: Failed to create inventory report: " & errorMessage
        End If

        Dim preview As New List(Of Dictionary(Of String, Object))()
        For i = 0 To Math.Min(items.Count, 3) - 1
            preview.Add(items(i))
        Next

        Return ToWorkspaceJson(New With {
            .path = If(String.IsNullOrWhiteSpace(rel), "", rel),
            .recursive = recursive,
            .max_items = maxItems,
            .output_format = outputFormat,
            .report_path = ToWorkspaceRelativePath(reportPath),
            .item_count = items.Count,
            .preview = preview
        })
    End Function

    Private Function HasChatAgentSessionWorkspace() As Boolean
        Return _chatAgentActive AndAlso
               Not String.IsNullOrWhiteSpace(_chatAgentTempDir) AndAlso
               Directory.Exists(_chatAgentTempDir)
    End Function

    ''' <summary>
    ''' Pushes the current chat-agent workspace root into the SharedLibrary PathPolicy.
    ''' Called whenever workspace state is mutated.
    ''' </summary>
    ''' <summary>
    ''' Pushes the current chat-agent workspace root into the SharedLibrary path/workspace state.
    ''' Called whenever workspace state is mutated.
    ''' </summary>
    Private Sub SyncWorkspaceToPathPolicy(Optional includeSessionTempFallback As Boolean = False)
        Try
            Dim ws = _chatAgentWorkspace

            Dim effectiveRoot As String = Nothing
            Dim persistUntilRevoked As Boolean = False
            Dim allowRead As Boolean = False
            Dim allowWrite As Boolean = False
            Dim allowMoveCopyRename As Boolean = False
            Dim allowDelete As Boolean = False
            Dim includeHiddenSystem As Boolean = False

            If ws IsNot Nothing AndAlso
               Not String.IsNullOrWhiteSpace(ws.RootPath) AndAlso
               Directory.Exists(ws.RootPath) Then

                effectiveRoot = Path.GetFullPath(ws.RootPath)
                persistUntilRevoked = ws.PersistUntilRevoked
                allowRead = ws.AllowRead
                allowWrite = ws.AllowWrite
                allowMoveCopyRename = ws.AllowMoveCopyRename
                allowDelete = ws.AllowDelete
                includeHiddenSystem = ws.IncludeHiddenSystem

            ElseIf includeSessionTempFallback AndAlso HasChatAgentSessionWorkspace() Then
                effectiveRoot = Path.GetFullPath(_chatAgentTempDir)
                persistUntilRevoked = False
                allowRead = True
                allowWrite = True
                allowMoveCopyRename = True
                allowDelete = True
                includeHiddenSystem = True
            End If

            If Not String.IsNullOrWhiteSpace(effectiveRoot) AndAlso Directory.Exists(effectiveRoot) Then
                If allowRead Then
                    SharedLibrary.Agents.PathPolicy.SetWorkspaceRoot(effectiveRoot)
                Else
                    SharedLibrary.Agents.PathPolicy.SetWorkspaceRoot(Nothing)
                End If

                SharedLibrary.Agents.WorkspaceTools.SetActive(
                    New SharedLibrary.Agents.WorkspaceState() With {
                        .RootPath = effectiveRoot,
                        .PersistUntilRevoked = persistUntilRevoked,
                        .AllowRead = allowRead,
                        .AllowWrite = allowWrite,
                        .AllowMoveCopyRename = allowMoveCopyRename,
                        .AllowDelete = allowDelete,
                        .IncludeHiddenSystem = includeHiddenSystem
                    })
            Else
                SharedLibrary.Agents.PathPolicy.SetWorkspaceRoot(Nothing)
                SharedLibrary.Agents.WorkspaceTools.SetActive(New SharedLibrary.Agents.WorkspaceState())
            End If
        Catch
            SharedLibrary.Agents.PathPolicy.SetWorkspaceRoot(Nothing)
            SharedLibrary.Agents.WorkspaceTools.SetActive(New SharedLibrary.Agents.WorkspaceState())
        End Try
    End Sub

    ''' <summary>
    ''' Removes a single chat agent file from the current session without clearing
    ''' the entire temporary workspace. Only the session copy is deleted.
    ''' </summary>
    Private Function ChatAgentRemoveFile(fileName As String) As Boolean
        If String.IsNullOrWhiteSpace(fileName) Then Return False

        Dim attToRemove = _chatAgentFiles.FirstOrDefault(
            Function(att)
                If att Is Nothing Then Return False

                If Not String.IsNullOrWhiteSpace(att.OriginalFileName) AndAlso
                   att.OriginalFileName.Equals(fileName, StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If

                If Not String.IsNullOrWhiteSpace(att.TempFilePath) AndAlso
                   Path.GetFileName(att.TempFilePath).Equals(fileName, StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If

                Return False
            End Function)

        If attToRemove Is Nothing Then Return False

        Dim tempPath As String = attToRemove.TempFilePath

        Try
            _chatAgentFiles.Remove(attToRemove)
        Catch
        End Try

        If _apCurrentAttachments IsNot Nothing Then
            Try
                _apCurrentAttachments.Remove(attToRemove)
            Catch
            End Try
        End If

        Try
            If Not String.IsNullOrWhiteSpace(tempPath) Then
                Dim fullTempPath = Path.GetFullPath(tempPath)

                If _chatAgentSurfacedFiles IsNot Nothing Then
                    _chatAgentSurfacedFiles.Remove(fullTempPath)
                End If

                If File.Exists(fullTempPath) Then
                    File.Delete(fullTempPath)
                End If
            End If
        Catch
        End Try

        Try
            If Not String.IsNullOrWhiteSpace(_chatAgentTempDir) AndAlso
               Directory.Exists(_chatAgentTempDir) AndAlso
               Directory.GetFileSystemEntries(_chatAgentTempDir).Length = 0 Then
                Directory.Delete(_chatAgentTempDir, recursive:=False)
                _chatAgentTempDir = Nothing
            End If
        Catch
        End Try

        Return True
    End Function

End Class
