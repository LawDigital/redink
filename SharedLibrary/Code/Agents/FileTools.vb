' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: FileTools.vb
' Purpose: Built-in shared tools for binary-safe file/directory operations across
'          the PathPolicy-governed roots (workspace, staging/temp, Desktop
'          fallback, and skill scripts/references). Complements TextTools, which
'          is UTF-8 text only.
'
'            file_copy       — copy a file or directory.
'            file_move       — move a file or directory.
'            file_rename     — rename a file or directory in place.
'            file_delete     — delete a file or directory (Recycle Bin by default).
'            file_make_dir   — create a directory.
'            file_remove_dir — remove a directory (Recycle Bin by default).
'
' Architecture:
'  - Every path is resolved through PathPolicy.Resolve(...); the policy enforces
'    workspace Read/Write permissions and skill-author write gating.
'  - Destructive verbs additionally honor the user-configured workspace
'    MoveCopyRename/Delete flags for targets under the workspace root.
'  - Binary-safe: uses File.Copy / File.Move (no encoding assumptions).
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports Microsoft.VisualBasic.FileIO
Imports Newtonsoft.Json
Imports SharedLibrary.SharedLibrary

Namespace Agents

    Public NotInheritable Class FileTools

        Private Sub New()
        End Sub

        Public Const ToolCopy As String = "file_copy"
        Public Const ToolMove As String = "file_move"
        Public Const ToolRename As String = "file_rename"
        Public Const ToolDelete As String = "file_delete"
        Public Const ToolMakeDir As String = "file_make_dir"
        Public Const ToolRemoveDir As String = "file_remove_dir"

        Public Shared Function IsFileTool(name As String) As Boolean
            If String.IsNullOrWhiteSpace(name) Then Return False
            Select Case name
                Case ToolCopy, ToolMove, ToolRename, ToolDelete, ToolMakeDir, ToolRemoveDir
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Public Shared Function BuildAll() As List(Of ModelConfig)
            Return New List(Of ModelConfig) From {
                BuildCopy(), BuildMove(), BuildRename(), BuildDelete(), BuildMakeDir(), BuildRemoveDir()
            }
        End Function

        ' --------------------------------------------------------------- dispatch

        Public Shared Function Execute(toolName As String, arguments As IDictionary(Of String, Object)) As String
            Try
                Select Case toolName
                    Case ToolCopy : Return ExecuteCopy(arguments)
                    Case ToolMove : Return ExecuteMove(arguments)
                    Case ToolRename : Return ExecuteRename(arguments)
                    Case ToolDelete : Return ExecuteDelete(arguments)
                    Case ToolMakeDir : Return ExecuteMakeDir(arguments)
                    Case ToolRemoveDir : Return ExecuteRemoveDir(arguments)
                    Case Else : Return Err_("unknown_file_tool", "Unknown tool: " & toolName)
                End Select
            Catch uae As UnauthorizedAccessException
                Return Err_("access_denied", uae.Message)
            Catch ex As Exception
                Return Err_("file_tool_failed", ex.Message)
            End Try
        End Function

        ' --------------------------------------------------------------- operations

        Private Shared Function ExecuteCopy(args As IDictionary(Of String, Object)) As String
            Dim src = PathPolicy.Resolve(GetStr(args, "source"), PathAccess.Read)
            Dim dst = PathPolicy.Resolve(GetStr(args, "destination"), PathAccess.Write)
            RequireWorkspaceMoveCopyRename(dst)
            Dim overwrite = GetBool(args, "overwrite", False)

            If Not File.Exists(src) AndAlso Not Directory.Exists(src) Then
                Return Err_("not_found", "Source not found.")
            End If

            Dim parent = Path.GetDirectoryName(dst)
            If Not String.IsNullOrWhiteSpace(parent) AndAlso Not Directory.Exists(parent) Then
                Directory.CreateDirectory(parent)
            End If

            If File.Exists(src) Then
                File.Copy(src, dst, overwrite)
            Else
                FileSystem.CopyDirectory(src, dst, overwrite)
            End If

            Return JsonConvert.SerializeObject(New With {Key .source = src, Key .destination = dst, Key .overwrite = overwrite})
        End Function

        Private Shared Function ExecuteMove(args As IDictionary(Of String, Object)) As String
            Dim src = PathPolicy.Resolve(GetStr(args, "source"), PathAccess.Write)
            Dim dst = PathPolicy.Resolve(GetStr(args, "destination"), PathAccess.Write)
            RequireWorkspaceMoveCopyRename(src)
            RequireWorkspaceMoveCopyRename(dst)

            If Not File.Exists(src) AndAlso Not Directory.Exists(src) Then
                Return Err_("not_found", "Source not found.")
            End If

            Dim parent = Path.GetDirectoryName(dst)
            If Not String.IsNullOrWhiteSpace(parent) AndAlso Not Directory.Exists(parent) Then
                Directory.CreateDirectory(parent)
            End If

            If File.Exists(src) Then
                File.Move(src, dst)
            Else
                Directory.Move(src, dst)
            End If

            Return JsonConvert.SerializeObject(New With {Key .source = src, Key .destination = dst, Key .moved = True})
        End Function

        Private Shared Function ExecuteRename(args As IDictionary(Of String, Object)) As String
            Dim src = PathPolicy.Resolve(GetStr(args, "path"), PathAccess.Write)
            RequireWorkspaceMoveCopyRename(src)

            Dim newName = SanitizeName(GetStr(args, "new_name"))
            If String.IsNullOrWhiteSpace(newName) Then Return Err_("missing_name", "new_name is required.")

            Dim parent = Path.GetDirectoryName(src)
            Dim dst = PathPolicy.Resolve(Path.Combine(parent, newName), PathAccess.Write)
            RequireWorkspaceMoveCopyRename(dst)

            If File.Exists(src) Then
                File.Move(src, dst)
            ElseIf Directory.Exists(src) Then
                Directory.Move(src, dst)
            Else
                Return Err_("not_found", "Path not found.")
            End If

            Return JsonConvert.SerializeObject(New With {Key .path = dst})
        End Function

        Private Shared Function ExecuteDelete(args As IDictionary(Of String, Object)) As String
            Dim p = PathPolicy.Resolve(GetStr(args, "path"), PathAccess.Write)
            RequireWorkspaceDelete(p)
            Dim toTrash = GetBool(args, "to_trash", True)

            If File.Exists(p) Then
                FileSystem.DeleteFile(p, UIOption.OnlyErrorDialogs,
                    If(toTrash, RecycleOption.SendToRecycleBin, RecycleOption.DeletePermanently),
                    UICancelOption.ThrowException)
            ElseIf Directory.Exists(p) Then
                Return Err_("is_directory", "Path is a directory. Use " & ToolRemoveDir & ".")
            Else
                Return Err_("not_found", "Path not found.")
            End If

            Return JsonConvert.SerializeObject(New With {Key .path = p, Key .to_trash = toTrash})
        End Function

        Private Shared Function ExecuteMakeDir(args As IDictionary(Of String, Object)) As String
            Dim p = PathPolicy.Resolve(GetStr(args, "path"), PathAccess.Write)
            Directory.CreateDirectory(p)
            Return JsonConvert.SerializeObject(New With {Key .path = p, Key .created = True})
        End Function

        Private Shared Function ExecuteRemoveDir(args As IDictionary(Of String, Object)) As String
            Dim p = PathPolicy.Resolve(GetStr(args, "path"), PathAccess.Write)
            RequireWorkspaceDelete(p)
            Dim toTrash = GetBool(args, "to_trash", True)

            If Not Directory.Exists(p) Then Return Err_("not_found", "Directory not found.")

            FileSystem.DeleteDirectory(p, UIOption.OnlyErrorDialogs,
                If(toTrash, RecycleOption.SendToRecycleBin, RecycleOption.DeletePermanently),
                UICancelOption.ThrowException)

            Return JsonConvert.SerializeObject(New With {Key .path = p, Key .to_trash = toTrash, Key .removed = True})
        End Function

        ' --------------------------------------------------------------- permission helpers

        ''' <summary>
        ''' Enforces the user-configured workspace MoveCopyRename permission for a target
        ''' that resolves under the workspace root. Non-workspace roots (staging, skills,
        ''' Desktop) are governed by their own gates and are not restricted here.
        ''' </summary>
        Private Shared Sub RequireWorkspaceMoveCopyRename(fullPath As String)
            If IsUnderWorkspace(fullPath) AndAlso Not PathPolicy.WorkspaceAllowMoveCopyRename Then
                Throw New UnauthorizedAccessException("Workspace move/copy/rename is disabled.")
            End If
        End Sub

        Private Shared Sub RequireWorkspaceDelete(fullPath As String)
            If IsUnderWorkspace(fullPath) AndAlso Not PathPolicy.WorkspaceAllowDelete Then
                Throw New UnauthorizedAccessException("Workspace delete is disabled.")
            End If
        End Sub

        Private Shared Function IsUnderWorkspace(fullPath As String) As Boolean
            Dim ws = PathPolicy.WorkspaceRoot
            If String.IsNullOrWhiteSpace(ws) OrElse String.IsNullOrWhiteSpace(fullPath) Then Return False
            Dim root = Path.GetFullPath(ws).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            Dim full = Path.GetFullPath(fullPath).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            If String.Equals(full, root, StringComparison.OrdinalIgnoreCase) Then Return True
            Return full.StartsWith(root & Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)
        End Function

        ' --------------------------------------------------------------- shared helpers

        Private Shared Function SanitizeName(name As String) As String
            If String.IsNullOrWhiteSpace(name) Then Return ""
            Dim invalid = Path.GetInvalidFileNameChars()
            Dim sb As New System.Text.StringBuilder(name.Length)
            For Each c In name
                If Array.IndexOf(invalid, c) >= 0 Then sb.Append("_"c) Else sb.Append(c)
            Next
            Return sb.ToString()
        End Function

        Private Shared Function Err_(code As String, message As String) As String
            Return JsonConvert.SerializeObject(New With {Key .error = code, Key .message = message})
        End Function

        Private Shared Function GetStr(args As IDictionary(Of String, Object), name As String) As String
            If args Is Nothing Then Return ""
            Dim v As Object = Nothing
            If Not args.TryGetValue(name, v) OrElse v Is Nothing Then Return ""
            Return System.Convert.ToString(v)
        End Function

        Private Shared Function GetBool(args As IDictionary(Of String, Object), name As String, defaultValue As Boolean) As Boolean
            If args Is Nothing Then Return defaultValue
            Dim v As Object = Nothing
            If Not args.TryGetValue(name, v) OrElse v Is Nothing Then Return defaultValue
            Try
                Return System.Convert.ToBoolean(v)
            Catch
                Select Case System.Convert.ToString(v).Trim().ToLowerInvariant()
                    Case "true", "1", "yes" : Return True
                    Case "false", "0", "no" : Return False
                    Case Else : Return defaultValue
                End Select
            End Try
        End Function

        ' --------------------------------------------------------------- factories

        Private Const RootsDescription As String =
            "Operates across the allowed roots: the agent workspace, the session staging/temp area, " &
            "and skill scripts/references. Reading skill files is always allowed; writing into a skill's " &
            "folder requires skill-author mode. Workspace operations honor the user's configured " &
            "read/write/move/delete permissions. Binary files are fully supported."

        Private Shared Function BuildCopy() As ModelConfig
            Dim def =
                "{""name"":""" & ToolCopy & """," &
                """description"":""Copy a file or directory (binary-safe). " & RootsDescription &
                " Typical use: copy a template from a skill's references into the workspace or staging area, or copy a produced file into a skill's references (author mode)."",""parameters"":{""type"":""object""," &
                """properties"":{" &
                """source"":{""type"":""string"",""description"":""Source file or directory path.""}," &
                """destination"":{""type"":""string"",""description"":""Destination path.""}," &
                """overwrite"":{""type"":""boolean"",""description"":""Overwrite existing target. Default false.""}}," &
                """required"":[""source"",""destination""]}}"
            Return New ModelConfig() With {
                .ToolName = ToolCopy, .ToolDefinition = def, .Tool = True, .ToolPriority = 914, .ToolErrorHandling = "skip",
                .ModelDescription = "File copy (binary-safe)",
                .ToolInstructionsPrompt = ToolCopy & ": Copy a file or directory between the workspace, staging area, and skill references/scripts. Supports binary files. Writing into a skill folder requires author mode; workspace writes require workspace write permission."
            }
        End Function

        Private Shared Function BuildMove() As ModelConfig
            Dim def =
                "{""name"":""" & ToolMove & """," &
                """description"":""Move a file or directory (binary-safe). " & RootsDescription & """,""parameters"":{""type"":""object""," &
                """properties"":{" &
                """source"":{""type"":""string"",""description"":""Source file or directory path.""}," &
                """destination"":{""type"":""string"",""description"":""Destination path.""}}," &
                """required"":[""source"",""destination""]}}"
            Return New ModelConfig() With {
                .ToolName = ToolMove, .ToolDefinition = def, .Tool = True, .ToolPriority = 914, .ToolErrorHandling = "skip",
                .ModelDescription = "File move (binary-safe)",
                .ToolInstructionsPrompt = ToolMove & ": Move a file or directory between the allowed roots. Workspace moves require workspace move/copy/rename permission."
            }
        End Function

        Private Shared Function BuildRename() As ModelConfig
            Dim def =
                "{""name"":""" & ToolRename & """," &
                """description"":""Rename a file or directory in place. " & RootsDescription & """,""parameters"":{""type"":""object""," &
                """properties"":{" &
                """path"":{""type"":""string"",""description"":""Existing file or directory path.""}," &
                """new_name"":{""type"":""string"",""description"":""New leaf name (no directory separators).""}}," &
                """required"":[""path"",""new_name""]}}"
            Return New ModelConfig() With {
                .ToolName = ToolRename, .ToolDefinition = def, .Tool = True, .ToolPriority = 914, .ToolErrorHandling = "skip",
                .ModelDescription = "File rename",
                .ToolInstructionsPrompt = ToolRename & ": Rename a file or directory in place. Workspace renames require workspace move/copy/rename permission."
            }
        End Function

        Private Shared Function BuildDelete() As ModelConfig
            Dim def =
                "{""name"":""" & ToolDelete & """," &
                """description"":""Delete a single file (to the Recycle Bin by default). " & RootsDescription & """,""parameters"":{""type"":""object""," &
                """properties"":{" &
                """path"":{""type"":""string"",""description"":""File path to delete.""}," &
                """to_trash"":{""type"":""boolean"",""description"":""Send to Recycle Bin (true) or delete permanently (false). Default true.""}}," &
                """required"":[""path""]}}"
            Return New ModelConfig() With {
                .ToolName = ToolDelete, .ToolDefinition = def, .Tool = True, .ToolPriority = 914, .ToolErrorHandling = "skip",
                .ModelDescription = "File delete",
                .ToolInstructionsPrompt = ToolDelete & ": Delete a single file. Workspace deletes require workspace delete permission."
            }
        End Function

        Private Shared Function BuildMakeDir() As ModelConfig
            Dim def =
                "{""name"":""" & ToolMakeDir & """," &
                """description"":""Create a directory (including intermediate directories). " & RootsDescription & """,""parameters"":{""type"":""object""," &
                """properties"":{" &
                """path"":{""type"":""string"",""description"":""Directory path to create.""}}," &
                """required"":[""path""]}}"
            Return New ModelConfig() With {
                .ToolName = ToolMakeDir, .ToolDefinition = def, .Tool = True, .ToolPriority = 914, .ToolErrorHandling = "skip",
                .ModelDescription = "Create directory",
                .ToolInstructionsPrompt = ToolMakeDir & ": Create a directory. Workspace creation requires workspace write permission; skill folders require author mode."
            }
        End Function

        Private Shared Function BuildRemoveDir() As ModelConfig
            Dim def =
                "{""name"":""" & ToolRemoveDir & """," &
                """description"":""Remove a directory and its contents (to the Recycle Bin by default). " & RootsDescription & """,""parameters"":{""type"":""object""," &
                """properties"":{" &
                """path"":{""type"":""string"",""description"":""Directory path to remove.""}," &
                """to_trash"":{""type"":""boolean"",""description"":""Send to Recycle Bin (true) or delete permanently (false). Default true.""}}," &
                """required"":[""path""]}}"
            Return New ModelConfig() With {
                .ToolName = ToolRemoveDir, .ToolDefinition = def, .Tool = True, .ToolPriority = 914, .ToolErrorHandling = "skip",
                .ModelDescription = "Remove directory",
                .ToolInstructionsPrompt = ToolRemoveDir & ": Remove a directory and its contents. Workspace removal requires workspace delete permission."
            }
        End Function

    End Class

End Namespace
