' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: WorkspaceTools.vb
' Purpose: Workspace-scoped file tools for the agent layer. Mirrors the most
'          useful subset of the Outlook chat-agent workspace tools but is
'          host-agnostic. Permissions are taken from the active WorkspaceState.
'
' Tools (all prefixed "workspace_"):
'  - workspace_get: return current workspace info and permissions.
'  - workspace_inventory: list files in workspace (glob + recursive).
'  - workspace_read / workspace_read_many: read UTF-8 text files in workspace.
'  - workspace_write: write/append/create_new UTF-8 text file in workspace.
'  - workspace_search: search content across files (substring or regex).
'  - workspace_copy / workspace_move / workspace_rename / workspace_delete:
'    file/folder operations (delete to Recycle Bin or permanent on request).
'  - workspace_make_dir: create a folder in workspace.
'  - workspace_extract_text / workspace_extract_text_many: extract readable text
'    from supported file formats (PDF, Word, Excel, etc.).
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic.FileIO
Imports Newtonsoft.Json
Imports SharedLibrary.SharedLibrary
Imports IO = System.IO
Imports Newtonsoft.Json.Linq

Namespace Agents

    Public NotInheritable Class WorkspaceTools

        Private Shared ReadOnly BinaryWorkspaceExtensions As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
            ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx",
            ".zip", ".rar", ".7z",
            ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".webp",
            ".mp3", ".wav", ".m4a", ".mp4", ".mov", ".avi", ".mkv",
            ".exe", ".dll", ".bin"
        }

        Private Shared Function IsBinaryWorkspaceExtension(path As String) As Boolean
            Dim ext As String = If(System.IO.Path.GetExtension(If(path, "")), "").Trim()
            If ext = "" Then Return False
            Return BinaryWorkspaceExtensions.Contains(ext)
        End Function

        Private Sub New()
        End Sub

        Private Shared _active As WorkspaceState = New WorkspaceState()

        ''' <summary>Sets the active workspace state. Hosts call this on activation/refresh.</summary>
        Public Shared Sub SetActive(state As WorkspaceState)
            _active = If(state, New WorkspaceState())
            ' Mirror into PathPolicy so generic text/skill tools also honor the workspace.
            If Not String.IsNullOrWhiteSpace(_active.RootPath) AndAlso Directory.Exists(_active.RootPath) AndAlso _active.AllowRead Then
                PathPolicy.SetWorkspaceRoot(_active.RootPath)
            Else
                PathPolicy.SetWorkspaceRoot(Nothing)
            End If
            ' Mirror the user-configured permissions so every PathPolicy-based tool
            ' (text_*, file_*) honors the same Read/Write/Move/Delete gates as workspace_*.
            PathPolicy.SetWorkspacePermissions(
                _active.AllowRead, _active.AllowWrite,
                _active.AllowMoveCopyRename, _active.AllowDelete)
        End Sub

        Public Shared ReadOnly Property Active As WorkspaceState
            Get
                Return _active
            End Get
        End Property

        Public Const ToolGet As String = "workspace_get"
        Public Const ToolInventory As String = "workspace_inventory"
        Public Const ToolRead As String = "workspace_read"
        Public Const ToolReadMany As String = "workspace_read_many"
        Public Const ToolWrite As String = "workspace_write"
        Public Const ToolSearch As String = "workspace_search"
        Public Const ToolCopy As String = "workspace_copy"
        Public Const ToolMove As String = "workspace_move"
        Public Const ToolRename As String = "workspace_rename"
        Public Const ToolDelete As String = "workspace_delete"
        Public Const ToolMakeDir As String = "workspace_make_dir"
        Public Const ToolExtractText As String = "workspace_extract_text"
        Public Const ToolExtractTextMany As String = "workspace_extract_text_many"

        Private Const MetaAliasesUsedKey As String = "__meta_aliases_used"
        Private Const MetaContentProvidedKey As String = "__meta_content_provided"
        Private Const MetaContentWasEmptyKey As String = "__meta_content_was_empty"

        Private NotInheritable Class ToolArgumentNormalizationResult
            Public Property Arguments As Dictionary(Of String, Object)
            Public Property ValidationErrorJson As String
        End Class

        Public Shared Function IsWorkspaceTool(name As String) As Boolean
            If String.IsNullOrWhiteSpace(name) Then Return False
            Select Case name
                Case ToolGet, ToolInventory, ToolRead, ToolReadMany, ToolWrite, ToolSearch,
                     ToolCopy, ToolMove, ToolRename, ToolDelete, ToolMakeDir,
                     ToolExtractText, ToolExtractTextMany
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Public Shared Function BuildAll() As List(Of ModelConfig)
            Return New List(Of ModelConfig) From {
                BuildGet(), BuildInventory(), BuildRead(), BuildReadMany(), BuildWrite(), BuildSearch(),
                BuildCopy(), BuildMove(), BuildRename(), BuildDelete(), BuildMakeDir(),
                BuildExtractText(), BuildExtractTextMany()
            }
        End Function


        ' --------------------------------------------------------------- dispatch
        Public Shared Function Execute(toolName As String, arguments As IDictionary(Of String, Object)) As String
            Try
                Dim normalized = NormalizeArguments(toolName, arguments)
                If normalized IsNot Nothing Then
                    If Not String.IsNullOrWhiteSpace(normalized.ValidationErrorJson) Then
                        Return normalized.ValidationErrorJson
                    End If

                    arguments = normalized.Arguments
                End If

                Select Case toolName
                    Case ToolGet : Return ExecuteGet()
                    Case ToolInventory : Return ExecuteInventory(arguments)
                    Case ToolRead : Return ExecuteRead(arguments)
                    Case ToolReadMany : Return ExecuteReadMany(arguments)
                    Case ToolWrite : Return ExecuteWrite(arguments)
                    Case ToolSearch : Return ExecuteSearch(arguments)
                    Case ToolCopy : Return ExecuteCopyOrMove(arguments, isMove:=False)
                    Case ToolMove : Return ExecuteCopyOrMove(arguments, isMove:=True)
                    Case ToolRename : Return ExecuteRename(arguments)
                    Case ToolDelete : Return ExecuteDelete(arguments)
                    Case ToolMakeDir : Return ExecuteMakeDir(arguments)
                    Case ToolExtractText, ToolExtractTextMany
                        Return Err_("host_only_tool", toolName & " must be executed by the host (Word/Outlook), not via the shared dispatcher.")
                    Case Else : Return Err_("unknown_workspace_tool", "Unknown tool '" & toolName & "'.")
                End Select
            Catch uae As UnauthorizedAccessException
                Return Err_("access_denied", uae.Message)
            Catch ex As Exception
                Return Err_("workspace_tool_failed", ex.Message)
            End Try
        End Function



        Private Shared Function NormalizeArguments(toolName As String,
                                           arguments As IDictionary(Of String, Object)) As ToolArgumentNormalizationResult
            Dim normalized = ToCaseInsensitiveDictionary(arguments)

            Select Case If(toolName, "").Trim().ToLowerInvariant()
                Case ToolWrite
                    Return NormalizeWorkspaceWriteArguments(normalized)

                Case Else
                    Return New ToolArgumentNormalizationResult() With {
                .Arguments = normalized
            }
            End Select
        End Function

        Private Shared Function NormalizeWorkspaceWriteArguments(arguments As Dictionary(Of String, Object)) As ToolArgumentNormalizationResult
            Dim aliasesUsed As New List(Of String)()

            Dim textValue As Object = Nothing
            Dim contentValue As Object = Nothing
            Dim hasText = TryGetArgumentValue(arguments, "text", textValue)
            Dim hasContent = TryGetArgumentValue(arguments, "content", contentValue)

            If hasText AndAlso hasContent Then
                Dim canonicalText = If(textValue, "").ToString()
                Dim aliasText = If(contentValue, "").ToString()

                If Not String.Equals(canonicalText, aliasText, StringComparison.Ordinal) Then
                    Return New ToolArgumentNormalizationResult() With {
                .Arguments = arguments,
                .ValidationErrorJson = BuildToolArgumentValidationError(
                    ToolWrite,
                    "conflicting_tool_arguments",
                    "workspace_write received conflicting values for text/content.",
                    unknown:=New String() {"text", "content"})
            }
                End If
            ElseIf hasContent Then
                arguments("text") = If(contentValue, "")
                aliasesUsed.Add("content->text")
            End If

            Dim pathValue As Object = Nothing
            Dim filenameValue As Object = Nothing
            Dim hasPath = TryGetArgumentValue(arguments, "path", pathValue)
            Dim hasFilename = TryGetArgumentValue(arguments, "filename", filenameValue)

            If Not hasPath AndAlso hasFilename Then
                Dim filenameAlias = If(filenameValue, "").ToString().Trim()
                Dim sanitized = SanitizeName(Path.GetFileName(filenameAlias))
                arguments("path") = sanitized
                aliasesUsed.Add("filename->path")
                hasPath = True
                pathValue = sanitized
            End If

            Dim modeValue As Object = Nothing
            Dim overwriteValue As Object = Nothing
            Dim hasMode = TryGetArgumentValue(arguments, "mode", modeValue)
            Dim hasOverwrite = TryGetArgumentValue(arguments, "overwrite", overwriteValue)

            Dim normalizedMode As String = ""

            If hasMode Then
                normalizedMode = NormalizeWorkspaceWriteMode(If(modeValue, "").ToString())
                If String.IsNullOrWhiteSpace(normalizedMode) Then
                    Return New ToolArgumentNormalizationResult() With {
                .Arguments = arguments,
                .ValidationErrorJson = BuildToolArgumentValidationError(
                    ToolWrite,
                    "invalid_tool_argument",
                    "workspace_write mode must be one of overwrite, append, or create.",
                    unknown:=New String() {"mode"})
            }
                End If
            End If

            If hasOverwrite Then
                Dim overwriteMode = If(GetBool(arguments, "overwrite", False), "overwrite", "create")

                If String.IsNullOrWhiteSpace(normalizedMode) Then
                    normalizedMode = overwriteMode
                    aliasesUsed.Add("overwrite->mode")
                ElseIf Not String.Equals(normalizedMode, overwriteMode, StringComparison.OrdinalIgnoreCase) Then
                    Return New ToolArgumentNormalizationResult() With {
                .Arguments = arguments,
                .ValidationErrorJson = BuildToolArgumentValidationError(
                    ToolWrite,
                    "conflicting_tool_arguments",
                    "workspace_write received conflicting values for mode/overwrite.",
                    unknown:=New String() {"mode", "overwrite"})
            }
                End If
            End If

            If String.IsNullOrWhiteSpace(normalizedMode) Then
                normalizedMode = "overwrite"
            End If
            arguments("mode") = normalizedMode

            If Not hasPath OrElse String.IsNullOrWhiteSpace(If(pathValue, "").ToString()) Then
                Return New ToolArgumentNormalizationResult() With {
            .Arguments = arguments,
            .ValidationErrorJson = BuildToolArgumentValidationError(
                ToolWrite,
                "missing_required_tool_argument",
                "workspace_write requires path and text/content.",
                missing:=New String() {"path"})
        }
            End If

            Dim normalizedTextValue As Object = Nothing
            Dim hasNormalizedText = TryGetArgumentValue(arguments, "text", normalizedTextValue)

            If Not hasNormalizedText Then
                Dim unexpected = GetUnexpectedWriteArguments(arguments)
                If unexpected.Count > 0 Then
                    Return New ToolArgumentNormalizationResult() With {
                .Arguments = arguments,
                .ValidationErrorJson = BuildToolArgumentValidationError(
                    ToolWrite,
                    "unknown_tool_argument",
                    "workspace_write did not receive usable text/content and ignored unexpected fields.",
                    missing:=New String() {"text"},
                    unknown:=unexpected,
                    aliasesUsed:=aliasesUsed)
            }
                End If

                Return New ToolArgumentNormalizationResult() With {
            .Arguments = arguments,
            .ValidationErrorJson = BuildToolArgumentValidationError(
                ToolWrite,
                "missing_required_tool_argument",
                "workspace_write requires text/content.",
                missing:=New String() {"text"},
                aliasesUsed:=aliasesUsed)
        }
            End If

            Dim normalizedText = If(normalizedTextValue, "").ToString()
            arguments("text") = normalizedText
            arguments(MetaAliasesUsedKey) = ToJArray(aliasesUsed)
            arguments(MetaContentProvidedKey) = True
            arguments(MetaContentWasEmptyKey) = (normalizedText.Length = 0)

            Return New ToolArgumentNormalizationResult() With {
        .Arguments = arguments
    }
        End Function

        Private Shared Function NormalizeWorkspaceWriteMode(value As String) As String
            Select Case If(value, "").Trim().ToLowerInvariant()
                Case "", "overwrite"
                    Return "overwrite"
                Case "append"
                    Return "append"
                Case "create", "create_new"
                    Return "create"
                Case Else
                    Return ""
            End Select
        End Function

        Private Shared Function ToCaseInsensitiveDictionary(arguments As IDictionary(Of String, Object)) As Dictionary(Of String, Object)
            Dim normalized As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)

            If arguments Is Nothing Then
                Return normalized
            End If

            For Each kvp In arguments
                If String.IsNullOrWhiteSpace(kvp.Key) Then Continue For
                normalized(kvp.Key) = kvp.Value
            Next

            Return normalized
        End Function

        Private Shared Function TryGetArgumentValue(arguments As IDictionary(Of String, Object),
                                            name As String,
                                            ByRef value As Object) As Boolean
            value = Nothing
            If arguments Is Nothing OrElse String.IsNullOrWhiteSpace(name) Then
                Return False
            End If

            Return arguments.TryGetValue(name, value)
        End Function

        Private Shared Function GetUnexpectedWriteArguments(arguments As IDictionary(Of String, Object)) As List(Of String)
            Dim unexpected As New List(Of String)()

            If arguments Is Nothing Then
                Return unexpected
            End If

            For Each key In arguments.Keys
                If String.IsNullOrWhiteSpace(key) Then Continue For
                If key.StartsWith("__", StringComparison.Ordinal) Then Continue For

                Select Case key.ToLowerInvariant()
                    Case "path", "text", "mode", "content", "overwrite", "filename"
                        Continue For
                    Case Else
                        unexpected.Add(key)
                End Select
            Next

            Return unexpected
        End Function

        Private Shared Function BuildToolArgumentValidationError(toolName As String,
                                                         errorCode As String,
                                                         message As String,
                                                         Optional missing As IEnumerable(Of String) = Nothing,
                                                         Optional unknown As IEnumerable(Of String) = Nothing,
                                                         Optional aliasesUsed As IEnumerable(Of String) = Nothing) As String
            Dim errorObj As New JObject(
        New JProperty("ok", False),
        New JProperty("error", New JObject(
            New JProperty("code", errorCode),
            New JProperty("tool", toolName),
            New JProperty("message", message))))

            Dim inner = DirectCast(errorObj("error"), JObject)

            Dim missingArray = ToJArray(missing)
            If missingArray.Count > 0 Then
                inner("missing") = missingArray
            End If

            Dim unknownArray = ToJArray(unknown)
            If unknownArray.Count > 0 Then
                inner("unknown") = unknownArray
            End If

            Dim aliasArray = ToJArray(aliasesUsed)
            If aliasArray.Count > 0 Then
                inner("aliasesUsed") = aliasArray
            End If

            Return errorObj.ToString(Formatting.None)
        End Function

        Private Shared Function ToJArray(values As IEnumerable(Of String)) As JArray
            Dim arr As New JArray()

            If values Is Nothing Then
                Return arr
            End If

            For Each value In values
                If value Is Nothing Then Continue For
                arr.Add(value)
            Next

            Return arr
        End Function


        ' --------------------------------------------------------------- guards

        Private Shared Function RequireConnected() As String
            If _active Is Nothing OrElse String.IsNullOrWhiteSpace(_active.RootPath) OrElse Not Directory.Exists(_active.RootPath) Then
                Return Err_("no_workspace", "No workspace is configured.")
            End If
            Return Nothing
        End Function

        Private Shared Function ResolveInsideWorkspace(relOrFull As String, Optional allowRoot As Boolean = False) As String
            If _active Is Nothing OrElse String.IsNullOrWhiteSpace(_active.RootPath) Then
                Throw New UnauthorizedAccessException("No workspace configured.")
            End If
            Dim root = Path.GetFullPath(_active.RootPath)
            Dim rel = If(relOrFull, "").Trim()
            If rel = "" Then
                If allowRoot Then Return root
                Throw New ArgumentException("Path is empty.")
            End If
            Dim full As String
            If Path.IsPathRooted(rel) Then
                full = Path.GetFullPath(rel)
            Else
                full = Path.GetFullPath(Path.Combine(root, rel))
            End If
            If Not full.Equals(root, StringComparison.OrdinalIgnoreCase) AndAlso
               Not full.StartsWith(root & Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) Then
                Throw New UnauthorizedAccessException("Path is outside the workspace.")
            End If
            Return full
        End Function

        ' --------------------------------------------------------------- execute

        Private Shared Function ExecuteGet() As String
            Dim connected = _active IsNot Nothing AndAlso
                            Not String.IsNullOrWhiteSpace(_active.RootPath) AndAlso
                            Directory.Exists(_active.RootPath)
            Return JsonConvert.SerializeObject(New With {
                Key .connected = connected,
                Key .rootPath = If(connected, _active.RootPath, ""),
                Key .allowRead = _active.AllowRead,
                Key .allowWrite = _active.AllowWrite,
                Key .allowMoveCopyRename = _active.AllowMoveCopyRename,
                Key .allowDelete = _active.AllowDelete,
                Key .fallbackRoot = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            })
        End Function

        Private Shared Function ExecuteInventory(args As IDictionary(Of String, Object)) As String
            Dim err = RequireConnected() : If err IsNot Nothing Then Return err
            If Not _active.AllowRead Then Return Err_("not_permitted", "Workspace read is disabled.")

            Dim relRoot = GetStr(args, "path")
            Dim glob = If(GetStr(args, "glob"), "")
            Dim recursive = GetBool(args, "recursive", True)
            Dim maxItems = Math.Min(Math.Max(GetInt(args, "max_items", 500), 1), 5000)

            Dim baseDir = ResolveInsideWorkspace(relRoot, allowRoot:=True)
            If Not Directory.Exists(baseDir) Then Return Err_("not_found", "Folder not found.")

            Dim searchOpt = If(recursive, IO.SearchOption.AllDirectories, IO.SearchOption.TopDirectoryOnly)
            Dim pattern = If(String.IsNullOrWhiteSpace(glob), "*", glob)
            Dim list As New List(Of Object)
            For Each f In Directory.EnumerateFileSystemEntries(baseDir, pattern, searchOpt)
                If list.Count >= maxItems Then Exit For
                Try
                    Dim isDir = Directory.Exists(f)
                    Dim attrs = File.GetAttributes(f)
                    If Not _active.IncludeHiddenSystem AndAlso
                       ((attrs And FileAttributes.Hidden) <> 0 OrElse (attrs And FileAttributes.System) <> 0) Then
                        Continue For
                    End If
                    Dim rel = f.Substring(_active.RootPath.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                    If isDir Then
                        list.Add(New With {Key .path = rel, Key .kind = "dir"})
                    Else
                        Dim fi As New FileInfo(f)
                        list.Add(New With {
                            Key .path = rel,
                            Key .kind = "file",
                            Key .size = fi.Length,
                            Key .mtime = fi.LastWriteTimeUtc
                        })
                    End If
                Catch
                End Try
            Next

            Return JsonConvert.SerializeObject(New With {
                Key .root = _active.RootPath,
                Key .basePath = baseDir,
                Key .count = list.Count,
                Key .items = list
            })
        End Function

        Private Shared Function ExecuteRead(args As IDictionary(Of String, Object)) As String
            Dim err = RequireConnected() : If err IsNot Nothing Then Return err
            If Not _active.AllowRead Then Return Err_("not_permitted", "Workspace read is disabled.")

            Dim p = ResolveInsideWorkspace(GetStr(args, "path"))
            If Not File.Exists(p) Then Return Err_("not_found", "File not found.")
            Dim fi As New FileInfo(p)
            If fi.Length > PathPolicy.MaxFileSizeBytes Then
                Return Err_("file_too_large", "File exceeds " & PathPolicy.MaxFileSizeBytes & " bytes.")
            End If
            Dim text = File.ReadAllText(p, Encoding.UTF8)
            Dim maxChars = GetInt(args, "max_chars", 0)
            Dim truncated = False
            If maxChars > 0 AndAlso text.Length > maxChars Then
                text = text.Substring(0, maxChars)
                truncated = True
            End If
            Return JsonConvert.SerializeObject(New With {
                Key .path = p, Key .size = fi.Length, Key .truncated = truncated, Key .text = text
            })
        End Function

        Private Shared Function ExecuteReadMany(args As IDictionary(Of String, Object)) As String
            Dim err = RequireConnected() : If err IsNot Nothing Then Return err
            If Not _active.AllowRead Then Return Err_("not_permitted", "Workspace read is disabled.")

            Dim paths = GetStringList(args, "paths")
            If paths.Count = 0 Then Return Err_("missing_paths", "paths is required.")

            Dim maxFiles = Math.Min(Math.Max(GetInt(args, "max_files", 20), 1), 100)
            Dim maxChars = GetInt(args, "max_chars_per_file", 0)

            Dim requestedCount = paths.Count
            Dim selected = paths.Take(maxFiles).ToList()
            Dim items As New List(Of Object)

            For Each relPath In selected
                Try
                    Dim p = ResolveInsideWorkspace(relPath)
                    If Not File.Exists(p) Then
                        items.Add(New With {
                            Key .path = relPath,
                            Key .error = "not_found",
                            Key .message = "File not found."
                        })
                        Continue For
                    End If

                    Dim fi As New FileInfo(p)
                    If fi.Length > PathPolicy.MaxFileSizeBytes Then
                        items.Add(New With {
                            Key .path = p,
                            Key .error = "file_too_large",
                            Key .size = fi.Length,
                            Key .message = "File exceeds " & PathPolicy.MaxFileSizeBytes & " bytes."
                        })
                        Continue For
                    End If

                    Dim text = File.ReadAllText(p, Encoding.UTF8)
                    Dim truncated = False
                    If maxChars > 0 AndAlso text.Length > maxChars Then
                        text = text.Substring(0, maxChars)
                        truncated = True
                    End If

                    items.Add(New With {
                        Key .path = p,
                        Key .size = fi.Length,
                        Key .truncated = truncated,
                        Key .text = text
                    })
                Catch uae As UnauthorizedAccessException
                    items.Add(New With {
                        Key .path = relPath,
                        Key .error = "access_denied",
                        Key .message = uae.Message
                    })
                Catch ex As Exception
                    items.Add(New With {
                        Key .path = relPath,
                        Key .error = "read_failed",
                        Key .message = ex.Message
                    })
                End Try
            Next

            Return JsonConvert.SerializeObject(New With {
                Key .requested_count = requestedCount,
                Key .processed_count = selected.Count,
                Key .items = items
            })
        End Function

        Private Shared Function ExecuteWrite(args As IDictionary(Of String, Object)) As String
            Dim err = RequireConnected() : If err IsNot Nothing Then Return err
            If Not _active.AllowWrite Then Return Err_("not_permitted", "Workspace write is disabled.")

            Dim relPath = GetStr(args, "path")
            Dim mode = GetStr(args, "mode").Trim().ToLowerInvariant()
            Dim text = GetStr(args, "text")
            Dim aliasesUsed = GetStringList(args, MetaAliasesUsedKey)
            Dim contentProvided = GetBool(args, MetaContentProvidedKey, False)
            Dim contentWasEmptyIntentionally = GetBool(args, MetaContentWasEmptyKey, False)

            Dim target = ResolveInsideWorkspace(relPath)

            If IsBinaryWorkspaceExtension(target) Then
                Return Err_(
            "binary_extension_not_supported",
            "workspace_write only creates UTF-8 text/code files. " &
            "Do not use it for .pdf, .docx, .xlsx, .pptx, image, archive, audio, video, or other binary/document formats. " &
            "Use a dedicated binary-producing tool and then copy the real output file into the workspace.")
            End If

            Dim dir = Path.GetDirectoryName(target)
            If Not String.IsNullOrWhiteSpace(dir) AndAlso Not Directory.Exists(dir) Then
                Directory.CreateDirectory(dir)
            End If

            Select Case mode
                Case "append"
                    File.AppendAllText(target, text, Encoding.UTF8)

                Case "create"
                    If File.Exists(target) Then
                        Return Err_("exists", "File already exists.")
                    End If
                    File.WriteAllText(target, text, Encoding.UTF8)

                Case Else
                    File.WriteAllText(target, text, Encoding.UTF8)
            End Select

            Dim fi As New FileInfo(target)

            ' Pick up edits to SKILL.md/AGENT.md immediately when writing into a resource root.
            AgentResources.RefreshIfResourcePath(target)

            Return JsonConvert.SerializeObject(New With {
        Key .path = target,
        Key .charsWritten = text.Length,
        Key .bytesWritten = Encoding.UTF8.GetByteCount(text),
        Key .fileSizeBytes = fi.Length,
        Key .mode = mode,
        Key .usedAliases = aliasesUsed.Count > 0,
        Key .aliasesUsed = aliasesUsed,
        Key .contentWasEmptyIntentionally = contentProvided AndAlso contentWasEmptyIntentionally
    })
        End Function

        Private Shared Function ExecuteSearch(args As IDictionary(Of String, Object)) As String
            Dim err = RequireConnected() : If err IsNot Nothing Then Return err
            If Not _active.AllowRead Then Return Err_("not_permitted", "Workspace read is disabled.")

            Dim query = GetStr(args, "query")
            If String.IsNullOrWhiteSpace(query) Then Return Err_("missing_query", "query is required.")
            Dim glob = If(GetStr(args, "glob"), "*.txt")
            Dim recursive = GetBool(args, "recursive", True)
            Dim useRegex = GetBool(args, "regex", False)
            Dim ignoreCase = GetBool(args, "ignore_case", True)
            Dim maxFiles = Math.Min(Math.Max(GetInt(args, "max_files", 50), 1), 500)
            Dim maxHitsPerFile = Math.Min(Math.Max(GetInt(args, "max_hits_per_file", 5), 1), 50)

            Dim opt As RegexOptions = RegexOptions.CultureInvariant
            If ignoreCase Then opt = opt Or RegexOptions.IgnoreCase
            Dim rx As Regex = Nothing
            If useRegex Then rx = New Regex(query, opt, TimeSpan.FromSeconds(2))

            Dim cmp = If(ignoreCase, StringComparison.OrdinalIgnoreCase, StringComparison.Ordinal)
            Dim searchOpt = If(recursive, IO.SearchOption.AllDirectories, IO.SearchOption.TopDirectoryOnly)
            Dim files = Directory.EnumerateFiles(_active.RootPath, glob, searchOpt).Take(maxFiles).ToList()

            Dim out As New List(Of Object)
            For Each f In files
                Try
                    Dim fi As New FileInfo(f)
                    If fi.Length > PathPolicy.MaxFileSizeBytes Then Continue For
                    Dim text = File.ReadAllText(f, Encoding.UTF8)
                    Dim hits As New List(Of Object)
                    If useRegex Then
                        For Each m As Match In rx.Matches(text)
                            If hits.Count >= maxHitsPerFile Then Exit For
                            hits.Add(BuildHit(text, m.Index, m.Length, m.Value))
                        Next
                    Else
                        Dim idx As Integer = 0
                        While idx < text.Length
                            Dim found = text.IndexOf(query, idx, cmp)
                            If found < 0 Then Exit While
                            hits.Add(BuildHit(text, found, query.Length, text.Substring(found, query.Length)))
                            If hits.Count >= maxHitsPerFile Then Exit While
                            idx = found + Math.Max(1, query.Length)
                        End While
                    End If
                    If hits.Count > 0 Then
                        Dim rel = f.Substring(_active.RootPath.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                        out.Add(New With {Key .path = rel, Key .hits = hits})
                    End If
                Catch
                End Try
            Next

            Return JsonConvert.SerializeObject(New With {
                Key .root = _active.RootPath,
                Key .files_searched = files.Count,
                Key .results = out
            })
        End Function

        Private Shared Function ExecuteCopyOrMove(args As IDictionary(Of String, Object), isMove As Boolean) As String
            Dim err = RequireConnected() : If err IsNot Nothing Then Return err
            If Not _active.AllowMoveCopyRename Then Return Err_("not_permitted", "Move/Copy/Rename is disabled.")
            Dim src = ResolveInsideWorkspace(GetStr(args, "source"))
            Dim dst = ResolveInsideWorkspace(GetStr(args, "destination"))
            If Not File.Exists(src) AndAlso Not Directory.Exists(src) Then Return Err_("not_found", "Source not found.")
            Dim parent = Path.GetDirectoryName(dst)
            If Not String.IsNullOrWhiteSpace(parent) AndAlso Not Directory.Exists(parent) Then Directory.CreateDirectory(parent)
            If File.Exists(src) Then
                If isMove Then File.Move(src, dst) Else File.Copy(src, dst, overwrite:=GetBool(args, "overwrite", False))
            Else
                If isMove Then
                    Directory.Move(src, dst)
                Else
                    FileSystem.CopyDirectory(src, dst, UIOption.OnlyErrorDialogs, UICancelOption.ThrowException)
                End If
            End If
            Return JsonConvert.SerializeObject(New With {Key .source = src, Key .destination = dst, Key .moved = isMove})
        End Function

        Private Shared Function ExecuteRename(args As IDictionary(Of String, Object)) As String
            Dim err = RequireConnected() : If err IsNot Nothing Then Return err
            If Not _active.AllowMoveCopyRename Then Return Err_("not_permitted", "Move/Copy/Rename is disabled.")
            Dim src = ResolveInsideWorkspace(GetStr(args, "path"))
            Dim newName = SanitizeName(GetStr(args, "new_name"))
            If String.IsNullOrWhiteSpace(newName) Then Return Err_("missing_name", "new_name is required.")
            Dim parent = Path.GetDirectoryName(src)
            Dim dst = Path.Combine(parent, newName)
            dst = ResolveInsideWorkspace(dst)
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
            Dim err = RequireConnected() : If err IsNot Nothing Then Return err
            If Not _active.AllowDelete Then Return Err_("not_permitted", "Delete is disabled.")
            Dim p = ResolveInsideWorkspace(GetStr(args, "path"))
            Dim toTrash = GetBool(args, "to_trash", True)
            If File.Exists(p) Then
                FileSystem.DeleteFile(p, UIOption.OnlyErrorDialogs,
                    If(toTrash, RecycleOption.SendToRecycleBin, RecycleOption.DeletePermanently),
                    UICancelOption.ThrowException)
            ElseIf Directory.Exists(p) Then
                FileSystem.DeleteDirectory(p, UIOption.OnlyErrorDialogs,
                    If(toTrash, RecycleOption.SendToRecycleBin, RecycleOption.DeletePermanently),
                    UICancelOption.ThrowException)
            Else
                Return Err_("not_found", "Path not found.")
            End If
            Return JsonConvert.SerializeObject(New With {Key .path = p, Key .to_trash = toTrash})
        End Function

        Private Shared Function ExecuteMakeDir(args As IDictionary(Of String, Object)) As String
            Dim err = RequireConnected() : If err IsNot Nothing Then Return err
            If Not _active.AllowWrite Then Return Err_("not_permitted", "Workspace write is disabled.")
            Dim p = ResolveInsideWorkspace(GetStr(args, "path"))
            Directory.CreateDirectory(p)
            Return JsonConvert.SerializeObject(New With {Key .path = p, Key .created = True})
        End Function

        ' --------------------------------------------------------------- helpers

        Private Shared Function BuildHit(text As String, index As Integer, length As Integer, match As String) As Object
            Dim winStart = Math.Max(0, index - 40)
            Dim winEnd = Math.Min(text.Length, index + length + 40)
            Dim ctx = text.Substring(winStart, winEnd - winStart).Replace(vbCr, " ").Replace(vbLf, " ")
            Return New With {Key .index = index, Key .length = length, Key .match = match, Key .context = ctx}
        End Function

        Private Shared Function UniquePath(p As String) As String
            If Not File.Exists(p) AndAlso Not Directory.Exists(p) Then Return p
            Dim dir = Path.GetDirectoryName(p)
            Dim name = Path.GetFileNameWithoutExtension(p)
            Dim ext = Path.GetExtension(p)
            For i = 2 To 1000
                Dim c = Path.Combine(dir, name & " (" & i.ToString() & ")" & ext)
                If Not File.Exists(c) AndAlso Not Directory.Exists(c) Then Return c
            Next
            Return Path.Combine(dir, name & "_" & Guid.NewGuid().ToString("N").Substring(0, 8) & ext)
        End Function

        Private Shared Function SanitizeName(name As String) As String
            If String.IsNullOrWhiteSpace(name) Then Return "untitled"
            Dim invalid = Path.GetInvalidFileNameChars()
            Dim sb As New StringBuilder(name.Length)
            For Each c In name
                If Array.IndexOf(invalid, c) >= 0 Then sb.Append("_"c) Else sb.Append(c)
            Next
            Return sb.ToString()
        End Function

        Private Shared Function Err_(code As String, message As String) As String
            Return JsonConvert.SerializeObject(New With {Key .error = code, Key .message = message})
        End Function

        ' --------------------------------------------------------------- argument helpers

        Private Shared Function GetStr(args As IDictionary(Of String, Object), name As String) As String
            If args Is Nothing Then Return ""
            Dim v As Object = Nothing
            If Not args.TryGetValue(name, v) OrElse v Is Nothing Then Return ""
            Return System.Convert.ToString(v)
        End Function

        Private Shared Function GetStringList(args As IDictionary(Of String, Object), name As String) As List(Of String)
            Dim list As New List(Of String)
            If args Is Nothing Then Return list

            Dim v As Object = Nothing
            If Not args.TryGetValue(name, v) OrElse v Is Nothing Then Return list

            If TypeOf v Is JArray Then
                For Each tk In CType(v, JArray)
                    Dim s = tk.ToString()
                    If Not String.IsNullOrWhiteSpace(s) Then list.Add(s)
                Next
                Return list
            End If

            If TypeOf v Is IEnumerable(Of Object) Then
                For Each o In CType(v, IEnumerable(Of Object))
                    Dim s = System.Convert.ToString(o)
                    If Not String.IsNullOrWhiteSpace(s) Then list.Add(s)
                Next
                Return list
            End If

            Dim single_ = System.Convert.ToString(v)
            If Not String.IsNullOrWhiteSpace(single_) Then list.Add(single_)
            Return list
        End Function

        Private Shared Function GetInt(args As IDictionary(Of String, Object), name As String, defaultValue As Integer) As Integer
            If args Is Nothing Then Return defaultValue
            Dim v As Object = Nothing
            If Not args.TryGetValue(name, v) OrElse v Is Nothing Then Return defaultValue
            Try : Return System.Convert.ToInt32(v) : Catch
                Dim n As Integer
                If Integer.TryParse(System.Convert.ToString(v), n) Then Return n
                Return defaultValue
            End Try
        End Function

        Private Shared Function GetBool(args As IDictionary(Of String, Object), name As String, defaultValue As Boolean) As Boolean
            If args Is Nothing Then Return defaultValue
            Dim v As Object = Nothing
            If Not args.TryGetValue(name, v) OrElse v Is Nothing Then Return defaultValue
            Try : Return System.Convert.ToBoolean(v) : Catch
                Select Case System.Convert.ToString(v).Trim().ToLowerInvariant()
                    Case "true", "1", "yes" : Return True
                    Case "false", "0", "no" : Return False
                    Case Else : Return defaultValue
                End Select
            End Try
        End Function

        ' --------------------------------------------------------------- factories

        Private Shared Function BuildGet() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolGet, .Tool = True, .ToolPriority = 910, .ToolErrorHandling = "skip",
                .ModelDescription = "Workspace (info)",
                .ToolDefinition = "{""name"":""" & ToolGet & """,""description"":""Return current workspace info and permissions. If 'connected' is false, no workspace is configured and writes go to the current session's staging/working area (delivered to the user at the end of the run), or to the Desktop when no session staging area is active."",""parameters"":{""type"":""object"",""properties"":{}}}",
                .ToolInstructionsPrompt = ToolGet & ": Inspect the current workspace state. Call this once to learn whether a workspace is configured and what permissions you have."
            }
        End Function

        Private Shared Function BuildInventory() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolInventory, .Tool = True, .ToolPriority = 911, .ToolErrorHandling = "skip",
                .ModelDescription = "Workspace (inventory)",
                .ToolDefinition = "{""name"":""" & ToolInventory & """,""description"":""List files and folders in the workspace (or a sub-folder). Supports glob filter and recursion."",""parameters"":{""type"":""object"",""properties"":{""path"":{""type"":""string"",""description"":""Optional sub-folder relative to the workspace root.""},""glob"":{""type"":""string"",""description"":""Wildcard filter such as '*.txt' or 'report-*.docx'. Default '*'.""},""recursive"":{""type"":""boolean"",""description"":""Recurse into subdirectories (default true).""},""max_items"":{""type"":""integer"",""description"":""Maximum entries (default 500, capped 5000).""}}}}",
                .ToolInstructionsPrompt = ToolInventory & ": Enumerate workspace contents to discover available files."
            }
        End Function

        Private Shared Function BuildRead() As ModelConfig
            Dim def =
                "{""name"":""" & ToolRead & """," &
                """description"":""Read one UTF-8 text file (not other formats!) inside the workspace. This tool is for a single file only and is not OK for many files. Use workspace_read_many when you need to read multiple text files. Use workspace_extract_text for non-plain-text formats.""," &
                """parameters"":{""type"":""object""," &
                """properties"":{" &
                """path"":{""type"":""string"",""description"":""Workspace-relative path.""}," &
                """max_chars"":{""type"":""integer"",""description"":""Optional cap on returned text length (0 = no cap).""}}," &
                """required"":[""path""]}}"

            Return New ModelConfig() With {
                .ToolName = ToolRead,
                .Tool = True,
                .ToolPriority = 912,
                .ToolErrorHandling = "skip",
                .ModelDescription = "Workspace (read)",
                .ToolDefinition = def,
                .ToolInstructionsPrompt = ToolRead & ": Read one text file from inside the workspace. Do not use this for many files; use " & ToolReadMany & " instead. For non-plain-text files, use " & ToolExtractText & "."
            }
        End Function

        Private Shared Function BuildReadMany() As ModelConfig
            Dim def =
                "{""name"":""" & ToolReadMany & """," &
                """description"":""Read multiple UTF-8 text files (not other formats!) inside the workspace in one call. Use this when you need many files instead of calling workspace_read repeatedly.""," &
                """parameters"":{""type"":""object""," &
                """properties"":{" &
                """paths"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Workspace-relative paths to read.""}," &
                """max_chars_per_file"":{""type"":""integer"",""description"":""Optional cap on returned text length for each file (0 = no cap).""}," &
                """max_files"":{""type"":""integer"",""description"":""Maximum files to process from 'paths' (default 20, capped 100).""}}," &
                """required"":[""paths""]}}"

            Return New ModelConfig() With {
                .ToolName = ToolReadMany,
                .Tool = True,
                .ToolPriority = 912,
                .ToolErrorHandling = "skip",
                .ModelDescription = "Workspace (read many)",
                .ToolDefinition = def,
                .ToolInstructionsPrompt = ToolReadMany & ": Read many text files from inside the workspace in one call."
            }
        End Function

        Private Shared Function BuildExtractText() As ModelConfig
            Dim def =
                "{""name"":""" & ToolExtractText & """," &
                """description"":""Extract readable text from one workspace file of any supported format (PDF with OCR fallback, DOCX, DOC, RTF, XLSX, XLS, PPTX, PPT, EML, MSG, TXT/CSV/JSON/XML/HTML/Markdown, and images/audio/video via the configured model). Supports incremental retrieval by character window and optional page-range hints when available. This tool is for a single file only and is not OK for many files. Use workspace_extract_text_many when you need to extract multiple files.""," &
                """parameters"":{""type"":""object""," &
                """properties"":{" &
                """path"":{""type"":""string"",""description"":""Workspace-relative path of the file to extract.""}," &
                """max_chars"":{""type"":""integer"",""description"":""Optional cap on returned characters (default 12000, max 500000).""}," &
                """start_char"":{""type"":""integer"",""description"":""Optional zero-based character offset for incremental extraction.""}," &
                """offset"":{""type"":""integer"",""description"":""Alias of start_char for incremental extraction.""}," &
                """start_page"":{""type"":""integer"",""description"":""Optional one-based start page hint when the extractor supports page ranges.""}," &
                """end_page"":{""type"":""integer"",""description"":""Optional one-based end page hint when the extractor supports page ranges.""}}," &
                """required"":[""path""]}}"

            Return New ModelConfig() With {
                .ToolName = ToolExtractText,
                .ToolDefinition = def,
                .ToolInstructionsPrompt = ToolExtractText & ": Extract text from one supported workspace file. Do not use this for many files; use " & ToolExtractTextMany & " instead.",
                .ModelDescription = "Workspace file text extractor",
                .Tool = True,
                .ToolPriority = 905,
                .ToolErrorHandling = "skip"
            }
        End Function

        Private Shared Function BuildExtractTextMany() As ModelConfig
            Dim def =
                "{""name"":""" & ToolExtractTextMany & """," &
                """description"":""Extract readable text from multiple workspace files of supported formats in one call. Use this instead of repeated workspace_extract_text calls when you need many files.""," &
                """parameters"":{""type"":""object""," &
                """properties"":{" &
                """paths"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Workspace-relative paths of the files to extract.""}," &
                """max_chars_per_file"":{""type"":""integer"",""description"":""Optional cap on returned characters for each file (default 100000, max 500000).""}," &
                """max_files"":{""type"":""integer"",""description"":""Maximum files to process from 'paths' (default 20, capped 100).""}}," &
                """required"":[""paths""]}}"

            Return New ModelConfig() With {
                .ToolName = ToolExtractTextMany,
                .ToolDefinition = def,
                .ToolInstructionsPrompt = ToolExtractTextMany & ": Extract text from many supported workspace files in one call.",
                .ModelDescription = "Workspace file text extractor (many)",
                .Tool = True,
                .ToolPriority = 905,
                .ToolErrorHandling = "skip"
            }
        End Function


        Private Shared Function BuildWrite() As ModelConfig
            Return New ModelConfig() With {
        .ToolName = ToolWrite,
        .Tool = True,
        .ToolPriority = 913,
        .ToolErrorHandling = "skip",
        .ModelDescription = "Workspace (write)",
        .ToolDefinition = "{""name"":""" & ToolWrite & """,""description"":""Write a UTF-8 text file inside the workspace. Never use this tool for binary/document formats such as PDF, DOCX, XLSX, PPTX, ZIP, images, audio, or video."",""parameters"":{""type"":""object"",""properties"":{""path"":{""type"":""string"",""description"":""Workspace-relative target path.""},""text"":{""type"":""string"",""description"":""Content to write. May be intentionally empty.""},""mode"":{""type"":""string"",""enum"":[""overwrite"",""append"",""create""],""description"":""Write mode. Default 'overwrite'.""}}, ""required"":[""path"",""text""]}}",
        .ToolInstructionsPrompt = ToolWrite & ": Write a text file inside the workspace. Never use this tool to create or overwrite PDF, Office, image, archive, audio, video, or other binary files."
    }
        End Function


        Private Shared Function BuildSearch() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolSearch, .Tool = True, .ToolPriority = 914, .ToolErrorHandling = "skip",
                .ModelDescription = "Workspace (search)",
                .ToolDefinition = "{""name"":""" & ToolSearch & """,""description"":""Search file contents across the workspace (text files only). Returns per-file matches with line context."",""parameters"":{""type"":""object"",""properties"":{""query"":{""type"":""string""},""glob"":{""type"":""string"",""description"":""File pattern (default '*.txt').""},""recursive"":{""type"":""boolean"",""description"":""Default true.""},""regex"":{""type"":""boolean"",""description"":""Treat query as regex (default false).""},""ignore_case"":{""type"":""boolean"",""description"":""Default true.""},""max_files"":{""type"":""integer""},""max_hits_per_file"":{""type"":""integer""}},""required"":[""query""]}}",
                .ToolInstructionsPrompt = ToolSearch & ": Search workspace files for a substring or regex."
            }
        End Function

        Private Shared Function BuildCopy() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolCopy, .Tool = True, .ToolPriority = 915, .ToolErrorHandling = "skip",
                .ModelDescription = "Workspace (copy)",
                .ToolDefinition = "{""name"":""" & ToolCopy & """,""description"":""Copy a file or folder inside the workspace."",""parameters"":{""type"":""object"",""properties"":{""source"":{""type"":""string""},""destination"":{""type"":""string""},""overwrite"":{""type"":""boolean"",""description"":""Default false.""}},""required"":[""source"",""destination""]}}",
                .ToolInstructionsPrompt = ToolCopy & ": Copy within the workspace."
            }
        End Function

        Private Shared Function BuildMove() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolMove, .Tool = True, .ToolPriority = 916, .ToolErrorHandling = "skip",
                .ModelDescription = "Workspace (move)",
                .ToolDefinition = "{""name"":""" & ToolMove & """,""description"":""Move a file or folder inside the workspace."",""parameters"":{""type"":""object"",""properties"":{""source"":{""type"":""string""},""destination"":{""type"":""string""}},""required"":[""source"",""destination""]}}",
                .ToolInstructionsPrompt = ToolMove & ": Move within the workspace."
            }
        End Function

        Private Shared Function BuildRename() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolRename, .Tool = True, .ToolPriority = 917, .ToolErrorHandling = "skip",
                .ModelDescription = "Workspace (rename)",
                .ToolDefinition = "{""name"":""" & ToolRename & """,""description"":""Rename a file or folder in place inside the workspace."",""parameters"":{""type"":""object"",""properties"":{""path"":{""type"":""string""},""new_name"":{""type"":""string""}},""required"":[""path"",""new_name""]}}",
                .ToolInstructionsPrompt = ToolRename & ": Rename within the workspace."
            }
        End Function

        Private Shared Function BuildDelete() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolDelete, .Tool = True, .ToolPriority = 918, .ToolErrorHandling = "skip",
                .ModelDescription = "Workspace (delete)",
                .ToolDefinition = "{""name"":""" & ToolDelete & """,""description"":""Delete a file or folder inside the workspace. By default the entry is moved to the Recycle Bin (recoverable)."",""parameters"":{""type"":""object"",""properties"":{""path"":{""type"":""string""},""to_trash"":{""type"":""boolean"",""description"":""Default true. Set false for permanent deletion.""}},""required"":[""path""]}}",
                .ToolInstructionsPrompt = ToolDelete & ": Delete (to Recycle Bin by default) within the workspace."
            }
        End Function

        Private Shared Function BuildMakeDir() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolMakeDir, .Tool = True, .ToolPriority = 919, .ToolErrorHandling = "skip",
                .ModelDescription = "Workspace (make dir)",
                .ToolDefinition = "{""name"":""" & ToolMakeDir & """,""description"":""Create a folder inside the workspace (creates intermediate folders as needed)."",""parameters"":{""type"":""object"",""properties"":{""path"":{""type"":""string""}},""required"":[""path""]}}",
                .ToolInstructionsPrompt = ToolMakeDir & ": Create a folder inside the workspace."
            }
        End Function

    End Class

End Namespace
