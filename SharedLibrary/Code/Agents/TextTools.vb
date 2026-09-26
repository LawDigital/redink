' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: TextTools.vb
' Purpose: Built-in tools for plain-text file operations:
'            text_read — read a UTF-8 text file (capped by PathPolicy size limit).
'            text_write — write/replace/append a UTF-8 text file.
'            text_search — search for substring or regex; returns hits with indices.
'
' Architecture:
'  - All file I/O goes through PathPolicy.Resolve(...) for workspace boundary.
'  - Tools are deterministic and side-effect-free except text_write.
'  - Search supports regex mode and case-sensitivity toggles.
'  - Write supports overwrite/append/create_new modes.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace Agents

    Partial Public NotInheritable Class TextTools

        Private Sub New()
        End Sub

        Public Const ToolRead As String = "text_read"
        Public Const ToolWrite As String = "text_write"
        Public Const ToolSearch As String = "text_search"

        Public Shared Function IsTextTool(name As String) As Boolean
            If String.IsNullOrWhiteSpace(name) Then Return False
            Select Case name
                Case ToolRead, ToolWrite, ToolSearch
                    Return True
                Case Else
                    If IsAnalyzeTextTool(name) Then
                        Return True
                    End If
                    Return IsExtendedTextTool(name)
            End Select
        End Function

        Public Shared Function BuildAll() As List(Of ModelConfig)
            Dim tools As New List(Of ModelConfig) From {
                BuildRead(),
                BuildWrite(),
                BuildSearch()
            }

            tools.AddRange(BuildAnalyzeTools())
            tools.AddRange(BuildExtendedTools())
            Return tools
        End Function

        ' --------------------------------------------------------------- dispatch

        Public Shared Function Execute(toolName As String, arguments As IDictionary(Of String, Object)) As String
            Return ExecuteAsync(toolName, arguments, Nothing, CancellationToken.None).GetAwaiter().GetResult()
        End Function

        Public Shared Async Function ExecuteAsync(toolName As String,
                                                  arguments As IDictionary(Of String, Object),
                                                  context As ISharedContext,
                                                  Optional cancellationToken As CancellationToken = Nothing) As Task(Of String)
            Try
                Select Case toolName
                    Case ToolRead
                        Return ExecuteRead(arguments)

                    Case ToolWrite
                        Return ExecuteWrite(arguments)

                    Case ToolSearch
                        Return ExecuteSearch(arguments)

                    Case Else
                        Dim analyzeResult As String =
                            Await ExecuteAnalyzeAsync(toolName, arguments, context, cancellationToken).ConfigureAwait(False)

                        If analyzeResult IsNot Nothing Then
                            Return analyzeResult
                        End If

                        Dim extendedResult As String =
                            Await ExecuteExtendedAsync(toolName, arguments, context, cancellationToken).ConfigureAwait(False)

                        If extendedResult IsNot Nothing Then
                            Return extendedResult
                        End If

                        Return JsonConvert.SerializeObject(New With {
                            Key .error = "unknown_text_tool",
                            Key .tool = toolName
                        })
                End Select
            Catch uae As UnauthorizedAccessException
                Return JsonConvert.SerializeObject(New With {
                    Key .error = "access_denied",
                    Key .message = uae.Message
                })
            Catch oce As OperationCanceledException
                Return JsonConvert.SerializeObject(New With {
                    Key .error = "cancelled",
                    Key .message = "The operation was cancelled."
                })
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    Key .error = "text_tool_failed",
                    Key .message = ex.Message
                })
            End Try
        End Function

        Private Shared Function ExecuteRead(args As IDictionary(Of String, Object)) As String
            Dim p As System.String = PathPolicy.Resolve(GetStr(args, "path"), PathAccess.Read)
            Try
                Dim startChar As System.Int32 = 0
                Dim explicitWindow As System.Boolean = args IsNot Nothing AndAlso
                    (args.ContainsKey("start_char") OrElse args.ContainsKey("offset"))
                Dim startValue As System.Int32 = 0
                Dim offsetValue As System.Int32 = 0
                If Not TryReadTextOffset(args, "start_char", startValue) OrElse
                   Not TryReadTextOffset(args, "offset", offsetValue) Then
                    Return JsonConvert.SerializeObject(New With {Key .error = "invalid_offset", Key .message = "Offsets must be non-negative integers in UTF-16 code units."})
                End If
                If args IsNot Nothing AndAlso args.ContainsKey("start_char") AndAlso args.ContainsKey("offset") AndAlso startValue <> offsetValue Then
                    Return JsonConvert.SerializeObject(New With {Key .error = "conflicting_offsets", Key .message = "start_char and offset must agree."})
                End If
                startChar = If(args IsNot Nothing AndAlso args.ContainsKey("start_char"), startValue, offsetValue)

                Dim snapshot As TextFileSnapshot = TextFileSnapshot.Read(p, expectedSha256:=GetStr(args, "expected_snapshot_sha256"))
                Dim totalChars As System.Int32 = snapshot.Content.Length
                If startChar > totalChars Then
                    Return JsonConvert.SerializeObject(New With {Key .error = "offset_out_of_range", Key .total_chars = totalChars, Key .start_char = startChar})
                End If
                If explicitWindow AndAlso startChar > 0 AndAlso startChar < totalChars AndAlso
                   System.Char.IsLowSurrogate(snapshot.Content(startChar)) AndAlso System.Char.IsHighSurrogate(snapshot.Content(startChar - 1)) Then
                    Return JsonConvert.SerializeObject(New With {Key .error = "invalid_offset", Key .message = "Offset splits a Unicode surrogate pair. Use next_offset from the preceding window."})
                End If

                Dim maxChars As System.Int32 = GetInt(args, "max_chars", 0)
                Dim count As System.Int32 = totalChars - startChar
                If maxChars > 0 Then count = System.Math.Min(count, maxChars)
                ' New paged calls never split a valid surrogate pair. A window may contain
                ' max_chars + 1 code units for this purpose. Legacy prefix calls are unchanged.
                If explicitWindow AndAlso count > 0 AndAlso startChar + count < totalChars AndAlso
                   System.Char.IsHighSurrogate(snapshot.Content(startChar + count - 1)) AndAlso
                   System.Char.IsLowSurrogate(snapshot.Content(startChar + count)) Then count += 1

                Dim nextPosition As System.Int32 = startChar + count
                Dim hasMore As System.Boolean = nextPosition < totalChars
                Dim nextOffset As System.Nullable(Of System.Int32) = Nothing
                If hasMore Then nextOffset = nextPosition
                Return JsonConvert.SerializeObject(New With {
                    Key .path = p,
                    Key .size = snapshot.SizeBytes,
                    Key .truncated = startChar > 0 OrElse hasMore,
                    Key .text = snapshot.Content.Substring(startChar, count),
                    Key .total_chars = totalChars,
                    Key .returned_chars = count,
                    Key .start_char = startChar,
                    Key .next_offset = nextOffset,
                    Key .has_more = hasMore,
                    Key .offset_unit = "utf16_code_units",
                    Key .snapshot_sha256 = snapshot.Sha256
                })
            Catch ex As TextFileInputException
                Return JsonConvert.SerializeObject(New With {
                    Key .error = ex.ErrorCode, Key .path = p, Key .message = ex.Message,
                    Key .size = ex.SizeBytes, Key .max = PathPolicy.MaxFileSizeBytes
                })
            End Try
        End Function

        Private Shared Function TryReadTextOffset(args As System.Collections.Generic.IDictionary(Of System.String, System.Object),
                                                   name As System.String, ByRef value As System.Int32) As System.Boolean
            value = 0
            If args Is Nothing OrElse Not args.ContainsKey(name) Then Return True
            Dim raw As System.Object = args(name)
            If raw Is Nothing OrElse TypeOf raw Is System.Boolean Then Return False
            Return System.Int32.TryParse(System.Convert.ToString(raw, System.Globalization.CultureInfo.InvariantCulture),
                System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, value) AndAlso value >= 0
        End Function

        Private Shared Function ExecuteWrite(args As IDictionary(Of String, Object)) As String
            Dim rawPath = GetStr(args, "path")
            Dim mode = (GetStr(args, "mode")).ToLowerInvariant() ' "" / "overwrite" / "append" / "create_new"
            Dim text = GetStr(args, "text")
            If text Is Nothing Then text = ""

            Dim target As String
            If String.IsNullOrWhiteSpace(rawPath) Then
                target = PathPolicy.NewWritablePath(If(GetStr(args, "filename"), "agent_output.txt"))
            Else
                target = PathPolicy.Resolve(rawPath, PathAccess.Write)
            End If

            Dim artifactMetadata As OptionalToolArtifactMetadata = Nothing
            Dim artifactFailureCode As String = ""
            Dim artifactFailureMessage As String = ""

            If Not ArtifactDelivery.TryPrepareOptionalToolArtifactMetadata(
                args,
                ArtifactStorageKind.Unknown,
                artifactMetadata,
                artifactFailureCode,
                artifactFailureMessage) Then

                Return JsonConvert.SerializeObject(New With {
                    Key .error = artifactFailureCode,
                    Key .message = artifactFailureMessage
                })
            End If

            Dim dir = Path.GetDirectoryName(target)
            If Not String.IsNullOrWhiteSpace(dir) AndAlso Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)

            Select Case mode
                Case "append"
                    File.AppendAllText(target, text, Encoding.UTF8)
                Case "create_new"
                    If File.Exists(target) Then
                        Return JsonConvert.SerializeObject(New With {Key .error = "exists", Key .path = target})
                    End If
                    File.WriteAllText(target, text, Encoding.UTF8)
                Case Else ' "overwrite" or default
                    File.WriteAllText(target, text, Encoding.UTF8)
            End Select

            Dim fi As New FileInfo(target)

            ' Pick up edits to SKILL.md/AGENT.md immediately when writing into a resource root.
            AgentResources.RefreshIfResourcePath(target)

            If artifactMetadata Is Nothing Then
                Return JsonConvert.SerializeObject(New With {
                    Key .path = target,
                    Key .size = fi.Length,
                    Key .mode = If(String.IsNullOrWhiteSpace(mode), "overwrite", mode)
                })
            End If

            Return JsonConvert.SerializeObject(New With {
                Key .path = target,
                Key .size = fi.Length,
                Key .mode = If(String.IsNullOrWhiteSpace(mode), "overwrite", mode),
                Key .produces_user_deliverable = artifactMetadata.ProducesUserDeliverable,
                Key .produces_intermediate_data = artifactMetadata.ProducesIntermediateData,
                Key .artifacts = New System.Object() {artifactMetadata.BuildArtifact(target)}
            })
        End Function

        Private Shared Function ExecuteSearch(args As IDictionary(Of String, Object)) As String
            Dim p = PathPolicy.Resolve(GetStr(args, "path"), PathAccess.Read)
            If Not File.Exists(p) Then
                Return JsonConvert.SerializeObject(New With {Key .error = "not_found", Key .path = p})
            End If
            Dim fi As New FileInfo(p)
            If fi.Length > PathPolicy.MaxFileSizeBytes Then
                Return JsonConvert.SerializeObject(New With {Key .error = "file_too_large", Key .path = p})
            End If

            Dim text = File.ReadAllText(p, Encoding.UTF8)
            Dim query = GetStr(args, "query")
            If String.IsNullOrWhiteSpace(query) Then
                Return JsonConvert.SerializeObject(New With {Key .error = "missing_query"})
            End If

            Dim useRegex = GetBool(args, "regex", False)
            Dim ignoreCase = GetBool(args, "ignore_case", True)
            Dim maxHits = GetInt(args, "max_hits", 50)
            If maxHits < 1 Then maxHits = 1
            If maxHits > 500 Then maxHits = 500

            Dim hits As New List(Of Object)
            If useRegex Then
                Dim opt As RegexOptions = RegexOptions.CultureInvariant
                If ignoreCase Then opt = opt Or RegexOptions.IgnoreCase
                Dim rx As New Regex(query, opt, TimeSpan.FromSeconds(2))
                For Each m As Match In rx.Matches(text)
                    If hits.Count >= maxHits Then Exit For
                    hits.Add(BuildHit(text, m.Index, m.Length, m.Value))
                Next
            Else
                Dim cmp = If(ignoreCase, StringComparison.OrdinalIgnoreCase, StringComparison.Ordinal)
                Dim idx = 0
                While idx < text.Length
                    Dim found = text.IndexOf(query, idx, cmp)
                    If found < 0 Then Exit While
                    hits.Add(BuildHit(text, found, query.Length, text.Substring(found, query.Length)))
                    If hits.Count >= maxHits Then Exit While
                    idx = found + Math.Max(1, query.Length)
                End While
            End If

            Return JsonConvert.SerializeObject(New With {
                Key .path = p,
                Key .hits = hits,
                Key .total = hits.Count
            })
        End Function

        Private Shared Function BuildHit(text As String, index As Integer, length As Integer, match As String) As Object
            Dim winStart = Math.Max(0, index - 40)
            Dim winEnd = Math.Min(text.Length, index + length + 40)
            Dim ctx = text.Substring(winStart, winEnd - winStart).Replace(vbCr, " ").Replace(vbLf, " ")
            Return New With {
                Key .index = index,
                Key .length = length,
                Key .match = match,
                Key .context = ctx
            }
        End Function

        ' --------------------------------------------------------------- factories

        Private Shared Function BuildRead() As ModelConfig
            Dim def =
                "{""name"":""" & ToolRead & """," &
                """description"":""Read a text file under PathPolicy. Legacy calls without offsets read a prefix. For paging supply start_char=0 and max_chars, then use next_offset while has_more=true. Offsets/char counts are UTF-16 code units; size remains bytes. snapshot_sha256 hashes exact file bytes including BOM. truncated means the response omits some file content, not that extraction failed."",""parameters"":{" &
                """type"":""object""," &
                """properties"":{" &
                """path"":{""type"":""string"",""description"":""Absolute or workspace-relative path.""}," &
                """start_char"":{""type"":""integer"",""minimum"":0,""description"":""Zero-based UTF-16 offset. Omit for legacy prefix behavior; use 0 to begin safe paging.""}," &
                """offset"":{""type"":""integer"",""minimum"":0,""description"":""Alias of start_char. If both are supplied they must agree.""}," &
                """expected_snapshot_sha256"":{""type"":""string"",""description"":""Optional SHA-256 from the previous window/export. A changed file is rejected rather than splicing different snapshots.""}," &
                """max_chars"":{""type"":""integer"",""description"":""Optional cap (0 or negative = no cap, legacy behavior). Explicit offset windows may add one UTF-16 unit to avoid splitting a surrogate pair. next_offset is null at EOF.""}}," &
                """required"":[""path""]}}"
            Return New ModelConfig() With {
                .ToolName = ToolRead,
                .ToolDefinition = def,
                .ToolInstructionsPrompt = ToolRead & ": Read a text file under PathPolicy. For large files use start_char/max_chars and chain next_offset with expected_snapshot_sha256. Do not retype an existing full source into another tool when a file input is exposed.",
                .ModelDescription = "Text (read)",
                .Tool = True,
                .ToolPriority = 920,
                .ToolErrorHandling = "skip"
            }
        End Function

        Private Shared Function BuildWrite() As ModelConfig
            Dim def =
                "{""name"":""" & ToolWrite & """," &
                """description"":""Write a UTF-8 text file. If 'path' is omitted, a new file is created in the default writable root using 'filename' as a suggestion. The default writable root is the connected workspace when one is set; otherwise the current session's staging/working area, which is delivered to the user at the end of the run. Prefer relative paths or an omitted path so intermediate edits stay in the working area rather than a fixed location. Modes: overwrite (default), append, create_new."",""parameters"":{" &
                """type"":""object""," &
                """properties"":{" &
                """path"":{""type"":""string"",""description"":""Absolute path, or a path relative to the default writable root (connected workspace, otherwise the session staging/working area). Omit to auto-name in the default writable root.""}," &
                """filename"":{""type"":""string"",""description"":""Suggested filename when 'path' is omitted.""}," &
                """text"":{""type"":""string"",""description"":""Content to write.""}," &
                """mode"":{""type"":""string"",""enum"":[""overwrite"",""append"",""create_new""],""description"":""Write mode (default 'overwrite').""}," &
                """artifact_id"":{""type"":""string"",""description"":""Optional opaque artifact id. When any artifact metadata is supplied, artifact_id/logical_deliverable_id/output_slot_id/artifact_state/artifact_delivery_intent are required together.""}," &
                """logical_deliverable_id"":{""type"":""string""}," &
                """output_slot_id"":{""type"":""string""}," &
                """supersedes_artifact_id"":{""type"":""string""}," &
                """artifact_state"":{""type"":""string"",""enum"":[""working"",""intermediate"",""final""]}," &
                """artifact_delivery_intent"":{""type"":""string"",""enum"":[""none"",""deliver_to_user"",""persist_only"",""deliver_and_persist""]}," &
                """storage_kind"":{""type"":""string"",""enum"":[""session_staging"",""connected_workspace"",""host_managed"",""unknown""]}," &
                """expected_artifacts"":{""type"":""array"",""items"":{""type"":""object"",""properties"":{""logical_deliverable_id"":{""type"":""string""},""output_slot_id"":{""type"":""string""}},""required"":[""logical_deliverable_id"",""output_slot_id""]}}}," &
                """required"":[""text""]}}"
            Return New ModelConfig() With {
                .ToolName = ToolWrite,
                .ToolDefinition = def,
                .ToolInstructionsPrompt = ToolWrite & ": Write a UTF-8 text file (sandboxed by path policy).",
                .ModelDescription = "Text (write)",
                .Tool = True,
                .ToolPriority = 921,
                .ToolErrorHandling = "skip",
                .CapabilityTags = "artifact_generation"
            }
        End Function

        Private Shared Function BuildSearch() As ModelConfig
            Dim def =
                "{""name"":""" & ToolSearch & """," &
                """description"":""Search the contents of a UTF-8 text file for a string or regex. Returns up to max_hits matches with byte/char index and a 40-char context window."",""parameters"":{" &
                """type"":""object""," &
                """properties"":{" &
                """path"":{""type"":""string"",""description"":""Absolute or workspace-relative path.""}," &
                """query"":{""type"":""string"",""description"":""Literal substring (default) or regex (when regex=true).""}," &
                """regex"":{""type"":""boolean"",""description"":""Treat query as .NET regex (default false).""}," &
                """ignore_case"":{""type"":""boolean"",""description"":""Case-insensitive matching (default true).""}," &
                """max_hits"":{""type"":""integer"",""description"":""Max number of hits to return (default 50, capped at 500).""}}," &
                """required"":[""path"",""query""]}}"
            Return New ModelConfig() With {
                .ToolName = ToolSearch,
                .ToolDefinition = def,
                .ToolInstructionsPrompt = ToolSearch & ": Search a text file for a substring or regex.",
                .ModelDescription = "Text (search)",
                .Tool = True,
                .ToolPriority = 922,
                .ToolErrorHandling = "skip"
            }
        End Function

        ' --------------------------------------------------------------- argument helpers

        Private Shared Function GetStr(args As IDictionary(Of String, Object), name As String) As String
            If args Is Nothing Then Return ""
            Dim v As Object = Nothing
            If Not args.TryGetValue(name, v) OrElse v Is Nothing Then Return ""
            Return System.Convert.ToString(v)
        End Function

        Private Shared Function GetInt(args As IDictionary(Of String, Object), name As String, defaultValue As Integer) As Integer
            If args Is Nothing Then Return defaultValue
            Dim v As Object = Nothing
            If Not args.TryGetValue(name, v) OrElse v Is Nothing Then Return defaultValue
            Try
                Return System.Convert.ToInt32(v)
            Catch
                Dim n As Integer
                If Integer.TryParse(System.Convert.ToString(v), n) Then Return n
                Return defaultValue
            End Try
        End Function

        Private Shared Function GetBool(args As IDictionary(Of String, Object), name As String, defaultValue As Boolean) As Boolean
            If args Is Nothing Then Return defaultValue
            Dim v As Object = Nothing
            If Not args.TryGetValue(name, v) OrElse v Is Nothing Then Return defaultValue
            Try
                Return System.Convert.ToBoolean(v)
            Catch
                Dim s = System.Convert.ToString(v)
                Select Case s.Trim().ToLowerInvariant()
                    Case "true", "1", "yes" : Return True
                    Case "false", "0", "no" : Return False
                    Case Else : Return defaultValue
                End Select
            End Try
        End Function

    End Class

End Namespace
