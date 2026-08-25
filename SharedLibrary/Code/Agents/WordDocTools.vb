' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: WordDocTools.vb
' Purpose: Agent tools that address the currently open Word document(s) via the
'          IWordDocumentHost bridge. Read-only by default (WordHostPolicy);
'          write verbs registered only when ActiveDocReadOnly = False.
'
' Tools:
'  - worddoc_list_open: list all open documents.
'  - worddoc_get_active: metadata for the active document.
'  - worddoc_extract_text: return plain text of document.
'  - worddoc_search: substring/regex search within document.
'  - worddoc_list_comments: list all comments on document.
'  - (write verbs when not read-only): worddoc_insert_text, worddoc_replace,
'    worddoc_delete, worddoc_comment_add, worddoc_format.
' =============================================================================

Option Strict On
Option Explicit On

Imports Newtonsoft.Json
Imports SharedLibrary.SharedLibrary

Namespace Agents

    Public NotInheritable Class WordDocTools

        Private Sub New()
        End Sub

        Public Const ToolListOpen As String = "worddoc_list_open"
        Public Const ToolGetActive As String = "worddoc_get_active"
        Public Const ToolExtract As String = "worddoc_extract_text"
        Public Const ToolSearch As String = "worddoc_search"
        Public Const ToolListComments As String = "worddoc_list_comments"
        Public Const ToolInsert As String = "worddoc_insert_text"
        Public Const ToolReplace As String = "worddoc_replace"
        Public Const ToolDelete As String = "worddoc_delete"
        Public Const ToolCommentAdd As String = "worddoc_comment_add"
        Public Const ToolFormat As String = "worddoc_format"

        Public Shared Function IsWordDocTool(name As String) As Boolean
            If String.IsNullOrWhiteSpace(name) Then Return False
            Select Case name
                Case ToolListOpen, ToolGetActive, ToolExtract, ToolSearch, ToolListComments,
                     ToolInsert, ToolReplace, ToolDelete, ToolCommentAdd, ToolFormat
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        ''' <summary>Read-only + (conditionally) write verbs.</summary>
        Public Shared Function BuildAll() As List(Of ModelConfig)
            Dim list As New List(Of ModelConfig) From {
                BuildListOpen(), BuildGetActive(), BuildExtract(), BuildSearch(), BuildListComments()
            }
            If Not WordHostPolicy.ActiveDocReadOnly Then
                list.Add(BuildInsert())
                list.Add(BuildReplace())
                list.Add(BuildDelete())
                list.Add(BuildCommentAdd())
                list.Add(BuildFormat())
            End If
            Return list
        End Function

        ' --------------------------------------------------------------- dispatch

        Public Shared Function Execute(toolName As String, arguments As IDictionary(Of String, Object)) As String
            Try
                Dim host = WordHostPolicy.Host
                If host Is Nothing OrElse Not host.HasActiveDocument() Then
                    Return Err_("no_active_document", "No Word document is available in this host.")
                End If

                Dim target = GetStr(arguments, "target")
                If String.Equals(target, "active", StringComparison.OrdinalIgnoreCase) Then target = ""

                Select Case toolName
                    Case ToolListOpen
                        Dim list = host.ListOpenDocuments()
                        Return JsonConvert.SerializeObject(New With {Key .documents = list})
                    Case ToolGetActive
                        Dim info = host.GetActiveDocument()
                        Return JsonConvert.SerializeObject(info)
                    Case ToolExtract
                        Dim text = host.ExtractText(target, GetInt(arguments, "max_chars", 0))
                        Return JsonConvert.SerializeObject(New With {Key .target = If(String.IsNullOrEmpty(target), "active", target), Key .text = text})
                    Case ToolSearch
                        Dim query = GetStr(arguments, "query")
                        If String.IsNullOrWhiteSpace(query) Then Return Err_("missing_query", "query is required.")
                        Return host.SearchJson(target, query, GetBool(arguments, "regex", False),
                                                GetBool(arguments, "ignore_case", True),
                                                GetInt(arguments, "max_hits", 50))
                    Case ToolListComments
                        Dim comments = host.ListComments(target)
                        Return JsonConvert.SerializeObject(New With {Key .target = If(String.IsNullOrEmpty(target), "active", target), Key .comments = comments})

                    Case ToolInsert, ToolReplace, ToolDelete, ToolCommentAdd, ToolFormat
                        If WordHostPolicy.ActiveDocReadOnly Then
                            Return Err_("read_only", "Active document is read-only for the agent. Enable writes in Red Ink settings to allow this.")
                        End If
                        Select Case toolName
                            Case ToolInsert
                                Return host.InsertTextJson(target, GetStr(arguments, "text"), If(GetStr(arguments, "location"), "end"),
                                                            GetBool(arguments, "track_changes", True))
                            Case ToolReplace
                                Return host.ReplaceJson(target, GetStr(arguments, "find"), GetStr(arguments, "text"),
                                                         GetBool(arguments, "only_first", True),
                                                         GetBool(arguments, "track_changes", True),
                                                         GetStr(arguments, "match_scope"))
                            Case ToolDelete
                                Return host.DeleteJson(target, GetStr(arguments, "find"),
                                                        GetBool(arguments, "track_changes", True),
                                                        GetStr(arguments, "match_scope"))
                            Case ToolCommentAdd
                                Return host.AddCommentJson(target, GetStr(arguments, "find"), GetStr(arguments, "text"),
                                                            If(GetStr(arguments, "author"), "Red Ink"),
                                                            If(GetStr(arguments, "initials"), "RI"))
                            Case ToolFormat
                                Return host.FormatJson(target,
                                                       GetStr(arguments, "find"),
                                                       GetStr(arguments, "style"),
                                                       GetNullableBool(arguments, "bold"),
                                                       GetNullableBool(arguments, "italic"),
                                                       GetNullableBool(arguments, "underline"),
                                                       GetInt(arguments, "size", 0),
                                                       GetStr(arguments, "color"),
                                                       GetStr(arguments, "align"))
                        End Select

                    Case Else
                        Return Err_("unknown_worddoc_tool", "Unknown tool '" & toolName & "'.")
                End Select
            Catch ex As Exception
                Return Err_("worddoc_tool_failed", ex.Message)
            End Try
            Return Err_("unknown_worddoc_tool", "Unreachable.")
        End Function

        ' --------------------------------------------------------------- factories

        Private Shared Function BuildListOpen() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolListOpen, .Tool = True, .ToolPriority = 870, .ToolErrorHandling = "skip",
                .ModelDescription = "Word doc (list open)",
                .ToolDefinition = "{""name"":""" & ToolListOpen & """,""description"":""List documents currently open in Word. Returns name, path, isActive, isReadOnly, isSaved."",""parameters"":{""type"":""object"",""properties"":{}}}",
                .ToolInstructionsPrompt = ToolListOpen & ": Discover documents currently open in Word."
            }
        End Function

        Private Shared Function BuildGetActive() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolGetActive, .Tool = True, .ToolPriority = 871, .ToolErrorHandling = "skip",
                .ModelDescription = "Word doc (active info)",
                .ToolDefinition = "{""name"":""" & ToolGetActive & """,""description"":""Return metadata of the active Word document."",""parameters"":{""type"":""object"",""properties"":{}}}",
                .ToolInstructionsPrompt = ToolGetActive & ": Get metadata of the active Word document."
            }
        End Function

        Private Shared Function BuildExtract() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolExtract, .Tool = True, .ToolPriority = 872, .ToolErrorHandling = "skip",
                .ModelDescription = "Word doc (extract text)",
                .ToolDefinition = "{""name"":""" & ToolExtract & """,""description"":""Extract plain text from the active or named open Word document."",""parameters"":{""type"":""object"",""properties"":{""target"":{""type"":""string"",""description"":""Empty or 'active' for the active doc; otherwise a document name or full path that is currently open in Word.""},""max_chars"":{""type"":""integer""}}}}",
                .ToolInstructionsPrompt = ToolExtract & ": Extract text from a currently open Word document."
            }
        End Function

        Private Shared Function BuildSearch() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolSearch, .Tool = True, .ToolPriority = 873, .ToolErrorHandling = "skip",
                .ModelDescription = "Word doc (search)",
                .ToolDefinition = "{""name"":""" & ToolSearch & """,""description"":""Search a currently open Word document for a substring or regex."",""parameters"":{""type"":""object"",""properties"":{""target"":{""type"":""string""},""query"":{""type"":""string""},""regex"":{""type"":""boolean""},""ignore_case"":{""type"":""boolean""},""max_hits"":{""type"":""integer""}},""required"":[""query""]}}",
                .ToolInstructionsPrompt = ToolSearch & ": Search the currently open Word document."
            }
        End Function

        Private Shared Function BuildListComments() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolListComments, .Tool = True, .ToolPriority = 874, .ToolErrorHandling = "skip",
                .ModelDescription = "Word doc (comments list)",
                .ToolDefinition = "{""name"":""" & ToolListComments & """,""description"":""List comments on the active or named open Word document."",""parameters"":{""type"":""object"",""properties"":{""target"":{""type"":""string""}}}}",
                .ToolInstructionsPrompt = ToolListComments & ": List comments on a currently open Word document."
            }
        End Function

        Private Shared Function BuildInsert() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolInsert, .Tool = True, .ToolPriority = 875, .ToolErrorHandling = "skip",
                .ModelDescription = "Word doc (insert text)",
                .ToolDefinition = "{""name"":""" & ToolInsert & """,""description"":""Insert text into the active or named open document. location: 'start' | 'end' | 'cursor' (default 'end'). track_changes defaults to true (insertion recorded as a tracked revision)."",""parameters"":{""type"":""object"",""properties"":{""target"":{""type"":""string""},""text"":{""type"":""string""},""location"":{""type"":""string"",""enum"":[""start"",""end"",""cursor""]},""track_changes"":{""type"":""boolean""}},""required"":[""text""]}}",
                .ToolInstructionsPrompt = ToolInsert & ": Insert text into the open Word document."
            }
        End Function

        Private Shared Function BuildReplace() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolReplace, .Tool = True, .ToolPriority = 876, .ToolErrorHandling = "skip",
                .ModelDescription = "Word doc (replace)",
                .ToolDefinition = "{""name"":""" & ToolReplace & """,""description"":""Replace the resilient-located 'find' span in the open document. Matching tolerates stray whitespace, paragraph marks, fields, footnote/endnote anchors and control characters, and works with anchors longer than 255 characters. track_changes defaults to true (tracked, footnote/field/format-safe surgical replacement). match_scope: '' (exact span, default) | 'sentence' | 'paragraph' expands the located span to the enclosing sentence/paragraph."",""parameters"":{""type"":""object"",""properties"":{""target"":{""type"":""string""},""find"":{""type"":""string""},""text"":{""type"":""string""},""only_first"":{""type"":""boolean""},""track_changes"":{""type"":""boolean""},""match_scope"":{""type"":""string"",""enum"":["""",""sentence"",""paragraph""]}},""required"":[""find"",""text""]}}",
                .ToolInstructionsPrompt = ToolReplace & ": Replace a resilient-located span in the open Word document (tracked by default)."
            }
        End Function

        Private Shared Function BuildDelete() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolDelete, .Tool = True, .ToolPriority = 876, .ToolErrorHandling = "skip",
                .ModelDescription = "Word doc (delete)",
                .ToolDefinition = "{""name"":""" & ToolDelete & """,""description"":""Delete the resilient-located 'find' span in the open document. Matching tolerates stray whitespace, paragraph marks, fields, footnote/endnote anchors and control characters, and works with anchors longer than 255 characters. track_changes defaults to true (tracked, footnote/field/format-safe surgical deletion). match_scope: '' (exact span, default) | 'sentence' | 'paragraph' deletes the enclosing sentence/paragraph."",""parameters"":{""type"":""object"",""properties"":{""target"":{""type"":""string""},""find"":{""type"":""string""},""track_changes"":{""type"":""boolean""},""match_scope"":{""type"":""string"",""enum"":["""",""sentence"",""paragraph""]}},""required"":[""find""]}}",
                .ToolInstructionsPrompt = ToolDelete & ": Delete a resilient-located word/sentence/paragraph in the open Word document (tracked by default)."
            }
        End Function
        Private Shared Function BuildCommentAdd() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolCommentAdd, .Tool = True, .ToolPriority = 877, .ToolErrorHandling = "skip",
                .ModelDescription = "Word doc (comment add)",
                .ToolDefinition = "{""name"":""" & ToolCommentAdd & """,""description"":""Add a Word comment anchored on a 'find' match in the open document."",""parameters"":{""type"":""object"",""properties"":{""target"":{""type"":""string""},""find"":{""type"":""string""},""text"":{""type"":""string""},""author"":{""type"":""string""},""initials"":{""type"":""string""}},""required"":[""find"",""text""]}}",
                .ToolInstructionsPrompt = ToolCommentAdd & ": Add a Word comment in the open document."
            }
        End Function

        Private Shared Function BuildFormat() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolFormat, .Tool = True, .ToolPriority = 878, .ToolErrorHandling = "skip",
                .ModelDescription = "Word doc (format)",
                .ToolDefinition = "{""name"":""" & ToolFormat & """,""description"":""Apply paragraph/run formatting on a 'find' match in the open document. Same fields as word_format."",""parameters"":{""type"":""object"",""properties"":{""target"":{""type"":""string""},""find"":{""type"":""string""},""style"":{""type"":""string""},""bold"":{""type"":""boolean""},""italic"":{""type"":""boolean""},""underline"":{""type"":""boolean""},""size"":{""type"":""integer""},""color"":{""type"":""string""},""align"":{""type"":""string""}},""required"":[""find""]}}",
                .ToolInstructionsPrompt = ToolFormat & ": Format a matched span in the open Word document."
            }
        End Function

        ' --------------------------------------------------------------- helpers

        Private Shared Function Err_(code As String, message As String) As String
            Return JsonConvert.SerializeObject(New With {Key .error = code, Key .message = message})
        End Function

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
            Try : Return System.Convert.ToInt32(v) : Catch
                Dim n As Integer
                If Integer.TryParse(System.Convert.ToString(v), n) Then Return n
                Return defaultValue
            End Try
        End Function

        Private Shared Function GetBool(args As IDictionary(Of String, Object), name As String, defaultValue As Boolean) As Boolean
            Dim nb = GetNullableBool(args, name)
            If nb.HasValue Then Return nb.Value
            Return defaultValue
        End Function

        Private Shared Function GetNullableBool(args As IDictionary(Of String, Object), name As String) As Boolean?
            If args Is Nothing Then Return Nothing
            Dim v As Object = Nothing
            If Not args.TryGetValue(name, v) OrElse v Is Nothing Then Return Nothing
            Try : Return System.Convert.ToBoolean(v) : Catch
                Select Case System.Convert.ToString(v).Trim().ToLowerInvariant()
                    Case "true", "1", "yes" : Return True
                    Case "false", "0", "no" : Return False
                    Case Else : Return Nothing
                End Select
            End Try
        End Function

    End Class

End Namespace
