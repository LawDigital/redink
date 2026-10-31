' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: MemoryTools.vb
' Purpose: ModelConfig factories and dispatcher for the session-memory tools.
'          The actual execution (mapping ToolCall → SessionMemory) is performed
'          by Execute(), which existing tooling loops route to when they encounter
'          a memory_* tool call.
'
' Tools:
'  - memory_put: Write a key/value pair with summary, tags, and metadata.
'  - memory_get: Retrieve a stored entry by key.
'  - memory_list: Enumerate all entries with summaries.
'  - memory_delete: Remove an entry by key.
'  - All use SessionMemory for persistence across restarts.
'
' KEY ARCHITECTURE
'   • Surface: MemoryTools (VB implementation).
'   • Entry points/helpers: BuildAll, IsMemoryTool, Execute.
' =============================================================================

Option Strict On
Option Explicit On

Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Agents

    Public NotInheritable Class MemoryTools

        Private Sub New()
        End Sub

        Public Const ToolPut As String = "memory_put"
        Public Const ToolGet As String = "memory_get"
        Public Const ToolList As String = "memory_list"
        Public Const ToolDelete As String = "memory_delete"

        ''' <summary>Returns the four memory tools as ModelConfig entries.</summary>
        Public Shared Function BuildAll() As List(Of SharedLibrary.ModelConfig)
            Return New List(Of SharedLibrary.ModelConfig) From {
                BuildPut(), BuildGet(), BuildList(), BuildDelete()
            }
        End Function

        Public Shared Function IsMemoryTool(name As String) As Boolean
            If String.IsNullOrWhiteSpace(name) Then Return False
            Select Case name
                Case ToolPut, ToolGet, ToolList, ToolDelete : Return True
                Case Else : Return False
            End Select
        End Function

        ''' <summary>
        ''' Dispatches a memory tool call. Caller passes the parsed ToolCall.Arguments
        ''' (the same dictionary the existing tooling loops already construct). The
        ''' result is a JSON string suitable for direct use as the tool response.
        ''' </summary>
        Public Shared Function Execute(toolName As String, arguments As IDictionary(Of String, Object)) As String
            Try
                Select Case toolName
                    Case ToolPut
                        Dim key = GetStr(arguments, "key")
                        Dim summary = GetStr(arguments, "summary")
                        Dim valueTok = GetToken(arguments, "value")
                        Dim tags = GetStringList(arguments, "tags")

                        Dim metadata As New SessionMemoryMetadata With {
                            .WorkflowId = CoalesceNonEmpty(GetStr(arguments, "workflowId"), WorkflowContinuity.CurrentWorkflowId),
                            .Source = CoalesceNonEmpty(GetStr(arguments, "source"), "model"),
                            .ContentKind = GetStr(arguments, "contentKind"),
                            .RelatedTool = GetStr(arguments, "relatedTool"),
                            .RelatedAgent = GetStr(arguments, "relatedAgent"),
                            .RelatedSkill = GetStr(arguments, "relatedSkill"),
                            .TrustLevel = GetStr(arguments, "trustLevel"),
                            .TrustedForRuntime = GetBool(arguments, "trustedForRuntime", False),
                            .CreatedAt = DateTime.UtcNow
                        }

                        Dim entry = SessionMemory.Put(key, summary, valueTok, tags, metadata)

                        Return JsonConvert.SerializeObject(New With {
                            Key .key = entry.Key,
                            Key .summary = entry.Summary,
                            Key .stub = SessionMemory.BuildStub(entry),
                            Key .metadata = entry.Metadata
                        })

                    Case ToolGet
                        Dim key = GetStr(arguments, "key")
                        Dim e = SessionMemory.Get(key)
                        If e Is Nothing Then
                            Return JsonConvert.SerializeObject(New With {Key .error = "not_found", Key .key = key})
                        End If
                        Return JsonConvert.SerializeObject(New With {
                            Key .key = e.Key,
                            Key .summary = e.Summary,
                            Key .value = e.Value,
                            Key .createdAt = e.CreatedAt,
                            Key .updatedAt = e.UpdatedAt,
                            Key .tags = e.Tags,
                            Key .metadata = e.Metadata
                        })

                    Case ToolList
                        Dim items = SessionMemory.List().
                            Select(Function(e) New With {
                                Key .key = e.Key,
                                Key .summary = e.Summary,
                                Key .createdAt = e.CreatedAt,
                                Key .updatedAt = e.UpdatedAt,
                                Key .tags = e.Tags,
                                Key .metadata = e.Metadata
                            }).
                            ToList()
                        Return JsonConvert.SerializeObject(items)

                    Case ToolGet
                        Dim key = GetStr(arguments, "key")
                        Dim e = SessionMemory.Get(key)
                        If e Is Nothing Then
                            Return JsonConvert.SerializeObject(New With {Key .error = "not_found", Key .key = key})
                        End If
                        Return JsonConvert.SerializeObject(New With {
                            Key .key = e.Key,
                            Key .summary = e.Summary,
                            Key .value = e.Value,
                            Key .createdAt = e.CreatedAt,
                            Key .tags = e.Tags
                        })

                    Case ToolList
                        Dim items = SessionMemory.List().
                            Select(Function(e) New With {
                                Key .key = e.Key,
                                Key .summary = e.Summary,
                                Key .createdAt = e.CreatedAt,
                                Key .tags = e.Tags
                            }).
                            ToList()
                        Return JsonConvert.SerializeObject(items)

                    Case ToolDelete
                        Dim key = GetStr(arguments, "key")
                        Dim ok = SessionMemory.Delete(key)
                        Return JsonConvert.SerializeObject(New With {Key .key = key, Key .deleted = ok})

                    Case Else
                        Return JsonConvert.SerializeObject(New With {Key .error = "unknown_memory_tool", Key .tool = toolName})
                End Select
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {Key .error = "memory_tool_failed", Key .message = ex.Message})
            End Try
        End Function

        ' --------------------------------------------------------------- factories

        ' PATCH 2: replace BuildPut() with this version.

        Private Shared Function BuildPut() As SharedLibrary.ModelConfig
            Dim def =
        "{""name"":""" & ToolPut & """," &
        """description"":""Store a value in the persistent session memory. Use this to offload large or future-only intermediate results so the running context stays lean. Returns a stub string ('[memory:KEY] SUMMARY') that you may reference later via memory_get. Workflow-related metadata is optional and advisory unless written by the host."",""parameters"":{" &
        """type"":""object""," &
        """properties"":{" &
        """key"":{""type"":""string"",""description"":""Optional stable identifier. If omitted, a fresh id is generated.""}," &
        """summary"":{""type"":""string"",""description"":""One-line description of what is stored (shown as stub in later turns).""}," &
        """value"":{""description"":""The value to store. Any JSON value (string, number, object, array).""}," &
        """tags"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Optional tags.""}," &
        """workflowId"":{""type"":""string"",""description"":""Optional workflow identifier. If omitted during an active workflow, the host may supply it automatically.""}," &
        """source"":{""type"":""string"",""description"":""Optional source classifier: host, tool, agent, model, user.""}," &
        """contentKind"":{""type"":""string"",""description"":""Optional content kind: runtime_state, tool_result, source_record, note, summary, draft, unknown.""}," &
        """relatedTool"":{""type"":""string"",""description"":""Optional related tool name.""}," &
        """relatedAgent"":{""type"":""string"",""description"":""Optional related agent name.""}," &
        """relatedSkill"":{""type"":""string"",""description"":""Optional related skill name.""}," &
        """trustLevel"":{""type"":""string"",""description"":""Optional trust label such as advisory or authoritative.""}," &
        """trustedForRuntime"":{""type"":""boolean"",""description"":""Optional flag marking the entry as trusted for runtime use.""}}," &
        """required"":[""summary"",""value""]}}"

            Return New SharedLibrary.ModelConfig() With {
        .ToolName = ToolPut,
        .ToolDefinition = def,
        .ToolInstructionsPrompt = ToolPut & ": Store data in session memory; receive a stub you can reference later via memory_get.",
        .ModelDescription = "Session Memory (store)",
        .Tool = True,
        .ToolPriority = 950,
        .ToolErrorHandling = "skip"
    }
        End Function

        Private Shared Function BuildGet() As SharedLibrary.ModelConfig
            Dim def =
                "{""name"":""" & ToolGet & """," &
                """description"":""Retrieve a previously stored value from session memory by its key."",""parameters"":{" &
                """type"":""object""," &
                """properties"":{""key"":{""type"":""string"",""description"":""The memory key.""}}," &
                """required"":[""key""]}}"
            Return New SharedLibrary.ModelConfig() With {
                .ToolName = ToolGet,
                .ToolDefinition = def,
                .ToolInstructionsPrompt = ToolGet & ": Retrieve a stored value from session memory by key.",
                .ModelDescription = "Session Memory (read)",
                .Tool = True,
                .ToolPriority = 951,
                .ToolErrorHandling = "skip"
            }
        End Function

        Private Shared Function BuildList() As SharedLibrary.ModelConfig
            Dim def =
                "{""name"":""" & ToolList & """," &
                """description"":""List all session-memory entries with key, summary, createdAt and tags."",""parameters"":{""type"":""object"",""properties"":{}}}"
            Return New SharedLibrary.ModelConfig() With {
                .ToolName = ToolList,
                .ToolDefinition = def,
                .ToolInstructionsPrompt = ToolList & ": List entries currently in session memory.",
                .ModelDescription = "Session Memory (list)",
                .Tool = True,
                .ToolPriority = 952,
                .ToolErrorHandling = "skip"
            }
        End Function

        Private Shared Function BuildDelete() As SharedLibrary.ModelConfig
            Dim def =
                "{""name"":""" & ToolDelete & """," &
                """description"":""Delete a session-memory entry by key."",""parameters"":{" &
                """type"":""object""," &
                """properties"":{""key"":{""type"":""string"",""description"":""The memory key to remove.""}}," &
                """required"":[""key""]}}"
            Return New SharedLibrary.ModelConfig() With {
                .ToolName = ToolDelete,
                .ToolDefinition = def,
                .ToolInstructionsPrompt = ToolDelete & ": Remove a session-memory entry.",
                .ModelDescription = "Session Memory (delete)",
                .Tool = True,
                .ToolPriority = 953,
                .ToolErrorHandling = "skip"
            }
        End Function

        ' --------------------------------------------------------------- argument helpers

        Private Shared Function GetBool(args As IDictionary(Of String, Object), name As String, defaultValue As Boolean) As Boolean
            If args Is Nothing Then Return defaultValue

            Dim v As Object = Nothing
            If Not args.TryGetValue(name, v) OrElse v Is Nothing Then
                Return defaultValue
            End If

            Try
                Return Convert.ToBoolean(v)
            Catch
                Dim s As String = Convert.ToString(v).Trim()

                Select Case s.ToLowerInvariant()
                    Case "true", "1", "yes", "y", "on"
                        Return True
                    Case "false", "0", "no", "n", "off"
                        Return False
                    Case Else
                        Return defaultValue
                End Select
            End Try
        End Function

        Private Shared Function CoalesceNonEmpty(ParamArray values() As String) As String
            If values Is Nothing Then Return ""

            For Each value In values
                If Not String.IsNullOrWhiteSpace(value) Then
                    Return value.Trim()
                End If
            Next

            Return ""
        End Function


        Private Shared Function GetStr(args As IDictionary(Of String, Object), name As String) As String
            If args Is Nothing Then Return ""
            Dim v As Object = Nothing
            If Not args.TryGetValue(name, v) OrElse v Is Nothing Then Return ""
            Return System.Convert.ToString(v)
        End Function

        Private Shared Function GetToken(args As IDictionary(Of String, Object), name As String) As JToken
            If args Is Nothing Then Return JValue.CreateNull()
            Dim v As Object = Nothing
            If Not args.TryGetValue(name, v) OrElse v Is Nothing Then Return JValue.CreateNull()
            If TypeOf v Is JToken Then Return CType(v, JToken)
            Try
                Return JToken.FromObject(v)
            Catch
                Return New JValue(System.Convert.ToString(v))
            End Try
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

    End Class

End Namespace