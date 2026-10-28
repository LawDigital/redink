' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' Host-independent original-request log contract. The first three lines of a new
' tooling log contain the complete, immutable host-supplied request. Model/tool
' output can never replace this header or supply its recipient metadata.

Option Explicit On
Option Strict On

Namespace Agents

    Public NotInheritable Class ToolingOriginalRequestLog

        Public Const BeginMarker As System.String = "[INKY_ORIGINAL_REQUEST_V1]"
        Public Const EndMarker As System.String = "[/INKY_ORIGINAL_REQUEST_V1]"

        Private Sub New()
        End Sub

        ''' <summary>
        ''' Immutable host input, captured before prompt enrichment, compaction, retries
        ''' or workflow restoration. A follow-up run has its own unchanged input and
        ''' workflow reference; nested sub-agents retain the parent header.
        ''' SenderEmail is descriptive evidence only, never a delivery authorization.
        ''' </summary>
        Public NotInheritable Class Record
            Public ReadOnly Property RequestText As System.String
            Public ReadOnly Property Subject As System.String
            Public ReadOnly Property SenderEmail As System.String
            Public ReadOnly Property Source As System.String
            Public ReadOnly Property MessageId As System.String
            Public ReadOnly Property WorkflowId As System.String
            Public ReadOnly Property IsWorkflowResume As System.Boolean
            Public ReadOnly Property RecordedUtc As System.DateTime

            Public Sub New(requestText As System.String,
                           Optional subject As System.String = "",
                           Optional senderEmail As System.String = "",
                           Optional source As System.String = "host_input",
                           Optional messageId As System.String = "",
                           Optional workflowId As System.String = "",
                           Optional isWorkflowResume As System.Boolean = False,
                           Optional recordedUtc As System.Nullable(Of System.DateTime) = Nothing)
                Me.RequestText = If(requestText, System.String.Empty)
                Me.Subject = If(subject, System.String.Empty)
                Me.SenderEmail = If(senderEmail, System.String.Empty)
                Me.Source = If(source, System.String.Empty)
                Me.MessageId = If(messageId, System.String.Empty)
                Me.WorkflowId = If(workflowId, System.String.Empty)
                Me.IsWorkflowResume = isWorkflowResume
                Me.RecordedUtc = If(recordedUtc.HasValue, recordedUtc.GetValueOrDefault().ToUniversalTime(), System.DateTime.UtcNow)
            End Sub
        End Class

        ''' <summary>
        ''' Keeps the selected caller input exactly as received, including whitespace
        ''' and dialog history. AutoPilot supplies a Record directly from its mail or
        ''' transcribed voicemail instead of using this already assembled prompt.
        ''' </summary>
        Public Shared Function CaptureHostInput(userText As System.String,
                                                otherPrompt As System.String,
                                                fullPromptOverride As System.String,
                                                workflowId As System.String,
                                                isWorkflowResume As System.Boolean) As Record
            If Not System.String.IsNullOrWhiteSpace(fullPromptOverride) Then
                Return New Record(fullPromptOverride, source:="host_full_prompt", workflowId:=workflowId, isWorkflowResume:=isWorkflowResume)
            End If
            If Not System.String.IsNullOrWhiteSpace(userText) Then
                Return New Record(userText, source:="host_user_text", workflowId:=workflowId, isWorkflowResume:=isWorkflowResume)
            End If
            Return New Record(If(otherPrompt, System.String.Empty), source:="host_other_prompt", workflowId:=workflowId, isWorkflowResume:=isWorkflowResume)
        End Function

        ''' <summary>
        ''' Returns a single JSON payload line, bounded by fixed first-line markers.
        ''' JSON escaping preserves the request without permitting forged log lines.
        ''' No truncation, trimming, summarization or UI logging is performed here.
        ''' </summary>
        Public Shared Function BuildHeader(originalRequest As Record) As System.String
            Dim snapshot As Record = If(originalRequest, New Record(System.String.Empty, source:="host_request_unavailable"))
            Dim payload As New Newtonsoft.Json.Linq.JObject From {
                {"schema", 1},
                {"request_text", snapshot.RequestText},
                {"request_chars", snapshot.RequestText.Length},
                {"request_sha256", ComputeRequestSha256(snapshot.RequestText)},
                {"subject", snapshot.Subject},
                {"sender_email", snapshot.SenderEmail},
                {"source", snapshot.Source},
                {"message_id", snapshot.MessageId},
                {"workflow_id", snapshot.WorkflowId},
                {"is_workflow_resume", snapshot.IsWorkflowResume},
                {"recorded_utc", snapshot.RecordedUtc.ToString("o", System.Globalization.CultureInfo.InvariantCulture)}
            }
            Dim settings As New Newtonsoft.Json.JsonSerializerSettings With {
                .Formatting = Newtonsoft.Json.Formatting.None,
                .StringEscapeHandling = Newtonsoft.Json.StringEscapeHandling.EscapeNonAscii
            }
            Return BeginMarker & System.Environment.NewLine &
                Newtonsoft.Json.JsonConvert.SerializeObject(payload, settings) & System.Environment.NewLine &
                EndMarker & System.Environment.NewLine
        End Function

        ''' <summary>
        ''' Reads only the fixed beginning of an individual tooling log. A marker
        ''' found later in model, tool, quoted mail or concatenated bundle content
        ''' is never accepted. Missing/legacy, malformed or truncated data fails
        ''' closed; callers must not infer recipients from log text or filenames.
        ''' The checksum detects accidental alteration; it is not authentication.
        ''' </summary>
        Public Shared Function TryReadHeader(logText As System.String, ByRef originalRequest As Record) As System.Boolean
            originalRequest = Nothing
            If System.String.IsNullOrEmpty(logText) Then Return False
            Try
                Using reader As New System.IO.StringReader(logText)
                    Dim firstLine As System.String = reader.ReadLine()
                    If firstLine IsNot Nothing AndAlso firstLine.Length > 0 AndAlso firstLine(0) = Microsoft.VisualBasic.ChrW(&HFEFF) Then
                        firstLine = firstLine.Substring(1)
                    End If
                    If Not System.String.Equals(firstLine, BeginMarker, System.StringComparison.Ordinal) Then Return False
                    Dim jsonLine As System.String = reader.ReadLine()
                    If System.String.IsNullOrEmpty(jsonLine) OrElse
                       Not System.String.Equals(reader.ReadLine(), EndMarker, System.StringComparison.Ordinal) Then Return False

                    Dim payload As Newtonsoft.Json.Linq.JObject
                    Using jsonInput As New System.IO.StringReader(jsonLine)
                        Using jsonReader As New Newtonsoft.Json.JsonTextReader(jsonInput)
                            jsonReader.DateParseHandling = Newtonsoft.Json.DateParseHandling.None
                            jsonReader.MaxDepth = 4
                            payload = Newtonsoft.Json.Linq.JObject.Load(jsonReader, New Newtonsoft.Json.Linq.JsonLoadSettings With {
                                .DuplicatePropertyNameHandling = Newtonsoft.Json.Linq.DuplicatePropertyNameHandling.Error
                            })
                            If jsonReader.Read() Then Return False
                        End Using
                    End Using

                    If payload("schema") Is Nothing OrElse payload("schema").Type <> Newtonsoft.Json.Linq.JTokenType.Integer OrElse
                       payload.Value(Of System.Int32)("schema") <> 1 Then Return False
                    For Each key As System.String In New System.String() {
                        "request_text", "request_sha256", "subject", "sender_email", "source", "message_id", "workflow_id", "recorded_utc"
                    }
                        If payload(key) Is Nothing OrElse payload(key).Type <> Newtonsoft.Json.Linq.JTokenType.String Then Return False
                    Next
                    If payload("request_chars") Is Nothing OrElse payload("request_chars").Type <> Newtonsoft.Json.Linq.JTokenType.Integer OrElse
                       payload("is_workflow_resume") Is Nothing OrElse payload("is_workflow_resume").Type <> Newtonsoft.Json.Linq.JTokenType.Boolean Then Return False

                    Dim requestText As System.String = payload.Value(Of System.String)("request_text")
                    If payload.Value(Of System.Int64)("request_chars") <> requestText.Length OrElse
                       Not System.String.Equals(payload.Value(Of System.String)("request_sha256"), ComputeRequestSha256(requestText), System.StringComparison.Ordinal) Then Return False
                    Dim recordedUtc As System.DateTime
                    If Not System.DateTime.TryParseExact(payload.Value(Of System.String)("recorded_utc"), "o",
                                                       System.Globalization.CultureInfo.InvariantCulture,
                                                       System.Globalization.DateTimeStyles.RoundtripKind, recordedUtc) OrElse
                       recordedUtc.Kind <> System.DateTimeKind.Utc Then Return False

                    originalRequest = New Record(requestText,
                                                 subject:=payload.Value(Of System.String)("subject"),
                                                 senderEmail:=payload.Value(Of System.String)("sender_email"),
                                                 source:=payload.Value(Of System.String)("source"),
                                                 messageId:=payload.Value(Of System.String)("message_id"),
                                                 workflowId:=payload.Value(Of System.String)("workflow_id"),
                                                 isWorkflowResume:=payload.Value(Of System.Boolean)("is_workflow_resume"),
                                                 recordedUtc:=recordedUtc)
                    Return True
                End Using
            Catch ex As System.Exception
                Return False
            End Try
        End Function

        Public Shared Function ComputeRequestSha256(requestText As System.String) As System.String
            Using algorithm As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
                Dim bytes As System.Byte() = System.Text.Encoding.UTF8.GetBytes(If(requestText, System.String.Empty))
                Return System.BitConverter.ToString(algorithm.ComputeHash(bytes)).Replace("-", System.String.Empty).ToLowerInvariant()
            End Using
        End Function

    End Class

End Namespace
