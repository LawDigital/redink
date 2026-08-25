' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: AskUserTool.vb
' Purpose: Model-driven request for clarification or required input from the
'          user during a tooling loop. Host- and UI-agnostic: the shared tool
'          only builds the request and formats the answer.
'
'          Interactivity safety (BINDING):
'            - MUST only ever block for input where a live user is present:
'              Word (DiscussInky / Form1) and Outlook Local Chat / Web Agent
'              (including when the Scheduler opens Local Chat in browser_prompt
'              mode).
'            - MUST NOT block in the e-mail Scheduler or AutoPilot, which run
'              unattended and reply by e-mail; blocking there would stall the
'              whole system. Hosts arm InteractivityProvider so Execute returns
'              a non-blocking result in those contexts.
' =============================================================================

Option Strict On
Option Explicit On

Imports Newtonsoft.Json.Linq

Namespace Agents

    ''' <summary>One selectable option offered to the user.</summary>
    Public NotInheritable Class AskUserOption
        Public Property Id As String
        Public Property Label As String
        Public Property Description As String
    End Class

    ''' <summary>Structured request passed to the host UI callback.</summary>
    Public NotInheritable Class AskUserRequest
        Public Property Question As String
        Public Property Options As List(Of AskUserOption)
        Public Property AllowFreeText As Boolean = True
        Public Property MultiSelect As Boolean = False
        ''' <summary>"text" (default), "integer", "number", or "choice".</summary>
        Public Property InputType As String = "text"
    End Class

    ''' <summary>Structured answer returned by the host UI callback.</summary>
    Public NotInheritable Class AskUserResult
        ''' <summary>"answered" or "cancelled".</summary>
        Public Property Status As String = "cancelled"
        Public Property SelectedOptionIds As List(Of String)
        Public Property FreeText As String
    End Class

    Public NotInheritable Class AskUserTool

        Private Sub New()
        End Sub

        Public Const ToolName As String = "ask_user"

        ''' <summary>
        ''' Optional host-supplied UI override. When Nothing, Execute uses the shared modal
        ''' dialog (SharedMethods.ShowAskUserDialog).
        ''' </summary>
        Public Shared Property Callback As Func(Of AskUserRequest, AskUserResult)

        ''' <summary>
        ''' Host-supplied predicate reporting whether a live, interactive user is present for
        ''' the current run. Word leaves this Nothing (always interactive). Outlook arms it so
        ''' e-mail Scheduler and AutoPilot runs are reported non-interactive. When it returns
        ''' False, Execute never blocks and returns a non-blocking "proceed" result.
        ''' </summary>
        Public Shared Property InteractivityProvider As Func(Of Boolean)

        Public Shared Function IsAskUserTool(name As String) As Boolean
            Return Not String.IsNullOrWhiteSpace(name) AndAlso
                   name.Trim().Equals(ToolName, StringComparison.OrdinalIgnoreCase)
        End Function

        ''' <summary>True when a live user is available to answer in the current run.</summary>
        Private Shared Function IsInteractive() As Boolean
            Dim p As Func(Of Boolean) = InteractivityProvider
            If p Is Nothing Then Return True
            Try
                Return p()
            Catch
                ' Fail safe: if the host cannot confirm interactivity, do not block.
                Return False
            End Try
        End Function

        Public Shared Function Build() As SharedLibrary.ModelConfig
            Dim def As String =
                "{""name"":""" & ToolName & """," &
                """description"":""Ask the user for information or direction needed to continue. This is the REQUIRED channel for obtaining user input: when you need a decision, choice, value, or clarification from the user, you MUST call this tool rather than asking in narrative prose or assuming a default. Use it when required information is missing, when multiple materially different interpretations are possible, or when a skill or workflow requires an explicit user choice or value. When a skill or workflow explicitly instructs you to ask the user (e.g. to disambiguate between several candidates), that instruction is mandatory and overrides any general preference for defaults: in that case you MUST call ask_user and MUST NOT skip it on the grounds that a probable default exists or that a chat sentence would be simpler. ONE QUESTION PER CALL (BINDING): the question field must contain exactly one decision/fact question. Never bundle several questions, numbered subquestions, or a questionnaire into one ask_user call. If several facts are missing, ask one question, wait for its answer, then issue a new ask_user call for the next still-material fact. For choice/yes-no/unknown questions, supply concrete clickable options and set input_type=choice. Choice buttons are suggestions, not an exhaustive list; the user must always be allowed to provide a free-form answer as an alternative. Absent such an explicit instruction, do not ask when the answer is already known, a harmless obvious default exists, or the uncertainty does not materially affect the result. If the run is non-interactive, this tool returns immediately without a user answer, so proceed with a clearly stated assumption instead.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """question"":{""type"":""string"",""description"":""One clear, concise question describing the actual decision to be made.""}," &
                """options"":{""type"":""array"",""description"":""Optional 2-6 concrete choices to help the user answer faster."",""items"":{""type"":""object"",""properties"":{" &
                """id"":{""type"":""string"",""description"":""Stable machine-readable option id, independent of the displayed wording.""}," &
                """label"":{""type"":""string"",""description"":""Short user-facing option text.""}," &
                """description"":{""type"":""string"",""description"":""Optional clarification when the label alone may be ambiguous. Omit this property when no clarification is needed.""}" &
                "}}}," &
                """allow_free_text"":{""type"":""boolean"",""description"":""When true (default), the user may ignore the options and answer freely. Options are suggestions, not an exhaustive list.""}," &
                """multi_select"":{""type"":""boolean"",""description"":""Set true only when selecting more than one option is semantically valid. Default false.""}," &
                """input_type"":{""type"":""string"",""enum"":[""text"",""integer"",""number"",""choice""],""description"":""Expected free-text answer kind when no option is chosen. Use 'integer' or 'number' to request a typed value, or 'choice' when one of the options must be selected. Default 'text'.""}" &
                "},""required"":[""question""],""additionalProperties"":false}}"

            Return New SharedLibrary.ModelConfig() With {
                .ToolName = ToolName,
                .ToolDefinition = def,
                .ToolInstructionsPrompt = ToolName & ": The required channel for getting input or direction from the user — when you need a decision, choice, value, or clarification, call this tool instead of asking in prose or silently assuming a default. ONE QUESTION PER CALL (BINDING): ask exactly one high-value question. Never put multiple questions, numbered subquestions, or a questionnaire into the question field. If several facts are missing, call ask_user once for the first still-material fact, wait for the answer, then call ask_user again for the next fact. When the answer is categorical (including yes/no/unknown), provide 2-6 concrete clickable options, set input_type=choice, and set allow_free_text=true so the user may always provide an alternative free-form answer. When a skill/workflow explicitly tells you to ask the user, that is mandatory and must be performed sequentially one ask_user call at a time. Otherwise, do not ask about minor uncertainty, do not repeat answered questions, and continue as soon as the answer is sufficient. If the run is non-interactive this tool returns without an answer; in that case proceed with a clearly stated assumption.",
                .ModelDescription = "User clarification / input (internal)",
                .Tool = True,
                .ToolPriority = 937,
                .ToolErrorHandling = "skip"
            }
        End Function

        Public Shared Function Execute(arguments As Dictionary(Of String, Object)) As String
            ' Non-blocking guard: never wait for input where no live user is present
            ' (e-mail Scheduler, AutoPilot). Returns a result the model can act on directly.
            If Not IsInteractive() Then
                Return New JObject(
                    New JProperty("status", "no_user"),
                    New JProperty("selected_option_ids", New JArray()),
                    New JProperty("free_text", Nothing),
                    New JProperty("guidance", "No live user can be asked in this run (unattended AutoPilot / e-mail Scheduler). Do not invent or guess a materially required fact. If the answer is required by the active skill/workflow, ask the clarification in the outgoing response and end the current run blocked so the requester can reply. Use a default only when the active workflow explicitly permits that default."),
                    New JProperty("error", New JObject(
                        New JProperty("code", "no_user_available"),
                        New JProperty("message", "This run is non-interactive; ask_user cannot collect an answer.")))
                ).ToString(Newtonsoft.Json.Formatting.None)
            End If

            Dim req As New AskUserRequest() With {
                .Question = GetString(arguments, "question"),
                .AllowFreeText = GetBool(arguments, "allow_free_text", True),
                .MultiSelect = GetBool(arguments, "multi_select", False),
                .InputType = GetInputType(arguments),
                .Options = ParseOptions(arguments)
            }

            ' Choice buttons are suggestions, never an exhaustive restriction.
            ' Preserve a free-text escape hatch even if a model/skill explicitly
            ' supplied allow_free_text=false.
If String.IsNullOrWhiteSpace(req.Question) Then
                Return New JObject(
                    New JProperty("status", "cancelled"),
                    New JProperty("selected_option_ids", New JArray()),
                    New JProperty("free_text", Nothing),
                    New JProperty("error", New JObject(
                        New JProperty("code", "missing_question"),
                        New JProperty("message", "ask_user requires a non-empty 'question'.")))
                ).ToString(Newtonsoft.Json.Formatting.None)
            End If

            If LooksLikeBundledQuestions(req.Question) Then
                Return New JObject(
                    New JProperty("status", "invalid_request"),
                    New JProperty("selected_option_ids", New JArray()),
                    New JProperty("free_text", Nothing),
                    New JProperty("error", New JObject(
                        New JProperty("code", "multiple_questions_not_allowed"),
                        New JProperty("message", "ask_user accepts exactly one question per call. Ask the first still-material fact now, wait for the answer, then call ask_user again for the next fact.")))
                ).ToString(Newtonsoft.Json.Formatting.None)
            End If

            If req.InputType.Equals("choice", StringComparison.OrdinalIgnoreCase) AndAlso
               (req.Options Is Nothing OrElse req.Options.Count < 2) Then

                Return New JObject(
                    New JProperty("status", "invalid_request"),
                    New JProperty("selected_option_ids", New JArray()),
                    New JProperty("free_text", Nothing),
                    New JProperty("error", New JObject(
                        New JProperty("code", "choice_options_required"),
                        New JProperty("message", "For input_type=choice, provide at least two concrete clickable options. Ask exactly one choice question in this call.")))
                ).ToString(Newtonsoft.Json.Formatting.None)
            End If

            Dim cb As Func(Of AskUserRequest, AskUserResult) = Callback
            If cb Is Nothing Then
                cb = AddressOf SharedLibrary.SharedMethods.ShowAskUserDialog
            End If

            Dim res As AskUserResult
            Try
                res = cb(req)
            Catch ex As System.Exception
                Return New JObject(
                    New JProperty("status", "cancelled"),
                    New JProperty("selected_option_ids", New JArray()),
                    New JProperty("free_text", Nothing),
                    New JProperty("error", New JObject(
                        New JProperty("code", "ask_user_failed"),
                        New JProperty("message", If(ex.Message, "The user prompt could not be shown."))))
                ).ToString(Newtonsoft.Json.Formatting.None)
            End Try

            If res Is Nothing Then
                res = New AskUserResult() With {.Status = "cancelled"}
            End If

            Dim ids As New JArray()
            If res.SelectedOptionIds IsNot Nothing Then
                For Each id In res.SelectedOptionIds
                    If Not String.IsNullOrWhiteSpace(id) Then ids.Add(id)
                Next
            End If

            Dim freeText As JToken =
                If(String.IsNullOrEmpty(res.FreeText), CType(JValue.CreateNull(), JToken), New JValue(res.FreeText))

            Dim statusText As String = If(res.Status, "cancelled")

            Dim outObj As New JObject(
                New JProperty("status", statusText),
                New JProperty("selected_option_ids", ids),
                New JProperty("free_text", freeText)
            )

            ' When the user dismisses the dialog without answering, the model must not
            ' re-ask or stall: it should pick the most likely answer itself and disclose it.
            If Not statusText.Equals("answered", StringComparison.OrdinalIgnoreCase) Then
                outObj("guidance") = "The user cancelled this question. Do not invent a materially required fact. If the active workflow requires the answer, leave the fact missing and end/continue according to that workflow's cancellation rule. Use an assumption only when the workflow explicitly permits it."
            End If

            Return outObj.ToString(Newtonsoft.Json.Formatting.None)
        End Function

        Private Shared Function LooksLikeBundledQuestions(question As String) As Boolean
            Dim text As String = If(question, "").Trim()
            If text = "" Then Return False

            Dim questionMarks As Integer = 0
            For Each ch As Char In text
                If ch = "?"c Then questionMarks += 1
            Next
            If questionMarks > 1 Then Return True

            Dim numberedMatches As System.Text.RegularExpressions.MatchCollection =
                System.Text.RegularExpressions.Regex.Matches(
                    text,
                    "(?m)^\s*\d{1,2}[\.\)]\s+"
                )
            If numberedMatches.Count > 1 Then Return True

            Dim bulletQuestions As System.Text.RegularExpressions.MatchCollection =
                System.Text.RegularExpressions.Regex.Matches(
                    text,
                    "(?m)^\s*[-*•]\s+.*\?\s*$"
                )
            Return bulletQuestions.Count > 1
        End Function

        Private Shared Function ParseOptions(args As Dictionary(Of String, Object)) As List(Of AskUserOption)
            Dim result As New List(Of AskUserOption)()
            If args Is Nothing OrElse Not args.ContainsKey("options") OrElse args("options") Is Nothing Then Return result

            Dim arr As JArray = Nothing
            Try
                Dim raw = args("options")
                Dim tok As JToken = TryCast(raw, JToken)
                If tok Is Nothing Then tok = JToken.FromObject(raw)
                arr = TryCast(tok, JArray)
            Catch
                Return result
            End Try

            If arr Is Nothing Then Return result
            For Each item In arr
                Dim o As JObject = TryCast(item, JObject)
                If o Is Nothing Then Continue For
                result.Add(New AskUserOption() With {
                    .Id = If(o("id")?.ToString(), ""),
                    .Label = If(o("label")?.ToString(), ""),
                    .Description = If(o("description") Is Nothing OrElse o("description").Type = JTokenType.Null,
                                      Nothing, o("description").ToString())
                })
            Next
            Return result
        End Function

        Private Shared Function GetInputType(args As Dictionary(Of String, Object)) As String
            Dim v As String = GetString(args, "input_type").Trim().ToLowerInvariant()
            Select Case v
                Case "integer", "number", "choice", "text"
                    Return v
                Case Else
                    Return "text"
            End Select
        End Function

        Private Shared Function GetString(args As Dictionary(Of String, Object), key As String) As String
            If args Is Nothing OrElse Not args.ContainsKey(key) OrElse args(key) Is Nothing Then Return ""
            Return Convert.ToString(args(key))
        End Function

        Private Shared Function GetBool(args As Dictionary(Of String, Object), key As String, fallback As Boolean) As Boolean
            If args Is Nothing OrElse Not args.ContainsKey(key) OrElse args(key) Is Nothing Then Return fallback
            Dim b As Boolean
            If Boolean.TryParse(Convert.ToString(args(key)), b) Then Return b
            Return fallback
        End Function

    End Class

End Namespace
