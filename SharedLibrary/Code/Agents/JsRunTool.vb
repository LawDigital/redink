' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: JsRunTool.vb
' Purpose: ModelConfig + dispatcher entry for the js_run tool. Executes
'          untrusted JavaScript inside a hidden WebView2 sandbox with
'          restricted network access and captured console output.
'
' Architecture:
'  - Tool definition for use in model configs (safe, human-readable).
'  - Dispatcher that marshals arguments to WebView2JsSandbox.RunAsync.
'  - Returns JSON envelope: { ok, result, logs } or { ok:false, error }.
'  - Security: network disabled by default; set allow_network=true per-call.
'  - Timeout: 500..120000 ms; default 15000 ms.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Threading
Imports System.Threading.Tasks
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace Agents

    Public NotInheritable Class JsRunTool

        Private Sub New()
        End Sub

        Public Const ToolName As String = "js_run"

        Public Shared Function IsJsTool(name As String) As Boolean
            Return Not String.IsNullOrWhiteSpace(name) AndAlso
                   String.Equals(name, ToolName, StringComparison.OrdinalIgnoreCase)
        End Function

        Public Shared Function IsDisabled(sharedContext As ISharedContext) As Boolean
            Return sharedContext IsNot Nothing AndAlso sharedContext.INI_JsRunDisable
        End Function

        Public Shared Function Build(sharedContext As ISharedContext) As ModelConfig
            If IsDisabled(sharedContext) Then
                Return Nothing
            End If

            Dim def =
"{""name"":""" & ToolName & """," &
"""description"":""Run sandboxed JavaScript inside a hidden WebView2. The 'code' parameter is executed as the BODY of an async function. Therefore, do NOT wrap it in 'async function ... { }' and do NOT invent wrapper parameters such as browser_mode. Always produce the final value with an explicit top-level 'return'. Use this tool for deterministic programmable operations such as exact word counts, character counts, line counts, regex extraction/counting, exact parsing, JSON reshaping, sorting, deduplication, and other rule-based text/data transformations. console.log/console.warn/console.error output is captured. This is a browser-style sandbox, NOT a Node.js runtime: do not use require(...), fs, process, __dirname, or __filename. This tool cannot access the host filesystem, cannot enumerate directories, and cannot read arbitrary local files; use the designated file/text/workspace tools to obtain data first, then use js_run only for in-memory computation on that data. Network access is DISABLED by default; set allow_network=true to permit fetch or controlled page navigation. Browser mode: set navigate_url to load a page into the hidden browser before the code runs against the live DOM. Optional wait_for_selector and wait_after_load_ms may be used. Security: only absolute http/https URLs are allowed; localhost, loopback, and private-network destinations are blocked. Default timeout 15s.""," &
"""parameters"":{""type"":""object""," &
"""properties"":{" &
"""code"":{""type"":""string"",""description"":""JavaScript source. IMPORTANT: this is already the BODY of an async function. Write statements directly and end with a top-level return of the final value.""}," &
"""timeout_ms"":{""type"":""integer"",""description"":""Wall-clock limit (500..120000; default 15000).""}," &
"""allow_network"":{""type"":""boolean"",""description"":""Permit network requests or browser navigation (default false).""}," &
"""navigate_url"":{""type"":""string"",""description"":""Optional absolute http/https URL to open in the hidden browser before the code runs. Requires allow_network=true.""}," &
"""wait_after_load_ms"":{""type"":""integer"",""description"":""Optional extra delay after page load before executing code (0..30000; default 1500 when navigate_url is set).""}," &
"""wait_for_selector"":{""type"":""string"",""description"":""Optional CSS selector to wait for before running the code. Shadow-DOM roots are searched recursively.""}}," &
"""required"":[""code""]}}"

            Return New ModelConfig() With {
        .ToolName = ToolName,
        .ToolDefinition = def,
        .ToolInstructionsPrompt =
            ToolName & ": Run sandboxed JavaScript and receive {ok, result, logs} or {ok:false, error}. " &
            "IMPORTANT: 'code' is already the BODY of an async function. Do not declare 'async function ...'. " &
            "Prefer this tool for deterministic programmable operations such as exact counting, regex extraction, parsing, deduplication, sorting, and rule-based text/data transformations. " &
            "This is a browser-style sandbox, not Node.js: do not use require(...), fs, process, __dirname, or __filename. " &
            "Do not use js_run to inspect the host filesystem, enumerate directories, or read arbitrary local files; first obtain the data through the designated file/text/workspace tools, then use js_run only for in-memory computation on that data. " &
            "Always return the final value explicitly at top level, for example: " &
            "'const links = [...document.querySelectorAll(""a[href]"")].map(a => a.href); return links;'. " &
            "For page DOM access, use allow_network=true and navigate_url='https://...'. Do not invent browser_mode.",
        .ModelDescription = "JS sandbox (WebView2)",
        .Tool = True,
        .ToolPriority = 860,
        .ToolErrorHandling = "skip"
    }
        End Function

        Public Shared Async Function ExecuteAsync(arguments As IDictionary(Of String, Object),
                                                  sharedContext As ISharedContext,
                                                  Optional cancellationToken As CancellationToken = Nothing) As Task(Of String)
            Try
                If IsDisabled(sharedContext) Then
                    Return JsonConvert.SerializeObject(New With {
                        Key .ok = False,
                        Key .error = "js_run_disabled",
                        Key .message = "JavaScript execution is disabled by configuration."})
                End If

                ' Pre-execution guard: the WebView2 sandbox is a browser context with no Node.js
                ' runtime, so require()/module/process/__dirname and Node's fs are guaranteed to
                ' throw "X is not defined". Rejecting such code before execution turns a certain
                ' failure into a structured, actionable hint (no behavioural regression: this code
                ' could never have succeeded in the sandbox).
                Dim nodeHint As String = DetectNodeApiUsage(GetStr(arguments, "code"))
                If nodeHint <> "" Then
                    Return JsonConvert.SerializeObject(New With {
                        Key .ok = False,
                        Key .error = "node_api_unavailable",
                        Key .message = nodeHint})
                End If

                Return Await WebView2JsSandbox.RunAsync(
                    code:=GetStr(arguments, "code"),
                    timeoutMs:=GetInt(arguments, "timeout_ms", 15000),
                    allowNetwork:=GetBool(arguments, "allow_network", False),
                    navigateUrl:=GetStr(arguments, "navigate_url"),
                    waitAfterLoadMs:=GetInt(arguments, "wait_after_load_ms", 1500),
                    waitForSelector:=GetStr(arguments, "wait_for_selector"),
                    cancellationToken:=cancellationToken).ConfigureAwait(False)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {Key .error = "js_run_failed", Key .message = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' Detects Node.js-only constructs that cannot exist in the WebView2 browser sandbox.
        ''' Returns a short, model-facing hint when found, or "" when the code is safe to run.
        ''' Only patterns that are guaranteed to fail in the sandbox are matched, so a positive
        ''' detection never blocks code that could otherwise have succeeded.
        ''' </summary>
        Friend Shared Function DetectNodeApiUsage(code As String) As String
            If String.IsNullOrWhiteSpace(code) Then Return ""

            Dim executableCode As System.String = MaskJavaScriptLiteralAndCommentText(code)

            ' Node module loading (the exact failure seen in production: "require is not defined").
            If Regex.IsMatch(executableCode, "(^|[^.\w])require\s*\(") Then
                Return "Node.js module loading (require(...)) is unavailable in this sandbox. " &
                       "To read local files use the designated file tools (for example text_read, text_search, " &
                       "or the workspace_* tools) instead of Node's fs module. Use js_run only for in-memory, " &
                       "rule-based computation on data you already have."
            End If

            ' Other Node-only globals that are guaranteed to be undefined in a browser context.
            If Regex.IsMatch(executableCode, "(^|[^.\w])module\s*\.\s*exports\b") OrElse
               Regex.IsMatch(executableCode, "(^|[^.\w])__dirname\b") OrElse
               Regex.IsMatch(executableCode, "(^|[^.\w])__filename\b") OrElse
               Regex.IsMatch(executableCode, "(^|[^.\w])process\s*\.\s*(env|argv|cwd|platform)\b") Then
                Return "This code relies on Node.js runtime globals (module.exports/__dirname/__filename/process) " &
                       "that do not exist in the sandboxed browser environment. Use js_run only for in-memory, " &
                       "rule-based computation, and use the file tools (text_read, text_search, workspace_*) for filesystem access."
            End If

            Return ""
        End Function

        ''' <summary>
        ''' Returns a same-shape analysis view with JavaScript string literals, template-literal
        ''' text, and comments masked out. The Node preflight is advisory rather than a security
        ''' boundary; masking ambiguous literal text avoids rejecting browser-valid JavaScript.
        ''' </summary>
        Private Shared Function MaskJavaScriptLiteralAndCommentText(code As System.String) As System.String
            If System.String.IsNullOrEmpty(code) Then Return If(code, System.String.Empty)

            Const StateCode As System.Int32 = 0
            Const StateSingleQuoted As System.Int32 = 1
            Const StateDoubleQuoted As System.Int32 = 2
            Const StateTemplateLiteral As System.Int32 = 3
            Const StateLineComment As System.Int32 = 4
            Const StateBlockComment As System.Int32 = 5
            Const StateRegularExpression As System.Int32 = 6

            Dim masked As New System.Text.StringBuilder(code.Length)
            Dim state As System.Int32 = StateCode
            Dim escaped As System.Boolean = False
            Dim regularExpressionCharacterClass As System.Boolean = False
            Dim canStartRegularExpression As System.Boolean = True
            Dim lexicalStateCertain As System.Boolean = True
            Dim templateExpressionDepth As System.Int32 = 0
            Dim templateExpressionDepthStack As New System.Collections.Generic.Stack(Of System.Int32)()
            Dim controlParenthesisStack As New System.Collections.Generic.Stack(Of System.Boolean)()
            Dim pendingControlParenthesis As System.Boolean = False
            Dim nextIdentifierIsMemberName As System.Boolean = False
            Dim index As System.Int32 = 0

            Do While index < code.Length
                Dim current As System.Char = code.Chars(index)
                Dim nextCharacter As System.Char = System.Char.MinValue
                If index + 1 < code.Length Then nextCharacter = code.Chars(index + 1)

                Select Case state
                    Case StateCode
                        If System.Char.IsWhiteSpace(current) Then
                            masked.Append(current)
                            index += 1
                            Continue Do
                        End If

                        If current = "'"c Then
                            masked.Append(" "c)
                            state = StateSingleQuoted
                            escaped = False
                            pendingControlParenthesis = False
                            index += 1
                            Continue Do
                        End If

                        If current = """"c Then
                            masked.Append(" "c)
                            state = StateDoubleQuoted
                            escaped = False
                            pendingControlParenthesis = False
                            index += 1
                            Continue Do
                        End If

                        If current = "`"c Then
                            masked.Append(" "c)
                            state = StateTemplateLiteral
                            escaped = False
                            pendingControlParenthesis = False
                            index += 1
                            Continue Do
                        End If

                        If current = "/"c AndAlso nextCharacter = "/"c Then
                            masked.Append("  ")
                            index += 2
                            state = StateLineComment
                            Continue Do
                        End If

                        If current = "/"c AndAlso nextCharacter = "*"c Then
                            masked.Append("  ")
                            index += 2
                            state = StateBlockComment
                            Continue Do
                        End If

                        If current = "/"c AndAlso canStartRegularExpression Then
                            masked.Append(" "c)
                            state = StateRegularExpression
                            escaped = False
                            regularExpressionCharacterClass = False
                            pendingControlParenthesis = False
                            index += 1
                            Continue Do
                        End If

                        If current = "/"c Then
                            ' Division and /= both require a right-hand expression. Keeping the slash
                            ' as code prevents a normal division operator from becoming a fake regex.
                            masked.Append(current)
                            canStartRegularExpression = True
                            pendingControlParenthesis = False
                            index += 1
                            Continue Do
                        End If

                        If IsJavaScriptIdentifierStart(current) Then
                            Dim identifierEnd As System.Int32 = index + 1
                            While identifierEnd < code.Length AndAlso
                                  IsJavaScriptIdentifierPart(code.Chars(identifierEnd))
                                identifierEnd += 1
                            End While

                            Dim identifier As System.String =
                                code.Substring(index, identifierEnd - index)
                            masked.Append(identifier)

                            Dim identifierIsMemberName As System.Boolean = nextIdentifierIsMemberName
                            nextIdentifierIsMemberName = False
                            If identifierIsMemberName Then
                                pendingControlParenthesis = False
                                canStartRegularExpression = False
                            Else
                                pendingControlParenthesis =
                                    IsJavaScriptControlParenthesisKeyword(identifier)
                                canStartRegularExpression =
                                    JavaScriptKeywordAllowsRegularExpressionAfter(identifier)
                            End If
                            index = identifierEnd
                            Continue Do
                        End If

                        If System.Char.IsDigit(current) Then
                            masked.Append(current)
                            canStartRegularExpression = False
                            pendingControlParenthesis = False
                            index += 1
                            Continue Do
                        End If

                        Select Case current
                            Case "("c
                                masked.Append(current)
                                controlParenthesisStack.Push(pendingControlParenthesis)
                                pendingControlParenthesis = False
                                canStartRegularExpression = True

                            Case ")"c
                                masked.Append(current)
                                Dim closesControlParenthesis As System.Boolean = False
                                If controlParenthesisStack.Count > 0 Then
                                    closesControlParenthesis = controlParenthesisStack.Pop()
                                End If
                                canStartRegularExpression = closesControlParenthesis
                                pendingControlParenthesis = False

                            Case "{"c
                                masked.Append(current)
                                If templateExpressionDepth > 0 Then
                                    templateExpressionDepth += 1
                                End If
                                canStartRegularExpression = True
                                pendingControlParenthesis = False

                            Case "}"c
                                If templateExpressionDepth > 0 Then
                                    templateExpressionDepth -= 1
                                    If templateExpressionDepth = 0 Then
                                        masked.Append(" "c)
                                        If templateExpressionDepthStack.Count > 0 Then
                                            templateExpressionDepth = templateExpressionDepthStack.Pop()
                                        End If
                                        state = StateTemplateLiteral
                                        escaped = False
                                        pendingControlParenthesis = False
                                        index += 1
                                        Continue Do
                                    End If
                                End If

                                masked.Append(current)
                                ' A closing brace can end an object value or a statement block. Treat
                                ' a following slash conservatively as regex-capable; a false negative
                                ' advisory hint is safer than rejecting browser-valid regex code.
                                canStartRegularExpression = True
                                pendingControlParenthesis = False

                            Case "["c
                                masked.Append(current)
                                canStartRegularExpression = True
                                pendingControlParenthesis = False

                            Case "]"c
                                masked.Append(current)
                                canStartRegularExpression = False
                                pendingControlParenthesis = False

                            Case "."c
                                masked.Append(current)
                                canStartRegularExpression = False
                                pendingControlParenthesis = False
                                nextIdentifierIsMemberName = True

                            Case "+"c, "-"c
                                If nextCharacter = current Then
                                    masked.Append(current)
                                    masked.Append(nextCharacter)
                                    index += 2
                                    pendingControlParenthesis = False
                                    Continue Do
                                End If

                                masked.Append(current)
                                canStartRegularExpression = True
                                pendingControlParenthesis = False

                            Case "="c, "!"c, "*"c, "%"c, "&"c, "|"c, "^"c,
                                 "<"c, ">"c, "?"c, ","c, ";"c, ":"c, "~"c
                                masked.Append(current)
                                canStartRegularExpression = True
                                pendingControlParenthesis = False

                            Case Else
                                masked.Append(current)
                                pendingControlParenthesis = False
                        End Select

                        index += 1

                    Case StateSingleQuoted, StateDoubleQuoted
                        If current = Microsoft.VisualBasic.ControlChars.Cr OrElse
                           current = Microsoft.VisualBasic.ControlChars.Lf Then
                            masked.Append(current)
                        Else
                            masked.Append(" "c)
                        End If

                        If escaped Then
                            escaped = False
                        ElseIf current = "\"c Then
                            escaped = True
                        ElseIf (state = StateSingleQuoted AndAlso current = "'"c) OrElse
                               (state = StateDoubleQuoted AndAlso current = """"c) Then
                            state = StateCode
                            canStartRegularExpression = False
                        End If

                        index += 1

                    Case StateTemplateLiteral
                        If escaped Then
                            If current = Microsoft.VisualBasic.ControlChars.Cr OrElse
                               current = Microsoft.VisualBasic.ControlChars.Lf Then
                                masked.Append(current)
                            Else
                                masked.Append(" "c)
                            End If
                            escaped = False
                            index += 1
                            Continue Do
                        End If

                        If current = "\"c Then
                            masked.Append(" "c)
                            escaped = True
                            index += 1
                            Continue Do
                        End If

                        If current = "`"c Then
                            masked.Append(" "c)
                            state = StateCode
                            canStartRegularExpression = False
                            index += 1
                            Continue Do
                        End If

                        If current = "$"c AndAlso nextCharacter = "{"c Then
                            masked.Append("  ")
                            index += 2
                            templateExpressionDepthStack.Push(templateExpressionDepth)
                            templateExpressionDepth = 1
                            state = StateCode
                            canStartRegularExpression = True
                            pendingControlParenthesis = False
                            Continue Do
                        End If

                        If current = Microsoft.VisualBasic.ControlChars.Cr OrElse
                           current = Microsoft.VisualBasic.ControlChars.Lf Then
                            masked.Append(current)
                        Else
                            masked.Append(" "c)
                        End If
                        index += 1

                    Case StateLineComment
                        If current = Microsoft.VisualBasic.ControlChars.Cr OrElse
                           current = Microsoft.VisualBasic.ControlChars.Lf Then
                            masked.Append(current)
                            state = StateCode
                        Else
                            masked.Append(" "c)
                        End If
                        index += 1

                    Case StateBlockComment
                        If current = "*"c AndAlso nextCharacter = "/"c Then
                            masked.Append("  ")
                            index += 2
                            state = StateCode
                            Continue Do
                        End If

                        If current = Microsoft.VisualBasic.ControlChars.Cr OrElse
                           current = Microsoft.VisualBasic.ControlChars.Lf Then
                            masked.Append(current)
                        Else
                            masked.Append(" "c)
                        End If
                        index += 1

                    Case StateRegularExpression
                        If current = Microsoft.VisualBasic.ControlChars.Cr OrElse
                           current = Microsoft.VisualBasic.ControlChars.Lf Then
                            masked.Append(current)
                            state = StateCode
                            lexicalStateCertain = False
                            canStartRegularExpression = True
                            regularExpressionCharacterClass = False
                            escaped = False
                            index += 1
                            Continue Do
                        End If

                        masked.Append(" "c)

                        If escaped Then
                            escaped = False
                        ElseIf current = "\"c Then
                            escaped = True
                        ElseIf current = "["c AndAlso Not regularExpressionCharacterClass Then
                            regularExpressionCharacterClass = True
                        ElseIf current = "]"c AndAlso regularExpressionCharacterClass Then
                            regularExpressionCharacterClass = False
                        ElseIf current = "/"c AndAlso Not regularExpressionCharacterClass Then
                            state = StateCode
                            canStartRegularExpression = False

                            Dim flagIndex As System.Int32 = index + 1
                            While flagIndex < code.Length AndAlso
                                  IsJavaScriptIdentifierPart(code.Chars(flagIndex))
                                masked.Append(" "c)
                                flagIndex += 1
                            End While

                            index = flagIndex
                            Continue Do
                        End If

                        index += 1
                End Select
            Loop

            If state = StateSingleQuoted OrElse
               state = StateDoubleQuoted OrElse
               state = StateTemplateLiteral OrElse
               state = StateBlockComment OrElse
               state = StateRegularExpression OrElse
               templateExpressionDepth > 0 OrElse
               templateExpressionDepthStack.Count > 0 Then
                lexicalStateCertain = False
            End If

            If Not lexicalStateCertain Then
                Return MaskAllJavaScriptTextPreservingLineBreaks(code)
            End If

            Return masked.ToString()
        End Function

        Private Shared Function IsJavaScriptIdentifierStart(value As System.Char) As System.Boolean
            Return value = "_"c OrElse
                   value = "$"c OrElse
                   System.Char.IsLetter(value)
        End Function

        Private Shared Function IsJavaScriptIdentifierPart(value As System.Char) As System.Boolean
            Return IsJavaScriptIdentifierStart(value) OrElse System.Char.IsDigit(value)
        End Function

        Private Shared Function JavaScriptKeywordAllowsRegularExpressionAfter(
            identifier As System.String
        ) As System.Boolean

            Select Case identifier
                Case "return", "throw", "case", "delete", "void", "typeof", "new",
                     "yield", "await", "instanceof", "in", "of", "else", "do"
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Private Shared Function IsJavaScriptControlParenthesisKeyword(
            identifier As System.String
        ) As System.Boolean

            Select Case identifier
                Case "if", "while", "for", "with", "switch", "catch"
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Private Shared Function MaskAllJavaScriptTextPreservingLineBreaks(
            code As System.String
        ) As System.String

            Dim masked As New System.Text.StringBuilder(code.Length)
            For Each current As System.Char In code
                If current = Microsoft.VisualBasic.ControlChars.Cr OrElse
                   current = Microsoft.VisualBasic.ControlChars.Lf Then
                    masked.Append(current)
                Else
                    masked.Append(" "c)
                End If
            Next
            Return masked.ToString()
        End Function

        Private Shared Function GetStr(args As IDictionary(Of String, Object), name As String) As String
            If args Is Nothing Then Return ""
            Dim v As Object = Nothing
            If Not args.TryGetValue(name, v) OrElse v Is Nothing Then Return ""
            Return System.Convert.ToString(v)
        End Function

        Private Shared Function GetInt(args As IDictionary(Of String, Object), name As String, dflt As Integer) As Integer
            If args Is Nothing Then Return dflt
            Dim v As Object = Nothing
            If Not args.TryGetValue(name, v) OrElse v Is Nothing Then Return dflt
            Try : Return System.Convert.ToInt32(v) : Catch
                Dim n As Integer
                If Integer.TryParse(System.Convert.ToString(v), n) Then Return n
                Return dflt
            End Try
        End Function

        Private Shared Function GetBool(args As IDictionary(Of String, Object), name As String, dflt As Boolean) As Boolean
            If args Is Nothing Then Return dflt
            Dim v As Object = Nothing
            If Not args.TryGetValue(name, v) OrElse v Is Nothing Then Return dflt
            Try : Return System.Convert.ToBoolean(v) : Catch
                Select Case System.Convert.ToString(v).Trim().ToLowerInvariant()
                    Case "true", "1", "yes" : Return True
                    Case "false", "0", "no" : Return False
                    Case Else : Return dflt
                End Select
            End Try
        End Function

    End Class

End Namespace
