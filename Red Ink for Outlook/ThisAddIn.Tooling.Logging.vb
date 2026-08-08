' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.Tooling.Logging.vb
' Purpose: Reduced file-based logger for tooling operations with nested session depth support.
'          Writes per-run log file with system prompt snapshots, LLM calls, and diagnostic traces.
'
' Architecture:
'  - Session Management:
'      - StartSession(): Enables logging (when INI_APIDebug=True) and increments _sessionDepth.
'          - Tracks nested tooling sessions (prevents sub-agent runs from tearing down parent log).
'          - WriteHeader(): Creates stable overwrite-per-run log file on Desktop.
'  - Logging API:
'      - LogStep(): Standard step logging (written to file and UI LogWindow).
'      - LogDiag(): Diagnostic logging (file only, not UI-visible).
'      - LogWarn(): Warning logging with optional details and exception.
'      - LogError(): Error logging with optional details and exception.
'      - LogModelConfigOnce(): Dumps ModelConfig properties (excluding sensitive fields).
'      - LogRawResponseStub(): Writes full LLM/tool responses with blank line separators.
'  - File Operations:
'      - Single stable filename per session (StableLogFileName, overwritten each run).
'      - Synchronized file writes via _lock to prevent interleaving.
'      - File path on Desktop (Environment.SpecialFolder.DesktopDirectory).
'  - Output Format:
'      - Header with session timestamp, version, and log location.
'      - Timestamped entries: [HH:mm:ss] category: message.
'      - Exception details include type, message, stack trace, and inner exception.
'  - Exception Serialization:
'      - WriteException(): Dumps full exception hierarchy (type, message, stack trace, inner exception).
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.Reflection
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Web.WebView2.Core
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn


    ''' <summary>
    ''' Reduced file-based logger for tooling operations.
    ''' - Single file per run (overwrites).
    ''' - Writes: LogWindow steps, warnings, errors, and pre-LLM call snapshots.
    ''' - Writes full raw LLM/tool responses with two empty lines before and after.
    ''' </summary>
    Partial Public Class ToolingFileLogger

        ''' <summary>Absolute filesystem path to the current tooling log file (if enabled).</summary>
        Private Shared _logPath As String = Nothing

        ''' <summary>Whether file logging is enabled (controlled by <c>INI_APIDebug</c> in <see cref="StartSession"/>).</summary>
        Private Shared _isEnabled As Boolean = False

        ''' <summary>Whether the session header has already been written in this run.</summary>
        Private Shared _started As Boolean = False

        ''' <summary>Synchronizes file writes to avoid interleaving.</summary>
        Private Shared ReadOnly _lock As New Object()

        ''' <summary>Stable log filename used for the tooling session log (overwritten each run).</summary>
        Private Shared ReadOnly StableLogFileName As String = $"{AN5}_Tooling_Log.txt"

        ''' <summary>Whether the current session writes into the local diagnostics folder (per-skill, keep-last-run).</summary>
        Private Shared _isDiagnosticsMode As Boolean = False

        ''' <summary>Sanitized name of the top-level skill invoked in this run (used for per-skill log filenames).</summary>
        Private Shared _topLevelSkillSlug As String = Nothing

        ''' <summary>Nested tooling-session depth. Prevents sub-agent runs from tearing down the parent log session.</summary>
        Private Shared _sessionDepth As Integer = 0

        ''' <summary>
        ''' Starts a tooling log session and writes the log header.
        ''' Logging is enabled only when <c>INI_APIDebug</c> is <c>True</c>.
        ''' </summary>
        Public Shared Sub StartSession()
            SyncLock _lock
                Dim authorActive As Boolean = SharedLibrary.Agents.SkillAuthorMode.IsActive
                Dim shouldEnable As Boolean = INI_APIDebug OrElse authorActive
                If Not shouldEnable Then Return

                _isEnabled = True
                _sessionDepth += 1

                If _started AndAlso Not String.IsNullOrWhiteSpace(_logPath) Then
                    WriteLine("STEP", $"Nested tooling session started (depth={_sessionDepth})")
                    Return
                End If

                ' Choose the base directory. In skill-author mode with a configured local
                ' resource root, write into <local>\diagnostics\ so the skill-author skill
                ' can read the last run's log. Otherwise fall back to the Desktop, but only
                ' when APIDebug is on (avoid cluttering the Desktop when it is off).
                Dim baseDir As String = Nothing
                _isDiagnosticsMode = False

                If authorActive Then
                    Dim localRoot As String = SharedLibrary.Agents.AgentResources.ConfiguredLocalPath
                    If Not String.IsNullOrWhiteSpace(localRoot) Then
                        Try
                            Dim diagDir As String = Path.Combine(localRoot, "diagnostics")
                            If Not Directory.Exists(diagDir) Then Directory.CreateDirectory(diagDir)
                            baseDir = diagDir
                            _isDiagnosticsMode = True
                        Catch
                            baseDir = Nothing
                        End Try
                    End If
                End If

                If baseDir Is Nothing Then
                    If Not INI_APIDebug Then
                        ' Author mode without a usable local root: skip logging entirely.
                        _isEnabled = False
                        _sessionDepth = 0
                        Return
                    End If
                    baseDir = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                End If

                _topLevelSkillSlug = Nothing
                _logPath = Path.Combine(baseDir, StableLogFileName)

                Try
                    WriteHeader()
                    _started = True
                Catch ex As Exception
                    _isEnabled = False
                    _started = False
                    _logPath = Nothing
                    _sessionDepth = 0
                    _isDiagnosticsMode = False
                End Try
            End SyncLock
        End Sub

        ''' <summary>
        ''' Writes the session header and overwrites any previous log file with the same name.
        ''' </summary>
        Private Shared Sub WriteHeader()
            Dim header As New StringBuilder()
            header.AppendLine("=" & New String("="c, 78))
            header.AppendLine($"{AN} - Tooling Log")
            header.AppendLine($"Started: {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}")
            header.AppendLine($"Version: {Version}")
            header.AppendLine($"File: {StableLogFileName} (overwritten each run)")
            header.AppendLine("=" & New String("="c, 78))
            header.AppendLine()

            File.WriteAllText(_logPath, header.ToString())
        End Sub



        Public Shared Sub LogStep(message As String)
            If Not _isEnabled OrElse String.IsNullOrWhiteSpace(_logPath) Then Return
            WriteLine("STEP", message)
        End Sub

        Public Shared Sub LogDiag(message As String)
            If Not _isEnabled OrElse String.IsNullOrWhiteSpace(_logPath) Then Return
            If String.IsNullOrWhiteSpace(message) Then Return
            WriteLine("DIAG", message)
        End Sub

        Public Shared Sub LogWarn(message As String, Optional details As String = "", Optional ex As Exception = Nothing)
            If Not _isEnabled OrElse String.IsNullOrWhiteSpace(_logPath) Then Return

            If Not String.IsNullOrWhiteSpace(message) Then
                WriteLine("WARN", message)
            End If

            If Not String.IsNullOrWhiteSpace(details) Then
                WriteLine("WARN", $"Details: {details}")
            End If

            If ex IsNot Nothing Then
                WriteException("WARN", ex)
            End If
        End Sub

        Public Shared Sub LogError(message As String, Optional details As String = "", Optional ex As Exception = Nothing)
            If Not _isEnabled OrElse String.IsNullOrWhiteSpace(_logPath) Then Return

            If Not String.IsNullOrWhiteSpace(message) Then
                WriteLine("ERR", message)
            End If

            If Not String.IsNullOrWhiteSpace(details) Then
                WriteLine("ERR", $"Details: {details}")
            End If

            If ex IsNot Nothing Then
                WriteException("ERR", ex)
            End If
        End Sub

        ''' <summary>
        ''' Writes exception details to the log under the specified category.
        ''' </summary>
        ''' <param name="category">Log category identifier (e.g., WARN/ERR/END).</param>
        ''' <param name="ex">Exception instance to serialize.</param>
        Private Shared Sub WriteException(category As String, ex As Exception)
            If ex Is Nothing Then Return
            WriteLine(category, $"Exception Type: {ex.GetType().FullName}")
            WriteLine(category, $"Exception Message: {ex.Message}")
            If Not String.IsNullOrWhiteSpace(ex.StackTrace) Then
                WriteLine(category, "Stack Trace:")
                WriteRaw(category, ex.StackTrace)
            End If
            If ex.InnerException IsNot Nothing Then
                WriteLine(category, $"Inner Exception Type: {ex.InnerException.GetType().FullName}")
                WriteLine(category, $"Inner Exception Message: {ex.InnerException.Message}")
            End If
        End Sub

        ''' <summary>
        ''' Logs every public instance property of <see cref="ModelConfig"/> (unmodified) for diagnostics,
        ''' skipping a fixed list of sensitive/high-volume properties.
        ''' </summary>
        ''' <param name="config">Configuration instance to log.</param>
        ''' <param name="label">Label written before the config dump.</param>
        Public Shared Sub LogModelConfigOnce(config As ModelConfig, label As String)
            If Not _isEnabled OrElse config Is Nothing Then Return

            WriteLine("CONF", $"{label}:")

            Try
                Dim excludedExact As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
                    "APIEncrypted",
                    "APIKey",
                    "APIKeyBack",
                    "APIKeyPrefix",
                    "DecodedAPI",
                    "TokenCount",
                    "MaxOutputToken",
                    "MergePrompt",
                    "QueryPrompt",
                    "TokenExpiry"
                }

                Dim excludedPrefixes As String() = {
                    "OAuth2",
                    "Parameter"
                }

                Dim props = GetType(ModelConfig).GetProperties(BindingFlags.Instance Or BindingFlags.Public).
                    Where(Function(p)
                              If excludedExact.Contains(p.Name) Then Return False
                              For Each prefix In excludedPrefixes
                                  If p.Name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) Then Return False
                              Next
                              Return True
                          End Function).
                    OrderBy(Function(p) p.Name).
                    ToList()

                For Each p In props
                    Dim v As Object = Nothing
                    Try
                        v = p.GetValue(config, Nothing)
                    Catch ex As Exception
                        WriteLine("CONF", $"  {p.Name}: <error reading>")
                        WriteLine("CONF", $"    {ex.GetType().Name}: {ex.Message}")
                        Continue For
                    End Try

                    Dim textValue As String
                    If v Is Nothing Then
                        textValue = ""
                    ElseIf TypeOf v Is DateTime Then
                        textValue = DirectCast(v, DateTime).ToString("yyyy-MM-dd HH:mm:ss.fff")
                    Else
                        textValue = v.ToString()
                    End If

                    WriteLine("CONF", $"  {p.Name}: {textValue}")
                Next
            Catch ex As Exception
                LogError("Failed to log ModelConfig.", ex:=ex)
            End Try
        End Sub

        ''' <summary>
        ''' Logs a snapshot of selected tooling-related INI variables prior to calling the main tool-enabled LLM.
        ''' Tool instructions and tool responses are logged as length stubs only (full content is already
        ''' recorded via <see cref="LogModelConfigOnce"/> and <see cref="LogRawResponse"/>).
        ''' </summary>
        Public Shared Sub LogPreMainLlmCallSnapshot()
            If Not _isEnabled Then Return
            WriteLine("LLM", "Pre LLM() snapshot (main tooling LLM):")
            WriteLine("LLM", $"  INI_Model_2: {SafeStr(INI_Model_2)}")
            WriteLine("LLM", $"  INI_APICall_2: {SafeStr(INI_APICall_2)}")

            Dim toolInstr = SafeStr(INI_APICall_ToolInstructions_2)
            WriteLine("LLM", $"  INI_APICall_ToolInstructions_2: ({toolInstr.Length} chars)")

            Dim toolResp = SafeStr(INI_APICall_ToolResponses_2)
            If toolResp.Length <= 500 Then
                WriteLine("LLM", $"  INI_APICall_ToolResponses_2: {toolResp}")
            Else
                Dim excerpt = toolResp.Substring(0, 500) & "..."
                WriteLine("LLM", $"  INI_APICall_ToolResponses_2: ({toolResp.Length} chars) {excerpt}")
            End If

            WriteLine("LLM", $"  INI_Response_2: {SafeStr(INI_Response_2)}")
        End Sub

        ''' <summary>
        ''' Logs a snapshot of selected variables prior to calling an external tool/service via <c>LLM</c>.
        ''' </summary>
        ''' <param name="ctx">Object expected to expose <c>INI_Model_2</c> and <c>INI_APICall_2</c> members.</param>
        Public Shared Sub LogPreToolLlmCallSnapshot(ctx As Object)
            If Not _isEnabled Then Return

            WriteLine("LLM", "Pre LLM() snapshot (tool/service call):")
            Try
                WriteLine("LLM", $"  INI_Model_2: {SafeGetMemberString(ctx, "INI_Model_2")}")
                WriteLine("LLM", $"  INI_APICall_2: {SafeGetMemberString(ctx, "INI_APICall_2")}")
            Catch ex As Exception
                LogError("Failed to capture tool LLM snapshot.", ex:=ex)
            End Try
        End Sub

        ''' <summary>
        ''' Logs raw response content (unmodified) with two blank lines before and after.
        ''' </summary>
        ''' <param name="source">Source label (e.g. "Main LLM()" or tool name).</param>
        ''' <param name="rawResponse">Raw response text.</param>
        Public Shared Sub LogRawResponse(source As String, rawResponse As String)
            If Not _isEnabled OrElse String.IsNullOrWhiteSpace(_logPath) Then Return

            WriteLine("RESP", $"Raw response ({source}) begins:")
            WriteRaw("RESP", vbCrLf & vbCrLf & SafeStr(rawResponse) & vbCrLf & vbCrLf)
            WriteLine("RESP", $"Raw response ({source}) ends.")
        End Sub

        ''' <summary>
        ''' Logs a brief stub of a raw response (length + short excerpt) without the full content.
        ''' Used for main LLM responses to keep the log file focused on tool calls.
        ''' </summary>
        ''' <param name="source">Source label (e.g. "Main LLM()").</param>
        ''' <param name="rawResponse">Raw response text.</param>
        ''' <param name="excerptLength">Maximum number of characters to include in the excerpt.</param>
        Public Shared Sub LogRawResponseStub(source As String, rawResponse As String, Optional excerptLength As Integer = 200)
            If Not _isEnabled OrElse String.IsNullOrWhiteSpace(_logPath) Then Return

            Dim safe As String = SafeStr(rawResponse)
            Dim charCount As Integer = safe.Length
            Dim excerpt As String = If(charCount <= excerptLength,
                safe,
                safe.Substring(0, excerptLength) & "...")

            WriteLine("RESP", $"Raw response ({source}): {charCount} chars")
            If charCount > 0 Then
                WriteLine("RESP", $"Excerpt: {excerpt}")
            End If
        End Sub

        ''' <summary>
        ''' Ends a tooling log session and writes summary/exception details (if provided).
        ''' </summary>
        ''' <param name="success">Whether the session completed successfully.</param>
        ''' <param name="summary">Optional summary string written to the log.</param>
        ''' <param name="ex">Optional exception written to the log.</param>
        Public Shared Sub EndSession(Optional success As Boolean = True, Optional summary As String = "", Optional ex As Exception = Nothing)
            SyncLock _lock
                If _sessionDepth > 0 Then
                    _sessionDepth -= 1
                End If

                If Not _isEnabled Then Return

                If _sessionDepth > 0 Then
                    WriteLine("END", $"Nested tooling session finished; remaining depth={_sessionDepth}")
                    If Not String.IsNullOrWhiteSpace(summary) Then
                        WriteLine("END", $"Nested summary: {summary}")
                    End If
                    If ex IsNot Nothing Then
                        WriteException("END", ex)
                    End If
                    Return
                End If

                WriteLine("END", $"Success: {success}")
                If Not String.IsNullOrWhiteSpace(summary) Then
                    WriteLine("END", $"Summary: {summary}")
                End If
                If ex IsNot Nothing Then
                    WriteException("END", ex)
                End If
                WriteLine("END", $"Ended: {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}")

                FinalizeDiagnosticsFiles()

                _isEnabled = False
                _started = False
                _logPath = Nothing
                _subAgentLogPath = Nothing
                _topLevelSkillSlug = Nothing
                _isDiagnosticsMode = False
                _sessionDepth = 0
            End SyncLock
        End Sub

        ''' <summary>
        ''' Writes a single timestamped log line.
        ''' </summary>
        ''' <param name="category">Category identifier (STEP/WARN/ERR/etc.).</param>
        ''' <param name="message">Message text.</param>
        Private Shared Sub WriteLine(category As String, message As String)
            SyncLock _lock
                Try
                    Dim entry = $"[{DateTime.Now:HH:mm:ss.fff}] [{category}] {message}"
                    File.AppendAllText(_logPath, entry & Environment.NewLine)
                Catch
                End Try
            End SyncLock
        End Sub

        ''' <summary>
        ''' Writes raw text to the log file without adding timestamps.
        ''' </summary>
        ''' <param name="category">Unused category label; retained to match caller signature.</param>
        ''' <param name="raw">Raw text to append.</param>
        Private Shared Sub WriteRaw(category As String, raw As String)
            SyncLock _lock
                Try
                    ' Raw is written as-is; only prefixed by one timestamp line already.
                    File.AppendAllText(_logPath, raw)
                Catch
                End Try
            End SyncLock
        End Sub

        ''' <summary>
        ''' Converts a potentially <c>Nothing</c> string into a non-null string for logging.
        ''' </summary>
        Private Shared Function SafeStr(value As String) As String
            If value Is Nothing Then Return ""
            Return value
        End Function

        ''' <summary>
        ''' Reads a string representation of a named property or field value via reflection.
        ''' </summary>
        ''' <param name="obj">Target object.</param>
        ''' <param name="memberName">Property or field name.</param>
        ''' <returns>Member value converted to string, or an empty string if missing.</returns>
        Private Shared Function SafeGetMemberString(obj As Object, memberName As String) As String
            If obj Is Nothing Then Return ""
            Dim t = obj.GetType()
            Dim p = t.GetProperty(memberName)
            If p IsNot Nothing Then
                Dim v = p.GetValue(obj, Nothing)
                Return If(v IsNot Nothing, v.ToString(), "")
            End If
            Dim f = t.GetField(memberName)
            If f IsNot Nothing Then
                Dim v = f.GetValue(obj)
                Return If(v IsNot Nothing, v.ToString(), "")
            End If
            Return ""
        End Function

        ''' <summary>
        ''' Returns whether file logging is currently enabled.
        ''' </summary>
        Public Shared ReadOnly Property IsEnabled As Boolean
            Get
                Return _isEnabled
            End Get
        End Property

        ''' <summary>
        ''' Returns the currently active log file path (empty/Nothing if not enabled).
        ''' </summary>
        Public Shared ReadOnly Property LogFilePath As String
            Get
                Return _logPath
            End Get
        End Property


        ' --- Sub-agent returns: full payload, separate file ----------
        Private Shared _subAgentLogPath As String = Nothing

        Public Shared Sub LogSubAgentReturn(source As String, rawResponse As String)
            If Not _isEnabled OrElse rawResponse Is Nothing Then Return
            Try
                EnsureSubAgentLogPath()
                If String.IsNullOrEmpty(_subAgentLogPath) Then Return

                Dim sb As New System.Text.StringBuilder()
                sb.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {If(source, "")}")
                sb.AppendLine(If(rawResponse, ""))
                sb.AppendLine("---")

                SyncLock _lock
                    IO.File.AppendAllText(_subAgentLogPath, sb.ToString(), System.Text.Encoding.UTF8)
                End SyncLock
            Catch
                ' Never throw from logger.
            End Try
        End Sub

        Private Shared Sub EnsureSubAgentLogPath()
            If Not String.IsNullOrEmpty(_subAgentLogPath) Then Return
            If String.IsNullOrEmpty(_logPath) Then Return
            Try
                Dim dir As String = IO.Path.GetDirectoryName(_logPath)
                _subAgentLogPath = IO.Path.Combine(dir, $"{AN5}_SubAgent_Returns.txt")
            Catch
                ' If the path cannot be resolved, sub-agent returns silently skip.
            End Try
        End Sub

        ''' <summary>
        ''' Records the top-level skill invoked in this run so diagnostics logs can be
        ''' finalized under per-skill filenames. Only the first (top-level) skill is kept.
        ''' </summary>
        Public Shared Sub SetTopLevelSkillName(skillName As String)
            SyncLock _lock
                If Not _isEnabled Then Return
                If Not String.IsNullOrWhiteSpace(_topLevelSkillSlug) Then Return
                _topLevelSkillSlug = SanitizeSlug(skillName)
            End SyncLock
        End Sub

        ''' <summary>Converts a skill name into a filename-safe slug (letters/digits/-/_).</summary>
        Private Shared Function SanitizeSlug(name As String) As String
            If String.IsNullOrWhiteSpace(name) Then Return Nothing
            Dim sb As New StringBuilder()
            For Each ch In name.Trim()
                If Char.IsLetterOrDigit(ch) OrElse ch = "-"c OrElse ch = "_"c Then
                    sb.Append(ch)
                ElseIf ch = " "c Then
                    sb.Append("-"c)
                End If
            Next
            Dim slug As String = sb.ToString().Trim("-"c, "_"c)
            If slug.Length > 60 Then slug = slug.Substring(0, 60)
            Return If(slug.Length = 0, Nothing, slug)
        End Function

        ''' <summary>
        ''' In diagnostics mode, renames the working tooling and sub-agent log files to
        ''' per-skill names (keep-last-run, overwritten). No-op outside diagnostics mode or
        ''' when no top-level skill was recorded.
        ''' </summary>
        Private Shared Sub FinalizeDiagnosticsFiles()
            If Not _isDiagnosticsMode Then Return
            If String.IsNullOrWhiteSpace(_topLevelSkillSlug) Then Return
            Try
                Dim dir As String = Path.GetDirectoryName(_logPath)

                Dim finalLog As String = Path.Combine(dir, $"{AN5}_Tooling_Log__{_topLevelSkillSlug}.txt")
                MoveOverwrite(_logPath, finalLog)

                If Not String.IsNullOrEmpty(_subAgentLogPath) AndAlso File.Exists(_subAgentLogPath) Then
                    Dim finalSub As String = Path.Combine(dir, $"{AN5}_SubAgent_Returns__{_topLevelSkillSlug}.txt")
                    MoveOverwrite(_subAgentLogPath, finalSub)
                End If
            Catch
                ' Never throw from logger.
            End Try
        End Sub

        ''' <summary>Moves <paramref name="src"/> to <paramref name="dest"/>, overwriting any existing file.</summary>
        Private Shared Sub MoveOverwrite(src As String, dest As String)
            Try
                If String.IsNullOrEmpty(src) OrElse Not File.Exists(src) Then Return
                If File.Exists(dest) Then File.Delete(dest)
                File.Move(src, dest)
            Catch
                ' Never throw from logger.
            End Try
        End Sub


    End Class


    Private Function BuildPromptDiagnosticStub(text As String,
                                            Optional maxExcerptChars As Integer = 120) As String
        Dim raw As String = If(text, "")
        Dim excerpt As String = Regex.Replace(raw, "\s+", " ").Trim()

        If excerpt.Length > maxExcerptChars Then
            excerpt = excerpt.Substring(0, maxExcerptChars) & "..."
        End If

        Dim hashText As String = ""

        Using sha = System.Security.Cryptography.SHA256.Create()
            Dim bytes = sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(raw))
            hashText = BitConverter.ToString(bytes).Replace("-", "").ToLowerInvariant()

            If hashText.Length > 16 Then
                hashText = hashText.Substring(0, 16)
            End If
        End Using

        Return $"len={raw.Length}; sha256={hashText}; excerpt=""{excerpt}"""
    End Function


    Private Sub LogLatestUserRequestDiagnostic(context As ToolExecutionContext, stage As String)
        If context Is Nothing Then
            Return
        End If

        context.Log(
            "latestUserRequestRaw[" & If(stage, "") & "] " &
            BuildPromptDiagnosticStub(context.LatestUserRequestRaw), "diag")
    End Sub

    Private Sub LogFinalResponseContractDiagnostics(context As ToolExecutionContext,
                                                    systemPrompt As String)
        If context Is Nothing Then
            Return
        End If

        Dim requiresTaskStatusFooter As Boolean =
            SharedLibrary.Agents.ToolingFinalResponseContractHelpers.RequiresTaskStatusFooter(
                context.FinalResponseContract)

        Dim normalizedSystemPrompt As String = If(systemPrompt, "")

        Dim hasActiveToolingContractHeader As Boolean =
            normalizedSystemPrompt.IndexOf(
                "ACTIVE-TOOLING CONTRACT:",
                StringComparison.OrdinalIgnoreCase) >= 0

        Dim hasUserFacingFinalAnswerContract As Boolean =
            normalizedSystemPrompt.IndexOf(
                "a user-facing final answer ending with exactly one <TASK_STATUS>",
                StringComparison.OrdinalIgnoreCase) >= 0

        Dim hasUserFacingFinalProseContract As Boolean =
            normalizedSystemPrompt.IndexOf(
                "a user-facing final prose answer ending with exactly one valid <TASK_STATUS>",
                StringComparison.OrdinalIgnoreCase) >= 0

        Dim hasUserFacingBlockedExplanationContract As Boolean =
            normalizedSystemPrompt.IndexOf(
                "a user-facing blocked explanation ending with exactly one valid <TASK_STATUS>",
                StringComparison.OrdinalIgnoreCase) >= 0

        Dim systemPromptContainsTaskStatusContract As Boolean =
            Not String.IsNullOrWhiteSpace(normalizedSystemPrompt) AndAlso
            ((hasActiveToolingContractHeader AndAlso hasUserFacingFinalAnswerContract) OrElse
             hasUserFacingFinalProseContract OrElse
             hasUserFacingBlockedExplanationContract)

        context.Log(
            $"finalResponseContract={SharedLibrary.Agents.ToolingFinalResponseContractHelpers.FormatToolingFinalResponseContract(context.FinalResponseContract)}; requiresTaskStatusFooter={If(requiresTaskStatusFooter, "true", "false")}; systemPromptContainsTaskStatusContract={If(systemPromptContainsTaskStatusContract, "true", "false")}",
            "diag")

        If requiresTaskStatusFooter <> systemPromptContainsTaskStatusContract Then
            context.LogWarn(
                "Final-response contract/prompt mismatch detected.",
                details:=$"host={context.HostKind}; finalResponseContract={SharedLibrary.Agents.ToolingFinalResponseContractHelpers.FormatToolingFinalResponseContract(context.FinalResponseContract)}; requiresTaskStatusFooter={If(requiresTaskStatusFooter, "true", "false")}; systemPromptContainsTaskStatusContract={If(systemPromptContainsTaskStatusContract, "true", "false")}")
        End If
    End Sub


End Class

