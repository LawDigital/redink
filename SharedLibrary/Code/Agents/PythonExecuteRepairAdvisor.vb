' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: PythonExecuteRepairAdvisor.vb
' Purpose: Host-agnostic retry-vs-repair advisor for the python_execute tooling
'          loop. Classifies each python_execute outcome, maintains a compact
'          per-session attempt history with normalized error fingerprints, and
'          annotates the model-facing failure payload with repair guidance so the
'          model stops guessing nonexistent APIs after deterministic errors.
'
' Architecture / How it works:
'  - Stateless entry point Annotate(...) is called by every host adapter
'    (Word tooling, Outlook Local Agent, Outlook AutoPilot) right after a
'    python_execute call returns a payload.
'  - Per-session state is keyed on the caller's session object via a
'    ConditionalWeakTable, so no host type has to change and state is released
'    with the session.
'  - Works with the 0.5.0 agent (classifies on error.code / error.source) and
'    an >= 0.5.1 agent (additionally reads error.exceptionType / error.objectType
'    / error.missingAttribute / error.message when present). Unknown extra fields
'    are ignored, so it never breaks on either agent version.
'  - Never throws: on any parsing problem it returns the original payload
'    unchanged, so it can never destabilize a tooling loop.
' =============================================================================


Option Explicit On
Option Strict On
Option Infer On

Namespace Agents

    ''' <summary>Internal classification of a python_execute outcome, distinct from the agent's raw retryable flag.</summary>
    Public Enum PythonExecuteOutcomeClass
        SUCCESS
        TRANSIENT_FAILURE
        CODE_REPAIR_REQUIRED
        DIAGNOSTIC_RUN_REQUIRED
        NON_RECOVERABLE_FAILURE
        REPAIR_BUDGET_EXHAUSTED
    End Enum

    ''' <summary>Configurable per-session limits governing automatic continuation.</summary>
    Public NotInheritable Class PythonExecuteRepairLimits
        Public Property MaxTransientRetries As System.Int32 = 2
        Public Property MaxCodeRepairs As System.Int32 = 3
        Public Property MaxDiagnosticRuns As System.Int32 = 1
    End Class

    ''' <summary>One recorded attempt in the compact repair history included in every repair prompt.</summary>
    Public NotInheritable Class PythonExecuteAttempt
        Public Property CodeExpression As System.String = System.String.Empty
        Public Property Modification As System.String = System.String.Empty
        Public Property ErrorMessage As System.String = System.String.Empty
        Public Property Fingerprint As System.String = System.String.Empty
        Public Property FingerprintChanged As System.Boolean
    End Class

    ''' <summary>Mutable per-session repair state. One instance per tooling-loop session.</summary>
    Public NotInheritable Class PythonExecuteRepairSession
        Public ReadOnly Property Limits As New PythonExecuteRepairLimits()
        Public ReadOnly Property Attempts As New System.Collections.Generic.List(Of PythonExecuteAttempt)()
        Public Property TransientRetriesUsed As System.Int32
        Public Property CodeRepairsUsed As System.Int32
        Public Property DiagnosticRunsUsed As System.Int32
        Public Property SameFingerprintStreak As System.Int32
        Public Property SymbolSwapStreak As System.Int32
        Public Property LastFingerprint As System.String
        Public Property LastCodeHash As System.String
        Public Property LastCodeText As System.String
        Public Property LastSymbol As System.String
        Public Property LastLine As System.Int32
        Public Property LastExceptionType As System.String
    End Class

    Public NotInheritable Class PythonExecuteRepairAdvisor

        Private Const MaxHistoryEntries As System.Int32 = 6
        Private Const MaxMessageChars As System.Int32 = 300
        Private Const MaxExpressionChars As System.Int32 = 200

        Private Shared ReadOnly SessionsLock As New System.Object()
        Private Shared ReadOnly Sessions As New System.Runtime.CompilerServices.ConditionalWeakTable(Of System.Object, PythonExecuteRepairSession)()
        Private Shared ReadOnly FallbackSessionKey As New System.Object()

        Private Sub New()
        End Sub

        ''' <summary>
        ''' Annotates the model-facing python_execute payload with retry-vs-repair semantics and a compact
        ''' attempt history. Safe to call for every outcome (success resets the session). Never throws;
        ''' returns the original payload unchanged when it cannot be parsed or augmented.
        ''' </summary>
        ''' <param name="sessionKey">The tooling-loop session object (e.g. ToolExecutionContext). May be Nothing.</param>
        ''' <param name="arguments">The original python_execute arguments (used to recover the failing source expression).</param>
        ''' <param name="payloadJson">The model-facing JSON payload returned by the core.</param>
        ''' <param name="success">Whether the core reported success.</param>
        Public Shared Function Annotate(
            sessionKey As System.Object,
            arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object),
            payloadJson As System.String,
            success As System.Boolean
        ) As System.String
            Dim ignoredTerminalReason As System.String = Nothing
            Return Annotate(sessionKey, arguments, payloadJson, success, ignoredTerminalReason)
        End Function

        ''' <summary>
        ''' Overload that also reports, via <paramref name="terminalReason"/>, a non-empty reason string when
        ''' the outcome is terminal (REPAIR_BUDGET_EXHAUSTED or NON_RECOVERABLE_FAILURE). Hosts use this to
        ''' stop offering the tool for the current turn (e.g. by requesting no-tool finalization). Returns an
        ''' empty reason for success and for recoverable (transient / repair / diagnostic) outcomes.
        ''' </summary>
        Public Shared Function Annotate(
            sessionKey As System.Object,
            arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object),
            payloadJson As System.String,
            success As System.Boolean,
            ByRef terminalReason As System.String
        ) As System.String

            terminalReason = System.String.Empty

            If System.String.IsNullOrWhiteSpace(payloadJson) Then
                Return payloadJson
            End If

            Try
                Dim session As PythonExecuteRepairSession = GetOrCreateSession(sessionKey)
                Dim payload As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(payloadJson)

                If success Then
                    ResetSession(session)
                    Return payloadJson
                End If

                Dim errorObj As Newtonsoft.Json.Linq.JObject = TryCast(payload("error"), Newtonsoft.Json.Linq.JObject)
                If errorObj Is Nothing Then
                    Return payloadJson
                End If

                Dim code As System.String = ReadString(errorObj("code"))
                Dim exceptionType As System.String = FirstNonEmpty(ReadString(errorObj("exceptionType")), MapExceptionType(code))
                Dim objectType As System.String = FirstNonEmpty(ReadString(errorObj("objectType")), ReadString(errorObj("object_type")))
                Dim agentRetryable As System.Boolean = ReadBoolean(errorObj("retryable"))
                Dim rawMessage As System.String = FirstNonEmpty(ReadString(errorObj("message")), FriendlyForCode(code))

                Dim fileName As System.String = "code.py"
                Dim line As System.Int32 = 0
                Dim symbol As System.String = System.String.Empty
                Dim sourceSymbol As System.String = System.String.Empty
                Dim sourceObj As Newtonsoft.Json.Linq.JObject = TryCast(errorObj("source"), Newtonsoft.Json.Linq.JObject)
                If sourceObj IsNot Nothing Then
                    fileName = FirstNonEmpty(ReadString(sourceObj("file")), "code.py")
                    line = ReadInt(sourceObj("line"))
                    sourceSymbol = ReadString(sourceObj("symbol"))
                End If

                ' Read the enhanced missing-symbol fields independently of error.source so agents that
                ' omit "source" but provide "missingAttribute"/"missing_symbol" are still handled. Legacy
                ' agents (only source.symbol) and agents that provide neither keep their existing behavior.
                symbol = FirstNonEmpty(
                    sourceSymbol,
                    FirstNonEmpty(
                        ReadString(errorObj("missingAttribute")),
                        ReadString(errorObj("missing_symbol"))))

                Dim codeText As System.String = ReadCodeArgument(arguments)
                Dim codeHash As System.String = ComputeHash(codeText)
                Dim expression As System.String = ExtractExpression(codeText, line, symbol)

                Dim fingerprint As System.String = BuildFingerprint(code, fileName, line, symbol, objectType)
                Dim fingerprintChanged As System.Boolean = session.LastFingerprint Is Nothing OrElse
                                                           Not System.String.Equals(session.LastFingerprint, fingerprint, System.StringComparison.Ordinal)

                Dim unchangedResubmission As System.Boolean =
                    session.LastCodeHash IsNot Nothing AndAlso
                    System.String.Equals(session.LastCodeHash, codeHash, System.StringComparison.Ordinal)

                UpdateProgressCounters(session, fingerprint, fingerprintChanged, exceptionType, line, symbol)

                Dim classification As PythonExecuteOutcomeClass = Classify(session, code, exceptionType, objectType, agentRetryable, unchangedResubmission)

                RecordAttempt(session, expression, unchangedResubmission, codeHash, rawMessage, fingerprint, fingerprintChanged)
                AdvanceBudget(session, classification)

                Dim regression As System.String = DetectSuspiciousRepair(session.LastCodeText, codeText)

                session.LastFingerprint = fingerprint
                session.LastCodeHash = codeHash
                session.LastCodeText = codeText
                session.LastSymbol = symbol
                session.LastLine = line
                session.LastExceptionType = exceptionType

                ' Surface a terminal reason so the host loop can stop offering the tool this turn.
                If classification = PythonExecuteOutcomeClass.REPAIR_BUDGET_EXHAUSTED Then
                    terminalReason = "python_execute repair budget exhausted or no progress across attempts (fingerprint=" & fingerprint & ")."
                ElseIf classification = PythonExecuteOutcomeClass.NON_RECOVERABLE_FAILURE Then
                    terminalReason = "python_execute reported a non-recoverable failure (code=" & code & ")."
                End If

                ' Override the raw retryable flag: only a genuinely transient outcome may repeat unchanged.
                errorObj("retryable") = New Newtonsoft.Json.Linq.JValue(classification = PythonExecuteOutcomeClass.TRANSIENT_FAILURE)
                errorObj("repairable") = New Newtonsoft.Json.Linq.JValue(
                    classification = PythonExecuteOutcomeClass.CODE_REPAIR_REQUIRED OrElse
                    classification = PythonExecuteOutcomeClass.DIAGNOSTIC_RUN_REQUIRED)

                Dim advisor As New Newtonsoft.Json.Linq.JObject(
                    New Newtonsoft.Json.Linq.JProperty("classification", classification.ToString()),
                    New Newtonsoft.Json.Linq.JProperty("exception_type", ToJsonValue(exceptionType)),
                    New Newtonsoft.Json.Linq.JProperty("object_type", ToJsonValue(objectType)),
                    New Newtonsoft.Json.Linq.JProperty("missing_symbol", ToJsonValue(symbol)),
                    New Newtonsoft.Json.Linq.JProperty("fingerprint", fingerprint),
                    New Newtonsoft.Json.Linq.JProperty("fingerprint_changed", fingerprintChanged),
                    New Newtonsoft.Json.Linq.JProperty("unchanged_resubmission", unchangedResubmission),
                    New Newtonsoft.Json.Linq.JProperty("allow_diagnostic_run", classification = PythonExecuteOutcomeClass.DIAGNOSTIC_RUN_REQUIRED),
                    New Newtonsoft.Json.Linq.JProperty("code_repairs_used", session.CodeRepairsUsed),
                    New Newtonsoft.Json.Linq.JProperty("max_code_repairs", session.Limits.MaxCodeRepairs),
                    New Newtonsoft.Json.Linq.JProperty("transient_retries_used", session.TransientRetriesUsed),
                    New Newtonsoft.Json.Linq.JProperty("max_transient_retries", session.Limits.MaxTransientRetries),
                    New Newtonsoft.Json.Linq.JProperty("diagnostic_runs_used", session.DiagnosticRunsUsed),
                    New Newtonsoft.Json.Linq.JProperty("max_diagnostic_runs", session.Limits.MaxDiagnosticRuns),
                    New Newtonsoft.Json.Linq.JProperty("guidance", BuildGuidance(classification, exceptionType, objectType, symbol)),
                    New Newtonsoft.Json.Linq.JProperty("suspicious_repair", ToJsonValue(regression)),
                    New Newtonsoft.Json.Linq.JProperty("attempt_history", BuildHistoryJson(session)))

                errorObj("advisor") = advisor
                Return payload.ToString(Newtonsoft.Json.Formatting.None)

            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine(ex.ToString())
                Return payloadJson
            End Try
        End Function

        ' ─────────────────────────────────────────────────────────────────────
        ' Classification
        ' ─────────────────────────────────────────────────────────────────────

        Private Shared Function Classify(
            session As PythonExecuteRepairSession,
            code As System.String,
            exceptionType As System.String,
            objectType As System.String,
            agentRetryable As System.Boolean,
            unchangedResubmission As System.Boolean
        ) As PythonExecuteOutcomeClass

            If IsFatalCode(code) Then
                Return PythonExecuteOutcomeClass.NON_RECOVERABLE_FAILURE
            End If

            ' No meaningful progress across attempts: stop rather than keep swapping guessed names.
            If session.SameFingerprintStreak >= 2 OrElse session.SymbolSwapStreak >= 3 Then
                Return PythonExecuteOutcomeClass.REPAIR_BUDGET_EXHAUSTED
            End If

            If IsDeterministicCode(code) Then
                If session.CodeRepairsUsed >= session.Limits.MaxCodeRepairs Then
                    Return PythonExecuteOutcomeClass.REPAIR_BUDGET_EXHAUSTED
                End If
                ' An unknown object/attribute mismatch earns one safe diagnostic run before another blind repair.
                If IsAttributeLike(code) AndAlso
                   System.String.IsNullOrEmpty(objectType) AndAlso
                   session.DiagnosticRunsUsed < session.Limits.MaxDiagnosticRuns AndAlso
                   Not unchangedResubmission Then
                    Return PythonExecuteOutcomeClass.DIAGNOSTIC_RUN_REQUIRED
                End If
                Return PythonExecuteOutcomeClass.CODE_REPAIR_REQUIRED
            End If

            ' Non-deterministic outcomes: honor the agent's transient signal within a bounded budget.
            If agentRetryable OrElse IsTransientCode(code) Then
                If session.TransientRetriesUsed >= session.Limits.MaxTransientRetries Then
                    Return PythonExecuteOutcomeClass.REPAIR_BUDGET_EXHAUSTED
                End If
                Return PythonExecuteOutcomeClass.TRANSIENT_FAILURE
            End If

            ' Everything else (unknown runtime failure) is treated as repairable, never an unchanged retry.
            If session.CodeRepairsUsed >= session.Limits.MaxCodeRepairs Then
                Return PythonExecuteOutcomeClass.REPAIR_BUDGET_EXHAUSTED
            End If
            Return PythonExecuteOutcomeClass.CODE_REPAIR_REQUIRED
        End Function

        Private Shared Sub AdvanceBudget(session As PythonExecuteRepairSession, classification As PythonExecuteOutcomeClass)
            Select Case classification
                Case PythonExecuteOutcomeClass.TRANSIENT_FAILURE
                    session.TransientRetriesUsed += 1
                Case PythonExecuteOutcomeClass.CODE_REPAIR_REQUIRED
                    session.CodeRepairsUsed += 1
                Case PythonExecuteOutcomeClass.DIAGNOSTIC_RUN_REQUIRED
                    session.DiagnosticRunsUsed += 1
            End Select
        End Sub

        Private Shared Sub UpdateProgressCounters(
            session As PythonExecuteRepairSession,
            fingerprint As System.String,
            fingerprintChanged As System.Boolean,
            exceptionType As System.String,
            line As System.Int32,
            symbol As System.String
        )
            If Not fingerprintChanged Then
                session.SameFingerprintStreak += 1
            Else
                session.SameFingerprintStreak = 0
            End If

            ' Detect "guess swapping": the same failure at the same site with only the missing name changing.
            Dim swapped As System.Boolean =
                session.LastExceptionType IsNot Nothing AndAlso
                System.String.Equals(session.LastExceptionType, exceptionType, System.StringComparison.Ordinal) AndAlso
                session.LastLine = line AndAlso line > 0 AndAlso
                Not System.String.Equals(If(session.LastSymbol, System.String.Empty), If(symbol, System.String.Empty), System.StringComparison.Ordinal)
            If swapped Then
                session.SymbolSwapStreak += 1
            Else
                session.SymbolSwapStreak = 0
            End If
        End Sub

        Private Shared Sub RecordAttempt(
            session As PythonExecuteRepairSession,
            expression As System.String,
            unchangedResubmission As System.Boolean,
            codeHash As System.String,
            rawMessage As System.String,
            fingerprint As System.String,
            fingerprintChanged As System.Boolean
        )
            Dim modification As System.String
            If session.LastCodeHash Is Nothing Then
                modification = "initial submission"
            ElseIf unchangedResubmission Then
                modification = "code unchanged (disproven approach resubmitted)"
            Else
                modification = "code changed"
            End If

            session.Attempts.Add(New PythonExecuteAttempt() With {
                .CodeExpression = Truncate(expression, MaxExpressionChars),
                .Modification = modification,
                .ErrorMessage = RedactSensitive(Truncate(rawMessage, MaxMessageChars)),
                .Fingerprint = fingerprint,
                .FingerprintChanged = fingerprintChanged
            })

            While session.Attempts.Count > MaxHistoryEntries
                session.Attempts.RemoveAt(0)
            End While
        End Sub

        Private Shared Sub ResetSession(session As PythonExecuteRepairSession)
            session.Attempts.Clear()
            session.TransientRetriesUsed = 0
            session.CodeRepairsUsed = 0
            session.DiagnosticRunsUsed = 0
            session.SameFingerprintStreak = 0
            session.SymbolSwapStreak = 0
            session.LastFingerprint = Nothing
            session.LastCodeHash = Nothing
            session.LastSymbol = Nothing
            session.LastLine = 0
            session.LastExceptionType = Nothing
        End Sub

        ' ─────────────────────────────────────────────────────────────────────
        ' Guidance / history payload
        ' ─────────────────────────────────────────────────────────────────────

        Private Shared Function BuildGuidance(
            classification As PythonExecuteOutcomeClass,
            exceptionType As System.String,
            objectType As System.String,
            symbol As System.String
        ) As System.String
            Select Case classification
                Case PythonExecuteOutcomeClass.TRANSIENT_FAILURE
                    Return "Transient failure. The identical request may be executed again unchanged."
                Case PythonExecuteOutcomeClass.NON_RECOVERABLE_FAILURE
                    Return "Non-recoverable failure. Do not retry or repair; report the situation instead."
                Case PythonExecuteOutcomeClass.REPAIR_BUDGET_EXHAUSTED
                    Return "Repair budget exhausted or no progress was made across attempts. Stop; do not submit another variation. Explain what is blocking the task."
                Case PythonExecuteOutcomeClass.DIAGNOSTIC_RUN_REQUIRED
                    Dim target As System.String = If(System.String.IsNullOrEmpty(symbol), "the value", "'" & symbol & "'")
                    Return "The real object type is unknown. Before changing the code, run a small diagnostic that prints only safe metadata for the failing value: print(type(value)); print(type(value).__module__); print(type(value).__qualname__); print(hasattr(value, " & Chr(34) & "attribute" & Chr(34) & ")). Then call the API that actually exists on that type instead of guessing " & target & "."
                Case Else ' CODE_REPAIR_REQUIRED
                    Dim head As System.String = "This is a deterministic Python error; resubmitting the same code will fail identically. Change the code."
                    If Not System.String.IsNullOrEmpty(exceptionType) AndAlso
                       (exceptionType = "AttributeError" OrElse exceptionType = "TypeError") Then
                        head &= " After an " & exceptionType & ", do not guess a similarly named attribute or method; use the actual API of " & If(System.String.IsNullOrEmpty(objectType), "the real object type", "'" & objectType & "'") & "."
                    End If
                    Return head & " State the concrete root cause, modify only the smallest necessary region, preserve all existing functionality and outputs, do not reuse any approach listed in attempt_history, do not mask the error with a broad try/except, and do not replace required behavior with dummy output."
            End Select
        End Function

        Private Shared Function BuildHistoryJson(session As PythonExecuteRepairSession) As Newtonsoft.Json.Linq.JArray
            Dim array As New Newtonsoft.Json.Linq.JArray()
            For Each attempt As PythonExecuteAttempt In session.Attempts
                array.Add(New Newtonsoft.Json.Linq.JObject(
                    New Newtonsoft.Json.Linq.JProperty("code_expression", attempt.CodeExpression),
                    New Newtonsoft.Json.Linq.JProperty("modification", attempt.Modification),
                    New Newtonsoft.Json.Linq.JProperty("error_message", attempt.ErrorMessage),
                    New Newtonsoft.Json.Linq.JProperty("fingerprint", attempt.Fingerprint),
                    New Newtonsoft.Json.Linq.JProperty("fingerprint_changed", attempt.FingerprintChanged)))
            Next
            Return array
        End Function

        ' ─────────────────────────────────────────────────────────────────────
        ' Error-code taxonomy (matches the safe-error vocabulary of the current agent)
        ' ─────────────────────────────────────────────────────────────────────

        Private Shared Function IsDeterministicCode(code As System.String) As System.Boolean
            Select Case code
                Case "PYTHON_SYNTAX_ERROR", "PYTHON_NAME_ERROR", "PYTHON_IMPORT_ERROR",
                     "PYTHON_ATTRIBUTE_ERROR", "PYTHON_TYPE_ERROR", "PYTHON_VALUE_ERROR",
                     "PYTHON_KEY_ERROR", "PYTHON_INDEX_ERROR"
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Private Shared Function IsAttributeLike(code As System.String) As System.Boolean
            Return code = "PYTHON_ATTRIBUTE_ERROR" OrElse code = "PYTHON_TYPE_ERROR"
        End Function

        Private Shared Function IsTransientCode(code As System.String) As System.Boolean
            Select Case code
                Case "LLM_RATE_LIMITED", "LLM_PROVIDER_UNAVAILABLE", "WEB_REQUEST_TIMEOUT",
                     "HOST_CALL_TIMEOUT", "ROOT_BUSY"
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Private Shared Function IsFatalCode(code As System.String) As System.Boolean
            Select Case code
                Case "EXECUTABLE_NOT_FOUND", "EXECUTABLE_HASH_MISMATCH", "EXECUTABLE_SIGNATURE_INVALID",
                     "EXECUTABLE_SIGNER_MISMATCH", "CONFIGURATION_INVALID", "SECURITY_INVARIANT_FAILED",
                     "SESSION_CANCELLED"
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Private Shared Function MapExceptionType(code As System.String) As System.String
            Select Case code
                Case "PYTHON_SYNTAX_ERROR" : Return "SyntaxError"
                Case "PYTHON_NAME_ERROR" : Return "NameError"
                Case "PYTHON_IMPORT_ERROR" : Return "ImportError"
                Case "PYTHON_ATTRIBUTE_ERROR" : Return "AttributeError"
                Case "PYTHON_TYPE_ERROR" : Return "TypeError"
                Case "PYTHON_VALUE_ERROR" : Return "ValueError"
                Case "PYTHON_KEY_ERROR" : Return "KeyError"
                Case "PYTHON_INDEX_ERROR" : Return "IndexError"
                Case "PYTHON_FILE_NOT_FOUND" : Return "FileNotFoundError"
                Case "PYTHON_PERMISSION_ERROR" : Return "PermissionError"
                Case "PYTHON_RUNTIME_ERROR" : Return "RuntimeError"
                Case Else : Return System.String.Empty
            End Select
        End Function

        Private Shared Function FriendlyForCode(code As System.String) As System.String
            If System.String.IsNullOrEmpty(code) Then Return System.String.Empty
            Return code
        End Function

        ' ─────────────────────────────────────────────────────────────────────
        ' Fingerprinting & sanitization
        ' ─────────────────────────────────────────────────────────────────────

        ''' <summary>
        ''' Compares the previous submission with the new one and returns a warning when the "repair"
        ''' looks destructive rather than minimal: removed output generation (publish_result / output_path),
        ''' removed validation, a newly introduced broad try/except, or a large net deletion. Returns an
        ''' empty string when the change looks like a legitimate targeted fix. Advisory only.
        ''' </summary>
        Private Shared Function DetectSuspiciousRepair(previousCode As System.String, newCode As System.String) As System.String
            If System.String.IsNullOrEmpty(previousCode) OrElse System.String.IsNullOrEmpty(newCode) Then
                Return System.String.Empty
            End If

            Dim warnings As New System.Collections.Generic.List(Of System.String)()

            Dim previousOutputs As System.Int32 = CountOccurrences(previousCode, "publish_result") + CountOccurrences(previousCode, "output_path")
            Dim newOutputs As System.Int32 = CountOccurrences(newCode, "publish_result") + CountOccurrences(newCode, "output_path")
            If previousOutputs > 0 AndAlso newOutputs < previousOutputs Then
                warnings.Add("output generation (publish_result/output_path) was removed or reduced")
            End If

            Dim previousDefs As System.Int32 = CountDefinitions(previousCode)
            Dim newDefs As System.Int32 = CountDefinitions(newCode)
            If newDefs < previousDefs Then
                warnings.Add("one or more function/class definitions were removed")
            End If

            If HasBroadExcept(newCode) AndAlso Not HasBroadExcept(previousCode) Then
                warnings.Add("a broad try/except was introduced, which can mask the real error")
            End If

            Dim previousLines As System.Int32 = CountNonEmptyLines(previousCode)
            Dim newLines As System.Int32 = CountNonEmptyLines(newCode)
            If previousLines >= 8 AndAlso newLines * 2 < previousLines Then
                warnings.Add("more than half of the code was deleted, which is unlikely to be a minimal repair")
            End If

            If warnings.Count = 0 Then
                Return System.String.Empty
            End If

            Return "The proposed change does not look like a minimal repair (" & System.String.Join("; ", warnings) &
                   "). Restore the removed functionality and outputs, then fix only the smallest region that caused the error."
        End Function

        Private Shared Function CountOccurrences(text As System.String, needle As System.String) As System.Int32
            If System.String.IsNullOrEmpty(text) OrElse System.String.IsNullOrEmpty(needle) Then Return 0
            Dim count As System.Int32 = 0
            Dim index As System.Int32 = text.IndexOf(needle, System.StringComparison.Ordinal)
            While index >= 0
                count += 1
                index = text.IndexOf(needle, index + needle.Length, System.StringComparison.Ordinal)
            End While
            Return count
        End Function

        Private Shared Function CountDefinitions(text As System.String) As System.Int32
            Return System.Text.RegularExpressions.Regex.Matches(
                text, "(?m)^\s*(?:async\s+)?(?:def|class)\s+[A-Za-z_]").Count
        End Function

        Private Shared Function HasBroadExcept(text As System.String) As System.Boolean
            Return System.Text.RegularExpressions.Regex.IsMatch(
                text, "(?m)^\s*except\s*(?::|\bException\b\s*(?:as\s+[A-Za-z_]\w*)?\s*:|BaseException\b)")
        End Function

        Private Shared Function CountNonEmptyLines(text As System.String) As System.Int32
            Dim count As System.Int32 = 0
            For Each raw As System.String In text.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Split(CChar(vbLf))
                If Not System.String.IsNullOrWhiteSpace(raw) Then count += 1
            Next
            Return count
        End Function

        Private Shared Function BuildFingerprint(code As System.String, fileName As System.String, line As System.Int32, symbol As System.String, objectType As System.String) As System.String
            Dim parts As New System.Collections.Generic.List(Of System.String) From {
                If(System.String.IsNullOrEmpty(code), "UNKNOWN", code),
                If(System.String.IsNullOrEmpty(fileName), "code.py", fileName),
                If(line > 0, line.ToString(System.Globalization.CultureInfo.InvariantCulture), "0"),
                If(System.String.IsNullOrEmpty(symbol), "-", symbol),
                If(System.String.IsNullOrEmpty(objectType), "-", objectType)
            }
            Return System.String.Join("|", parts)
        End Function

        Private Shared Function ExtractExpression(codeText As System.String, line As System.Int32, symbol As System.String) As System.String
            If System.String.IsNullOrEmpty(codeText) OrElse line < 1 Then
                Return If(symbol, System.String.Empty)
            End If
            Dim lines As System.String() = codeText.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Split(CChar(vbLf))
            If line > lines.Length Then
                Return If(symbol, System.String.Empty)
            End If
            Return lines(line - 1).Trim()
        End Function

        Private Shared Function ReadCodeArgument(arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object)) As System.String
            If arguments Is Nothing Then Return System.String.Empty
            Dim value As System.Object = Nothing
            If arguments.TryGetValue("code", value) AndAlso value IsNot Nothing Then
                Return System.Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture)
            End If
            Return System.String.Empty
        End Function

        Private Shared Function ComputeHash(value As System.String) As System.String
            If value Is Nothing Then value = System.String.Empty
            Using sha As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
                Dim bytes As System.Byte() = sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(value))
                Return System.BitConverter.ToString(bytes).Replace("-", System.String.Empty)
            End Using
        End Function

        ''' <summary>Removes host paths, UNC prefixes, GUIDs and long hex runs so no session-specific data leaks to the model.</summary>
        Private Shared Function RedactSensitive(value As System.String) As System.String
            If System.String.IsNullOrEmpty(value) Then Return System.String.Empty
            Dim result As System.String = value
            result = System.Text.RegularExpressions.Regex.Replace(result, "[A-Za-z]:\\[^\s""']*", "<redacted-path>")
            result = System.Text.RegularExpressions.Regex.Replace(result, "\\\\[^\s""']+", "<redacted-path>")
            result = System.Text.RegularExpressions.Regex.Replace(result, "[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}", "<redacted-id>")
            result = System.Text.RegularExpressions.Regex.Replace(result, "\b[0-9a-fA-F]{32,}\b", "<redacted-id>")
            Return result
        End Function

        ' ─────────────────────────────────────────────────────────────────────
        ' Small helpers
        ' ─────────────────────────────────────────────────────────────────────

        Private Shared Function GetOrCreateSession(sessionKey As System.Object) As PythonExecuteRepairSession
            Dim key As System.Object = If(sessionKey, FallbackSessionKey)
            SyncLock SessionsLock
                Dim session As PythonExecuteRepairSession = Nothing
                If Not Sessions.TryGetValue(key, session) Then
                    session = New PythonExecuteRepairSession()
                    Sessions.Add(key, session)
                End If
                Return session
            End SyncLock
        End Function

        Private Shared Function ReadString(token As Newtonsoft.Json.Linq.JToken) As System.String
            If token Is Nothing OrElse token.Type = Newtonsoft.Json.Linq.JTokenType.Null Then Return System.String.Empty
            Return System.Convert.ToString(token, System.Globalization.CultureInfo.InvariantCulture)
        End Function

        Private Shared Function ReadInt(token As Newtonsoft.Json.Linq.JToken) As System.Int32
            If token Is Nothing OrElse token.Type = Newtonsoft.Json.Linq.JTokenType.Null Then Return 0
            Dim parsed As System.Int32
            If System.Int32.TryParse(ReadString(token), System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, parsed) Then Return parsed
            Return 0
        End Function

        Private Shared Function ReadBoolean(token As Newtonsoft.Json.Linq.JToken) As System.Boolean
            If token Is Nothing OrElse token.Type = Newtonsoft.Json.Linq.JTokenType.Null Then Return False
            Dim parsed As System.Boolean
            If System.Boolean.TryParse(ReadString(token), parsed) Then Return parsed
            Return False
        End Function

        Private Shared Function FirstNonEmpty(ParamArray candidates As System.String()) As System.String
            If candidates IsNot Nothing Then
                For Each candidate As System.String In candidates
                    If Not System.String.IsNullOrEmpty(candidate) Then Return candidate
                Next
            End If
            Return System.String.Empty
        End Function

        Private Shared Function ToJsonValue(value As System.String) As Newtonsoft.Json.Linq.JToken
            If System.String.IsNullOrEmpty(value) Then Return Newtonsoft.Json.Linq.JValue.CreateNull()
            Return New Newtonsoft.Json.Linq.JValue(value)
        End Function

        Private Shared Function Truncate(value As System.String, maxChars As System.Int32) As System.String
            If System.String.IsNullOrEmpty(value) Then Return System.String.Empty
            If value.Length <= maxChars Then Return value
            Return value.Substring(0, maxChars) & "…"
        End Function

    End Class

End Namespace
