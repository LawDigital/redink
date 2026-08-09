' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: PythonExecuteRepairAdvisorSelfTests.vb
' Purpose: Self-tests for PythonExecuteRepairAdvisor: verifies retry-vs-repair
'          classification, error fingerprinting, diagnostic runs, rejected API
'          guessing, minimal-repair (functionality-preservation) detection, and
'          repair-budget exhaustion. Follows the existing DEBUG self-test pattern.
' =============================================================================

#If DEBUG Then

Option Strict On
Option Explicit On

Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.Agents

Namespace AgentsXX

    Public NotInheritable Class PythonExecuteRepairAdvisorSelfTests

        Private Sub New()
        End Sub

        Public Shared Sub RunAll()
            RunNamedTest(NameOf(TestDeterministicErrorIsNotRetryable), AddressOf TestDeterministicErrorIsNotRetryable)
            RunNamedTest(NameOf(TestUnchangedResubmissionIsFlagged), AddressOf TestUnchangedResubmissionIsFlagged)
            RunNamedTest(NameOf(TestUnknownAttributeErrorRequestsDiagnosticRun), AddressOf TestUnknownAttributeErrorRequestsDiagnosticRun)
            RunNamedTest(NameOf(TestRepeatedFingerprintExhaustsBudget), AddressOf TestRepeatedFingerprintExhaustsBudget)
            RunNamedTest(NameOf(TestGuessSwappingIsRejected), AddressOf TestGuessSwappingIsRejected)
            RunNamedTest(NameOf(TestSuspiciousRepairRemovingOutputIsFlagged), AddressOf TestSuspiciousRepairRemovingOutputIsFlagged)
            RunNamedTest(NameOf(TestTransientFailureRemainsRetryable), AddressOf TestTransientFailureRemainsRetryable)
            RunNamedTest(NameOf(TestSuccessResetsSession), AddressOf TestSuccessResetsSession)
            RunNamedTest(NameOf(TestFingerprintIsStableAndSanitized), AddressOf TestFingerprintIsStableAndSanitized)
            RunNamedTest(NameOf(TestSymbolFromLegacySourceSymbol), AddressOf TestSymbolFromLegacySourceSymbol)
            RunNamedTest(NameOf(TestSymbolFromEnhancedMissingAttribute), AddressOf TestSymbolFromEnhancedMissingAttribute)
            RunNamedTest(NameOf(TestSymbolFromEnhancedMissingSymbol), AddressOf TestSymbolFromEnhancedMissingSymbol)
            RunNamedTest(NameOf(TestSymbolAbsentWhenNoSourceAndNoEnhancedFields), AddressOf TestSymbolAbsentWhenNoSourceAndNoEnhancedFields)
            RunNamedTest(NameOf(TestUnchangedResubmissionBlockedBeforeWorker), AddressOf TestUnchangedResubmissionBlockedBeforeWorker)
            RunNamedTest(NameOf(TestChangedCodeIsNotBlockedBeforeWorker), AddressOf TestChangedCodeIsNotBlockedBeforeWorker)
            RunNamedTest(NameOf(TestBoundaryHintForPublishedTuple), AddressOf TestBoundaryHintForPublishedTuple)
            RunNamedTest(NameOf(TestBoundaryHintForWindowsPathFilename), AddressOf TestBoundaryHintForWindowsPathFilename)
            RunNamedTest(NameOf(TestBoundaryHintForHeaderTextAttribute), AddressOf TestBoundaryHintForHeaderTextAttribute)
            RunNamedTest(NameOf(TestTransientBudgetIsExhausted), AddressOf TestTransientBudgetIsExhausted)
            RunNamedTest(NameOf(TestDiagnosticBudgetThenCodeRepair), AddressOf TestDiagnosticBudgetThenCodeRepair)
            RunNamedTest(NameOf(TestSuspiciousRepairIsRejectedNotJustFlagged), AddressOf TestSuspiciousRepairIsRejectedNotJustFlagged)
            RunNamedTest(NameOf(TestPostconditionRejectsNoObservableOutcome), AddressOf TestPostconditionRejectsNoObservableOutcome)
            RunNamedTest(NameOf(TestPostconditionRejectsEmptyOutputFile), AddressOf TestPostconditionRejectsEmptyOutputFile)
            RunNamedTest(NameOf(TestPostconditionAcceptsValidSuccess), AddressOf TestPostconditionAcceptsValidSuccess)
            RunNamedTest(NameOf(TestStableSessionHistoryGrowsAcrossRetries), AddressOf TestStableSessionHistoryGrowsAcrossRetries)
            RunNamedTest(NameOf(TestOptionalDiagnosticFieldsSurviveAnnotation), AddressOf TestOptionalDiagnosticFieldsSurviveAnnotation)
            RunNamedTest(NameOf(TestSuspiciousRepairBlockedBeforeWorker), AddressOf TestSuspiciousRepairBlockedBeforeWorker)
            RunNamedTest(NameOf(TestRemovalOfPublishResultRejected), AddressOf TestRemovalOfPublishResultRejected)
            RunNamedTest(NameOf(TestRemovalOfOutputPathRejected), AddressOf TestRemovalOfOutputPathRejected)
            RunNamedTest(NameOf(TestRemovalOfFunctionRejected), AddressOf TestRemovalOfFunctionRejected)
            RunNamedTest(NameOf(TestBroadExceptRejected), AddressOf TestBroadExceptRejected)
            RunNamedTest(NameOf(TestBareExceptRejected), AddressOf TestBareExceptRejected)
            RunNamedTest(NameOf(TestNarrowExceptNotRejected), AddressOf TestNarrowExceptNotRejected)
            RunNamedTest(NameOf(TestMajorityDeletionRejected), AddressOf TestMajorityDeletionRejected)
            RunNamedTest(NameOf(TestSmallValidRepairAllowed), AddressOf TestSmallValidRepairAllowed)
            RunNamedTest(NameOf(TestPostconditionAcceptsNonEmptyOutputFile), AddressOf TestPostconditionAcceptsNonEmptyOutputFile)
            RunNamedTest(NameOf(TestPostconditionExemptsDiagnosticRun), AddressOf TestPostconditionExemptsDiagnosticRun)
            RunNamedTest(NameOf(TestPostconditionDoesNotReplaceWorkerFailure), AddressOf TestPostconditionDoesNotReplaceWorkerFailure)
            RunNamedTest(NameOf(TestPostconditionEntersHistoryAndBudget), AddressOf TestPostconditionEntersHistoryAndBudget)
            RunNamedTest(NameOf(TestSuccessAfterPostconditionRepairResetsState), AddressOf TestSuccessAfterPostconditionRepairResetsState)
            RunNamedTest(NameOf(TestInputReferenceFailureRequestsArgumentRepair), AddressOf TestInputReferenceFailureRequestsArgumentRepair)
            RunNamedTest(NameOf(TestRequestInvalidIsNotCodeRepair), AddressOf TestRequestInvalidIsNotCodeRepair)
            RunNamedTest(NameOf(TestArgumentUnchangedResubmissionFlagged), AddressOf TestArgumentUnchangedResubmissionFlagged)
        End Sub

        Public Shared Function RunAllAndReturnStatus() As String
            Try
                RunAll()
                Return "PythonExecuteRepairAdvisor self-tests passed."
            Catch ex As Exception
                Debug.WriteLine("[PythonExecuteRepairAdvisorSelfTests] FAILED :: " & ex.ToString())
                Throw
            End Try
        End Function

        ' ── Helpers ──────────────────────────────────────────────────────────

        Private Shared Function Args(code As String) As System.Collections.Generic.Dictionary(Of String, Object)
            Return New System.Collections.Generic.Dictionary(Of String, Object)(StringComparer.Ordinal) From {{"code", code}}
        End Function

        Private Shared Function FailurePayload(errorCode As String, line As Integer, symbol As String,
                                               Optional retryable As Boolean = False,
                                               Optional objectType As String = Nothing,
                                               Optional message As String = Nothing) As String
            Dim errorObj As New JObject(
                New JProperty("code", errorCode),
                New JProperty("phase", "execute"),
                New JProperty("retryable", retryable),
                New JProperty("source", New JObject(
                    New JProperty("file", "code.py"),
                    New JProperty("line", line),
                    New JProperty("column", Nothing),
                    New JProperty("function", Nothing),
                    New JProperty("symbol", If(symbol Is Nothing, CType(JValue.CreateNull(), JToken), New JValue(symbol))))),
                New JProperty("stack", New JArray()))
            If objectType IsNot Nothing Then errorObj("objectType") = objectType
            If message IsNot Nothing Then errorObj("message") = message
            Return New JObject(
                New JProperty("status", "failed"),
                New JProperty("result", JValue.CreateNull()),
                New JProperty("output_files", New JArray()),
                New JProperty("error", errorObj)).ToString(Formatting.None)
        End Function

        ''' <summary>
        ''' Builds a failure payload with fine-grained control over the enhanced symbol fields, so tests can
        ''' exercise legacy (source.symbol), enhanced (missingAttribute / missing_symbol), and no-source cases.
        ''' </summary>
        Private Shared Function FailurePayloadCustom(errorCode As String, line As Integer,
                                                     Optional includeSource As Boolean = True,
                                                     Optional sourceSymbol As String = Nothing,
                                                     Optional missingAttribute As String = Nothing,
                                                     Optional missingSymbol As String = Nothing) As String
            Dim errorObj As New JObject(
                New JProperty("code", errorCode),
                New JProperty("phase", "execute"),
                New JProperty("retryable", False),
                New JProperty("stack", New JArray()))
            If includeSource Then
                errorObj("source") = New JObject(
                    New JProperty("file", "code.py"),
                    New JProperty("line", line),
                    New JProperty("column", Nothing),
                    New JProperty("function", Nothing),
                    New JProperty("symbol", If(sourceSymbol Is Nothing, CType(JValue.CreateNull(), JToken), New JValue(sourceSymbol))))
            End If
            If missingAttribute IsNot Nothing Then errorObj("missingAttribute") = missingAttribute
            If missingSymbol IsNot Nothing Then errorObj("missing_symbol") = missingSymbol
            Return New JObject(
                New JProperty("status", "failed"),
                New JProperty("result", JValue.CreateNull()),
                New JProperty("output_files", New JArray()),
                New JProperty("error", errorObj)).ToString(Formatting.None)
        End Function

        Private Shared Function AnnotatedError(sessionKey As Object, code As String, payload As String, success As Boolean,
                                               ByRef terminalReason As String) As JObject
            Dim annotated As String = PythonExecuteRepairAdvisor.Annotate(
                sessionKey, Args(code), payload, success, terminalReason)
            Return CType(JObject.Parse(annotated)("error"), JObject)
        End Function

        ' ── Tests ────────────────────────────────────────────────────────────

        Private Shared Sub TestDeterministicErrorIsNotRetryable()
            Dim session As New Object()
            Dim terminal As String = Nothing
            Dim err = AnnotatedError(session, "print(unknown_name)",
                FailurePayload("PYTHON_NAME_ERROR", 1, "unknown_name", retryable:=True), False, terminal)

            AssertFalse(err.Value(Of Boolean)("retryable"), "A deterministic Python error must never be retryable, even if the agent said retryable=true.")
            AssertTrue(err("advisor") IsNot Nothing, "Advisor block must be attached.")
            AssertEqual("CODE_REPAIR_REQUIRED", err("advisor").Value(Of String)("classification"), "Deterministic error should require code repair.")
        End Sub

        Private Shared Sub TestUnchangedResubmissionIsFlagged()
            Dim session As New Object()
            Dim terminal As String = Nothing
            Dim code As String = "print(x.foo)"
            PythonExecuteRepairAdvisor.Annotate(session, Args(code), FailurePayload("PYTHON_NAME_ERROR", 1, "x"), False, terminal)
            Dim err = AnnotatedError(session, code, FailurePayload("PYTHON_NAME_ERROR", 1, "x"), False, terminal)

            AssertTrue(err("advisor").Value(Of Boolean)("unchanged_resubmission"), "Resubmitting identical code must be flagged.")
        End Sub

        Private Shared Sub TestUnknownAttributeErrorRequestsDiagnosticRun()
            Dim session As New Object()
            Dim terminal As String = Nothing
            ' No objectType provided -> the type is unknown -> one diagnostic run is offered.
            Dim err = AnnotatedError(session, "print(value.tostring())",
                FailurePayload("PYTHON_ATTRIBUTE_ERROR", 3, "tostring"), False, terminal)

            AssertEqual("DIAGNOSTIC_RUN_REQUIRED", err("advisor").Value(Of String)("classification"), "Unknown-type AttributeError should request a diagnostic run.")
            AssertTrue(err("advisor").Value(Of Boolean)("allow_diagnostic_run"), "Diagnostic run flag should be set.")
        End Sub

        Private Shared Sub TestRepeatedFingerprintExhaustsBudget()
            Dim session As New Object()
            Dim terminal As String = Nothing
            Dim payload As String = FailurePayload("PYTHON_TYPE_ERROR", 5, "join", objectType:="builtins.list")

            ' Same fingerprint repeated: after enough no-progress repeats the loop must stop.
            For i As Integer = 1 To 3
                PythonExecuteRepairAdvisor.Annotate(session, Args("a = " & i.ToString()), payload, False, terminal)
            Next

            AssertFalse(String.IsNullOrEmpty(terminal), "A repeated fingerprint with no progress must yield a terminal reason.")
        End Sub

        Private Shared Sub TestGuessSwappingIsRejected()
            Dim session As New Object()
            Dim terminal As String = Nothing
            ' Same error type at the same line, only the missing symbol changes each time = guess swapping.
            PythonExecuteRepairAdvisor.Annotate(session, Args("h.text"), FailurePayload("PYTHON_ATTRIBUTE_ERROR", 2, "text", objectType:="docx._Header"), False, terminal)
            PythonExecuteRepairAdvisor.Annotate(session, Args("h.value"), FailurePayload("PYTHON_ATTRIBUTE_ERROR", 2, "value", objectType:="docx._Header"), False, terminal)
            PythonExecuteRepairAdvisor.Annotate(session, Args("h.content"), FailurePayload("PYTHON_ATTRIBUTE_ERROR", 2, "content", objectType:="docx._Header"), False, terminal)
            Dim err = AnnotatedError(session, "h.body", FailurePayload("PYTHON_ATTRIBUTE_ERROR", 2, "body", objectType:="docx._Header"), False, terminal)

            AssertEqual("REPAIR_BUDGET_EXHAUSTED", err("advisor").Value(Of String)("classification"), "Repeated single-name guessing should stop the loop.")
            AssertFalse(String.IsNullOrEmpty(terminal), "Guess swapping should be terminal.")
        End Sub

        Private Shared Sub TestSuspiciousRepairRemovingOutputIsFlagged()
            Dim session As New Object()
            Dim terminal As String = Nothing
            Dim first As String =
                "def build():" & vbLf &
                "    data = compute()" & vbLf &
                "    agent_api.publish_result(data)" & vbLf &
                "    agent_api.output_path('out.docx')" & vbLf &
                "build()"
            PythonExecuteRepairAdvisor.Annotate(session, Args(first), FailurePayload("PYTHON_VALUE_ERROR", 2, "compute"), False, terminal)

            ' "Repair" that deletes the output generation and function.
            Dim second As String = "print('done')"
            Dim err = AnnotatedError(session, second, FailurePayload("PYTHON_VALUE_ERROR", 1, Nothing), False, terminal)

            Dim suspicious = err("advisor")("suspicious_repair")
            AssertTrue(suspicious IsNot Nothing AndAlso suspicious.Type <> JTokenType.Null, "Removing output generation must be flagged as a suspicious repair.")
        End Sub

        Private Shared Sub TestTransientFailureRemainsRetryable()
            Dim session As New Object()
            Dim terminal As String = Nothing
            Dim err = AnnotatedError(session, "x = agent_api.web_get('https://example.org')",
                FailurePayload("WEB_REQUEST_TIMEOUT", 1, Nothing, retryable:=True), False, terminal)

            AssertTrue(err.Value(Of Boolean)("retryable"), "A transient failure should remain retryable.")
            AssertEqual("TRANSIENT_FAILURE", err("advisor").Value(Of String)("classification"), "Transient code should classify as transient.")
            AssertTrue(String.IsNullOrEmpty(terminal), "A transient failure is not terminal.")
        End Sub

        Private Shared Sub TestSuccessResetsSession()
            Dim session As New Object()
            Dim terminal As String = Nothing
            PythonExecuteRepairAdvisor.Annotate(session, Args("h.text"), FailurePayload("PYTHON_ATTRIBUTE_ERROR", 2, "text"), False, terminal)

            Dim successPayload As String = New JObject(
                New JProperty("status", "success"),
                New JProperty("result", New JObject(New JProperty("kind", "text"), New JProperty("value", "ok"))),
                New JProperty("output_files", New JArray()),
                New JProperty("error", JValue.CreateNull())).ToString(Formatting.None)
            PythonExecuteRepairAdvisor.Annotate(session, Args("h.text"), successPayload, True, terminal)

            ' After a success the same code is no longer an "unchanged resubmission".
            Dim err = AnnotatedError(session, "h.text", FailurePayload("PYTHON_ATTRIBUTE_ERROR", 2, "text"), False, terminal)
            AssertFalse(err("advisor").Value(Of Boolean)("unchanged_resubmission"), "Success must reset the per-session history.")
        End Sub

        Private Shared Sub TestFingerprintIsStableAndSanitized()
            Dim session1 As New Object()
            Dim session2 As New Object()
            Dim terminal As String = Nothing
            Dim payloadWithPath As String = FailurePayload("PYTHON_ATTRIBUTE_ERROR", 45, "tostring",
                objectType:="lxml.etree._Element",
                message:="'_Element' object has no attribute 'tostring' at C:\Users\temp\abcd1234\code.py")

            Dim errA = AnnotatedError(session1, "x", payloadWithPath, False, terminal)
            Dim errB = AnnotatedError(session2, "x", payloadWithPath, False, terminal)

            AssertEqual("PYTHON_ATTRIBUTE_ERROR|code.py|45|tostring|lxml.etree._Element",
                errA("advisor").Value(Of String)("fingerprint"), "Fingerprint format mismatch.")
            AssertEqual(errA("advisor").Value(Of String)("fingerprint"),
                errB("advisor").Value(Of String)("fingerprint"), "Fingerprint must be stable across sessions for the same failure.")

            Dim history = CType(errA("advisor")("attempt_history"), JArray)
            AssertTrue(history.Count >= 1, "Attempt history should contain the recorded attempt.")
            Dim recordedMessage As String = history(0).Value(Of String)("error_message")
            AssertFalse(recordedMessage.Contains("C:\"), "Host paths must be redacted from recorded messages.")
        End Sub

        Private Shared Sub TestSymbolFromLegacySourceSymbol()
            Dim session As New Object()
            Dim terminal As String = Nothing
            Dim err = AnnotatedError(session, "print(h.text)",
                FailurePayloadCustom("PYTHON_ATTRIBUTE_ERROR", 1, includeSource:=True, sourceSymbol:="text"), False, terminal)

            AssertEqual("text", err("advisor").Value(Of String)("missing_symbol"), "Legacy source.symbol should populate the missing symbol.")
            AssertEqual("PYTHON_ATTRIBUTE_ERROR|code.py|1|text|-", err("advisor").Value(Of String)("fingerprint"), "Legacy symbol must appear in the fingerprint.")
        End Sub

        Private Shared Sub TestSymbolFromEnhancedMissingAttribute()
            Dim session As New Object()
            Dim terminal As String = Nothing
            Dim err = AnnotatedError(session, "print(h.tostring())",
                FailurePayloadCustom("PYTHON_ATTRIBUTE_ERROR", 2, includeSource:=True, sourceSymbol:=Nothing, missingAttribute:="tostring"), False, terminal)

            AssertEqual("tostring", err("advisor").Value(Of String)("missing_symbol"), "Enhanced missingAttribute should populate the missing symbol when source.symbol is absent.")
        End Sub

        Private Shared Sub TestSymbolFromEnhancedMissingSymbol()
            Dim session As New Object()
            Dim terminal As String = Nothing
            Dim err = AnnotatedError(session, "print(unknown_name)",
                FailurePayloadCustom("PYTHON_NAME_ERROR", 3, includeSource:=True, sourceSymbol:=Nothing, missingSymbol:="unknown_name"), False, terminal)

            AssertEqual("unknown_name", err("advisor").Value(Of String)("missing_symbol"), "Enhanced missing_symbol should populate the missing symbol.")
        End Sub

        Private Shared Sub TestSymbolAbsentWhenNoSourceAndNoEnhancedFields()
            Dim session As New Object()
            Dim terminal As String = Nothing
            ' Payload without "source" and without the enhanced fields: symbol falls back to empty ("-").
            Dim err = AnnotatedError(session, "print(x)",
                FailurePayloadCustom("PYTHON_NAME_ERROR", 0, includeSource:=False), False, terminal)

            Dim missing = err("advisor")("missing_symbol")
            AssertTrue(missing Is Nothing OrElse missing.Type = JTokenType.Null, "Missing symbol should be null when no source and no enhanced fields are present.")
            AssertEqual("PYTHON_NAME_ERROR|code.py|0|-|-", err("advisor").Value(Of String)("fingerprint"), "Fingerprint should use placeholders when no symbol is available.")
        End Sub

        ''' <summary>
        ''' Builds a SUCCESS payload matching the core schema (result object plus output_files with a "bytes"
        ''' field), so task-postcondition tests can control whether an observable/valid outcome exists.
        ''' </summary>
        Private Shared Function SuccessPayload(Optional includeResult As Boolean = True, Optional outputBytes As Integer() = Nothing) As String
            Dim resultToken As JToken
            If includeResult Then
                resultToken = New JObject(New JProperty("kind", "text"), New JProperty("value", "ok"))
            Else
                resultToken = JValue.CreateNull()
            End If
            Dim outputs As New JArray()
            If outputBytes IsNot Nothing Then
                Dim index As Integer = 0
                For Each size As Integer In outputBytes
                    outputs.Add(New JObject(
                        New JProperty("name", "out" & index.ToString() & ".bin"),
                        New JProperty("media_type", "application/octet-stream"),
                        New JProperty("bytes", size),
                        New JProperty("sha256", "")))
                    index += 1
                Next
            End If
            Return New JObject(
                New JProperty("status", "success"),
                New JProperty("exit_code", 0),
                New JProperty("duration_ms", 1),
                New JProperty("diagnostic_id", Guid.NewGuid().ToString("D")),
                New JProperty("human_log_available", False),
                New JProperty("result", resultToken),
                New JProperty("output_files", outputs),
                New JProperty("error", JValue.CreateNull())).ToString(Formatting.None)
        End Function

        ' ── Pre-execution unchanged-resubmission guard (#4) ───────────────────

        Private Shared Sub TestUnchangedResubmissionBlockedBeforeWorker()
            Dim session As New Object()
            Dim terminal As String = Nothing
            Dim code As String = "value = section.header.text"
            ' A deterministic failure with a known object type classifies as CODE_REPAIR_REQUIRED and records the hash.
            PythonExecuteRepairAdvisor.Annotate(session, Args(code),
                FailurePayload("PYTHON_ATTRIBUTE_ERROR", 1, "text", objectType:="docx.section._Header"), False, terminal)

            Dim rejection As String = Nothing
            Dim blocked As Boolean = PythonExecuteRepairAdvisor.ShouldRejectUnchangedResubmission(session, Args(code), rejection)

            AssertTrue(blocked, "Identical code after a deterministic failure must be rejected before the worker starts.")
            Dim err = CType(JObject.Parse(rejection)("error"), JObject)
            AssertEqual("UNCHANGED_RESUBMISSION_REJECTED", err.Value(Of String)("code"), "Rejection payload must carry the pre-execution rejection code.")
            AssertTrue(err("advisor").Value(Of Boolean)("unchanged_resubmission"), "Rejection must mark unchanged_resubmission.")
            AssertFalse(err("advisor").Value(Of Boolean)("worker_invoked"), "Rejection must record that no worker was invoked.")
        End Sub

        Private Shared Sub TestChangedCodeIsNotBlockedBeforeWorker()
            Dim session As New Object()
            Dim terminal As String = Nothing
            PythonExecuteRepairAdvisor.Annotate(session, Args("value = section.header.text"),
                FailurePayload("PYTHON_ATTRIBUTE_ERROR", 1, "text", objectType:="docx.section._Header"), False, terminal)

            Dim rejection As String = Nothing
            Dim blocked As Boolean = PythonExecuteRepairAdvisor.ShouldRejectUnchangedResubmission(
                session, Args("value = section.header.paragraphs[0].text"), rejection)

            AssertFalse(blocked, "A genuinely changed program must not be blocked by the pre-execution guard.")
        End Sub

        ' ── Targeted boundary-error hints (#8) ────────────────────────────────

        Private Shared Sub TestBoundaryHintForPublishedTuple()
            Dim session As New Object()
            Dim terminal As String = Nothing
            Dim err = AnnotatedError(session, "agent_api.publish_result({'items': [(1, 'a')]})",
                FailurePayload("PYTHON_TYPE_ERROR", 1, Nothing, message:="Published result contains unsupported type tuple"), False, terminal)

            Dim guidance = err("advisor").Value(Of String)("guidance")
            AssertTrue(guidance.IndexOf("tuple", StringComparison.OrdinalIgnoreCase) >= 0, "A published-tuple error must yield a tuple-conversion hint.")
        End Sub

        Private Shared Sub TestBoundaryHintForWindowsPathFilename()
            Dim session As New Object()
            Dim terminal As String = Nothing
            Dim err = AnnotatedError(session, "doc = SimpleDocTemplate(path)",
                FailurePayload("PYTHON_TYPE_ERROR", 1, Nothing, message:="Cannot use WindowsPath('out.pdf') as a filename or file"), False, terminal)

            Dim guidance = err("advisor").Value(Of String)("guidance")
            AssertTrue(guidance.IndexOf("str(path)", StringComparison.OrdinalIgnoreCase) >= 0, "A WindowsPath filename error must yield a str(path) conversion hint.")
        End Sub

        Private Shared Sub TestBoundaryHintForHeaderTextAttribute()
            Dim session As New Object()
            Dim terminal As String = Nothing
            Dim err = AnnotatedError(session, "text = section.header.text",
                FailurePayload("PYTHON_ATTRIBUTE_ERROR", 1, "text", objectType:="docx.section._Header",
                               message:="'_Header' object has no attribute 'text'"), False, terminal)

            Dim guidance = err("advisor").Value(Of String)("guidance")
            AssertTrue(guidance.IndexOf("paragraphs", StringComparison.OrdinalIgnoreCase) >= 0, "A header .text error must point to section.header.paragraphs.")
        End Sub

        ' ── Budget enforcement (#3) ───────────────────────────────────────────

        Private Shared Sub TestTransientBudgetIsExhausted()
            Dim session As New Object()
            Dim terminal As String = Nothing
            ' Vary the line so the fingerprint changes each time and exhaustion is driven by the transient budget,
            ' not by the same-fingerprint streak.
            PythonExecuteRepairAdvisor.Annotate(session, Args("x = fetch(1)"), FailurePayload("WEB_REQUEST_TIMEOUT", 1, Nothing, retryable:=True), False, terminal)
            PythonExecuteRepairAdvisor.Annotate(session, Args("x = fetch(2)"), FailurePayload("WEB_REQUEST_TIMEOUT", 2, Nothing, retryable:=True), False, terminal)
            Dim err = AnnotatedError(session, "x = fetch(3)", FailurePayload("WEB_REQUEST_TIMEOUT", 3, Nothing, retryable:=True), False, terminal)

            AssertEqual("REPAIR_BUDGET_EXHAUSTED", err("advisor").Value(Of String)("classification"), "Transient retries must be bounded by the transient budget.")
            AssertFalse(String.IsNullOrEmpty(terminal), "An exhausted transient budget must be terminal.")
        End Sub

        Private Shared Sub TestDiagnosticBudgetThenCodeRepair()
            Dim session As New Object()
            Dim terminal As String = Nothing
            ' First unknown-type attribute error earns the single diagnostic run.
            Dim err1 = AnnotatedError(session, "print(value.tostring())", FailurePayload("PYTHON_ATTRIBUTE_ERROR", 3, "tostring"), False, terminal)
            AssertEqual("DIAGNOSTIC_RUN_REQUIRED", err1("advisor").Value(Of String)("classification"), "The first unknown-type AttributeError should request a diagnostic run.")

            ' Once the single diagnostic budget is spent, the next unknown-type error must require a code repair.
            Dim err2 = AnnotatedError(session, "print(value.foo())", FailurePayload("PYTHON_ATTRIBUTE_ERROR", 7, "foo"), False, terminal)
            AssertEqual("CODE_REPAIR_REQUIRED", err2("advisor").Value(Of String)("classification"), "After the diagnostic budget is spent, further unknown-type errors must require a code repair, not another diagnostic.")
        End Sub

        ' ── Suspicious-repair rejection (#10) ─────────────────────────────────

        Private Shared Sub TestSuspiciousRepairIsRejectedNotJustFlagged()
            Dim session As New Object()
            Dim terminal As String = Nothing
            Dim first As String =
                "def build():" & vbLf &
                "    data = compute()" & vbLf &
                "    agent_api.publish_result(data)" & vbLf &
                "    agent_api.output_path('out.docx')" & vbLf &
                "build()"
            PythonExecuteRepairAdvisor.Annotate(session, Args(first), FailurePayload("PYTHON_VALUE_ERROR", 2, "compute"), False, terminal)

            ' A "repair" that deletes the output generation and the function is a degrading change.
            Dim second As String = "print('done')"
            Dim err = AnnotatedError(session, second, FailurePayload("PYTHON_VALUE_ERROR", 1, Nothing), False, terminal)

            AssertTrue(err("advisor").Value(Of Boolean)("repair_rejected"), "A degrading repair must be rejected, not merely flagged.")
            AssertFalse(err.Value(Of Boolean)("repairable"), "A rejected degrading repair must not be advertised as repairable.")
            AssertFalse(String.IsNullOrEmpty(terminal), "A rejected degrading repair must stop the automatic loop (terminal).")
        End Sub

        ' ── Task-postcondition validation, distinct from worker success (#11) ─

        Private Shared Sub TestPostconditionRejectsNoObservableOutcome()
            Dim incompletePayload As String = Nothing
            Dim incomplete As Boolean = PythonExecuteRepairAdvisor.TryBuildIncompleteTaskPayload(
                New Object(), Args("x = 1"), SuccessPayload(includeResult:=False, outputBytes:=Nothing), incompletePayload)

            AssertTrue(incomplete, "A success that publishes neither a result nor an output file must be treated as incomplete.")
            Dim err = CType(JObject.Parse(incompletePayload)("error"), JObject)
            AssertEqual("TASK_POSTCONDITION_FAILED", err.Value(Of String)("code"), "Incomplete-task payload must carry the postcondition failure code.")
            AssertEqual("observable_outcome_required", err.Value(Of String)("postcondition"), "No-outcome failure must name the observable_outcome_required postcondition.")
            AssertEqual("failed", JObject.Parse(incompletePayload).Value(Of String)("status"), "Incomplete-task payload must report a failed status.")
        End Sub

        Private Shared Sub TestPostconditionRejectsEmptyOutputFile()
            Dim incompletePayload As String = Nothing
            Dim incomplete As Boolean = PythonExecuteRepairAdvisor.TryBuildIncompleteTaskPayload(
                New Object(), Args("x = 1"), SuccessPayload(includeResult:=True, outputBytes:=New Integer() {0}), incompletePayload)

            AssertTrue(incomplete, "A success declaring an empty (0-byte) output file must be treated as incomplete.")
            Dim err = CType(JObject.Parse(incompletePayload)("error"), JObject)
            AssertEqual("non_empty_output", err.Value(Of String)("postcondition"), "Empty-file failure must name the non_empty_output postcondition.")
            AssertEqual("out0.bin", err.Value(Of String)("output"), "Empty-file failure must name the relative output file.")
            AssertFalse(err.Value(Of String)("output").Contains(":\"), "The postcondition payload must not expose an absolute host path.")
        End Sub

        Private Shared Sub TestPostconditionAcceptsValidSuccess()
            Dim incompletePayload As String = Nothing
            Dim incomplete As Boolean = PythonExecuteRepairAdvisor.TryBuildIncompleteTaskPayload(
                New Object(), Args("x = 1"), SuccessPayload(includeResult:=True, outputBytes:=New Integer() {128}), incompletePayload)

            AssertFalse(incomplete, "A success with an observable result and a non-empty output file must be accepted.")
            AssertTrue(String.IsNullOrEmpty(incompletePayload), "A valid success must not emit an incomplete-task payload.")
        End Sub

        ' ── Stable session state and additive-field survival ──────────────────

        Private Shared Sub TestStableSessionHistoryGrowsAcrossRetries()
            Dim session As New Object()
            Dim terminal As String = Nothing
            PythonExecuteRepairAdvisor.Annotate(session, Args("a = 1"), FailurePayload("PYTHON_VALUE_ERROR", 1, Nothing), False, terminal)
            PythonExecuteRepairAdvisor.Annotate(session, Args("a = 2"), FailurePayload("PYTHON_VALUE_ERROR", 2, Nothing), False, terminal)
            Dim err = AnnotatedError(session, "a = 3", FailurePayload("PYTHON_VALUE_ERROR", 3, Nothing), False, terminal)

            Dim history = CType(err("advisor")("attempt_history"), JArray)
            AssertEqual("3", history.Count.ToString(), "One stable session must accumulate all three attempts in its history.")
        End Sub

        Private Shared Sub TestOptionalDiagnosticFieldsSurviveAnnotation()
            Dim errorObj As New JObject(
                New JProperty("code", "PYTHON_ATTRIBUTE_ERROR"),
                New JProperty("phase", "execute"),
                New JProperty("retryable", False),
                New JProperty("exceptionType", "AttributeError"),
                New JProperty("objectType", "docx.section._Header"),
                New JProperty("missingAttribute", "text"),
                New JProperty("message", "'_Header' object has no attribute 'text'"),
                New JProperty("futureField", "keepme"),
                New JProperty("source", New JObject(
                    New JProperty("file", "code.py"),
                    New JProperty("line", 1),
                    New JProperty("symbol", JValue.CreateNull()))),
                New JProperty("stack", New JArray()))
            Dim payload As String = New JObject(
                New JProperty("status", "failed"),
                New JProperty("result", JValue.CreateNull()),
                New JProperty("output_files", New JArray()),
                New JProperty("error", errorObj)).ToString(Formatting.None)

            Dim terminal As String = Nothing
            Dim annotated As String = PythonExecuteRepairAdvisor.Annotate(New Object(), Args("section.header.text"), payload, False, terminal)
            Dim err = CType(JObject.Parse(annotated)("error"), JObject)

            AssertEqual("AttributeError", err.Value(Of String)("exceptionType"), "exceptionType must reach the model.")
            AssertEqual("docx.section._Header", err.Value(Of String)("objectType"), "objectType must reach the model.")
            AssertEqual("text", err.Value(Of String)("missingAttribute"), "missingAttribute must reach the model.")
            AssertEqual("keepme", err.Value(Of String)("futureField"), "Unknown additive fields must survive annotation.")
            AssertTrue(err("advisor") IsNot Nothing, "The advisor block must be attached alongside the optional fields.")
        End Sub

        ''' <summary>Seeds a session with a prior deterministic code-repair baseline for the suspicious-repair guard.</summary>
        Private Shared Sub SeedRepairBaseline(session As Object, code As String)
            Dim terminal As String = Nothing
            PythonExecuteRepairAdvisor.Annotate(session, Args(code), FailurePayload("PYTHON_VALUE_ERROR", 2, "compute"), False, terminal)
        End Sub

        Private Shared ReadOnly BaselineProgram As String =
            "def build():" & vbLf &
            "    data = compute()" & vbLf &
            "    agent_api.publish_result(data)" & vbLf &
            "    agent_api.output_path('out.docx')" & vbLf &
            "build()"

        ' ── Suspicious-repair rejection before worker startup (#10) ───────────

        Private Shared Sub TestSuspiciousRepairBlockedBeforeWorker()
            Dim session As New Object()
            SeedRepairBaseline(session, BaselineProgram)

            Dim rejection As String = Nothing
            Dim blocked As Boolean = PythonExecuteRepairAdvisor.ShouldRejectSuspiciousRepair(session, Args("print('done')"), rejection)

            AssertTrue(blocked, "A destructive repair must be blocked before the worker starts.")
            Dim err = CType(JObject.Parse(rejection)("error"), JObject)
            AssertEqual("SUSPICIOUS_REPAIR_REJECTED", err.Value(Of String)("code"), "Pre-execution rejection must carry the suspicious-repair code.")
            AssertFalse(err("advisor").Value(Of Boolean)("worker_invoked"), "A blocked suspicious repair must not invoke the worker.")
        End Sub

        Private Shared Sub TestRemovalOfPublishResultRejected()
            Dim session As New Object()
            SeedRepairBaseline(session, BaselineProgram)
            Dim degraded As String =
                "def build():" & vbLf &
                "    data = compute_fixed()" & vbLf &
                "    agent_api.output_path('out.docx')" & vbLf &
                "build()"
            Dim rejection As String = Nothing
            AssertTrue(PythonExecuteRepairAdvisor.ShouldRejectSuspiciousRepair(session, Args(degraded), rejection),
                       "Removing publish_result must be rejected.")
        End Sub

        Private Shared Sub TestRemovalOfOutputPathRejected()
            Dim session As New Object()
            SeedRepairBaseline(session, BaselineProgram)
            Dim degraded As String =
                "def build():" & vbLf &
                "    data = compute_fixed()" & vbLf &
                "    agent_api.publish_result(data)" & vbLf &
                "build()"
            Dim rejection As String = Nothing
            AssertTrue(PythonExecuteRepairAdvisor.ShouldRejectSuspiciousRepair(session, Args(degraded), rejection),
                       "Removing output_path must be rejected.")
        End Sub

        Private Shared Sub TestRemovalOfFunctionRejected()
            Dim session As New Object()
            Dim baseline As String =
                "def helper():" & vbLf &
                "    return 1" & vbLf &
                "def build():" & vbLf &
                "    agent_api.publish_result(helper())" & vbLf &
                "    agent_api.output_path('out.docx')" & vbLf &
                "build()"
            SeedRepairBaseline(session, baseline)
            Dim degraded As String =
                "def build():" & vbLf &
                "    agent_api.publish_result(1)" & vbLf &
                "    agent_api.output_path('out.docx')" & vbLf &
                "build()"
            Dim rejection As String = Nothing
            AssertTrue(PythonExecuteRepairAdvisor.ShouldRejectSuspiciousRepair(session, Args(degraded), rejection),
                       "Removing a function/class definition must be rejected.")
        End Sub

        Private Shared Sub TestBroadExceptRejected()
            Dim session As New Object()
            SeedRepairBaseline(session, BaselineProgram)
            Dim degraded As String =
                "def build():" & vbLf &
                "    try:" & vbLf &
                "        data = compute()" & vbLf &
                "    except Exception:" & vbLf &
                "        data = None" & vbLf &
                "    agent_api.publish_result(data)" & vbLf &
                "    agent_api.output_path('out.docx')" & vbLf &
                "build()"
            Dim rejection As String = Nothing
            AssertTrue(PythonExecuteRepairAdvisor.ShouldRejectSuspiciousRepair(session, Args(degraded), rejection),
                       "A newly introduced broad 'except Exception' must be rejected.")
        End Sub

        Private Shared Sub TestBareExceptRejected()
            Dim session As New Object()
            SeedRepairBaseline(session, BaselineProgram)
            Dim degraded As String =
                "def build():" & vbLf &
                "    try:" & vbLf &
                "        data = compute()" & vbLf &
                "    except:" & vbLf &
                "        data = None" & vbLf &
                "    agent_api.publish_result(data)" & vbLf &
                "    agent_api.output_path('out.docx')" & vbLf &
                "build()"
            Dim rejection As String = Nothing
            AssertTrue(PythonExecuteRepairAdvisor.ShouldRejectSuspiciousRepair(session, Args(degraded), rejection),
                       "A newly introduced bare 'except:' must be rejected.")
        End Sub

        Private Shared Sub TestNarrowExceptNotRejected()
            Dim session As New Object()
            SeedRepairBaseline(session, BaselineProgram)
            Dim narrow As String =
                "def build():" & vbLf &
                "    try:" & vbLf &
                "        data = compute()" & vbLf &
                "    except ValueError:" & vbLf &
                "        data = fallback()" & vbLf &
                "    agent_api.publish_result(data)" & vbLf &
                "    agent_api.output_path('out.docx')" & vbLf &
                "build()"
            Dim rejection As String = Nothing
            AssertFalse(PythonExecuteRepairAdvisor.ShouldRejectSuspiciousRepair(session, Args(narrow), rejection),
                        "A narrow, specific exception handler must not be rejected.")
        End Sub

        Private Shared Sub TestMajorityDeletionRejected()
            Dim session As New Object()
            Dim baseline As String =
                "def build():" & vbLf &
                "    a = step_one()" & vbLf &
                "    b = step_two(a)" & vbLf &
                "    c = step_three(b)" & vbLf &
                "    d = step_four(c)" & vbLf &
                "    e = step_five(d)" & vbLf &
                "    agent_api.publish_result(e)" & vbLf &
                "    agent_api.output_path('out.docx')" & vbLf &
                "build()"
            SeedRepairBaseline(session, baseline)
            Dim gutted As String =
                "agent_api.publish_result(1)" & vbLf &
                "agent_api.output_path('out.docx')"
            Dim rejection As String = Nothing
            AssertTrue(PythonExecuteRepairAdvisor.ShouldRejectSuspiciousRepair(session, Args(gutted), rejection),
                       "Deleting more than half of the program must be rejected.")
        End Sub

        Private Shared Sub TestSmallValidRepairAllowed()
            Dim session As New Object()
            SeedRepairBaseline(session, BaselineProgram)
            Dim fixedProgram As String =
                "def build():" & vbLf &
                "    data = compute_fixed()" & vbLf &
                "    agent_api.publish_result(data)" & vbLf &
                "    agent_api.output_path('out.docx')" & vbLf &
                "build()"
            Dim rejection As String = Nothing
            AssertFalse(PythonExecuteRepairAdvisor.ShouldRejectSuspiciousRepair(session, Args(fixedProgram), rejection),
                        "A small, output-preserving repair must be allowed.")
        End Sub

        ' ── Postcondition acceptance, exemption, and ordering (#11) ───────────

        Private Shared Sub TestPostconditionAcceptsNonEmptyOutputFile()
            Dim incompletePayload As String = Nothing
            Dim incomplete As Boolean = PythonExecuteRepairAdvisor.TryBuildIncompleteTaskPayload(
                New Object(), Args("x = 1"), SuccessPayload(includeResult:=False, outputBytes:=New Integer() {64}), incompletePayload)

            AssertFalse(incomplete, "A success with a non-empty output file (and no direct result) must be accepted.")
        End Sub

        Private Shared Sub TestPostconditionExemptsDiagnosticRun()
            Dim session As New Object()
            Dim terminal As String = Nothing
            ' Drive the session into DIAGNOSTIC_RUN_REQUIRED (unknown-type AttributeError, no objectType).
            PythonExecuteRepairAdvisor.Annotate(session, Args("print(value.tostring())"),
                FailurePayload("PYTHON_ATTRIBUTE_ERROR", 3, "tostring"), False, terminal)

            Dim incompletePayload As String = Nothing
            Dim incomplete As Boolean = PythonExecuteRepairAdvisor.TryBuildIncompleteTaskPayload(
                session, Args("print(type(value))"), SuccessPayload(includeResult:=False, outputBytes:=Nothing), incompletePayload)

            AssertFalse(incomplete, "A diagnostic run must be exempt from the observable-outcome postcondition.")
        End Sub

        Private Shared Sub TestPostconditionDoesNotReplaceWorkerFailure()
            Dim incompletePayload As String = Nothing
            Dim incomplete As Boolean = PythonExecuteRepairAdvisor.TryBuildIncompleteTaskPayload(
                New Object(), Args("x = 1"), FailurePayload("PYTHON_VALUE_ERROR", 1, Nothing), incompletePayload)

            AssertFalse(incomplete, "A genuine worker failure must not be replaced by a postcondition failure.")
        End Sub

        Private Shared Sub TestPostconditionEntersHistoryAndBudget()
            Dim session As New Object()
            Dim incompletePayload As String = Nothing
            PythonExecuteRepairAdvisor.TryBuildIncompleteTaskPayload(
                New Object(), Args("x = 1"), SuccessPayload(includeResult:=False, outputBytes:=Nothing), incompletePayload)

            Dim terminal As String = Nothing
            Dim annotated As String = PythonExecuteRepairAdvisor.Annotate(session, Args("x = 1"), incompletePayload, False, terminal)
            Dim err = CType(JObject.Parse(annotated)("error"), JObject)

            AssertEqual("CODE_REPAIR_REQUIRED", err("advisor").Value(Of String)("classification"), "TASK_POSTCONDITION_FAILED must require a code repair.")
            AssertFalse(err.Value(Of Boolean)("retryable"), "TASK_POSTCONDITION_FAILED must not be retryable unchanged.")
            AssertTrue(err.Value(Of Boolean)("repairable"), "TASK_POSTCONDITION_FAILED must be repairable.")
            AssertEqual("1", err("advisor").Value(Of Integer)("code_repairs_used").ToString(), "A postcondition failure must consume a code-repair budget slot.")
            AssertTrue(CType(err("advisor")("attempt_history"), JArray).Count >= 1, "A postcondition failure must enter the attempt history.")
        End Sub

        Private Shared Sub TestSuccessAfterPostconditionRepairResetsState()
            Dim session As New Object()
            Dim terminal As String = Nothing
            Dim incompletePayload As String = Nothing
            PythonExecuteRepairAdvisor.TryBuildIncompleteTaskPayload(
                New Object(), Args("x = 1"), SuccessPayload(includeResult:=False, outputBytes:=Nothing), incompletePayload)
            PythonExecuteRepairAdvisor.Annotate(session, Args("x = 1"), incompletePayload, False, terminal)

            PythonExecuteRepairAdvisor.Annotate(session, Args("x = 2"),
                SuccessPayload(includeResult:=True, outputBytes:=New Integer() {64}), True, terminal)

            Dim err = AnnotatedError(session, "x = 3", FailurePayload("PYTHON_VALUE_ERROR", 1, Nothing), False, terminal)
            AssertEqual("1", CType(err("advisor")("attempt_history"), JArray).Count.ToString(),
                        "A valid success after a postcondition repair must reset the advisor state.")
        End Sub

        ''' <summary>Builds a tool-argument / input-reference failure payload (the Python code was not executed).</summary>
        Private Shared Function ArgumentFailurePayload(code As String, Optional guidance As String = Nothing) As String
            Dim errorObj As New JObject(
                New JProperty("code", code),
                New JProperty("phase", "initializing"),
                New JProperty("retryable", False),
                New JProperty("source", JValue.CreateNull()),
                New JProperty("message", "The supplied path is an internal published-result path and cannot be used directly as input_files."),
                New JProperty("stack", New JArray()))
            If guidance IsNot Nothing Then errorObj("guidance") = guidance
            Return New JObject(
                New JProperty("status", "failed"),
                New JProperty("result", JValue.CreateNull()),
                New JProperty("output_files", New JArray()),
                New JProperty("error", errorObj)).ToString(Formatting.None)
        End Function

        Private Shared Function ArgsWithInput(code As String, inputFile As String) As System.Collections.Generic.Dictionary(Of String, Object)
            Return New System.Collections.Generic.Dictionary(Of String, Object)(StringComparer.Ordinal) From {
                {"code", code},
                {"input_files", New JArray(inputFile)}}
        End Function

        ' ── Input / tool-argument failures are not Python-code repairs (#12/#13) ──

        Private Shared Sub TestInputReferenceFailureRequestsArgumentRepair()
            Dim session As New Object()
            Dim terminal As String = Nothing
            Dim annotated As String = PythonExecuteRepairAdvisor.Annotate(
                session, ArgsWithInput("doc = 1", "results/abc/Vollmacht_2.docx"),
                ArgumentFailurePayload("INPUT_REFERENCE_INVALID",
                    "Use an attachment reference, a workspace-relative path, or an explicitly reusable published-file handle."),
                False, terminal)
            Dim err = CType(JObject.Parse(annotated)("error"), JObject)

            AssertFalse(err.Value(Of Boolean)("retryable"), "An input-reference failure must not be retryable unchanged.")
            AssertTrue(err.Value(Of Boolean)("repairable"), "An input-reference failure must be repairable.")
            AssertEqual("TOOL_ARGUMENT_REPAIR_REQUIRED", err("advisor").Value(Of String)("classification"), "An input-reference failure must not be classified as a Python-code repair.")
            AssertEqual("tool_arguments", err("advisor").Value(Of String)("failure_domain"), "An input-reference failure must be scoped to the tool arguments.")
            Dim guidance = err("advisor").Value(Of String)("guidance")
            AssertTrue(guidance.IndexOf("input_files", StringComparison.OrdinalIgnoreCase) >= 0, "Guidance must direct the model to correct input_files.")
            AssertTrue(guidance.IndexOf("not executed", StringComparison.OrdinalIgnoreCase) >= 0, "Guidance must state the Python code was not executed.")
        End Sub

        Private Shared Sub TestRequestInvalidIsNotCodeRepair()
            Dim session As New Object()
            Dim terminal As String = Nothing
            Dim annotated As String = PythonExecuteRepairAdvisor.Annotate(
                session, Args("x = 1"), ArgumentFailurePayload("REQUEST_INVALID"), False, terminal)
            Dim err = CType(JObject.Parse(annotated)("error"), JObject)

            AssertEqual("TOOL_ARGUMENT_REPAIR_REQUIRED", err("advisor").Value(Of String)("classification"), "REQUEST_INVALID must be treated as a tool-argument failure, not a Python-code repair.")
        End Sub

        Private Shared Sub TestArgumentUnchangedResubmissionFlagged()
            Dim session As New Object()
            Dim terminal As String = Nothing
            Dim args = ArgsWithInput("doc = 1", "results/abc/Vollmacht_2.docx")
            PythonExecuteRepairAdvisor.Annotate(session, args, ArgumentFailurePayload("INPUT_REFERENCE_INVALID"), False, terminal)
            Dim annotated As String = PythonExecuteRepairAdvisor.Annotate(session, args, ArgumentFailurePayload("INPUT_REFERENCE_INVALID"), False, terminal)
            Dim err = CType(JObject.Parse(annotated)("error"), JObject)

            AssertTrue(err("advisor").Value(Of Boolean)("unchanged_resubmission"), "Resubmitting identical tool arguments must be flagged as an unchanged resubmission.")
            Dim history = CType(err("advisor")("attempt_history"), JArray)
            AssertFalse(history(history.Count - 1).Value(Of Boolean)("args_changed"), "An unchanged-argument resubmission must record args_changed=false.")
        End Sub

        ' ── Test scaffolding (matches existing self-test files) ───────────────

        Private Shared Sub RunNamedTest(name As String, test As Action)
            Debug.WriteLine("[PythonExecuteRepairAdvisorSelfTests] RUN  " & name)
            Try
                test.Invoke()
                Debug.WriteLine("[PythonExecuteRepairAdvisorSelfTests] PASS " & name)
            Catch ex As Exception
                Debug.WriteLine("[PythonExecuteRepairAdvisorSelfTests] FAIL " & name & " :: " & ex.ToString())
                Throw
            End Try
        End Sub

        Private Shared Sub AssertTrue(condition As Boolean, message As String)
            If Not condition Then Throw New InvalidOperationException(message)
        End Sub

        Private Shared Sub AssertFalse(condition As Boolean, message As String)
            If condition Then Throw New InvalidOperationException(message)
        End Sub

        Private Shared Sub AssertEqual(expected As String, actual As String, message As String)
            If Not String.Equals(expected, actual, StringComparison.Ordinal) Then
                Throw New InvalidOperationException($"{message} Expected='{expected}', Actual='{actual}'.")
            End If
        End Sub

    End Class

End Namespace

#End If
