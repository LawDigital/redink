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
