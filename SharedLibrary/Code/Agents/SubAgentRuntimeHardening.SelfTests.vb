' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SubAgentRuntimeHardening.SelfTests.vb
' Purpose: Self-tests for SubAgentRuntimeHardening envelope normalization,
'          empty-output detection, and error envelope handling. Validates
'          correct parse behavior for all final-output patterns.
'
' Tests:
'  - Envelope preserved: {summary, result} recognized as complete.
'  - Direct object preserved: {x:1} treated as json_object result.
'  - Array preserved: [{x:1}] treated as json_array result.
'  - Plain text fallback: "plain text" recognized as text result.
'  - Empty output retries: single retry on empty, then agent_empty_result if still empty.
'  - Error envelope: {error:{code,...}} recognized as error resultKind.
'  - Allowed tool scoping: tools list respected from agent descriptor.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Threading
Imports System.Threading.Tasks
Imports Newtonsoft.Json.Linq

Namespace AgentsXX

    Public NotInheritable Class SubAgentRuntimeHardeningSelfTests

        Private Sub New()
        End Sub

        Public Shared Sub RunAll()
            TestEnvelopePreserved()
            TestDirectObjectPreserved()
            TestArrayPreserved()
            TestPlainTextFallback()
            TestEmptyOutputRetriesExactlyOnce()
            TestRetryFailureReturnsAgentEmptyResult()
            TestErrorEnvelopeIsNotSuccess()
            TestStructuredErrorMetadataExtraction()
            TestEmptyMainModelResponsePayload()
            TestModelEmptyResponseRetryUsesCompactRecoveryPrompt()
            TestAllowedToolScoping()
            TestSharedBehaviorParity()
            ToolCallSequencingSelfTests.RunAll()
        End Sub

        Public Shared Function RunAllAndReturnStatus() As String
            RunAll()
            Return "SubAgentRuntimeHardening self-tests passed."
        End Function

        Private Shared Sub TestEnvelopePreserved()
            Dim env = SubAgentRuntimeHardening.NormalizeFinalOutput("{""summary"":""ok"",""result"":{""x"":1}}", jsonRequired:=True)
            AssertTrue(env.ResultKind = "envelope", "Expected envelope resultKind.")
            AssertTrue(TypeOf env.Result Is JObject, "Expected JObject result.")
            AssertTrue(env.Result.Value(Of Integer)("x") = 1, "Expected preserved object payload.")
        End Sub

        Private Shared Sub TestDirectObjectPreserved()
            Dim env = SubAgentRuntimeHardening.NormalizeFinalOutput("{""x"":1}", jsonRequired:=True)
            AssertTrue(env.ResultKind = "json_object", "Expected json_object resultKind.")
            AssertTrue(TypeOf env.Result Is JObject, "Expected direct object result.")
        End Sub

        Private Shared Sub TestArrayPreserved()
            Dim env = SubAgentRuntimeHardening.NormalizeFinalOutput("[{""x"":1}]", jsonRequired:=True)
            AssertTrue(env.ResultKind = "json_array", "Expected json_array resultKind.")
            AssertTrue(TypeOf env.Result Is JArray, "Expected direct array result.")
        End Sub

        Private Shared Sub TestPlainTextFallback()
            Dim env = SubAgentRuntimeHardening.NormalizeFinalOutput("plain text", jsonRequired:=False)
            AssertTrue(env.ResultKind = "text", "Expected text fallback resultKind.")
            AssertTrue(env.Result IsNot Nothing AndAlso env.Result.Type = JTokenType.String, "Expected string result token.")
        End Sub

        Private Shared Sub TestEmptyOutputRetriesExactlyOnce()
            Dim host As New FakeHost({"   ", "{""summary"":""ok"",""result"":{""done"":true}}"})
            Dim agent As New AgentDescriptor() With {
                .Name = "fake_agent_retry_once",
                .AllowedTools = New List(Of String) From {"tool_a"}
            }

            Dim payload As String =
                SubAgentRunner.InvokeResolvedAsync(host, agent, "task", storeResultInMemory:=False).
                    GetAwaiter().
                    GetResult()

            Dim obj = JObject.Parse(payload)

            AssertTrue(host.CallCount = 2, "Expected exactly one retry.")
            AssertTrue(obj.Value(Of String)("resultKind") = "envelope", "Expected successful envelope after retry.")
        End Sub

        Private Shared Sub TestRetryFailureReturnsAgentEmptyResult()
            Dim host As New FakeHost({"   ", "{}"})
            Dim agent As New AgentDescriptor() With {
                .Name = "fake_agent_retry_fail"
            }

            Dim payload As String =
                SubAgentRunner.InvokeResolvedAsync(host, agent, "task", storeResultInMemory:=False).
                    GetAwaiter().
                    GetResult()

            Dim obj = JObject.Parse(payload)

            AssertTrue(host.CallCount = 2, "Expected exactly one retry before failure.")
            AssertTrue(obj.Value(Of String)("resultKind") = "error", "Expected error resultKind.")
            AssertTrue(obj("error") IsNot Nothing AndAlso obj("error").Value(Of String)("code") = "agent_empty_result", "Expected agent_empty_result.")
        End Sub

        Private Shared Sub TestErrorEnvelopeIsNotSuccess()
            Dim host As New FakeHost({
                "{""summary"":""failed"",""result"":null,""resultKind"":""error"",""error"":{""code"":""fake_error"",""phase"":""final_output_parse"",""message"":""failed""}}"
            })
            Dim agent As New AgentDescriptor() With {
                .Name = "fake_agent_error"
            }

            Dim payload As String =
                SubAgentRunner.InvokeResolvedAsync(host, agent, "task", storeResultInMemory:=False).
                    GetAwaiter().
                    GetResult()

            Dim obj = JObject.Parse(payload)

            AssertTrue(obj.Value(Of String)("resultKind") = "error", "Expected preserved error resultKind.")
            AssertTrue(obj("error") IsNot Nothing AndAlso obj("error").Value(Of String)("code") = "fake_error", "Expected preserved error code.")
            AssertTrue(obj("memory_key") Is Nothing, "Error payload must not be treated as normal success.")
        End Sub

        Private Shared Sub TestStructuredErrorMetadataExtraction()
            Dim payload As System.String =
                "{""status"":""error"",""error"":{""code"":""PLAYWRIGHT_DRIVER_UNAVAILABLE"",""message"":""driver assets missing"",""retryable"":false}}"

            Dim errorCode As System.String = ""
            Dim resultKind As System.String = ""
            Dim errorMessage As System.String = ""
            Dim retryable As System.Boolean = True

            AssertTrue(
                SubAgentRuntimeHardening.TryGetEnvelopeErrorInfo(payload, errorCode, resultKind),
                "Expected structured error payload to be recognized.")
            AssertTrue(errorCode = "PLAYWRIGHT_DRIVER_UNAVAILABLE", "Expected error code to be preserved.")
            AssertTrue(
                SubAgentRuntimeHardening.TryGetEnvelopeErrorMessage(payload, errorMessage),
                "Expected structured error message to be available.")
            AssertTrue(errorMessage = "driver assets missing", "Expected exact structured error message.")
            AssertTrue(
                SubAgentRuntimeHardening.TryGetEnvelopeRetryable(payload, retryable),
                "Expected explicit retryable metadata to be available.")
            AssertTrue(Not retryable, "Explicit retryable=false must remain false.")
        End Sub

        Private Shared Sub TestEmptyMainModelResponsePayload()
            Dim obj = JObject.Parse(SubAgentRuntimeHardening.BuildModelEmptyResponsePayload())

            AssertTrue(obj.Value(Of String)("status") = "blocked", "Expected blocked status.")
            AssertTrue(obj("error") IsNot Nothing AndAlso obj("error").Value(Of String)("code") = "model_empty_response", "Expected model_empty_response.")
        End Sub


        Private Shared Sub TestModelEmptyResponseRetryUsesCompactRecoveryPrompt()
            Dim host As New FakeHost({
        SubAgentRuntimeHardening.BuildModelEmptyResponsePayload(
            lastToolName:="workspace_extract_text",
            lastToolResultSummary:="workspace_extract_text succeeded with a compacted excerpt and more content is available via next_offset.",
            compactedToolResponse:=True,
            retryHint:="Return the final JSON object, or request a smaller follow-up window."),
        "{""summary"":""ok"",""result"":{""done"":true}}"
    })

            Dim agent As New AgentDescriptor() With {
        .Name = "fake_agent_compact_retry",
        .AllowedTools = New List(Of String) From {"workspace_extract_text", "workspace_read", "js_run", "tool_loader"}
    }

            Dim payload As String =
        SubAgentRunner.InvokeResolvedAsync(host, agent, "task", storeResultInMemory:=False).
            GetAwaiter().
            GetResult()

            Dim obj = JObject.Parse(payload)

            AssertTrue(host.CallCount = 2, "Expected one retry after model_empty_response.")
            AssertTrue(obj("error") Is Nothing, "Retry should succeed.")
        End Sub


        Private Shared Sub TestAllowedToolScoping()
            Dim reg As New ToolRegistry()

            reg.RegisterEager(New SharedLibrary.ModelConfig With {.Tool = True, .ToolName = "tool_a"})
            reg.RegisterEager(New SharedLibrary.ModelConfig With {.Tool = True, .ToolName = "tool_b"})

            Dim narrowed = reg.Narrow(New String() {"tool_a"})
            Dim emptyScoped = reg.Narrow(Array.Empty(Of String)())

            AssertTrue(narrowed.Contains("tool_a"), "Expected allowed tool in narrowed registry.")
            AssertTrue(Not narrowed.Contains("tool_b"), "Expected disallowed tool to be absent.")
            AssertTrue(Not emptyScoped.Contains("tool_a"), "Expected empty allow-list to produce no tools.")
        End Sub

        Private Shared Sub TestSharedBehaviorParity()
            Dim wordHost As New FakeHost({"{""summary"":""ok"",""result"":{""v"":1}}"})
            Dim outlookHost As New FakeHost({"{""summary"":""ok"",""result"":{""v"":1}}"})
            Dim agent As New AgentDescriptor() With {
                .Name = "fake_agent_parity",
                .AllowedTools = New List(Of String) From {"tool_x"}
            }

            Dim wordPayload As String =
                SubAgentRunner.InvokeResolvedAsync(wordHost, agent, "task", storeResultInMemory:=False).
                    GetAwaiter().
                    GetResult()

            Dim outlookPayload As String =
                SubAgentRunner.InvokeResolvedAsync(outlookHost, agent, "task", storeResultInMemory:=False).
                    GetAwaiter().
                    GetResult()

            AssertTrue(wordPayload = outlookPayload, "Expected shared runtime behavior parity.")
        End Sub

        Private Shared Sub AssertTrue(condition As Boolean, message As String)
            If Not condition Then
                Throw New InvalidOperationException(message)
            End If
        End Sub

        Private NotInheritable Class FakeHost
            Implements ISubAgentHost

            Private ReadOnly _responses As Queue(Of String)

            Public Property CallCount As Integer
            Public Property LastAllowedToolNames As IReadOnlyList(Of String)

            Public Sub New(responses As IEnumerable(Of String))
                _responses = New Queue(Of String)(responses)
            End Sub

            Public Function RunIsolatedToolingLoopAsync(request As SubAgentRunRequest,
                                                        cancellationToken As CancellationToken) As Task(Of String) _
                                                        Implements ISubAgentHost.RunIsolatedToolingLoopAsync
                CallCount += 1
                LastAllowedToolNames = request.AllowedToolNames

                Dim nextResponse As String = ""
                If _responses.Count > 0 Then
                    nextResponse = _responses.Dequeue()
                End If

                Return Task.FromResult(nextResponse)
            End Function
        End Class

    End Class

End Namespace