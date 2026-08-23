' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: WorkflowContinuitySelfTests.vb
' Purpose:
'   DEBUG-only self-tests for workflow-continuity state, recovery, and completion
'   invariants used by the shared tooling loop.
'
' Architecture:
'   Developer-only test harness excluded from SharedLibrary.vbproj. It exercises
'   WorkflowContinuity behavior directly without becoming part of release binaries.
' =============================================================================

#If DEBUG Then

Option Strict On
Option Explicit On

Imports System.IO
Imports System.Threading
Imports System.Threading.Tasks
Imports Newtonsoft.Json.Linq

Namespace AgentsXX

    Public NotInheritable Class WorkflowContinuitySelfTests

        Private Sub New()
        End Sub

        Public Shared Sub RunAll()
            RunNamedTest(NameOf(TestWorkflowIdCreatedForWorkflowStart), AddressOf TestWorkflowIdCreatedForWorkflowStart)
            RunNamedTest(NameOf(TestWorkflowIdPropagatesToSubAgentRequest), AddressOf TestWorkflowIdPropagatesToSubAgentRequest)
            RunNamedTest(NameOf(TestWorkflowIdAppearsInWorkflowLogLabel), AddressOf TestWorkflowIdAppearsInWorkflowLogLabel)
            RunNamedTest(NameOf(TestRuntimeStateUpdatesAfterSuccessfulToolCall), AddressOf TestRuntimeStateUpdatesAfterSuccessfulToolCall)
            RunNamedTest(NameOf(TestRuntimeStateUpdatesAfterFailedToolCall), AddressOf TestRuntimeStateUpdatesAfterFailedToolCall)
            RunNamedTest(NameOf(TestCheckpointWrittenAfterKeyTransition), AddressOf TestCheckpointWrittenAfterKeyTransition)
            RunNamedTest(NameOf(TestSessionMemoryEntryCarriesWorkflowId), AddressOf TestSessionMemoryEntryCarriesWorkflowId)
            RunNamedTest(NameOf(TestMemoryStubInjectedWithoutFullMemoryContent), AddressOf TestMemoryStubInjectedWithoutFullMemoryContent)
            RunNamedTest(NameOf(TestSourceStubInjectedWithoutFullSourceContent), AddressOf TestSourceStubInjectedWithoutFullSourceContent)
            RunNamedTest(NameOf(TestFullMemoryAndSourceContentNotAutoInjected), AddressOf TestFullMemoryAndSourceContentNotAutoInjected)
            RunNamedTest(NameOf(TestFollowUpTurnIncludesRecentWorkflowMemoryStubs), AddressOf TestFollowUpTurnIncludesRecentWorkflowMemoryStubs)
            RunNamedTest(NameOf(TestToolingPromptInstructsMemoryGetForPriorWorkflowResults), AddressOf TestToolingPromptInstructsMemoryGetForPriorWorkflowResults)
            RunNamedTest(NameOf(TestPromptInstructsUserFacingProseToFollowLatestUserRequestLanguage), AddressOf TestPromptInstructsUserFacingProseToFollowLatestUserRequestLanguage)
            RunNamedTest(NameOf(TestSubAgentReceivesOnlyExplicitlyPassedStub), AddressOf TestSubAgentReceivesOnlyExplicitlyPassedStub)
            RunNamedTest(NameOf(TestHostAuthoredRuntimeStateIsAuthoritative), AddressOf TestHostAuthoredRuntimeStateIsAuthoritative)
            RunNamedTest(NameOf(TestAgentModelMemoryIsAdvisory), AddressOf TestAgentModelMemoryIsAdvisory)
            RunNamedTest(NameOf(TestWordOutlookGenericBehaviorParity), AddressOf TestWordOutlookGenericBehaviorParity)
        End Sub

        Private Shared Sub TestWorkflowIdCreatedForWorkflowStart()
            Dim workflowId As String = WorkflowContinuity.CreateWorkflowId()
            Dim state = WorkflowContinuity.StartWorkflow(workflowId, "Word")

            AssertTrue(workflowId <> "", "workflowId should not be empty.")
            AssertEqual(workflowId, state.WorkflowId, "workflowId mismatch.")
            AssertEqual("workflow_started", state.CurrentPhase, "workflow start phase mismatch.")
        End Sub

        Private Shared Sub TestWorkflowIdPropagatesToSubAgentRequest()
            Dim workflowId As String = WorkflowContinuity.CreateWorkflowId()
            Dim host As New CaptureHost()
            Dim agent As New AgentDescriptor() With {.Name = "fake_agent"}

            Dim task = SubAgentRunner.InvokeResolvedAsync(
                host,
                agent,
                "test task",
                contextBlob:="explicit stub",
                storeResultInMemory:=False,
                workflowId:=workflowId,
                cancellationToken:=CancellationToken.None)

            task.GetAwaiter().GetResult()

            AssertEqual(workflowId, host.LastRequest.WorkflowId, "workflowId did not propagate to sub-agent request.")
        End Sub

        Private Shared Sub TestWorkflowIdAppearsInWorkflowLogLabel()
            Dim workflowId As String = WorkflowContinuity.CreateWorkflowId()
            Dim label As String = WorkflowContinuity.BuildWorkflowLogLabel(workflowId, "tool_call_succeeded", "fake_tool", "fake_agent")

            AssertTrue(label.Contains(workflowId), "workflowId must appear in the workflow log label.")
            AssertTrue(label.Contains("tool_call_succeeded"), "phase must appear in the workflow log label.")
        End Sub

        Private Shared Sub TestRuntimeStateUpdatesAfterSuccessfulToolCall()
            Dim workflowId As String = WorkflowContinuity.CreateWorkflowId()
            WorkflowContinuity.StartWorkflow(workflowId, "Word")

            WorkflowContinuity.NoteToolCallResult(
                workflowId,
                "Word",
                "fake_tool",
                succeeded:=True,
                resultRef:="mem_123",
                outputReference:="output/ref",
                sourceRefs:=New String() {"src_1"},
                retryCount:=0)

            Dim state = WorkflowContinuity.GetState(workflowId)

            AssertEqual("fake_tool", state.LastSuccessfulTool, "LastSuccessfulTool mismatch.")
            AssertTrue(state.ToolCallSuccessCount = 1, "Success count mismatch.")
            AssertFalse(state.UnresolvedToolFailure, "Success should clear unresolved failure.")
            AssertEqual("mem_123", state.LastStructuredToolResultRef, "Structured result ref mismatch.")
        End Sub

        Private Shared Sub TestRuntimeStateUpdatesAfterFailedToolCall()
            Dim workflowId As String = WorkflowContinuity.CreateWorkflowId()
            WorkflowContinuity.StartWorkflow(workflowId, "Outlook")

            WorkflowContinuity.NoteToolCallResult(
                workflowId,
                "Outlook",
                "fake_tool",
                succeeded:=False,
                resultRef:="",
                outputReference:="",
                sourceRefs:=Array.Empty(Of String)(),
                retryCount:=2)

            Dim state = WorkflowContinuity.GetState(workflowId)

            AssertEqual("fake_tool", state.LastFailedTool, "LastFailedTool mismatch.")
            AssertTrue(state.ToolCallFailureCount = 1, "Failure count mismatch.")
            AssertTrue(state.UnresolvedToolFailure, "Failure should set unresolvedToolFailure.")
            AssertTrue(state.RetryCount = 2, "RetryCount mismatch.")
        End Sub

        Private Shared Sub TestCheckpointWrittenAfterKeyTransition()
            Dim workflowId As String = WorkflowContinuity.CreateWorkflowId()
            WorkflowContinuity.StartWorkflow(workflowId, "Word")

            Dim path As String = WorkflowContinuity.GetCheckpointPath(workflowId)
            AssertTrue(File.Exists(path), "Checkpoint file should exist after workflow start.")
        End Sub

        Private Shared Sub TestSessionMemoryEntryCarriesWorkflowId()
            Dim workflowId As String = WorkflowContinuity.CreateWorkflowId()

            Using WorkflowContinuity.BeginWorkflowScope(workflowId, "Word")
                Dim entry = SessionMemory.Put(
                    key:="wf_test_" & Guid.NewGuid().ToString("N"),
                    summary:="workflow memory summary",
                    value:=JObject.Parse("{""note"":""body""}"),
                    metadata:=New SessionMemoryMetadata With {
                        .ContentKind = "note",
                        .Source = "model"
                    })

                Try
                    AssertEqual(workflowId, entry.Metadata.WorkflowId, "WorkflowId should be attached to memory entry.")
                Finally
                    SessionMemory.Delete(entry.Key)
                End Try
            End Using
        End Sub

        Private Shared Sub TestMemoryStubInjectedWithoutFullMemoryContent()
            Dim workflowId As String = WorkflowContinuity.CreateWorkflowId()
            Dim key As String = "wf_mem_" & Guid.NewGuid().ToString("N")

            Using WorkflowContinuity.BeginWorkflowScope(workflowId, "Word")
                Dim entry = SessionMemory.Put(
                    key:=key,
                    summary:="Short advisory note",
                    value:=JValue.CreateString("FULL_MEMORY_CONTENT_SHOULD_NOT_APPEAR"),
                    metadata:=New SessionMemoryMetadata With {
                        .ContentKind = "note",
                        .Source = "model"
                    })

                Try
                    Dim block = WorkflowContinuity.BuildPromptContextBlock(workflowId, 4, 4)
                    AssertTrue(block.Contains("[memory:" & key & "]"), "Memory stub should be injected.")
                    AssertFalse(block.Contains("FULL_MEMORY_CONTENT_SHOULD_NOT_APPEAR"), "Full memory content must not be auto-injected.")
                Finally
                    SessionMemory.Delete(entry.Key)
                End Try
            End Using
        End Sub

        Private Shared Sub TestSourceStubInjectedWithoutFullSourceContent()
            Dim workflowId As String = WorkflowContinuity.CreateWorkflowId()
            Dim key As String = "wf_src_" & Guid.NewGuid().ToString("N")

            Dim sourceValue As New JObject(
                New JProperty("sourceId", "src_001"),
                New JProperty("title", "Neutral Source"),
                New JProperty("sourceType", "external_reference"),
                New JProperty("reference", "https://should-not-be-auto-injected.example"),
                New JProperty("retrievedAt", DateTime.UtcNow.ToString("o")),
                New JProperty("summary", "Short source summary"),
                New JProperty("fullBody", "FULL_SOURCE_CONTENT_SHOULD_NOT_APPEAR"))

            Using WorkflowContinuity.BeginWorkflowScope(workflowId, "Outlook")
                Dim entry = SessionMemory.Put(
                    key:=key,
                    summary:="Neutral source record",
                    value:=sourceValue,
                    metadata:=New SessionMemoryMetadata With {
                        .ContentKind = "source_record",
                        .Source = "tool",
                        .RelatedTool = "fake_research_tool"
                    })

                Try
                    Dim block = WorkflowContinuity.BuildPromptContextBlock(workflowId, 4, 4)
                    AssertTrue(block.Contains("[memory:" & key & "]"), "Source stub should include the memory stub.")
                    AssertTrue(block.Contains("Neutral Source"), "Source title should appear in the stub.")
                    AssertFalse(block.Contains("https://should-not-be-auto-injected.example"), "Reference URL must not be auto-injected.")
                    AssertFalse(block.Contains("FULL_SOURCE_CONTENT_SHOULD_NOT_APPEAR"), "Full source content must not be auto-injected.")
                Finally
                    SessionMemory.Delete(entry.Key)
                End Try
            End Using
        End Sub

        Private Shared Sub TestFullMemoryAndSourceContentNotAutoInjected()
            Dim workflowId As String = WorkflowContinuity.CreateWorkflowId()
            Dim memoryKey As String = "wf_mem2_" & Guid.NewGuid().ToString("N")
            Dim sourceKey As String = "wf_src2_" & Guid.NewGuid().ToString("N")

            Using WorkflowContinuity.BeginWorkflowScope(workflowId, "Word")
                Dim memoryEntry = SessionMemory.Put(
                    key:=memoryKey,
                    summary:="Memory stub",
                    value:=JValue.CreateString("SECRET_MEMORY_BODY"),
                    metadata:=New SessionMemoryMetadata With {
                        .ContentKind = "summary",
                        .Source = "model"
                    })

                Dim sourceEntry = SessionMemory.Put(
                    key:=sourceKey,
                    summary:="Source stub",
                    value:=New JObject(
                        New JProperty("sourceId", "src_002"),
                        New JProperty("title", "Source Title"),
                        New JProperty("reference", "SECRET_REFERENCE"),
                        New JProperty("summary", "Source summary"),
                        New JProperty("payload", "SECRET_SOURCE_BODY")),
                    metadata:=New SessionMemoryMetadata With {
                        .ContentKind = "source_record",
                        .Source = "tool"
                    })

                Try
                    Dim block = WorkflowContinuity.BuildPromptContextBlock(workflowId, 4, 4)
                    AssertFalse(block.Contains("SECRET_MEMORY_BODY"), "Full memory body must not be auto-injected.")
                    AssertFalse(block.Contains("SECRET_REFERENCE"), "Full source reference must not be auto-injected.")
                    AssertFalse(block.Contains("SECRET_SOURCE_BODY"), "Full source body must not be auto-injected.")
                Finally
                    SessionMemory.Delete(memoryEntry.Key)
                    SessionMemory.Delete(sourceEntry.Key)
                End Try
            End Using
        End Sub

        Private Shared Sub TestFollowUpTurnIncludesRecentWorkflowMemoryStubs()
            Dim previousWorkflowId As String = WorkflowContinuity.CreateWorkflowId()
            Dim currentWorkflowId As String = WorkflowContinuity.CreateWorkflowId()
            Dim key As String = "wf_followup_" & Guid.NewGuid().ToString("N")

            WorkflowContinuity.StartWorkflow(previousWorkflowId, "Word")
            WorkflowContinuity.StartWorkflow(currentWorkflowId, "Word")

            Dim entry = SessionMemory.Put(
                key:=key,
                summary:="Stored workflow result",
                value:=JValue.CreateString("FULL_WORKFLOW_MEMORY_SHOULD_NOT_APPEAR"),
                metadata:=New SessionMemoryMetadata With {
                    .WorkflowId = previousWorkflowId,
                    .ContentKind = "summary",
                    .Source = "tool",
                    .RelatedTool = "fake_tool",
                    .RelatedAgent = "fake_agent"
                })

            Try
                Dim block As String = WorkflowContinuity.BuildPromptContextBlock(currentWorkflowId, 4, 4)

                AssertTrue(block.Contains("Recent workflow memory stubs"), "Follow-up prompt should include recent workflow memory stubs.")
                AssertTrue(block.Contains("[memory:" & key & "]"), "Follow-up prompt should include the recent memory stub key.")
                AssertTrue(block.Contains("contentKind=summary"), "Follow-up prompt should include contentKind.")
                AssertTrue(block.Contains("source=tool"), "Follow-up prompt should include source.")
                AssertTrue(block.Contains("workflowId=" & previousWorkflowId), "Follow-up prompt should include workflowId.")
                AssertTrue(block.Contains("relatedTool=fake_tool"), "Follow-up prompt should include relatedTool when available.")
                AssertTrue(block.Contains("relatedAgent=fake_agent"), "Follow-up prompt should include relatedAgent when available.")
                AssertFalse(block.Contains("FULL_WORKFLOW_MEMORY_SHOULD_NOT_APPEAR"), "Follow-up prompt must not inject full memory content.")
            Finally
                SessionMemory.Delete(entry.Key)
            End Try
        End Sub

        Private Shared Sub TestToolingPromptInstructsMemoryGetForPriorWorkflowResults()
            Dim toolingPrompt As String = SharedLibrary.SharedMethods.Default_SP_Add_Tooling
            Dim agentLayerPrompt As String = SharedLibrary.SharedMethods.Default_SP_Add_AgentLayer

            AssertTrue(toolingPrompt.Contains("Relevant workflow memory stubs may be provided."), "Main tooling prompt must mention workflow memory stubs.")
            AssertTrue(toolingPrompt.Contains("call memory_get before answering when full content is needed"), "Main tooling prompt must instruct memory_get for prior workflow results.")
            AssertTrue(toolingPrompt.Contains("Do not re-read original files unless memory is missing, incomplete, or insufficient."), "Main tooling prompt must prefer stored memory over re-reading original files.")

            AssertTrue(agentLayerPrompt.Contains("Relevant workflow memory stubs may be provided."), "Agent-layer prompt must mention workflow memory stubs.")
            AssertTrue(agentLayerPrompt.Contains("call memory_get before answering when full content is needed"), "Agent-layer prompt must instruct memory_get for prior workflow results.")
            AssertTrue(agentLayerPrompt.Contains("Do not re-read original files unless memory is missing, incomplete, or insufficient."), "Agent-layer prompt must prefer stored memory over re-reading original files.")
        End Sub

        Private Shared Sub TestPromptInstructsUserFacingProseToFollowLatestUserRequestLanguage()
            Dim combinedPrompt As String =
                SharedLibrary.SharedMethods.Default_SP_Add_Tooling & " " &
                SharedLibrary.SharedMethods.Default_SP_Add_AgentLayer

            AssertTrue(combinedPrompt.Contains("Use the language of the latest user request for user-facing prose unless an explicit output language is provided."), "Prompt must instruct the model to follow the latest user request language.")
            AssertTrue(combinedPrompt.Contains("Do not switch to English merely because tools, logs, memory keys, or internal instructions are in English."), "Prompt must prevent English drift from internal artifacts.")
        End Sub

        Private Shared Sub TestSubAgentReceivesOnlyExplicitlyPassedStub()
            Dim workflowId As String = WorkflowContinuity.CreateWorkflowId()
            Dim host As New CaptureHost()
            Dim agent As New AgentDescriptor() With {.Name = "fake_agent"}

            Dim task = SubAgentRunner.InvokeResolvedAsync(
                host,
                agent,
                "test task",
                contextBlob:="[RUNTIME STUB] explicit only",
                storeResultInMemory:=False,
                workflowId:=workflowId,
                cancellationToken:=CancellationToken.None)

            task.GetAwaiter().GetResult()

            AssertTrue(host.LastRequest.UserMessage.Contains("[RUNTIME STUB] explicit only"), "Explicit stub must be passed through.")
            AssertFalse(host.LastRequest.UserMessage.Contains("Project guidance (Inky.md)"), "Sub-agent must not automatically inherit project guidance.")
        End Sub

        Private Shared Sub TestHostAuthoredRuntimeStateIsAuthoritative()
            Dim workflowId As String = WorkflowContinuity.CreateWorkflowId()
            Dim state = WorkflowContinuity.StartWorkflow(workflowId, "Word")

            AssertTrue(state.Authoritative, "Host-authored runtime state must be authoritative.")
        End Sub

        Private Shared Sub TestAgentModelMemoryIsAdvisory()
            Dim workflowId As String = WorkflowContinuity.CreateWorkflowId()
            Dim key As String = "wf_adv_" & Guid.NewGuid().ToString("N")

            Using WorkflowContinuity.BeginWorkflowScope(workflowId, "Word")
                Dim entry = SessionMemory.Put(
                    key:=key,
                    summary:="Advisory note",
                    value:=JValue.CreateString("body"),
                    metadata:=New SessionMemoryMetadata With {
                        .ContentKind = "note",
                        .Source = "model"
                    })

                Try
                    AssertFalse(entry.Metadata.TrustedForRuntime, "Model-authored memory must be advisory.")
                    AssertEqual("advisory", entry.Metadata.TrustLevel, "TrustLevel mismatch.")
                Finally
                    SessionMemory.Delete(entry.Key)
                End Try
            End Using
        End Sub

        Private Shared Sub TestWordOutlookGenericBehaviorParity()
            Dim workflowIdWord As String = WorkflowContinuity.CreateWorkflowId()
            Dim workflowIdOutlook As String = WorkflowContinuity.CreateWorkflowId()

            Dim wordState = WorkflowContinuity.StartWorkflow(workflowIdWord, "Word")
            Dim outlookState = WorkflowContinuity.StartWorkflow(workflowIdOutlook, "Outlook")

            AssertTrue(wordState.Authoritative = outlookState.Authoritative, "Authoritative parity mismatch.")
            AssertEqual("workflow_started", wordState.CurrentPhase, "Word currentPhase mismatch.")
            AssertEqual("workflow_started", outlookState.CurrentPhase, "Outlook currentPhase mismatch.")
        End Sub

        Private Shared Sub RunNamedTest(name As String, test As Action)
            Debug.WriteLine("[WorkflowContinuitySelfTests] RUN  " & name)

            Try
                test.Invoke()
                Debug.WriteLine("[WorkflowContinuitySelfTests] PASS " & name)
            Catch ex As Exception
                Debug.WriteLine("[WorkflowContinuitySelfTests] FAIL " & name & " :: " & ex.ToString())
                Throw
            End Try
        End Sub

        Private Shared Sub AssertTrue(condition As Boolean, message As String)
            If Not condition Then
                Throw New InvalidOperationException(message)
            End If
        End Sub

        Private Shared Sub AssertFalse(condition As Boolean, message As String)
            If condition Then
                Throw New InvalidOperationException(message)
            End If
        End Sub

        Private Shared Sub AssertEqual(expected As String, actual As String, message As String)
            If Not String.Equals(expected, actual, StringComparison.Ordinal) Then
                Throw New InvalidOperationException($"{message} Expected='{expected}', Actual='{actual}'.")
            End If
        End Sub

        Private NotInheritable Class CaptureHost
            Implements ISubAgentHost

            Public Property LastRequest As SubAgentRunRequest

            Public Function RunIsolatedToolingLoopAsync(request As SubAgentRunRequest,
                                                        cancellationToken As CancellationToken) As Task(Of String) _
                Implements ISubAgentHost.RunIsolatedToolingLoopAsync

                LastRequest = request
                Return Task.FromResult("{""summary"":""ok"",""result"":{""status"":""done""}}")
            End Function
        End Class

    End Class

End Namespace

#End If