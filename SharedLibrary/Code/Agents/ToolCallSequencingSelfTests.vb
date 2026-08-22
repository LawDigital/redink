' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ToolCallSequencingSelfTests.vb
' Purpose: Comprehensive self-tests for ToolCallSequencing validation logic.
'          Tests cover tool call batching, task status parsing, memory grounding,
'          workflow continuation, and error recovery across main loop and sub-agents.
'
' Test Categories:
'  - Batching: independent read-only calls batch; dependent or mutating calls stop.
'  - Task status: Complete/Blocked/ContinueTurn parsing and validation.
'  - Memory grounding: required vs. optional modes; subset disclosure logic.
'  - Continuation: retry exhaustion produces host-generated blocked message.
'  - Workflow: sub-agent failures don't poison later invocations.
'  - Workspace tools: write modes, overwrite aliases, missing text validation.
'  - Note: Tests are currently disabled (xRunAll); see RunAll() for status.
' =============================================================================

#If DEBUG Then

Option Strict On
Option Explicit On

Imports System.IO
Imports Newtonsoft.Json.Linq

Namespace AgentsXX

    Public NotInheritable Class ToolCallSequencingSelfTests

        Private Sub New()
        End Sub

        Public Shared Sub RunAll()
            Debug.WriteLine("[ToolCallSequencingSelfTests] DISABLED")
        End Sub

        Public Shared Function RunAllAndReturnStatus() As String
            Return "ToolCallSequencing self-tests are disabled."
        End Function

        Public Shared Sub xRunAll()
            Debug.WriteLine("[ToolCallSequencingSelfTests] START")

            RunNamedTest(NameOf(TestIndependentReadOnlyCallsBatchTogether), AddressOf TestIndependentReadOnlyCallsBatchTogether)
            RunNamedTest(NameOf(TestMutatingCallStopsBatch), AddressOf TestMutatingCallStopsBatch)
            RunNamedTest(NameOf(TestReadOnlyThenAgentStopsAfterAgent), AddressOf TestReadOnlyThenAgentStopsAfterAgent)
            RunNamedTest(NameOf(TestAgentThenOtherStopsAtAgent), AddressOf TestAgentThenOtherStopsAtAgent)
            RunNamedTest(NameOf(TestSkillThenOtherStopsAtSkill), AddressOf TestSkillThenOtherStopsAtSkill)
            RunNamedTest(NameOf(TestUnknownToolActsAsBarrier), AddressOf TestUnknownToolActsAsBarrier)
            RunNamedTest(NameOf(TestFailedToolSetsUnresolvedToolFailure), AddressOf TestFailedToolSetsUnresolvedToolFailure)
            RunNamedTest(NameOf(TestRetryInvariantRestoresDroppedDesign), AddressOf TestRetryInvariantRestoresDroppedDesign)
            RunNamedTest(NameOf(TestRetryInvariantRejectsDesignReplacement), AddressOf TestRetryInvariantRejectsDesignReplacement)
            RunNamedTest(NameOf(TestRetryInvariantSurvivesDifferentToolCall), AddressOf TestRetryInvariantSurvivesDifferentToolCall)
            RunNamedTest(NameOf(TestSameToolRetryDoesNotPrematurelyRecoverFailure), AddressOf TestSameToolRetryDoesNotPrematurelyRecoverFailure)
            RunNamedTest(NameOf(TestRetryInvariantReleasedAfterOriginalToolSuccess), AddressOf TestRetryInvariantReleasedAfterOriginalToolSuccess)
            RunNamedTest(NameOf(TestRetryInvariantCapturesHostResolvedDefault), AddressOf TestRetryInvariantCapturesHostResolvedDefault)
            RunNamedTest(NameOf(TestWordTemplateContractParsesPropertyLists), AddressOf TestWordTemplateContractParsesPropertyLists)
            RunNamedTest(NameOf(TestWordTemplateContractNormalizesQuotedValues), AddressOf TestWordTemplateContractNormalizesQuotedValues)
            RunNamedTest(NameOf(TestWordTemplateContractSlotNameIsSemanticFree), AddressOf TestWordTemplateContractSlotNameIsSemanticFree)
            RunNamedTest(NameOf(TestWordTemplateContractStillAcceptsLegacyTables), AddressOf TestWordTemplateContractStillAcceptsLegacyTables)
            RunNamedTest(NameOf(TestWordTemplateContractRejectsUnknownBodySemantic), AddressOf TestWordTemplateContractRejectsUnknownBodySemantic)
            RunNamedTest(NameOf(TestWordRenderingRulesIgnoreHumanProse), AddressOf TestWordRenderingRulesIgnoreHumanProse)
            RunNamedTest(NameOf(TestActiveToolingAcceptsToolCallTurn), AddressOf TestActiveToolingAcceptsToolCallTurn)
            RunNamedTest(NameOf(TestActiveToolingAcceptsFinalCompleteTurn), AddressOf TestActiveToolingAcceptsFinalCompleteTurn)
            RunNamedTest(NameOf(TestActiveToolingAcceptsFinalBlockedTurn), AddressOf TestActiveToolingAcceptsFinalBlockedTurn)
            RunNamedTest(NameOf(TestActiveToolingRejectsTextWithoutTaskStatus), AddressOf TestActiveToolingRejectsTextWithoutTaskStatus)
            RunNamedTest(NameOf(TestActiveToolingRejectsProgressNarration), AddressOf TestActiveToolingRejectsProgressNarration)
            RunNamedTest(NameOf(TestActiveToolingRejectsContinueStatusAsFinal), AddressOf TestActiveToolingRejectsContinueStatusAsFinal)
            RunNamedTest(NameOf(TestActiveToolingRejectsEmptyResponse), AddressOf TestActiveToolingRejectsEmptyResponse)
            RunNamedTest(NameOf(TestActiveToolingRejectsMalformedTaskStatus), AddressOf TestActiveToolingRejectsMalformedTaskStatus)
            RunNamedTest(NameOf(TestActiveToolingRejectsMultipleTaskStatusFooters), AddressOf TestActiveToolingRejectsMultipleTaskStatusFooters)
            RunNamedTest(NameOf(TestActiveToolingRejectsTaskStatusNotAtEnd), AddressOf TestActiveToolingRejectsTaskStatusNotAtEnd)
            RunNamedTest(NameOf(TestActiveToolingRejectsRawInternalJson), AddressOf TestActiveToolingRejectsRawInternalJson)
            RunNamedTest(NameOf(TestActiveToolingRejectsCompleteWithUnresolvedToolFailure), AddressOf TestActiveToolingRejectsCompleteWithUnresolvedToolFailure)
            RunNamedTest(NameOf(TestActiveToolingRepairPromptIsStrict), AddressOf TestActiveToolingRepairPromptIsStrict)
            RunNamedTest(NameOf(TestMemoryGroundingClassifierParsesRequiredDecision), AddressOf TestMemoryGroundingClassifierParsesRequiredDecision)
            RunNamedTest(NameOf(TestMemoryGroundingClassifierParsesRequiredDecisionWrappedInJsonFence), AddressOf TestMemoryGroundingClassifierParsesRequiredDecisionWrappedInJsonFence)
            RunNamedTest(NameOf(TestMemoryGroundingClassifierParsesRequiredDecisionWrappedInPlainFence), AddressOf TestMemoryGroundingClassifierParsesRequiredDecisionWrappedInPlainFence)
            RunNamedTest(NameOf(TestMemoryGroundingClassifierRejectsProsePlusJson), AddressOf TestMemoryGroundingClassifierRejectsProsePlusJson)
            RunNamedTest(NameOf(TestMemoryGroundingClassifierInvalidJsonDefaultsSafely), AddressOf TestMemoryGroundingClassifierInvalidJsonDefaultsSafely)
            RunNamedTest(NameOf(TestMemoryGroundingClassifierParsesOptionalDecision), AddressOf TestMemoryGroundingClassifierParsesOptionalDecision)
            RunNamedTest(NameOf(TestMemoryGroundingClassifierParsesNoneDecision), AddressOf TestMemoryGroundingClassifierParsesNoneDecision)
            RunNamedTest(NameOf(TestMemoryGroundingClassifierPromptRemainsGeneric), AddressOf TestMemoryGroundingClassifierPromptRemainsGeneric)
            RunNamedTest(NameOf(TestRequiredMemoryGroundingRejectsFinalCompleteWithoutMemoryAccess), AddressOf TestRequiredMemoryGroundingRejectsFinalCompleteWithoutMemoryAccess)
            RunNamedTest(NameOf(TestRequiredMemoryGroundingRejectsFinalCompleteAfterMemoryListWithEntries), AddressOf TestRequiredMemoryGroundingRejectsFinalCompleteAfterMemoryListWithEntries)
            RunNamedTest(NameOf(TestRequiredMemoryGroundingRepairPromptEscalatesToMemoryGetAfterList), AddressOf TestRequiredMemoryGroundingRepairPromptEscalatesToMemoryGetAfterList)
            RunNamedTest(NameOf(TestRequiredMemoryGroundingAcceptsFinalCompleteAfterMemoryGet), AddressOf TestRequiredMemoryGroundingAcceptsFinalCompleteAfterMemoryGet)
            RunNamedTest(NameOf(TestRequiredMemoryGroundingAcceptsFinalCompleteAfterEmptyMemoryList), AddressOf TestRequiredMemoryGroundingAcceptsFinalCompleteAfterEmptyMemoryList)
            RunNamedTest(NameOf(TestOptionalMemoryGroundingDoesNotRejectFinalCompleteWithoutMemoryAccess), AddressOf TestOptionalMemoryGroundingDoesNotRejectFinalCompleteWithoutMemoryAccess)
            RunNamedTest(NameOf(TestMemoryGroundingModeNoneAllowsFinalCompleteWithoutMemoryTools), AddressOf TestMemoryGroundingModeNoneAllowsFinalCompleteWithoutMemoryTools)
            RunNamedTest(NameOf(TestMemoryStubsAloneDoNotSatisfyRequiredGrounding), AddressOf TestMemoryStubsAloneDoNotSatisfyRequiredGrounding)
            RunNamedTest(NameOf(TestRequiredMemoryGroundingRepairPromptIsStrict), AddressOf TestRequiredMemoryGroundingRepairPromptIsStrict)
            RunNamedTest(NameOf(TestPromptInstructsMemoryGetForExplicitMemoryGrounding), AddressOf TestPromptInstructsMemoryGetForExplicitMemoryGrounding)
            RunNamedTest(NameOf(TestPromptInstructsUserFacingProseToFollowLatestUserRequestLanguage), AddressOf TestPromptInstructsUserFacingProseToFollowLatestUserRequestLanguage)
            RunNamedTest(NameOf(TestInvalidNarrationAfterStructuredResultIsRejected), AddressOf TestInvalidNarrationAfterStructuredResultIsRejected)
            RunNamedTest(NameOf(TestStructuredResultRemainsAvailableDuringRepair), AddressOf TestStructuredResultRemainsAvailableDuringRepair)
            RunNamedTest(NameOf(TestRetryExhaustionProducesHostGeneratedBlockedMessage), AddressOf TestRetryExhaustionProducesHostGeneratedBlockedMessage)
            RunNamedTest(NameOf(TestNoRawBlockedJsonIsSurfaced), AddressOf TestNoRawBlockedJsonIsSurfaced)
            RunNamedTest(NameOf(TestFallbackUsesRuntimeUserLanguage), AddressOf TestFallbackUsesRuntimeUserLanguage)
            RunNamedTest(NameOf(TestToolCallCountIsReportedOnlyAsToolCallCount), AddressOf TestToolCallCountIsReportedOnlyAsToolCallCount)
            RunNamedTest(NameOf(TestRepairPromptRemainsGeneric), AddressOf TestRepairPromptRemainsGeneric)
            RunNamedTest(NameOf(TestActiveToolingWordOutlookParity), AddressOf TestActiveToolingWordOutlookParity)
            RunNamedTest(NameOf(TestContinuationGuardRetryExhaustionWithUnresolvedFailureReturnsBlocked), AddressOf TestContinuationGuardRetryExhaustionWithUnresolvedFailureReturnsBlocked)
            RunNamedTest(NameOf(TestSharedBehaviorParity), AddressOf TestSharedBehaviorParity)
            RunNamedTest(NameOf(TestSubAgentEmptyModelResponseRetriesOnceAndBatchContinues), AddressOf TestSubAgentEmptyModelResponseRetriesOnceAndBatchContinues)
            RunNamedTest(NameOf(TestStructuredSubAgentFailureReturnsToParentAndLaterWorkCanContinue), AddressOf TestStructuredSubAgentFailureReturnsToParentAndLaterWorkCanContinue)
            RunNamedTest(NameOf(TestRepeatedSubAgentInvocationRetainsAllowedTools), AddressOf TestRepeatedSubAgentInvocationRetainsAllowedTools)
            RunNamedTest(NameOf(TestFailedSubAgentInitializationDoesNotPoisonLaterInvocation), AddressOf TestFailedSubAgentInitializationDoesNotPoisonLaterInvocation)
            RunNamedTest(NameOf(TestSequencingBarrierDoesNotClearSubAgentAllowedRegistry), AddressOf TestSequencingBarrierDoesNotClearSubAgentAllowedRegistry)
            RunNamedTest(NameOf(TestOutlookAndWordRepeatedInvocationParity), AddressOf TestOutlookAndWordRepeatedInvocationParity)
            RunNamedTest(NameOf(TestAuthoritativeRegistryUnavailableFailsBeforeModelCall), AddressOf TestAuthoritativeRegistryUnavailableFailsBeforeModelCall)
            RunNamedTest(NameOf(TestWorkspaceWriteWithCanonicalTextWritesFullContent), AddressOf TestWorkspaceWriteWithCanonicalTextWritesFullContent)
            RunNamedTest(NameOf(TestWorkspaceWriteWithLegacyContentAliasWritesFullContent), AddressOf TestWorkspaceWriteWithLegacyContentAliasWritesFullContent)
            RunNamedTest(NameOf(TestWorkspaceWriteOverwriteAliasMapsToOverwriteMode), AddressOf TestWorkspaceWriteOverwriteAliasMapsToOverwriteMode)
            RunNamedTest(NameOf(TestWorkspaceWriteMissingTextFailsValidation), AddressOf TestWorkspaceWriteMissingTextFailsValidation)
            RunNamedTest(NameOf(TestWorkspaceWriteUnknownContentFieldDoesNotSilentlyCreateTinyFile), AddressOf TestWorkspaceWriteUnknownContentFieldDoesNotSilentlyCreateTinyFile)
            RunNamedTest(NameOf(TestWorkspaceToolSchemasMatchDeclaredContracts), AddressOf TestWorkspaceToolSchemasMatchDeclaredContracts)
            RunNamedTest(NameOf(TestSkippedFailureDoesNotShrinkRegistrySnapshot), AddressOf TestSkippedFailureDoesNotShrinkRegistrySnapshot)
            RunNamedTest(NameOf(TestPreviousSubAgentCallDoesNotMutateParentRegistrySnapshot), AddressOf TestPreviousSubAgentCallDoesNotMutateParentRegistrySnapshot)
            RunNamedTest(NameOf(TestRequiredMemoryGroundingSmallListRequiresAllRetrievedKeys), AddressOf TestRequiredMemoryGroundingSmallListRequiresAllRetrievedKeys)
            RunNamedTest(NameOf(TestRequiredMemoryGroundingSmallListPromptMentionsRemainingKeys), AddressOf TestRequiredMemoryGroundingSmallListPromptMentionsRemainingKeys)
            RunNamedTest(NameOf(TestRequiredMemoryGroundingAcceptsFinalCompleteAfterAllSmallListKeysRetrieved), AddressOf TestRequiredMemoryGroundingAcceptsFinalCompleteAfterAllSmallListKeysRetrieved)
            RunNamedTest(NameOf(TestRequiredMemoryGroundingLargeListAllowsSubsetWithDisclosureGuidance), AddressOf TestRequiredMemoryGroundingLargeListAllowsSubsetWithDisclosureGuidance)
            RunNamedTest(NameOf(TestRequiredMemoryGroundingRejectsPartialRetrievalWithoutSubsetDisclosure), AddressOf TestRequiredMemoryGroundingRejectsPartialRetrievalWithoutSubsetDisclosure)
            RunNamedTest(NameOf(TestRequiredMemoryGroundingAcceptsSubsetDisclosureForSmallList), AddressOf TestRequiredMemoryGroundingAcceptsSubsetDisclosureForSmallList)
            RunNamedTest(NameOf(TestRequiredMemoryGroundingAcceptsSubsetDisclosureForLargeList), AddressOf TestRequiredMemoryGroundingAcceptsSubsetDisclosureForLargeList)
            WorkflowContinuitySelfTests.RunAll()
            Debug.WriteLine("[ToolCallSequencingSelfTests] PASS")
        End Sub

        Public Shared Function xRunAllAndReturnStatus() As String
            Try
                RunAll()
                Return "ToolCallSequencing self-tests passed."
            Catch ex As Exception
                Debug.WriteLine("[ToolCallSequencingSelfTests] FAILED :: " & ex.ToString())
                Throw
            End Try
        End Function

        Private Shared Sub TestIndependentReadOnlyCallsBatchTogether()
            Dim plan = ToolCallSequencing.BuildExecutionPlan(
                New String() {"workspace_read", "workspace_search"})

            AssertEqual(2, plan.ExecutedCount, "Read-only batch should execute together.")
            AssertEqual(0, plan.DeferredCount, "Read-only batch should not defer calls.")
            AssertTrue(plan.IsFullyBatchSafe, "Read-only batch should be fully safe.")
        End Sub

        Private Shared Sub TestMutatingCallStopsBatch()
            Dim plan = ToolCallSequencing.BuildExecutionPlan(
                New String() {"workspace_write", "workspace_read"})

            AssertTrue(plan.Calls(0).WillExecute, "First mutating call should execute.")
            AssertFalse(plan.Calls(1).WillExecute, "Later call after mutating barrier should defer.")
            AssertEqual(ToolCallSequencing.ToolCallClassification.Mutating, plan.Calls(0).Classification, "Mutating classification mismatch.")
        End Sub

        Private Shared Sub TestReadOnlyThenAgentStopsAfterAgent()
            Dim plan = ToolCallSequencing.BuildExecutionPlan(
                New String() {"workspace_read", "agent_fake", "workspace_search"})

            AssertTrue(plan.Calls(0).WillExecute, "Leading read-only call should execute.")
            AssertTrue(plan.Calls(1).WillExecute, "First barrier call should execute.")
            AssertFalse(plan.Calls(2).WillExecute, "Call after agent barrier should defer.")
            AssertEqual(ToolCallSequencing.ToolCallClassification.Agent, plan.Calls(1).Classification, "Agent classification mismatch.")
        End Sub

        Private Shared Sub TestAgentThenOtherStopsAtAgent()
            Dim plan = ToolCallSequencing.BuildExecutionPlan(
                New String() {"agent_fake", "workspace_read"})

            AssertTrue(plan.Calls(0).WillExecute, "Agent call should execute.")
            AssertFalse(plan.Calls(1).WillExecute, "Call after agent should defer.")
        End Sub

        Private Shared Sub TestSkillThenOtherStopsAtSkill()
            Dim plan = ToolCallSequencing.BuildExecutionPlan(
                New String() {"skill_fake", "workspace_read"})

            AssertTrue(plan.Calls(0).WillExecute, "Skill call should execute.")
            AssertFalse(plan.Calls(1).WillExecute, "Call after skill should defer.")
        End Sub

        Private Shared Sub TestUnknownToolActsAsBarrier()
            Dim plan = ToolCallSequencing.BuildExecutionPlan(
                New String() {"custom_unknown_tool", "workspace_read"})

            AssertEqual(ToolCallSequencing.ToolCallClassification.Unknown, plan.Calls(0).Classification, "Unknown classification mismatch.")
            AssertTrue(plan.Calls(0).WillExecute, "Unknown first call should execute.")
            AssertFalse(plan.Calls(1).WillExecute, "Call after unknown barrier should defer.")
        End Sub

        Private Shared Sub TestFailedToolSetsUnresolvedToolFailure()
            Dim state As New ToolCallSequencing.ToolingRunState()

            state.NoteToolFailure("fake_tool", "fake_error", "Failure")

            AssertTrue(state.HasUnresolvedToolFailure, "Failed tool should set unresolved failure.")
            AssertEqual("fake_tool", state.LastToolName, "Last tool name mismatch.")
            AssertEqual("fake_error", state.LastErrorCode, "Last error code mismatch.")
        End Sub

        Private Shared Sub TestContinuationGuardRetryExhaustionWithUnresolvedFailureReturnsBlocked()
            Dim state As New ToolCallSequencing.ToolingRunState()
            state.NoteToolFailure("fake_tool", "fake_error", "Failure")

            Dim shouldBlock = ToolCallSequencing.ShouldBlockTextOnlyFinalization(
                state,
                retryCount:=5,
                maxRetryCount:=5,
                hasValidFinalAnswer:=False)

            AssertTrue(shouldBlock, "Retry exhaustion with unresolved failure should block finalization.")

            Dim payload = ToolCallSequencing.BuildBlockedResultPayload(
                ToolCallSequencing.UnresolvedToolFailureCode,
                "finalization",
                "The tooling run ended with an unresolved tool failure.",
                state.LastToolName,
                state.LastErrorCode,
                state.LastErrorMessage)

            Dim obj = JObject.Parse(payload)

            AssertEqual("blocked", obj.Value(Of String)("status"), "Blocked payload status mismatch.")
            AssertEqual(ToolCallSequencing.UnresolvedToolFailureCode, obj("error")?.Value(Of String)("code"), "Blocked payload code mismatch.")
        End Sub

        Private Shared Sub TestSharedBehaviorParity()
            Dim left = ToolCallSequencing.BuildExecutionPlan(
                New String() {"workspace_read", "agent_fake", "workspace_search"})

            Dim right = ToolCallSequencing.BuildExecutionPlan(
                New String() {"workspace_read", "agent_fake", "workspace_search"})

            AssertEqual(left.ExecutedCount, right.ExecutedCount, "Executed count parity mismatch.")
            AssertEqual(left.DeferredCount, right.DeferredCount, "Deferred count parity mismatch.")
            AssertEqual(left.Calls(1).Classification, right.Calls(1).Classification, "Classification parity mismatch.")
        End Sub

        Private Shared Sub RunNamedTest(name As String, test As Action)
            Debug.WriteLine("[ToolCallSequencingSelfTests] RUN  " & name)

            Try
                Dim testTask = System.Threading.Tasks.Task.Run(
                    Sub()
                        test.Invoke()
                    End Sub)

                If Not testTask.Wait(TimeSpan.FromSeconds(20)) Then
                    Debug.WriteLine("[ToolCallSequencingSelfTests] TIMEOUT " & name)
                    Throw New TimeoutException("Timed out after 20 seconds.")
                End If

                If testTask.IsFaulted Then
                    Throw testTask.Exception.GetBaseException()
                End If

                Debug.WriteLine("[ToolCallSequencingSelfTests] PASS " & name)
            Catch ex As Exception
                Debug.WriteLine("[ToolCallSequencingSelfTests] FAIL " & name & " :: " & ex.ToString())
                Throw
            End Try
        End Sub


        Private Shared Sub AssertTrue(condition As Boolean, message As String)
            If Not condition Then
                Throw New InvalidOperationException(message)
            End If
        End Sub

        Private Shared Sub TestSubAgentEmptyModelResponseRetriesOnceAndBatchContinues()
            Dim ag As New AgentDescriptor() With {
                .Name = "mock_subagent"
            }

            Dim host As New MockSubAgentHost()
            host.AddResponse("item-1", BuildMockSuccessPayload("done:item-1"))
            host.AddResponse("item-2", SubAgentRuntimeHardening.BuildModelEmptyResponsePayload())
            host.AddResponse("item-2", BuildMockSuccessPayload("done:item-2"))
            host.AddResponse("item-3", BuildMockSuccessPayload("done:item-3"))

            Dim results = RunMockBatchSkill(host, ag, New String() {"item-1", "item-2", "item-3"})

            AssertEqual(3, results.Count, "Batch skill should complete all items.")
            AssertEqual("done:item-2", results(1), "Second item should succeed after one retry.")
            AssertEqual(4, host.CallCount, "Second item should require exactly one extra sub-agent call.")
        End Sub

        Private Shared Sub TestStructuredSubAgentFailureReturnsToParentAndLaterWorkCanContinue()
            Dim ag As New AgentDescriptor() With {
                .Name = "mock_subagent"
            }

            Dim host As New MockSubAgentHost()
            host.AddResponse("item-1", BuildMockSuccessPayload("done:item-1"))
            host.AddResponse("item-2", SubAgentRuntimeHardening.BuildModelEmptyResponsePayload())
            host.AddResponse("item-2", SubAgentRuntimeHardening.BuildModelEmptyResponsePayload())
            host.AddResponse("item-3", BuildMockSuccessPayload("done:item-3"))

            Dim results = RunMockBatchSkillAllowErrors(host, ag, New String() {"item-1", "item-2", "item-3"})

            AssertEqual(3, results.Count, "Parent batch should continue after structured sub-agent error.")
            AssertEqual("done:item-1", results(0), "First item mismatch.")
            AssertEqual("ERR:model_empty_response", results(1), "Second item should return structured error to parent.")
            AssertEqual("done:item-3", results(2), "Later work should continue after the failed sub-agent item.")

            Dim runState As New ToolCallSequencing.ToolingRunState()
            runState.NoteToolFailure("agent_mock", SubAgentRuntimeHardening.ModelEmptyResponseCode, "empty", skippedByPolicy:=True, returnedToParent:=True)
            AssertTrue(runState.RequiresParentRecovery, "Skipped structured agent failure should require parent recovery.")

            runState.NoteRecoveryByLaterToolCall("mock_parent_fallback")
            AssertFalse(runState.HasUnresolvedToolFailure, "Later parent tool call should recover skipped failure.")

            Dim unrecovered As New ToolCallSequencing.ToolingRunState()
            unrecovered.NoteToolFailure("agent_mock", SubAgentRuntimeHardening.ModelEmptyResponseCode, "empty", skippedByPolicy:=True, returnedToParent:=True)

            Dim shouldBlock = ToolCallSequencing.ShouldBlockTextOnlyFinalization(
                unrecovered,
                retryCount:=5,
                maxRetryCount:=5,
                hasValidFinalAnswer:=False)

            AssertTrue(shouldBlock, "Skipped failure should become fatal only after guard exhaustion.")
        End Sub

        Private Shared Sub TestRepeatedSubAgentInvocationRetainsAllowedTools()
            Dim host As New MockRepeatedToolScopeHost(
                "Word",
                BuildFakeToolRegistry({"fake_tool_alpha", "fake_tool_beta"}))

            Dim ag As New AgentDescriptor() With {
                .Name = "neutral_fake"
            }
            ag.AllowedTools.Add("fake_tool_alpha")
            ag.AllowedTools.Add("fake_tool_beta")

            For i As Integer = 1 To 3
                Dim payload = SubAgentRunner.InvokeResolvedAsync(
                    host,
                    ag,
                    "repeat-" & i.ToString(),
                    contextBlob:=Nothing,
                    storeResultInMemory:=False,
                    cancellationToken:=System.Threading.CancellationToken.None).GetAwaiter().GetResult()

                Dim obj = JObject.Parse(payload)
                AssertTrue(obj("error") Is Nothing, $"Invocation {i} should succeed.")
            Next

            AssertEqual(3, host.Invocations.Count, "Expected three recorded invocations.")

            Dim first = host.Invocations(0)
            Dim second = host.Invocations(1)
            Dim third = host.Invocations(2)

            AssertToolNameSequenceEqual(New String() {"fake_tool_alpha", "fake_tool_beta"}, first.ResolvedTools, "First invocation resolved tools mismatch.")
            AssertToolNameSequenceEqual(first.ResolvedTools, second.ResolvedTools, "Second invocation should resolve the same tools as the first.")
            AssertToolNameSequenceEqual(first.ResolvedTools, third.ResolvedTools, "Third invocation should resolve the same tools as the first.")

            AssertToolNameSequenceEqual(first.FinalCallableTools, second.FinalCallableTools, "Second invocation callable scope mismatch.")
            AssertToolNameSequenceEqual(first.FinalCallableTools, third.FinalCallableTools, "Third invocation callable scope mismatch.")

            AssertEqual(1, first.InvocationIndex, "First invocation index mismatch.")
            AssertEqual(2, second.InvocationIndex, "Second invocation index mismatch.")
            AssertEqual(3, third.InvocationIndex, "Third invocation index mismatch.")

            AssertEqual(1, first.AgentInvocationCount, "First same-agent invocation count mismatch.")
            AssertEqual(2, second.AgentInvocationCount, "Second same-agent invocation count mismatch.")
            AssertEqual(3, third.AgentInvocationCount, "Third same-agent invocation count mismatch.")

            AssertTrue(first.SnapshotExists, "First invocation should have a parent registry snapshot.")
            AssertTrue(second.SnapshotExists, "Second invocation should have a parent registry snapshot.")
            AssertTrue(third.SnapshotExists, "Third invocation should have a parent registry snapshot.")

            AssertEqual(first.SnapshotToolCount, second.SnapshotToolCount, "Second invocation snapshot count mismatch.")
            AssertEqual(first.SnapshotToolCount, third.SnapshotToolCount, "Third invocation snapshot count mismatch.")

        End Sub

        Private Shared Sub TestFailedSubAgentInitializationDoesNotPoisonLaterInvocation()
            Dim host As New MockRepeatedToolScopeHost(
                "Word",
                BuildFakeToolRegistry({"fake_tool_alpha"}))

            Dim failingAgent As New AgentDescriptor() With {
                .Name = "neutral_fake"
            }
            failingAgent.AllowedTools.Add("missing_tool")

            Dim failingPayload = SubAgentRunner.InvokeResolvedAsync(
                host,
                failingAgent,
                "fail-once",
                contextBlob:=Nothing,
                storeResultInMemory:=False,
                cancellationToken:=System.Threading.CancellationToken.None).GetAwaiter().GetResult()

            Dim failingObj = JObject.Parse(failingPayload)
            AssertEqual(Agents.SubAgentRuntimeHardening.RequiredToolMissingCode,
                        failingObj("error")?.Value(Of String)("code"),
                        "Initialization failure should return the structured tool-scope error.")

            Dim succeedingAgent As New AgentDescriptor() With {
                .Name = "neutral_fake"
            }
            succeedingAgent.AllowedTools.Add("fake_tool_alpha")

            Dim succeedingPayload = SubAgentRunner.InvokeResolvedAsync(
                host,
                succeedingAgent,
                "succeed-next",
                contextBlob:=Nothing,
                storeResultInMemory:=False,
                cancellationToken:=System.Threading.CancellationToken.None).GetAwaiter().GetResult()

            Dim succeedingObj = JObject.Parse(succeedingPayload)
            AssertTrue(succeedingObj("error") Is Nothing, "Later invocation should still succeed after an initialization failure.")

            Dim lastInvocation = host.Invocations(host.Invocations.Count - 1)
            AssertToolNameSequenceEqual(New String() {"fake_tool_alpha"}, lastInvocation.ResolvedTools, "Successful invocation should resolve the expected tool.")
        End Sub

        Private Shared Sub TestSequencingBarrierDoesNotClearSubAgentAllowedRegistry()
            Dim host As New MockRepeatedToolScopeHost(
                "Word",
                BuildFakeToolRegistry({"fake_tool_alpha", "fake_tool_beta"}))

            Dim ag As New AgentDescriptor() With {
                .Name = "neutral_fake"
            }
            ag.AllowedTools.Add("fake_tool_alpha")
            ag.AllowedTools.Add("fake_tool_beta")

            Dim firstPayload = SubAgentRunner.InvokeResolvedAsync(
                host,
                ag,
                "before-barrier",
                contextBlob:=Nothing,
                storeResultInMemory:=False,
                cancellationToken:=System.Threading.CancellationToken.None).GetAwaiter().GetResult()

            Dim firstObj = JObject.Parse(firstPayload)
            AssertTrue(firstObj("error") Is Nothing, "First invocation should succeed.")

            Dim plan = ToolCallSequencing.BuildExecutionPlan(
                New String() {"workspace_read", "agent_neutral_fake", "workspace_search"})

            AssertTrue(plan.Calls(1).WillExecute, "Barrier call should execute.")
            AssertFalse(plan.Calls(2).WillExecute, "Later call should defer after the barrier.")

            Dim secondPayload = SubAgentRunner.InvokeResolvedAsync(
                host,
                ag,
                "after-barrier",
                contextBlob:=Nothing,
                storeResultInMemory:=False,
                cancellationToken:=System.Threading.CancellationToken.None).GetAwaiter().GetResult()

            Dim secondObj = JObject.Parse(secondPayload)
            AssertTrue(secondObj("error") Is Nothing, "Second invocation should still succeed after the sequencing barrier.")

            AssertToolNameSequenceEqual(
                host.Invocations(0).ResolvedTools,
                host.Invocations(1).ResolvedTools,
                "Sequencing barriers must not clear the allowed-tool registry.")
        End Sub

        Private Shared Sub TestOutlookAndWordRepeatedInvocationParity()
            Dim wordHost As New MockRepeatedToolScopeHost(
                "Word",
                BuildFakeToolRegistry({"fake_tool_alpha", "fake_tool_beta"}))

            Dim outlookHost As New MockRepeatedToolScopeHost(
                "Outlook",
                BuildFakeToolRegistry({"fake_tool_alpha", "fake_tool_beta"}))

            Dim ag As New AgentDescriptor() With {
                .Name = "neutral_fake"
            }
            ag.AllowedTools.Add("fake_tool_alpha")
            ag.AllowedTools.Add("fake_tool_beta")

            For i As Integer = 1 To 2
                Dim wordPayload = SubAgentRunner.InvokeResolvedAsync(
                    wordHost,
                    ag,
                    "word-" & i.ToString(),
                    contextBlob:=Nothing,
                    storeResultInMemory:=False,
                    cancellationToken:=System.Threading.CancellationToken.None).GetAwaiter().GetResult()

                Dim outlookPayload = SubAgentRunner.InvokeResolvedAsync(
                    outlookHost,
                    ag,
                    "outlook-" & i.ToString(),
                    contextBlob:=Nothing,
                    storeResultInMemory:=False,
                    cancellationToken:=System.Threading.CancellationToken.None).GetAwaiter().GetResult()

                AssertTrue(JObject.Parse(wordPayload)("error") Is Nothing, "Word parity invocation should succeed.")
                AssertTrue(JObject.Parse(outlookPayload)("error") Is Nothing, "Outlook parity invocation should succeed.")
            Next

            AssertEqual(wordHost.Invocations.Count, outlookHost.Invocations.Count, "Parity host invocation count mismatch.")

            For i As Integer = 0 To wordHost.Invocations.Count - 1
                AssertToolNameSequenceEqual(wordHost.Invocations(i).ResolvedTools, outlookHost.Invocations(i).ResolvedTools, $"Resolved-tools parity mismatch at invocation {i}.")
                AssertToolNameSequenceEqual(wordHost.Invocations(i).MissingTools, outlookHost.Invocations(i).MissingTools, $"Missing-tools parity mismatch at invocation {i}.")
                AssertToolNameSequenceEqual(wordHost.Invocations(i).FinalSelectedTools, outlookHost.Invocations(i).FinalSelectedTools, $"Final-selected-tools parity mismatch at invocation {i}.")
                AssertToolNameSequenceEqual(wordHost.Invocations(i).FinalCallableTools, outlookHost.Invocations(i).FinalCallableTools, $"Final-callable-tools parity mismatch at invocation {i}.")
                AssertEqual(wordHost.Invocations(i).AgentInvocationCount, outlookHost.Invocations(i).AgentInvocationCount, $"Same-agent invocation count parity mismatch at invocation {i}.")
            Next
        End Sub

        Private Shared Sub TestAuthoritativeRegistryUnavailableFailsBeforeModelCall()
            Dim host As New MockRepeatedToolScopeHost(
        "Word",
        authoritativeRegistry:=Nothing,
        authoritativeRegistryAvailable:=False)

            Dim ag As New AgentDescriptor() With {
        .Name = "neutral_fake"
    }
            ag.AllowedTools.Add("workspace_extract_text")
            ag.AllowedTools.Add("workspace_read")
            ag.AllowedTools.Add("js_run")

            Dim payload = SubAgentRunner.InvokeResolvedAsync(
        host,
        ag,
        "fail-before-model",
        contextBlob:=Nothing,
        storeResultInMemory:=False,
        cancellationToken:=System.Threading.CancellationToken.None).GetAwaiter().GetResult()

            Dim obj = JObject.Parse(payload)

            AssertEqual(
        Agents.SubAgentRuntimeHardening.ParentRegistryMissingCode,
        obj("error")?.Value(Of String)("code"),
        "Initialization failure should return the parent-registry-missing error.")

            AssertEqual(
        Agents.SubAgentRuntimeHardening.ParentRegistryMissingPhase,
        obj("error")?.Value(Of String)("phase"),
        "Initialization failure phase mismatch.")
        End Sub

        Private Shared Sub TestRetryInvariantRestoresDroppedDesign()
            Dim state As New ToolCallSequencing.ToolingRunState()
            Dim firstArgs As New System.Collections.Generic.Dictionary(Of String, Object)(System.StringComparer.OrdinalIgnoreCase) From {
                {"design_name", "Example_Design"},
                {"content", "first"}
            }
            Dim restored As String = ""
            Dim validationError As String = ""
            AssertTrue(
                ToolCallSequencing.EnforceRetryInvariantArguments("create_artifact", firstArgs, state, restored, validationError),
                "Initial invariant capture should succeed.")

            state.NoteToolFailure("create_artifact", "template_failure", "failed")

            Dim retryArgs As New System.Collections.Generic.Dictionary(Of String, Object)(System.StringComparer.OrdinalIgnoreCase) From {
                {"content", "retry"}
            }
            restored = ""
            validationError = ""
            AssertTrue(
                ToolCallSequencing.EnforceRetryInvariantArguments("create_artifact", retryArgs, state, restored, validationError),
                "Retry should restore a dropped design constraint.")
            AssertEqual("Example_Design", retryArgs("design_name").ToString(), "Retry design was not restored.")
            AssertTrue(restored.IndexOf("design_name=Example_Design", System.StringComparison.OrdinalIgnoreCase) >= 0, "Restore summary should report the restored design.")
        End Sub

        Private Shared Sub TestRetryInvariantRejectsDesignReplacement()
            Dim state As New ToolCallSequencing.ToolingRunState()
            Dim firstArgs As New System.Collections.Generic.Dictionary(Of String, Object)(System.StringComparer.OrdinalIgnoreCase) From {
                {"design_name", "Example_Design"}
            }
            Dim restored As String = ""
            Dim validationError As String = ""
            AssertTrue(
                ToolCallSequencing.EnforceRetryInvariantArguments("create_artifact", firstArgs, state, restored, validationError),
                "Initial invariant capture should succeed.")
            state.NoteToolFailure("create_artifact", "template_failure", "failed")

            Dim retryArgs As New System.Collections.Generic.Dictionary(Of String, Object)(System.StringComparer.OrdinalIgnoreCase) From {
                {"design_name", "Different_Design"}
            }
            restored = ""
            validationError = ""
            AssertFalse(
                ToolCallSequencing.EnforceRetryInvariantArguments("create_artifact", retryArgs, state, restored, validationError),
                "Retry must reject replacement of the required design.")
            AssertTrue(validationError.IndexOf("remains binding", System.StringComparison.OrdinalIgnoreCase) >= 0, "Retry rejection should explain that the original design remains binding.")
        End Sub

        Private Shared Sub TestRetryInvariantSurvivesDifferentToolCall()
            Dim state As New ToolCallSequencing.ToolingRunState()
            Dim firstArgs As New System.Collections.Generic.Dictionary(Of String, Object)(System.StringComparer.OrdinalIgnoreCase) From {
                {"design_name", "Example_Design"}
            }
            Dim restored As String = ""
            Dim validationError As String = ""
            AssertTrue(
                ToolCallSequencing.EnforceRetryInvariantArguments("create_artifact", firstArgs, state, restored, validationError),
                "Initial invariant capture should succeed.")
            state.NoteToolFailure("create_artifact", "template_failure", "failed")

            state.NoteRecoveryByLaterToolCall("helper_tool")
            Dim helperArgs As New System.Collections.Generic.Dictionary(Of String, Object)(System.StringComparer.OrdinalIgnoreCase)()
            AssertTrue(
                ToolCallSequencing.EnforceRetryInvariantArguments("helper_tool", helperArgs, state, restored, validationError),
                "A different recovery tool call must not erase the failed artifact's retry invariants.")
            state.NoteSuccessfulProgress("helper_tool")

            Dim retryArgs As New System.Collections.Generic.Dictionary(Of String, Object)(System.StringComparer.OrdinalIgnoreCase)()
            restored = ""
            validationError = ""
            AssertTrue(
                ToolCallSequencing.EnforceRetryInvariantArguments("create_artifact", retryArgs, state, restored, validationError),
                "Artifact retry should still restore the design after another tool call.")
            AssertEqual("Example_Design", retryArgs("design_name").ToString(), "Different tool call erased the saved design invariant.")
        End Sub

        Private Shared Sub TestSameToolRetryDoesNotPrematurelyRecoverFailure()
            Dim state As New ToolCallSequencing.ToolingRunState()
            Dim firstArgs As New System.Collections.Generic.Dictionary(Of String, Object)(System.StringComparer.OrdinalIgnoreCase) From {
                {"design_name", "Example_Design"}
            }
            Dim restored As String = ""
            Dim validationError As String = ""
            AssertTrue(
                ToolCallSequencing.EnforceRetryInvariantArguments("create_artifact", firstArgs, state, restored, validationError),
                "Initial invariant capture should succeed.")
            state.NoteToolFailure("create_artifact", "template_failure", "failed", skippedByPolicy:=True, returnedToParent:=True)

            state.NoteRecoveryByLaterToolCall("create_artifact")
            AssertTrue(state.HasUnresolvedToolFailure, "Issuing the same failed tool again must remain a retry until success.")
        End Sub

        Private Shared Sub TestRetryInvariantReleasedAfterOriginalToolSuccess()
            Dim state As New ToolCallSequencing.ToolingRunState()
            Dim firstArgs As New System.Collections.Generic.Dictionary(Of String, Object)(System.StringComparer.OrdinalIgnoreCase) From {
                {"design_name", "Example_Design"}
            }
            Dim restored As String = ""
            Dim validationError As String = ""
            AssertTrue(
                ToolCallSequencing.EnforceRetryInvariantArguments("create_artifact", firstArgs, state, restored, validationError),
                "Initial invariant capture should succeed.")
            state.NoteToolFailure("create_artifact", "template_failure", "failed")

            Dim retryArgs As New System.Collections.Generic.Dictionary(Of String, Object)(System.StringComparer.OrdinalIgnoreCase)()
            AssertTrue(
                ToolCallSequencing.EnforceRetryInvariantArguments("create_artifact", retryArgs, state, restored, validationError),
                "Retry should restore the design.")
            state.NoteSuccessfulProgress("create_artifact")

            Dim laterIndependentArgs As New System.Collections.Generic.Dictionary(Of String, Object)(System.StringComparer.OrdinalIgnoreCase)()
            restored = ""
            validationError = ""
            AssertTrue(
                ToolCallSequencing.EnforceRetryInvariantArguments("create_artifact", laterIndependentArgs, state, restored, validationError),
                "A later independent call after success should not be forced to reuse the old design.")
            AssertFalse(laterIndependentArgs.ContainsKey("design_name"), "Successful completion should release the prior retry design lock.")
        End Sub

        Private Shared Sub TestRetryInvariantCapturesHostResolvedDefault()
            Dim state As New ToolCallSequencing.ToolingRunState()
            Dim firstArgs As New System.Collections.Generic.Dictionary(Of String, Object)(System.StringComparer.OrdinalIgnoreCase) From {
                {"content", "first"}
            }
            Dim restored As String = ""
            Dim validationError As String = ""
            AssertTrue(
                ToolCallSequencing.EnforceRetryInvariantArguments("create_artifact", firstArgs, state, restored, validationError),
                "Initial call without an explicit design should pass the retry gate.")

            ' Simulate deterministic host-side default resolution after the dispatcher gate.
            firstArgs("design_name") = "Resolved_Default"
            ToolCallSequencing.CaptureRetryInvariantArguments("create_artifact", firstArgs, state)
            state.NoteToolFailure("create_artifact", "template_failure", "failed")

            Dim replacementArgs As New System.Collections.Generic.Dictionary(Of String, Object)(System.StringComparer.OrdinalIgnoreCase) From {
                {"design_name", "Different_Design"}
            }
            restored = ""
            validationError = ""
            AssertFalse(
                ToolCallSequencing.EnforceRetryInvariantArguments("create_artifact", replacementArgs, state, restored, validationError),
                "A retry must not replace a host-resolved default design after failure.")

            Dim droppedArgs As New System.Collections.Generic.Dictionary(Of String, Object)(System.StringComparer.OrdinalIgnoreCase)()
            restored = ""
            validationError = ""
            AssertTrue(
                ToolCallSequencing.EnforceRetryInvariantArguments("create_artifact", droppedArgs, state, restored, validationError),
                "A retry that drops a host-resolved default should restore it.")
            AssertEqual("Resolved_Default", droppedArgs("design_name").ToString(), "Host-resolved default design was not restored on retry.")
        End Sub

        Private Shared Sub TestWordTemplateContractParsesPropertyLists()
            Dim guidance As String = System.String.Join(vbLf, New String() {
                "# Example",
                "## Word template slots",
                "- placeholder: [[RI:Recipient]]",
                "  source: template_fields.recipient",
                "  purpose: Person or group to whom the document is addressed.",
                "  required: yes",
                "",
                "- placeholder: [[RI:Body]]",
                "  source: markdown_content",
                "  purpose: Main substantive document body.",
                "  required: yes",
                "",
                "## Word body styles",
                "- paragraph: Body Style",
                "- heading1: Heading One",
                "- bullet1: Bullet One"
            })

            Dim contract As Agents.WordTemplateBindingContract = Nothing
            Dim validationError As String = ""
            AssertTrue(
                Agents.WordTemplateBindingContractParser.TryParse(guidance, "Example.md", contract, validationError),
                "Property-list Word template contract should parse.")
            AssertTrue(contract IsNot Nothing, "Property-list contract should be returned.")
            AssertEqual(2, contract.Slots.Count, "Unexpected property-list slot count.")
            AssertEqual("text", contract.Slots(0).ContentMode, "template_fields content mode should default to text.")
            AssertEqual("markdown", contract.Slots(1).ContentMode, "markdown_content content mode should default to markdown.")
            AssertEqual("Person or group to whom the document is addressed.", contract.Slots(0).Purpose, "Slot purpose mismatch.")
            AssertTrue(contract.BuildPromptSummary().IndexOf("Person or group", System.StringComparison.OrdinalIgnoreCase) >= 0, "Prompt summary should expose slot purpose.")
            AssertEqual(3, contract.BodyStyles.Count, "Unexpected property-list body-style count.")
            AssertEqual("Body Style", contract.BuildNativeParagraphStyleMap()("paragraph"), "Paragraph native style mapping mismatch.")
        End Sub

        Private Shared Sub TestWordTemplateContractNormalizesQuotedValues()
            Dim guidance As String = System.String.Join(vbLf, New String() {
                "## Word template slots",
                "- placeholder: `[[RI:Recipient]]`",
                "  source: ""template_fields.recipient""",
                "  purpose: 'Recipient field'",
                "  required: `yes`",
                "",
                "- placeholder: [[RI:Body]]",
                "  source: markdown_content",
                "  purpose: Main body",
                "  required: yes",
                "",
                "## Word body styles",
                "- paragraph: `Body Style`",
                "- heading1: ""Heading One"""
            })

            Dim contract As Agents.WordTemplateBindingContract = Nothing
            Dim validationError As String = ""
            AssertTrue(
                Agents.WordTemplateBindingContractParser.TryParse(guidance, "Quoted.md", contract, validationError),
                "Matching quotes/backticks should be tolerated as input.")
            AssertEqual("[[RI:Recipient]]", contract.Slots(0).Placeholder, "Quoted placeholder was not normalized.")
            AssertEqual("template_fields.recipient", contract.Slots(0).Source, "Quoted source was not normalized.")
            AssertEqual("Recipient field", contract.Slots(0).Purpose, "Quoted purpose was not normalized.")
            AssertEqual("Body Style", contract.BuildNativeParagraphStyleMap()("paragraph"), "Quoted style name was not normalized.")
            AssertEqual("Heading One", contract.BuildNativeParagraphStyleMap()("heading1"), "Quoted heading style was not normalized.")
        End Sub

        Private Shared Sub TestWordTemplateContractSlotNameIsSemanticFree()
            Dim bodyGuidance As String = System.String.Join(vbLf, New String() {
                "## Word template slots",
                "- placeholder: [[RI:Body]]",
                "  source: markdown_content",
                "  content: markdown",
                "  required: yes"
            })
            Dim textGuidance As String = bodyGuidance.Replace("[[RI:Body]]", "[[RI:Text]]")

            Dim bodyContract As Agents.WordTemplateBindingContract = Nothing
            Dim textContract As Agents.WordTemplateBindingContract = Nothing
            Dim validationError As String = ""
            AssertTrue(
                Agents.WordTemplateBindingContractParser.TryParse(bodyGuidance, "Body.md", bodyContract, validationError),
                "Body-named slot should parse.")
            validationError = ""
            AssertTrue(
                Agents.WordTemplateBindingContractParser.TryParse(textGuidance, "Text.md", textContract, validationError),
                "Text-named slot should parse.")
            AssertTrue(bodyContract.Slots(0).UsesMarkdownContent, "Body slot source should define markdown semantics.")
            AssertTrue(textContract.Slots(0).UsesMarkdownContent, "Text slot source should define markdown semantics.")
            AssertEqual("Body", bodyContract.Slots(0).SlotId, "Body slot id mismatch.")
            AssertEqual("Text", textContract.Slots(0).SlotId, "Text slot id mismatch.")
        End Sub

        Private Shared Sub TestWordTemplateContractStillAcceptsLegacyTables()
            Dim guidance As String = System.String.Join(vbLf, New String() {
                "## Word template slots",
                "| Placeholder | Source | Content | Required |",
                "| --- | --- | --- | --- |",
                "| [[RI:Body]] | markdown_content | markdown | yes |",
                "",
                "## Word body styles",
                "| Semantic | Word style |",
                "| --- | --- |",
                "| paragraph | Legacy Body |"
            })

            Dim contract As Agents.WordTemplateBindingContract = Nothing
            Dim validationError As String = ""
            AssertTrue(
                Agents.WordTemplateBindingContractParser.TryParse(guidance, "Legacy.md", contract, validationError),
                "Legacy Markdown-table Word contract should remain readable.")
            AssertEqual(1, contract.Slots.Count, "Legacy slot count mismatch.")
            AssertEqual(1, contract.BodyStyles.Count, "Legacy body-style count mismatch.")
        End Sub

        Private Shared Sub TestWordTemplateContractRejectsUnknownBodySemantic()
            Dim guidance As String = System.String.Join(vbLf, New String() {
                "## Word body styles",
                "- heading7: Unsupported"
            })

            Dim contract As Agents.WordTemplateBindingContract = Nothing
            Dim validationError As String = ""
            AssertFalse(
                Agents.WordTemplateBindingContractParser.TryParse(guidance, "Invalid.md", contract, validationError),
                "Unknown body-style semantic must be rejected.")
            AssertTrue(validationError.IndexOf("Word body styles", System.StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                       validationError.IndexOf("semantic", System.StringComparison.OrdinalIgnoreCase) >= 0,
                       "Unknown semantic rejection should be explicit.")
        End Sub

        Private Shared Sub TestWordRenderingRulesIgnoreHumanProse()
            Dim guidance As String = System.String.Join(vbLf, New String() {
                "# Example",
                "## Word rendering rules",
                "- heading_numbering: native",
                "",
                "A design-specific companion may override this with `heading_numbering: preserve` when required.",
                "This prose contains another colon: it is documentation, not configuration."
            })

            Dim contract As Agents.WordTemplateBindingContract = Nothing
            Dim validationError As String = ""
            AssertTrue(
                Agents.WordTemplateBindingContractParser.TryParse(guidance, "Rendering.md", contract, validationError),
                "Human prose in the rendering-rules section must not be parsed as configuration.")
            AssertTrue(contract IsNot Nothing, "Rendering-rule contract should be returned.")
            AssertEqual("native", contract.HeadingNumberingMode, "Explicit heading_numbering rule should remain authoritative.")
        End Sub

        Private Shared Sub TestActiveToolingAcceptsToolCallTurn()
            Dim result = ToolCallSequencing.ValidateActiveToolingTurn(
                responseText:="tool turn",
                hasToolCalls:=True,
                hasUnresolvedToolFailure:=False)

            AssertEqual(ToolCallSequencing.ActiveToolingTurnKind.ToolCallTurn, result.TurnKind, "Tool call turn should be accepted.")
        End Sub

        Private Shared Sub TestActiveToolingAcceptsFinalCompleteTurn()
            Dim result = ToolCallSequencing.ValidateActiveToolingTurn(
                BuildFinalTurnText("Completed successfully.", "complete", "done"),
                hasToolCalls:=False,
                hasUnresolvedToolFailure:=False)

            AssertEqual(ToolCallSequencing.ActiveToolingTurnKind.FinalCompleteTurn, result.TurnKind, "Final complete turn should be accepted.")
        End Sub

        Private Shared Sub TestActiveToolingAcceptsFinalBlockedTurn()
            Dim result = ToolCallSequencing.ValidateActiveToolingTurn(
                BuildFinalTurnText("Stopped safely.", "blocked", "stopped"),
                hasToolCalls:=False,
                hasUnresolvedToolFailure:=False)

            AssertEqual(ToolCallSequencing.ActiveToolingTurnKind.FinalBlockedTurn, result.TurnKind, "Final blocked turn should be accepted.")
        End Sub

        Private Shared Sub TestActiveToolingRejectsTextWithoutTaskStatus()
            Dim result = ToolCallSequencing.ValidateActiveToolingTurn(
                "Plain text without a footer.",
                hasToolCalls:=False,
                hasUnresolvedToolFailure:=False)

            AssertEqual(ToolCallSequencing.ActiveToolingTurnKind.InvalidTurn, result.TurnKind, "Text without TASK_STATUS must be rejected.")
            AssertEqual("missing_task_status", result.InvalidReason, "Missing TASK_STATUS reason mismatch.")
        End Sub

        Private Shared Sub TestActiveToolingRejectsProgressNarration()
            Dim result = ToolCallSequencing.ValidateActiveToolingTurn(
                "I will continue with the remaining items next.",
                hasToolCalls:=False,
                hasUnresolvedToolFailure:=False)

            AssertEqual(ToolCallSequencing.ActiveToolingTurnKind.InvalidTurn, result.TurnKind, "Progress narration must be rejected.")
        End Sub

        Private Shared Sub TestActiveToolingRejectsContinueStatusAsFinal()
            Dim result = ToolCallSequencing.ValidateActiveToolingTurn(
                BuildFinalTurnText("Continuing.", "continue", "more_work"),
                hasToolCalls:=False,
                hasUnresolvedToolFailure:=False)

            AssertEqual(ToolCallSequencing.ActiveToolingTurnKind.InvalidTurn, result.TurnKind, "TASK_STATUS continue must be rejected as final.")
            AssertEqual("task_status_continue_not_final", result.InvalidReason, "Continue-status reason mismatch.")
        End Sub

        Private Shared Sub TestActiveToolingRejectsEmptyResponse()
            Dim result = ToolCallSequencing.ValidateActiveToolingTurn(
                "   ",
                hasToolCalls:=False,
                hasUnresolvedToolFailure:=False)

            AssertEqual(ToolCallSequencing.ActiveToolingTurnKind.InvalidTurn, result.TurnKind, "Empty response must be rejected.")
            AssertEqual("empty_response", result.InvalidReason, "Empty-response reason mismatch.")
        End Sub

        Private Shared Sub TestActiveToolingRejectsMalformedTaskStatus()
            Dim result = ToolCallSequencing.ValidateActiveToolingTurn(
                "Bad footer <TASK_STATUS>{""status"":""complete""}</TASK_STATUS>",
                hasToolCalls:=False,
                hasUnresolvedToolFailure:=False)

            AssertEqual(ToolCallSequencing.ActiveToolingTurnKind.InvalidTurn, result.TurnKind, "Malformed TASK_STATUS must be rejected.")
            AssertEqual("task_status_missing_reason", result.InvalidReason, "Malformed TASK_STATUS reason mismatch.")
        End Sub

        Private Shared Sub TestActiveToolingRejectsMultipleTaskStatusFooters()
            Dim text As String =
                BuildFinalTurnText("First.", "blocked", "one") &
                BuildFinalTurnText("Second.", "blocked", "two")

            Dim result = ToolCallSequencing.ValidateActiveToolingTurn(
                text,
                hasToolCalls:=False,
                hasUnresolvedToolFailure:=False)

            AssertEqual(ToolCallSequencing.ActiveToolingTurnKind.InvalidTurn, result.TurnKind, "Multiple TASK_STATUS footers must be rejected.")
            AssertEqual("multiple_task_status", result.InvalidReason, "Multiple footer reason mismatch.")
        End Sub

        Private Shared Sub TestActiveToolingRejectsTaskStatusNotAtEnd()
            Dim text As String =
                BuildFinalTurnText("Almost done.", "blocked", "stopped") & " trailing text"

            Dim result = ToolCallSequencing.ValidateActiveToolingTurn(
                text,
                hasToolCalls:=False,
                hasUnresolvedToolFailure:=False)

            AssertEqual(ToolCallSequencing.ActiveToolingTurnKind.InvalidTurn, result.TurnKind, "TASK_STATUS not at end must be rejected.")
            AssertEqual("task_status_not_at_end", result.InvalidReason, "Footer-at-end reason mismatch.")
        End Sub

        Private Shared Sub TestActiveToolingRejectsRawInternalJson()
            Dim result = ToolCallSequencing.ValidateActiveToolingTurn(
                "{""status"":""blocked"",""error"":{""code"":""invalid_text_only_finalization""}}",
                hasToolCalls:=False,
                hasUnresolvedToolFailure:=False)

            AssertEqual(ToolCallSequencing.ActiveToolingTurnKind.InvalidTurn, result.TurnKind, "Raw internal JSON must be rejected.")
            AssertEqual("raw_internal_json", result.InvalidReason, "Raw-JSON reason mismatch.")
        End Sub

        Private Shared Sub TestActiveToolingRejectsCompleteWithUnresolvedToolFailure()
            Dim state As New ToolCallSequencing.ToolingRunState()
            state.NoteToolFailure("fake_tool", "fake_error", "failed")

            Dim result = ToolCallSequencing.ValidateActiveToolingTurn(
                BuildFinalTurnText("Done.", "complete", "finished"),
                hasToolCalls:=False,
                hasUnresolvedToolFailure:=state.HasUnresolvedToolFailure)

            AssertEqual(ToolCallSequencing.ActiveToolingTurnKind.InvalidTurn, result.TurnKind, "Complete with unresolved failure must be rejected.")
            AssertEqual("complete_with_unresolved_tool_failure", result.InvalidReason, "Unresolved-failure reason mismatch.")
        End Sub

        Private Shared Sub TestActiveToolingRepairPromptIsStrict()
            Dim prompt As String = ToolCallSequencing.BuildActiveToolingRepairPrompt()

            AssertTrue(prompt.Contains("the next required tool call"), "Repair prompt must require the next tool call.")
            AssertTrue(prompt.Contains("TASK_STATUS continue is invalid during active tooling"), "Repair prompt must reject continue status.")
            AssertTrue(prompt.Contains("Raw internal JSON is invalid"), "Repair prompt must reject raw JSON.")
        End Sub

        Private Shared Sub TestRetryExhaustionProducesHostGeneratedBlockedMessage()
            Dim state As New ToolCallSequencing.ToolingRunState() With {
                .LastSuccessfulToolCall = "fake_read",
                .LastStateFilePath = "state/output.json",
                .ActiveToolingSession = True,
                .HasOpenToolWorkflow = True
            }

            Dim message As String = ToolCallSequencing.BuildUserSafeBlockedFinalMessage(
                state,
                ToolCallSequencing.InvalidTextOnlyFinalizationCode,
                "The tooling run ended because the model did not return a valid next tool call or a valid final status.",
                successCount:=3,
                failedCount:=0,
                appendTaskStatusFooter:=True)

            AssertTrue(message.Contains("Last successful tool call: fake_read."), "Blocked message must include the last successful tool.")
            AssertTrue(message.Contains("Partial output or state reference: state/output.json."), "Blocked message must mention partial state when known.")
            AssertTrue(message.Contains("Successful tool-call count: 3. Failed tool-call count: 0."), "Blocked message must include tool counts.")
            AssertTrue(message.Contains("<TASK_STATUS>{""status"":""blocked"""), "Blocked message must include a blocked TASK_STATUS footer.")
        End Sub

        Private Shared Sub TestNoRawBlockedJsonIsSurfaced()
            Dim state As New ToolCallSequencing.ToolingRunState() With {
                .ActiveToolingSession = True,
                .HasOpenToolWorkflow = True
            }

            Dim message As String = ToolCallSequencing.BuildUserSafeBlockedFinalMessage(
                state,
                ToolCallSequencing.InvalidTextOnlyFinalizationCode,
                "The tooling run ended because the model did not return a valid next tool call or a valid final status.",
                successCount:=0,
                failedCount:=0,
                appendTaskStatusFooter:=True)

            AssertFalse(message.TrimStart().StartsWith("{", StringComparison.Ordinal), "User-facing blocked message must not be raw JSON.")
            AssertTrue(message.Contains("No unresolved tool failure was recorded before response-contract enforcement stopped the run."), "unresolvedToolFailure=false must not be misreported as a tool failure.")
        End Sub

        Private Shared Sub TestActiveToolingWordOutlookParity()
            Dim wordResult = ToolCallSequencing.ValidateActiveToolingTurn(
                BuildFinalTurnText("Blocked.", "blocked", "stopped"),
                hasToolCalls:=False,
                hasUnresolvedToolFailure:=False)

            Dim outlookResult = ToolCallSequencing.ValidateActiveToolingTurn(
                BuildFinalTurnText("Blocked.", "blocked", "stopped"),
                hasToolCalls:=False,
                hasUnresolvedToolFailure:=False)

            AssertEqual(wordResult.TurnKind, outlookResult.TurnKind, "Word/Outlook parity turn-kind mismatch.")
            AssertEqual(wordResult.TaskStatusSummary, outlookResult.TaskStatusSummary, "Word/Outlook parity TASK_STATUS mismatch.")
        End Sub


        Private Shared Sub TestMemoryGroundingClassifierParsesRequiredDecision()
            Dim decision =
                ToolCallSequencing.ParseMemoryGroundingIntentClassifierDecision(
                    "{""memoryGroundingMode"":""required"",""reason"":""explicit memory grounding requested"",""shouldExposeRecentMemoryStubs"":true}")

            AssertTrue(decision.IsValid, "Required classifier decision should parse.")
            AssertEqual("required",
                        ToolCallSequencing.FormatMemoryGroundingMode(decision.MemoryGroundingMode),
                        "Required classifier mode mismatch.")
            AssertTrue(decision.ShouldExposeRecentMemoryStubs, "Required classifier decision should expose recent memory stubs.")
        End Sub

        Private Shared Sub TestMemoryGroundingClassifierParsesRequiredDecisionWrappedInJsonFence()
            Dim normalizedOutput As String = ""
            Dim parseError As String = ""

            Dim decision =
                ToolCallSequencing.ParseMemoryGroundingIntentClassifierDecision(
                    "```json" & vbCrLf &
                    "{""memoryGroundingMode"":""required"",""reason"":""explicit memory grounding requested"",""shouldExposeRecentMemoryStubs"":true}" & vbCrLf &
                    "```",
                    normalizedOutput,
                    parseError)

            AssertTrue(decision.IsValid, "JSON fenced classifier output should parse.")
            AssertEqual("required",
                        ToolCallSequencing.FormatMemoryGroundingMode(decision.MemoryGroundingMode),
                        "JSON fenced classifier mode mismatch.")
            AssertTrue(normalizedOutput.StartsWith("{"), "Normalized fenced output should start with a JSON object.")
            AssertEqual("", parseError, "JSON fenced classifier output should not produce a parse error.")
        End Sub

        Private Shared Sub TestMemoryGroundingClassifierParsesRequiredDecisionWrappedInPlainFence()
            Dim normalizedOutput As String = ""
            Dim parseError As String = ""

            Dim decision =
                ToolCallSequencing.ParseMemoryGroundingIntentClassifierDecision(
                    "```" & vbCrLf &
                    "{""memoryGroundingMode"":""required"",""reason"":""explicit memory grounding requested"",""shouldExposeRecentMemoryStubs"":true}" & vbCrLf &
                    "```",
                    normalizedOutput,
                    parseError)

            AssertTrue(decision.IsValid, "Plain fenced classifier output should parse.")
            AssertEqual("required",
                        ToolCallSequencing.FormatMemoryGroundingMode(decision.MemoryGroundingMode),
                        "Plain fenced classifier mode mismatch.")
            AssertTrue(normalizedOutput.StartsWith("{"), "Normalized plain-fenced output should start with a JSON object.")
            AssertEqual("", parseError, "Plain fenced classifier output should not produce a parse error.")
        End Sub

        Private Shared Sub TestMemoryGroundingClassifierRejectsProsePlusJson()
            Dim normalizedOutput As String = ""
            Dim parseError As String = ""

            Dim decision =
                ToolCallSequencing.ParseMemoryGroundingIntentClassifierDecision(
                    "Here is the result:" & vbCrLf &
                    "{""memoryGroundingMode"":""required"",""reason"":""explicit memory grounding requested"",""shouldExposeRecentMemoryStubs"":true}",
                    normalizedOutput,
                    parseError)

            AssertFalse(decision.IsValid, "Prose plus JSON must not be accepted.")
            AssertTrue(parseError <> "", "Rejected prose plus JSON should produce a parse error.")
        End Sub

        Private Shared Sub TestMemoryGroundingClassifierInvalidJsonDefaultsSafely()
            Dim normalizedOutput As String = ""
            Dim parseError As String = ""

            Dim decision =
                ToolCallSequencing.ParseMemoryGroundingIntentClassifierDecision(
                    "{""memoryGroundingMode"":""required"",""reason"":""x"",""shouldExposeRecentMemoryStubs"":tru}",
                    normalizedOutput,
                    parseError)

            AssertFalse(decision.IsValid, "Invalid JSON must not parse.")
            AssertEqual("none",
                        ToolCallSequencing.FormatMemoryGroundingMode(decision.MemoryGroundingMode),
                        "Invalid JSON must default safely to none.")
            AssertTrue(parseError <> "", "Invalid JSON should report a parse error.")
        End Sub


        Private Shared Sub TestMemoryGroundingClassifierParsesOptionalDecision()
            Dim decision =
                ToolCallSequencing.ParseMemoryGroundingIntentClassifierDecision(
                    "{""memoryGroundingMode"":""optional"",""reason"":""stored context may help"",""shouldExposeRecentMemoryStubs"":true}")

            AssertTrue(decision.IsValid, "Optional classifier decision should parse.")
            AssertEqual("optional",
                        ToolCallSequencing.FormatMemoryGroundingMode(decision.MemoryGroundingMode),
                        "Optional classifier mode mismatch.")
            AssertTrue(decision.ShouldExposeRecentMemoryStubs, "Optional classifier decision should expose recent memory stubs.")
        End Sub

        Private Shared Sub TestMemoryGroundingClassifierParsesNoneDecision()
            Dim decision =
                ToolCallSequencing.ParseMemoryGroundingIntentClassifierDecision(
                    "{""memoryGroundingMode"":""none"",""reason"":""unrelated new task"",""shouldExposeRecentMemoryStubs"":false}")

            AssertTrue(decision.IsValid, "None classifier decision should parse.")
            AssertEqual("none",
                        ToolCallSequencing.FormatMemoryGroundingMode(decision.MemoryGroundingMode),
                        "None classifier mode mismatch.")
            AssertFalse(decision.ShouldExposeRecentMemoryStubs, "None classifier decision should not expose recent memory stubs.")
        End Sub

        Private Shared Sub TestMemoryGroundingClassifierPromptRemainsGeneric()
            Dim prompt As String = ToolCallSequencing.BuildMemoryGroundingIntentClassifierSystemPrompt()

            AssertTrue(prompt.Contains("Return EXACTLY one raw JSON object and nothing else."), "Classifier prompt must require raw JSON only.")
            AssertTrue(prompt.Contains("Do NOT use Markdown. Do NOT use code fences."), "Classifier prompt must forbid Markdown fences.")
            AssertTrue(prompt.Contains("Base the decision on semantic meaning"), "Classifier prompt should be semantic, not keyword-based.")
            AssertFalse(prompt.Contains("German"), "Classifier prompt should not hardcode language heuristics.")
            AssertFalse(prompt.Contains("English"), "Classifier prompt should not hardcode language heuristics.")
        End Sub

        Private Shared Sub TestRequiredMemoryGroundingSmallListRequiresAllRetrievedKeys()
            Dim state As New ToolCallSequencing.ToolingRunState() With {
                .MemoryGroundingMode = ToolCallSequencing.MemoryGroundingMode.Required
            }

            ToolCallSequencing.NoteMemoryGroundingToolResult(
                state,
                Agents.MemoryTools.ToolList,
                "[{""key"":""mem_1"",""summary"":""stub""},{""key"":""mem_2"",""summary"":""stub""},{""key"":""mem_3"",""summary"":""stub""}]",
                succeeded:=True)

            ToolCallSequencing.NoteMemoryGroundingToolResult(
                state,
                Agents.MemoryTools.ToolGet,
                "{""key"":""mem_1"",""summary"":""stub"",""value"":{""note"":""body""}}",
                succeeded:=True)

            Dim result = ToolCallSequencing.ValidateActiveToolingTurn(
                BuildFinalTurnText("Done.", "complete", "finished"),
                hasToolCalls:=False,
                hasUnresolvedToolFailure:=False,
                runState:=state)

            AssertEqual(ToolCallSequencing.ActiveToolingTurnKind.InvalidTurn, result.TurnKind, "Small listed memory sets should keep requiring remaining memory_get calls.")
            AssertEqual(ToolCallSequencing.MemoryListDoneButMemoryGetRequiredCode, result.InvalidReason, "Small-list partial retrieval reason mismatch.")
        End Sub

        Private Shared Sub TestRequiredMemoryGroundingSmallListPromptMentionsRemainingKeys()
            Dim state As New ToolCallSequencing.ToolingRunState() With {
                .MemoryGroundingMode = ToolCallSequencing.MemoryGroundingMode.Required
            }

            ToolCallSequencing.NoteMemoryGroundingToolResult(
                state,
                Agents.MemoryTools.ToolList,
                "[{""key"":""mem_1"",""summary"":""stub""},{""key"":""mem_2"",""summary"":""stub""},{""key"":""mem_3"",""summary"":""stub""}]",
                succeeded:=True)

            ToolCallSequencing.NoteMemoryGroundingToolResult(
                state,
                Agents.MemoryTools.ToolGet,
                "{""key"":""mem_1"",""summary"":""stub"",""value"":{""note"":""body""}}",
                succeeded:=True)

            Dim prompt As String = ToolCallSequencing.BuildRequiredMemoryGroundingRepairPrompt(state)

            AssertTrue(prompt.Contains("Retrieve all remaining listed memory key"), "Prompt should require remaining keys for small lists.")
            AssertTrue(prompt.Contains("mem_2"), "Prompt should include unretrieved keys.")
            AssertTrue(prompt.Contains("mem_3"), "Prompt should include unretrieved keys.")
        End Sub

        Private Shared Sub TestRequiredMemoryGroundingAcceptsFinalCompleteAfterAllSmallListKeysRetrieved()
            Dim state As New ToolCallSequencing.ToolingRunState() With {
                .MemoryGroundingMode = ToolCallSequencing.MemoryGroundingMode.Required
            }

            ToolCallSequencing.NoteMemoryGroundingToolResult(
                state,
                Agents.MemoryTools.ToolList,
                "[{""key"":""mem_1"",""summary"":""stub""},{""key"":""mem_2"",""summary"":""stub""},{""key"":""mem_3"",""summary"":""stub""}]",
                succeeded:=True)

            ToolCallSequencing.NoteMemoryGroundingToolResult(
                state,
                Agents.MemoryTools.ToolGet,
                "{""key"":""mem_1"",""summary"":""stub"",""value"":{""note"":""body1""}}",
                succeeded:=True)
            ToolCallSequencing.NoteMemoryGroundingToolResult(
                state,
                Agents.MemoryTools.ToolGet,
                "{""key"":""mem_2"",""summary"":""stub"",""value"":{""note"":""body2""}}",
                succeeded:=True)
            ToolCallSequencing.NoteMemoryGroundingToolResult(
                state,
                Agents.MemoryTools.ToolGet,
                "{""key"":""mem_3"",""summary"":""stub"",""value"":{""note"":""body3""}}",
                succeeded:=True)

            Dim result = ToolCallSequencing.ValidateActiveToolingTurn(
                BuildFinalTurnText("Done.", "complete", "finished"),
                hasToolCalls:=False,
                hasUnresolvedToolFailure:=False,
                runState:=state)

            AssertEqual(ToolCallSequencing.ActiveToolingTurnKind.FinalCompleteTurn, result.TurnKind, "Final complete should be accepted after all small-list memory keys are retrieved.")
        End Sub

        Private Shared Sub TestRequiredMemoryGroundingRejectsPartialRetrievalWithoutSubsetDisclosure()
            Dim state As New ToolCallSequencing.ToolingRunState() With {
                .MemoryGroundingMode = ToolCallSequencing.MemoryGroundingMode.Required
            }

            ToolCallSequencing.NoteMemoryGroundingToolResult(
                state,
                Agents.MemoryTools.ToolList,
                "[{""key"":""mem_1"",""summary"":""stub""},{""key"":""mem_2"",""summary"":""stub""},{""key"":""mem_3"",""summary"":""stub""},{""key"":""mem_4"",""summary"":""stub""},{""key"":""mem_5"",""summary"":""stub""},{""key"":""mem_6"",""summary"":""stub""},{""key"":""mem_7"",""summary"":""stub""},{""key"":""mem_8"",""summary"":""stub""},{""key"":""mem_9"",""summary"":""stub""},{""key"":""mem_10"",""summary"":""stub""},{""key"":""mem_11"",""summary"":""stub""}]",
                succeeded:=True)

            ToolCallSequencing.NoteMemoryGroundingToolResult(
                state,
                Agents.MemoryTools.ToolGet,
                "{""key"":""mem_1"",""summary"":""stub"",""value"":{""note"":""body1""}}",
                succeeded:=True)

            Dim result = ToolCallSequencing.ValidateActiveToolingTurn(
                BuildFinalTurnText("Done.", "complete", "finished"),
                hasToolCalls:=False,
                hasUnresolvedToolFailure:=False,
                runState:=state)

            AssertEqual(ToolCallSequencing.ActiveToolingTurnKind.InvalidTurn, result.TurnKind, "Partial retrieval without subset disclosure must be rejected.")
            AssertEqual(ToolCallSequencing.PartialMemoryRetrievalRequiresSubsetDisclosureCode, result.InvalidReason, "Partial retrieval rejection reason mismatch.")
        End Sub

        Private Shared Sub TestRequiredMemoryGroundingAcceptsSubsetDisclosureForSmallList()
            Dim state As New ToolCallSequencing.ToolingRunState() With {
                .MemoryGroundingMode = ToolCallSequencing.MemoryGroundingMode.Required
            }

            ToolCallSequencing.NoteMemoryGroundingToolResult(
                state,
                Agents.MemoryTools.ToolList,
                "[{""key"":""mem_1"",""summary"":""stub""},{""key"":""mem_2"",""summary"":""stub""},{""key"":""mem_3"",""summary"":""stub""}]",
                succeeded:=True)

            ToolCallSequencing.NoteMemoryGroundingToolResult(
                state,
                Agents.MemoryTools.ToolGet,
                "{""key"":""mem_1"",""summary"":""stub"",""value"":{""note"":""body1""}}",
                succeeded:=True)

            Dim responseText As String =
                "Done." & vbCrLf &
                "<TASK_STATUS>{""status"":""complete"",""reason"":""finished"",""memoryGroundingScope"":""subset""}</TASK_STATUS>"

            Dim result = ToolCallSequencing.ValidateActiveToolingTurn(
                responseText,
                hasToolCalls:=False,
                hasUnresolvedToolFailure:=False,
                runState:=state)

            AssertEqual(ToolCallSequencing.ActiveToolingTurnKind.FinalCompleteTurn, result.TurnKind, "Subset disclosure should allow partial retrieval finalization.")
        End Sub

        Private Shared Sub TestRequiredMemoryGroundingAcceptsSubsetDisclosureForLargeList()
            Dim state As New ToolCallSequencing.ToolingRunState() With {
                .MemoryGroundingMode = ToolCallSequencing.MemoryGroundingMode.Required
            }

            ToolCallSequencing.NoteMemoryGroundingToolResult(
                state,
                Agents.MemoryTools.ToolList,
                "[{""key"":""mem_1"",""summary"":""stub""},{""key"":""mem_2"",""summary"":""stub""},{""key"":""mem_3"",""summary"":""stub""},{""key"":""mem_4"",""summary"":""stub""},{""key"":""mem_5"",""summary"":""stub""},{""key"":""mem_6"",""summary"":""stub""},{""key"":""mem_7"",""summary"":""stub""},{""key"":""mem_8"",""summary"":""stub""},{""key"":""mem_9"",""summary"":""stub""},{""key"":""mem_10"",""summary"":""stub""},{""key"":""mem_11"",""summary"":""stub""}]",
                succeeded:=True)

            ToolCallSequencing.NoteMemoryGroundingToolResult(
                state,
                Agents.MemoryTools.ToolGet,
                "{""key"":""mem_1"",""summary"":""stub"",""value"":{""note"":""body1""}}",
                succeeded:=True)

            Dim responseText As String =
                "Done." & vbCrLf &
                "<TASK_STATUS>{""status"":""complete"",""reason"":""finished"",""memoryGroundingScope"":""subset""}</TASK_STATUS>"

            Dim result = ToolCallSequencing.ValidateActiveToolingTurn(
                responseText,
                hasToolCalls:=False,
                hasUnresolvedToolFailure:=False,
                runState:=state)

            AssertEqual(ToolCallSequencing.ActiveToolingTurnKind.FinalCompleteTurn, result.TurnKind, "Large-list partial retrieval should be allowed when subset disclosure is explicit.")
        End Sub

        Private Shared Sub TestRequiredMemoryGroundingLargeListAllowsSubsetWithDisclosureGuidance()
            Dim state As New ToolCallSequencing.ToolingRunState() With {
                .MemoryGroundingMode = ToolCallSequencing.MemoryGroundingMode.Required
            }

            ToolCallSequencing.NoteMemoryGroundingToolResult(
                state,
                Agents.MemoryTools.ToolList,
                "[{""key"":""mem_1"",""summary"":""stub""},{""key"":""mem_2"",""summary"":""stub""},{""key"":""mem_3"",""summary"":""stub""},{""key"":""mem_4"",""summary"":""stub""},{""key"":""mem_5"",""summary"":""stub""},{""key"":""mem_6"",""summary"":""stub""},{""key"":""mem_7"",""summary"":""stub""},{""key"":""mem_8"",""summary"":""stub""},{""key"":""mem_9"",""summary"":""stub""},{""key"":""mem_10"",""summary"":""stub""},{""key"":""mem_11"",""summary"":""stub""}]",
                succeeded:=True)

            ToolCallSequencing.NoteMemoryGroundingToolResult(
                state,
                Agents.MemoryTools.ToolGet,
                "{""key"":""mem_1"",""summary"":""stub"",""value"":{""note"":""body1""}}",
                succeeded:=True)

            Dim prompt As String = ToolCallSequencing.BuildRequiredMemoryGroundingRepairPrompt(state)

            AssertTrue(state.FinalAnswerBasedOnSubset, "Subset flag should be true when only part of a large list was retrieved.")
            AssertTrue(prompt.Contains("retrieved subset"), "Large-list guidance should require subset disclosure when not all listed keys were retrieved.")
        End Sub

        Private Shared Sub TestRequiredMemoryGroundingRejectsFinalCompleteWithoutMemoryAccess()
            Dim state As New ToolCallSequencing.ToolingRunState() With {
                .MemoryGroundingMode = ToolCallSequencing.MemoryGroundingMode.Required
            }

            Dim result = ToolCallSequencing.ValidateActiveToolingTurn(
                BuildFinalTurnText("Done.", "complete", "finished"),
                hasToolCalls:=False,
                hasUnresolvedToolFailure:=False,
                runState:=state)

            AssertEqual(ToolCallSequencing.ActiveToolingTurnKind.InvalidTurn, result.TurnKind, "Required memory grounding must reject final complete without memory access.")
            AssertEqual(ToolCallSequencing.MissingRequiredMemoryAccessCode, result.InvalidReason, "Missing memory access reason mismatch.")
        End Sub

        Private Shared Sub TestRequiredMemoryGroundingRejectsFinalCompleteAfterMemoryListWithEntries()
            Dim state As New ToolCallSequencing.ToolingRunState() With {
                .MemoryGroundingMode = ToolCallSequencing.MemoryGroundingMode.Required
            }

            ToolCallSequencing.NoteMemoryGroundingToolResult(
                state,
                Agents.MemoryTools.ToolList,
                "[{""key"":""mem_1"",""summary"":""stub""}]",
                succeeded:=True)

            Dim result = ToolCallSequencing.ValidateActiveToolingTurn(
                BuildFinalTurnText("Done.", "complete", "finished"),
                hasToolCalls:=False,
                hasUnresolvedToolFailure:=False,
                runState:=state)

            AssertEqual(ToolCallSequencing.ActiveToolingTurnKind.InvalidTurn, result.TurnKind, "Required memory grounding must reject final complete after memory_list when entries exist but memory_get has not been called.")
            AssertEqual(ToolCallSequencing.MemoryListDoneButMemoryGetRequiredCode, result.InvalidReason, "memory_list progression rejection reason mismatch.")
        End Sub

        Private Shared Sub TestRequiredMemoryGroundingRepairPromptEscalatesToMemoryGetAfterList()
            Dim state As New ToolCallSequencing.ToolingRunState() With {
                .MemoryGroundingMode = ToolCallSequencing.MemoryGroundingMode.Required
            }

            ToolCallSequencing.NoteMemoryGroundingToolResult(
                state,
                Agents.MemoryTools.ToolList,
                "[{""key"":""mem_1"",""summary"":""stub""},{""key"":""mem_2"",""summary"":""stub""}]",
                succeeded:=True)

            Dim prompt As String = ToolCallSequencing.BuildRequiredMemoryGroundingRepairPrompt(state)

            AssertTrue(prompt.Contains("Memory grounding is required."), "Repair prompt must acknowledge required memory grounding.")
            AssertTrue(prompt.Contains("memory_list found relevant entries."), "Repair prompt must acknowledge that memory_list found entries.")
            AssertTrue(prompt.Contains("Call memory_get for the relevant memory key"), "Repair prompt must escalate to memory_get after memory_list returns entries.")
            AssertTrue(prompt.Contains("mem_1"), "Repair prompt should include available memory keys when present.")
            AssertTrue(prompt.Contains("mem_2"), "Repair prompt should include available memory keys when present.")
            AssertFalse(prompt.Contains("Call memory_list or memory_get before finalizing"), "Repair prompt must not stay at the generic first-step wording after memory_list already returned entries.")
        End Sub

        Private Shared Sub TestRequiredMemoryGroundingAcceptsFinalCompleteAfterMemoryGet()
            Dim state As New ToolCallSequencing.ToolingRunState() With {
                .MemoryGroundingMode = ToolCallSequencing.MemoryGroundingMode.Required
            }

            ToolCallSequencing.NoteMemoryGroundingToolResult(
                state,
                Agents.MemoryTools.ToolGet,
                "{""key"":""mem_1"",""summary"":""stub"",""value"":{""note"":""body""}}",
                succeeded:=True)

            Dim result = ToolCallSequencing.ValidateActiveToolingTurn(
                BuildFinalTurnText("Done.", "complete", "finished"),
                hasToolCalls:=False,
                hasUnresolvedToolFailure:=False,
                runState:=state)

            AssertEqual(ToolCallSequencing.ActiveToolingTurnKind.FinalCompleteTurn, result.TurnKind, "Required memory grounding must accept final complete after memory_get.")
        End Sub

        Private Shared Sub TestRequiredMemoryGroundingAcceptsFinalCompleteAfterEmptyMemoryList()
            Dim state As New ToolCallSequencing.ToolingRunState() With {
                .MemoryGroundingMode = ToolCallSequencing.MemoryGroundingMode.Required
            }

            ToolCallSequencing.NoteMemoryGroundingToolResult(
                state,
                Agents.MemoryTools.ToolList,
                "[]",
                succeeded:=True)

            Dim result = ToolCallSequencing.ValidateActiveToolingTurn(
                BuildFinalTurnText("Done.", "complete", "finished"),
                hasToolCalls:=False,
                hasUnresolvedToolFailure:=False,
                runState:=state)

            AssertEqual(ToolCallSequencing.ActiveToolingTurnKind.FinalCompleteTurn, result.TurnKind, "Required memory grounding must accept final complete after an empty memory_list result.")
        End Sub

        Private Shared Sub TestOptionalMemoryGroundingDoesNotRejectFinalCompleteWithoutMemoryAccess()
            Dim state As New ToolCallSequencing.ToolingRunState() With {
                .MemoryGroundingMode = ToolCallSequencing.MemoryGroundingMode.OptionalMode
            }

            Dim result = ToolCallSequencing.ValidateActiveToolingTurn(
                BuildFinalTurnText("Done.", "complete", "finished"),
                hasToolCalls:=False,
                hasUnresolvedToolFailure:=False,
                runState:=state)

            AssertEqual(ToolCallSequencing.ActiveToolingTurnKind.FinalCompleteTurn, result.TurnKind, "Optional memory grounding must not reject final complete solely for missing memory access.")
        End Sub

        Private Shared Sub TestMemoryGroundingModeNoneAllowsFinalCompleteWithoutMemoryTools()
            Dim state As New ToolCallSequencing.ToolingRunState() With {
                .MemoryGroundingMode = ToolCallSequencing.MemoryGroundingMode.None
            }

            Dim result = ToolCallSequencing.ValidateActiveToolingTurn(
                BuildFinalTurnText("Done.", "complete", "finished"),
                hasToolCalls:=False,
                hasUnresolvedToolFailure:=False,
                runState:=state)

            AssertEqual(ToolCallSequencing.ActiveToolingTurnKind.FinalCompleteTurn, result.TurnKind, "Memory grounding mode none must allow final complete without memory tools.")
        End Sub

        Private Shared Sub TestMemoryStubsAloneDoNotSatisfyRequiredGrounding()
            Dim state As New ToolCallSequencing.ToolingRunState() With {
                .MemoryGroundingMode = ToolCallSequencing.MemoryGroundingMode.Required,
                .LastStructuredToolResult = "[memory:mem_stub_only] stub only"
            }

            Dim result = ToolCallSequencing.ValidateActiveToolingTurn(
                BuildFinalTurnText("Done.", "complete", "finished"),
                hasToolCalls:=False,
                hasUnresolvedToolFailure:=False,
                runState:=state)

            AssertEqual(ToolCallSequencing.ActiveToolingTurnKind.InvalidTurn, result.TurnKind, "Memory stubs alone must not satisfy required memory grounding.")
            AssertEqual(ToolCallSequencing.MissingRequiredMemoryAccessCode, result.InvalidReason, "Memory stub grounding rejection reason mismatch.")
        End Sub

        Private Shared Sub TestRequiredMemoryGroundingRepairPromptIsStrict()
            Dim prompt As String = ToolCallSequencing.BuildRequiredMemoryGroundingRepairPrompt()

            AssertTrue(prompt.Contains("Memory grounding is required."), "Repair prompt must state the missing memory grounding requirement.")
            AssertTrue(prompt.Contains("Call memory_list or memory_get before finalizing"), "Repair prompt must require memory_list or memory_get.")
            AssertTrue(prompt.Contains("return a valid blocked response"), "Repair prompt must allow a blocked response when memory cannot be accessed.")
        End Sub

        Private Shared Sub TestPromptInstructsMemoryGetForExplicitMemoryGrounding()
            Dim combinedPrompt As String =
                SharedLibrary.SharedMethods.Default_SP_Add_Tooling & " " &
                SharedLibrary.SharedMethods.Default_SP_Add_AgentLayer

            AssertTrue(
                combinedPrompt.Contains("If the user or runtime explicitly requires a memory-grounded answer, you must call memory_list or memory_get before finalizing unless the full relevant memory content has already been retrieved in this turn."),
                "Prompt must require memory_list or memory_get for explicit memory grounding.")

            AssertTrue(
                combinedPrompt.Contains("Do not claim that an answer is based on Memory from chat context alone."),
                "Prompt must reject claiming Memory grounding from chat context alone.")
        End Sub

        Private Shared Sub TestPromptInstructsUserFacingProseToFollowLatestUserRequestLanguage()
            Dim combinedPrompt As String =
                SharedLibrary.SharedMethods.Default_SP_Add_Tooling & " " &
                SharedLibrary.SharedMethods.Default_SP_Add_AgentLayer

            AssertTrue(
                combinedPrompt.Contains("Use the language of the latest user request for user-facing prose unless an explicit output language is provided."),
                "Prompt must instruct user-facing prose to follow the latest user request language.")

            AssertTrue(
                combinedPrompt.Contains("Do not switch to English merely because tools, logs, memory keys, or internal instructions are in English."),
                "Prompt must prevent English drift from internal artifacts.")
        End Sub


        Private Shared Sub TestInvalidNarrationAfterStructuredResultIsRejected()
            Dim state As New ToolCallSequencing.ToolingRunState()

            ToolCallSequencing.NoteToolResultForRepair(
                state,
                "fake_tool",
                "{""summary"":""ok"",""result"":{""output_reference"":""partial/ref.bin""},""resultKind"":""json_object""}",
                "json_object")

            Dim result = ToolCallSequencing.ValidateActiveToolingTurn(
                "I will continue with the next step.",
                hasToolCalls:=False,
                hasUnresolvedToolFailure:=False)

            AssertEqual(ToolCallSequencing.ActiveToolingTurnKind.InvalidTurn, result.TurnKind, "Invalid narration after a structured tool result must be rejected.")
            AssertEqual("missing_task_status", result.InvalidReason, "Structured-result narration rejection reason mismatch.")
            AssertEqual("fake_tool", state.LastStructuredToolName, "Structured result should remain attached to the last successful tool.")
        End Sub

        Private Shared Sub TestStructuredResultRemainsAvailableDuringRepair()
            Dim state As New ToolCallSequencing.ToolingRunState() With {
                .UserLanguage = "en-US"
            }

            ToolCallSequencing.NoteToolResultForRepair(
                state,
                "fake_tool",
                "{""summary"":""ok"",""result"":{""output_reference"":""partial/ref.bin""},""resultKind"":""json_object""}",
                "json_object")

            Dim prompt As String = ToolCallSequencing.BuildActiveToolingRepairPrompt(state)

            AssertEqual("fake_tool", state.LastStructuredToolName, "Structured result tool name mismatch.")
            AssertTrue(state.LastStructuredToolResult.Contains("""summary"":""ok"""), "Structured result payload should remain available during repair.")
            AssertEqual("partial/ref.bin", state.LastKnownOutputReference, "Structured output reference should remain available during repair.")
            AssertTrue(prompt.Contains("structured tool result"), "Repair prompt should explicitly preserve the structured tool result.")
        End Sub

        Private Shared Sub TestFallbackUsesRuntimeUserLanguage()
            Dim state As New ToolCallSequencing.ToolingRunState() With {
                .UserLanguage = "de-DE",
                .LastSuccessfulToolCall = "fake_tool"
            }

            Dim message As String = ToolCallSequencing.BuildUserSafeBlockedFinalMessage(
                state,
                ToolCallSequencing.InvalidTextOnlyFinalizationCode,
                "internal-only",
                successCount:=2,
                failedCount:=1,
                appendTaskStatusFooter:=True)

            AssertTrue(message.Contains("Der Tool-Ablauf wurde beendet"), "Blocked fallback should use the runtime user language value.")
            AssertTrue(message.Contains("Letzter erfolgreicher Tool-Aufruf: fake_tool."), "Localized fallback should include the last successful tool.")
        End Sub

        Private Shared Sub TestToolCallCountIsReportedOnlyAsToolCallCount()
            Dim state As New ToolCallSequencing.ToolingRunState() With {
                .UserLanguage = "en-US"
            }

            Dim message As String = ToolCallSequencing.BuildUserSafeBlockedFinalMessage(
                state,
                ToolCallSequencing.InvalidTextOnlyFinalizationCode,
                "internal-only",
                successCount:=4,
                failedCount:=1,
                appendTaskStatusFooter:=True)

            AssertTrue(message.Contains("Successful tool-call count: 4."), "Blocked fallback must label the successful count as a tool-call count.")
            AssertTrue(message.Contains("Failed tool-call count: 1."), "Blocked fallback must label the failed count as a tool-call count.")
            AssertTrue(message.Contains("tool-call counts only"), "Blocked fallback must explicitly avoid conflating tool-call counts with business completion.")
            AssertTrue(message.IndexOf("items processed", StringComparison.OrdinalIgnoreCase) < 0, "Blocked fallback must not claim business-item completion.")
        End Sub

        Private Shared Sub TestRepairPromptRemainsGeneric()
            Dim prompt As String = ToolCallSequencing.BuildActiveToolingRepairPrompt()

            AssertTrue(prompt.IndexOf("Word", StringComparison.OrdinalIgnoreCase) < 0, "Repair prompt must not contain workflow-specific product names.")
            AssertTrue(prompt.IndexOf("Outlook", StringComparison.OrdinalIgnoreCase) < 0, "Repair prompt must not contain workflow-specific product names.")
            AssertTrue(prompt.IndexOf(".docx", StringComparison.OrdinalIgnoreCase) < 0, "Repair prompt must not contain document-type examples.")
            AssertTrue(prompt.IndexOf(".xlsx", StringComparison.OrdinalIgnoreCase) < 0, "Repair prompt must not contain document-type examples.")
            AssertTrue(prompt.IndexOf("table", StringComparison.OrdinalIgnoreCase) < 0, "Repair prompt must remain generic and avoid workflow examples.")
        End Sub

        Private Shared Function BuildFinalTurnText(body As String, status As String, reason As String) As String
            Return body & " " & ToolCallSequencing.BuildTaskStatusFooter(status, reason)
        End Function

        Private Shared Function BuildFakeToolRegistry(toolNames As IEnumerable(Of String)) As Agents.ToolRegistry
            Dim tools As New List(Of SharedLibrary.ModelConfig)()

            For Each toolName In toolNames
                Dim normalizedName As String = If(toolName, "").Trim()
                If normalizedName = "" Then Continue For

                tools.Add(New SharedLibrary.ModelConfig() With {
                    .Tool = True,
                    .ToolName = normalizedName,
                    .ToolInstructionsPrompt = normalizedName & ": test tool",
                    .ToolDefinition =
                        "{""name"":""" & normalizedName & """," &
                        """description"":""Test tool.""," &
                        """parameters"":{""type"":""object"",""properties"":{},""additionalProperties"":false}}",
                    .ToolPriority = 0,
                    .ToolErrorHandling = "skip"
                })
            Next

            Return Agents.ToolRegistryBuilder.FromModelConfigs(tools, "test")
        End Function

        Private Shared Sub AssertToolNameSequenceEqual(expected As IEnumerable(Of String),
                                                       actual As IEnumerable(Of String),
                                                       message As String)
            Dim expectedText As String = String.Join("|", If(expected, Array.Empty(Of String)()))
            Dim actualText As String = String.Join("|", If(actual, Array.Empty(Of String)()))
            AssertEqual(expectedText, actualText, message)
        End Sub


        Private NotInheritable Class MockRepeatedToolScopeHost
            Implements ISubAgentHost

            Private ReadOnly _hostKind As String
            Private ReadOnly _parentToolRegistrySnapshot As Agents.ToolRegistry
            Private ReadOnly _authoritativeRegistryAvailable As Boolean
            Private ReadOnly _parentRunId As String = Guid.NewGuid().ToString("N")
            Private ReadOnly _agentInvocationCounts As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
            Private _invocationIndex As Integer

            Public ReadOnly Invocations As New List(Of MockToolScopeInvocationRecord)()

            Public Sub New(hostKind As String,
                   authoritativeRegistry As Agents.ToolRegistry,
                   Optional authoritativeRegistryAvailable As Boolean = True)
                _hostKind = hostKind
                _authoritativeRegistryAvailable = authoritativeRegistryAvailable

                If authoritativeRegistryAvailable AndAlso authoritativeRegistry IsNot Nothing Then
                    _parentToolRegistrySnapshot = authoritativeRegistry.Snapshot()
                Else
                    _parentToolRegistrySnapshot = Nothing
                End If
            End Sub

            Public Function RunIsolatedToolingLoopAsync(request As SubAgentRunRequest,
                                                cancellationToken As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of String) _
        Implements ISubAgentHost.RunIsolatedToolingLoopAsync

                _invocationIndex += 1

                Dim sameAgentInvocationCount As Integer = 1

                If _agentInvocationCounts.TryGetValue(request.AgentName, sameAgentInvocationCount) Then
                    sameAgentInvocationCount += 1
                    _agentInvocationCounts(request.AgentName) = sameAgentInvocationCount
                Else
                    sameAgentInvocationCount = 1
                    _agentInvocationCounts(request.AgentName) = 1
                End If

                Dim authoritativeSnapshot As Agents.ToolRegistry =
            If(_parentToolRegistrySnapshot Is Nothing,
               Nothing,
               _parentToolRegistrySnapshot.Snapshot())

                Dim snapshotExists As Boolean = (authoritativeSnapshot IsNot Nothing)
                Dim snapshotToolCount As Integer =
            If(authoritativeSnapshot Is Nothing,
               0,
               authoritativeSnapshot.ListNames().Count)

                Dim scopeInit = Agents.SubAgentToolScopeInitializer.Initialize(
            authoritativeSnapshot,
            request.AllowedToolNames)

                Invocations.Add(New MockToolScopeInvocationRecord() With {
            .HostKind = _hostKind,
            .ParentRunId = _parentRunId,
            .InvocationIndex = _invocationIndex,
            .AgentName = request.AgentName,
            .AgentInvocationCount = sameAgentInvocationCount,
            .SnapshotExists = snapshotExists,
            .SnapshotToolCount = snapshotToolCount,
            .RequestedTools = New List(Of String)(If(scopeInit.RequestedToolNames, New List(Of String)())),
            .ResolvedTools = New List(Of String)(scopeInit.ResolvedToolNames),
            .MissingTools = New List(Of String)(scopeInit.MissingToolNames),
            .FinalSelectedTools = New List(Of String)(scopeInit.FinalSelectedToolNames),
            .FinalCallableTools = New List(Of String)(scopeInit.FinalCallableToolNames)
        })

                If Not snapshotExists Then
                    Return System.Threading.Tasks.Task.FromResult(
                Agents.SubAgentRuntimeHardening.BuildParentRegistryMissingPayload(
                    requestedToolNames:=scopeInit.RequestedToolNames))
                End If

                If scopeInit.HasRequestedTools AndAlso
           (scopeInit.HasMissingRequestedTools OrElse scopeInit.HasMissingFinalToolNames) Then

                    Dim missingRequiredToolNames As List(Of String) = scopeInit.MissingFinalToolNames
                    If missingRequiredToolNames.Count = 0 Then
                        missingRequiredToolNames = New List(Of String)(scopeInit.MissingToolNames)
                    End If

                    Return System.Threading.Tasks.Task.FromResult(
                Agents.SubAgentRuntimeHardening.BuildRequiredToolMissingPayload(
                    missingRequiredToolNames,
                    requestedToolNames:=scopeInit.RequestedToolNames,
                    resolvedToolNames:=scopeInit.ResolvedToolNames))
                End If

                Dim payload As New JObject(
            New JProperty("summary", String.Join(", ", scopeInit.FinalCallableToolNames)),
            New JProperty("result", New JObject(
                New JProperty("hostKind", _hostKind),
                New JProperty("parentRunId", _parentRunId),
                New JProperty("invocationIndex", _invocationIndex),
                New JProperty("agentInvocationCount", sameAgentInvocationCount),
                New JProperty("snapshotExists", snapshotExists),
                New JProperty("snapshotToolCount", snapshotToolCount),
                New JProperty("requestedTools", New JArray(If(scopeInit.RequestedToolNames, New List(Of String)()).ToArray())),
                New JProperty("resolvedTools", New JArray(scopeInit.ResolvedToolNames.ToArray())),
                New JProperty("missingTools", New JArray(scopeInit.MissingToolNames.ToArray())),
                New JProperty("finalSelectedTools", New JArray(scopeInit.FinalSelectedToolNames.ToArray())),
                New JProperty("finalCallableTools", New JArray(scopeInit.FinalCallableToolNames.ToArray()))
            ))
        )

                Return System.Threading.Tasks.Task.FromResult(payload.ToString())
            End Function
        End Class


        Private NotInheritable Class MockToolScopeInvocationRecord
            Public Property HostKind As String
            Public Property ParentRunId As String
            Public Property InvocationIndex As Integer
            Public Property AgentName As String
            Public Property AgentInvocationCount As Integer
            Public Property RequestedTools As List(Of String)
            Public Property ResolvedTools As List(Of String)
            Public Property MissingTools As List(Of String)
            Public Property FinalSelectedTools As List(Of String)
            Public Property FinalCallableTools As List(Of String)

            Public Property SnapshotExists As Boolean
            Public Property SnapshotToolCount As Integer
        End Class

        Private Shared Function RunMockBatchSkill(host As ISubAgentHost,
                                                  ag As AgentDescriptor,
                                                  items As IEnumerable(Of String)) As System.Collections.Generic.List(Of String)
            Dim results As New System.Collections.Generic.List(Of String)()

            For Each item In items
                Dim payload = SubAgentRunner.InvokeResolvedAsync(
                    host,
                    ag,
                    item,
                    contextBlob:=Nothing,
                    storeResultInMemory:=False,
                    cancellationToken:=System.Threading.CancellationToken.None).GetAwaiter().GetResult()

                Dim obj = JObject.Parse(payload)
                If obj("error") IsNot Nothing Then
                    Throw New InvalidOperationException("Mock batch skill expected success.")
                End If

                results.Add(If(obj.Value(Of String)("summary"), ""))
            Next

            Return results
        End Function

        Private Shared Function RunMockBatchSkillAllowErrors(host As ISubAgentHost,
                                                             ag As AgentDescriptor,
                                                             items As IEnumerable(Of String)) As System.Collections.Generic.List(Of String)
            Dim results As New System.Collections.Generic.List(Of String)()

            For Each item In items
                Dim payload = SubAgentRunner.InvokeResolvedAsync(
                    host,
                    ag,
                    item,
                    contextBlob:=Nothing,
                    storeResultInMemory:=False,
                    cancellationToken:=System.Threading.CancellationToken.None).GetAwaiter().GetResult()

                Dim obj = JObject.Parse(payload)
                Dim errObj As JObject = TryCast(obj("error"), JObject)

                If errObj IsNot Nothing Then
                    results.Add("ERR:" & If(errObj.Value(Of String)("code"), ""))
                Else
                    results.Add(If(obj.Value(Of String)("summary"), ""))
                End If
            Next

            Return results
        End Function

        Private Shared Function BuildMockSuccessPayload(summary As String) As String
            Return New JObject(
                New JProperty("summary", summary),
                New JProperty("result", summary)
            ).ToString()
        End Function

        Private NotInheritable Class MockSubAgentHost
            Implements ISubAgentHost

            Private ReadOnly _responses As New System.Collections.Generic.Dictionary(Of String, System.Collections.Generic.Queue(Of String))(StringComparer.Ordinal)
            Public Property CallCount As Integer

            Public Sub AddResponse(task As String, response As String)
                Dim q As System.Collections.Generic.Queue(Of String) = Nothing

                If Not _responses.TryGetValue(task, q) Then
                    q = New System.Collections.Generic.Queue(Of String)()
                    _responses(task) = q
                End If

                q.Enqueue(response)
            End Sub

            Public Function RunIsolatedToolingLoopAsync(request As SubAgentRunRequest,
                                                        cancellationToken As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of String) _
                Implements ISubAgentHost.RunIsolatedToolingLoopAsync

                CallCount += 1

                Dim taskKey As String = ExtractTaskKey(request.UserMessage)
                Dim q As System.Collections.Generic.Queue(Of String) = Nothing

                If _responses.TryGetValue(taskKey, q) AndAlso q.Count > 0 Then
                    Return System.Threading.Tasks.Task.FromResult(q.Dequeue())
                End If

                Return System.Threading.Tasks.Task.FromResult(SubAgentRuntimeHardening.BuildModelEmptyResponsePayload())
            End Function

            Private Shared Function ExtractTaskKey(userMessage As String) As String
                Dim text As String = If(userMessage, "").Replace(vbCrLf, vbLf)
                Dim lines = text.Split({vbLf}, StringSplitOptions.None)
                Dim foundTask As Boolean = False
                Dim sb As New System.Text.StringBuilder()

                For Each line In lines
                    If Not foundTask Then
                        If line.Trim().Equals("Task:", StringComparison.OrdinalIgnoreCase) Then
                            foundTask = True
                        End If
                        Continue For
                    End If

                    If line.Trim() = "" Then Exit For

                    If sb.Length > 0 Then sb.AppendLine()
                    sb.Append(line)
                Next

                Return sb.ToString().Trim()
            End Function
        End Class

        Private Shared Sub AssertFalse(condition As Boolean, message As String)
            If condition Then
                Throw New InvalidOperationException(message)
            End If
        End Sub

        Private Shared Sub AssertEqual(expected As Integer, actual As Integer, message As String)
            If expected <> actual Then
                Throw New InvalidOperationException($"{message} Expected='{expected}', Actual='{actual}'.")
            End If
        End Sub

        Private Shared Sub AssertEqual(expected As String, actual As String, message As String)
            If Not String.Equals(expected, actual, StringComparison.Ordinal) Then
                Throw New InvalidOperationException($"{message} Expected='{expected}', Actual='{actual}'.")
            End If
        End Sub

        Private Shared Sub AssertEqual(expected As ToolCallSequencing.ToolCallClassification,
                                       actual As ToolCallSequencing.ToolCallClassification,
                                       message As String)
            If expected <> actual Then
                Throw New InvalidOperationException($"{message} Expected='{expected}', Actual='{actual}'.")
            End If
        End Sub

        Private NotInheritable Class WorkspaceToolSchemaContract
            Public Property Properties As List(Of String)
            Public Property Required As List(Of String)
        End Class

        Private Shared Sub TestWorkspaceWriteWithCanonicalTextWritesFullContent()
            WithTemporaryWorkspace(
                Sub(root)
                    Dim result = JObject.Parse(
                        WorkspaceTools.Execute(
                            WorkspaceTools.ToolWrite,
                            New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase) From {
                                {"path", "canonical.txt"},
                                {"text", "alpha"},
                                {"mode", "overwrite"}
                            }))

                    AssertTrue(result("error") Is Nothing, "Canonical workspace_write should succeed.")
                    AssertEqual("overwrite", result.Value(Of String)("mode"), "Canonical mode mismatch.")
                    AssertEqual(5, result.Value(Of Integer)("charsWritten"), "Canonical charsWritten mismatch.")

                    Dim readBack = JObject.Parse(
                        WorkspaceTools.Execute(
                            WorkspaceTools.ToolRead,
                            New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase) From {
                                {"path", "canonical.txt"}
                            }))

                    AssertEqual("alpha", readBack.Value(Of String)("text"), "Canonical write read-back mismatch.")
                End Sub)
        End Sub

        Private Shared Sub TestWorkspaceWriteWithLegacyContentAliasWritesFullContent()
            WithTemporaryWorkspace(
                Sub(root)
                    Dim result = JObject.Parse(
                        WorkspaceTools.Execute(
                            WorkspaceTools.ToolWrite,
                            New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase) From {
                                {"path", "legacy.txt"},
                                {"content", "bravo"},
                                {"mode", "overwrite"}
                            }))

                    AssertTrue(result("error") Is Nothing, "Legacy content alias should succeed.")
                    AssertTrue(result.Value(Of Boolean)("usedAliases"), "Legacy content alias should be reported.")
                    AssertTrue(JsonArrayContains(result("aliasesUsed"), "content->text"), "content alias mapping was not reported.")

                    Dim readBack = JObject.Parse(
                        WorkspaceTools.Execute(
                            WorkspaceTools.ToolRead,
                            New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase) From {
                                {"path", "legacy.txt"}
                            }))

                    AssertEqual("bravo", readBack.Value(Of String)("text"), "Legacy content alias read-back mismatch.")
                End Sub)
        End Sub

        Private Shared Sub TestWorkspaceWriteOverwriteAliasMapsToOverwriteMode()
            WithTemporaryWorkspace(
                Sub(root)
                    Dim firstWrite = JObject.Parse(
                        WorkspaceTools.Execute(
                            WorkspaceTools.ToolWrite,
                            New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase) From {
                                {"path", "mode-map.txt"},
                                {"text", "old"},
                                {"mode", "overwrite"}
                            }))

                    AssertTrue(firstWrite("error") Is Nothing, "Initial write should succeed.")

                    Dim secondWrite = JObject.Parse(
                        WorkspaceTools.Execute(
                            WorkspaceTools.ToolWrite,
                            New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase) From {
                                {"path", "mode-map.txt"},
                                {"content", "new"},
                                {"overwrite", True}
                            }))

                    AssertTrue(secondWrite("error") Is Nothing, "overwrite alias write should succeed.")
                    AssertEqual("overwrite", secondWrite.Value(Of String)("mode"), "overwrite alias should normalize to overwrite mode.")
                    AssertTrue(JsonArrayContains(secondWrite("aliasesUsed"), "overwrite->mode"), "overwrite alias mapping was not reported.")

                    Dim readBack = JObject.Parse(
                        WorkspaceTools.Execute(
                            WorkspaceTools.ToolRead,
                            New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase) From {
                                {"path", "mode-map.txt"}
                            }))

                    AssertEqual("new", readBack.Value(Of String)("text"), "overwrite alias did not replace content.")
                End Sub)
        End Sub

        Private Shared Sub TestWorkspaceWriteMissingTextFailsValidation()
            WithTemporaryWorkspace(
        Sub(root)
            Dim result = JObject.Parse(
                WorkspaceTools.Execute(
                    WorkspaceTools.ToolWrite,
                    New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase) From {
                        {"path", "missing.txt"}
                    }))

            AssertFalse(result.Value(Of Boolean)("ok"), "Missing text should fail validation.")
            AssertEqual("missing_required_tool_argument",
                        result("error")?.Value(Of String)("code"),
                        "Missing text error code mismatch.")
            AssertTrue(JsonArrayContains(result("error")?("missing"), "text"), "Missing text field was not reported.")
            AssertFalse(File.Exists(Path.Combine(root, "missing.txt")), "Missing text must not create a file.")
        End Sub)
        End Sub

        Private Shared Sub TestWorkspaceWriteUnknownContentFieldDoesNotSilentlyCreateTinyFile()
            WithTemporaryWorkspace(
        Sub(root)
            Dim targetPath = Path.Combine(root, "unknown.txt")

            Dim result = JObject.Parse(
                WorkspaceTools.Execute(
                    WorkspaceTools.ToolWrite,
                    New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase) From {
                        {"path", "unknown.txt"},
                        {"body", "payload that must not be discarded"}
                    }))

            AssertFalse(result.Value(Of Boolean)("ok"), "Unknown content field should fail validation.")
            AssertEqual("unknown_tool_argument",
                        result("error")?.Value(Of String)("code"),
                        "Unknown content field error code mismatch.")
            AssertTrue(JsonArrayContains(result("error")?("unknown"), "body"), "Unknown field was not reported.")
            AssertFalse(File.Exists(targetPath), "Unknown content field must not create an empty or tiny file.")
        End Sub)
        End Sub

        Private Shared Sub TestWorkspaceToolSchemasMatchDeclaredContracts()
            Dim expected = GetExpectedWorkspaceToolContracts()
            Dim actual = WorkspaceTools.BuildAll()

            AssertEqual(expected.Count, actual.Count, "Workspace tool count mismatch.")

            For Each tool In actual
                Dim def = JObject.Parse(tool.ToolDefinition)
                Dim name = def.Value(Of String)("name")

                AssertTrue(expected.ContainsKey(name), "Unexpected workspace tool definition: " & name)

                Dim contract = expected(name)
                Dim actualProperties = GetJsonPropertyNames(def("parameters")?("properties"))
                Dim actualRequired = GetJsonStringArray(def("parameters")?("required"))

                AssertStringSetEqual(contract.Properties, actualProperties, "Schema property mismatch for " & name & ".")
                AssertStringSetEqual(contract.Required, actualRequired, "Schema required mismatch for " & name & ".")
            Next
        End Sub

        Private Shared Function GetExpectedWorkspaceToolContracts() As Dictionary(Of String, WorkspaceToolSchemaContract)
            Return New Dictionary(Of String, WorkspaceToolSchemaContract)(StringComparer.OrdinalIgnoreCase) From {
                {
                    WorkspaceTools.ToolGet,
                    New WorkspaceToolSchemaContract() With {
                        .Properties = New List(Of String)(),
                        .Required = New List(Of String)()
                    }
                },
                {
                    WorkspaceTools.ToolInventory,
                    New WorkspaceToolSchemaContract() With {
                        .Properties = New List(Of String) From {"path", "glob", "recursive", "max_items"},
                        .Required = New List(Of String)()
                    }
                },
                {
                    WorkspaceTools.ToolRead,
                    New WorkspaceToolSchemaContract() With {
                        .Properties = New List(Of String) From {"path", "max_chars"},
                        .Required = New List(Of String) From {"path"}
                    }
                },
                {
                    WorkspaceTools.ToolReadMany,
                    New WorkspaceToolSchemaContract() With {
                        .Properties = New List(Of String) From {"paths", "max_chars_per_file", "max_files"},
                        .Required = New List(Of String) From {"paths"}
                    }
                },
                {
                    WorkspaceTools.ToolWrite,
                    New WorkspaceToolSchemaContract() With {
                        .Properties = New List(Of String) From {"path", "text", "mode"},
                        .Required = New List(Of String) From {"path", "text"}
                    }
                },
                {
                    WorkspaceTools.ToolSearch,
                    New WorkspaceToolSchemaContract() With {
                        .Properties = New List(Of String) From {"query", "glob", "recursive", "regex", "ignore_case", "max_files", "max_hits_per_file"},
                        .Required = New List(Of String) From {"query"}
                    }
                },
                {
                    WorkspaceTools.ToolCopy,
                    New WorkspaceToolSchemaContract() With {
                        .Properties = New List(Of String) From {"source", "destination", "overwrite"},
                        .Required = New List(Of String) From {"source", "destination"}
                    }
                },
                {
                    WorkspaceTools.ToolMove,
                    New WorkspaceToolSchemaContract() With {
                        .Properties = New List(Of String) From {"source", "destination"},
                        .Required = New List(Of String) From {"source", "destination"}
                    }
                },
                {
                    WorkspaceTools.ToolRename,
                    New WorkspaceToolSchemaContract() With {
                        .Properties = New List(Of String) From {"path", "new_name"},
                        .Required = New List(Of String) From {"path", "new_name"}
                    }
                },
                {
                    WorkspaceTools.ToolDelete,
                    New WorkspaceToolSchemaContract() With {
                        .Properties = New List(Of String) From {"path", "to_trash"},
                        .Required = New List(Of String) From {"path"}
                    }
                },
                {
                    WorkspaceTools.ToolMakeDir,
                    New WorkspaceToolSchemaContract() With {
                        .Properties = New List(Of String) From {"path"},
                        .Required = New List(Of String) From {"path"}
                    }
                },
                {
                    WorkspaceTools.ToolExtractText,
                    New WorkspaceToolSchemaContract() With {
                        .Properties = New List(Of String) From {"path", "max_chars", "start_char", "offset", "start_page", "end_page"},
                        .Required = New List(Of String) From {"path"}
                    }
                },
                {
                    WorkspaceTools.ToolExtractTextMany,
                    New WorkspaceToolSchemaContract() With {
                        .Properties = New List(Of String) From {"paths", "max_chars_per_file", "max_files"},
                        .Required = New List(Of String) From {"paths"}
                    }
                }
            }
        End Function

        Private Shared Sub WithTemporaryWorkspace(test As Action(Of String))
            Dim previous = WorkspaceTools.Active
            Dim root = Path.Combine(Path.GetTempPath(), "RedInkWorkspaceToolsTests_" & Guid.NewGuid().ToString("N"))

            Directory.CreateDirectory(root)

            Try
                WorkspaceTools.SetActive(New WorkspaceState() With {
                    .RootPath = root,
                    .AllowRead = True,
                    .AllowWrite = True,
                    .AllowMoveCopyRename = True,
                    .AllowDelete = True,
                    .IncludeHiddenSystem = False
                })

                test.Invoke(root)
            Finally
                WorkspaceTools.SetActive(previous)

                Try
                    If Directory.Exists(root) Then
                        Directory.Delete(root, recursive:=True)
                    End If
                Catch
                End Try
            End Try
        End Sub

        Private Shared Function JsonArrayContains(token As JToken, expected As String) As Boolean
            Dim arr = TryCast(token, JArray)
            If arr Is Nothing Then Return False

            For Each item As JToken In arr
                If String.Equals(item.ToString(), expected, StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If
            Next

            Return False
        End Function

        Private Shared Function GetJsonPropertyNames(token As JToken) As List(Of String)
            Dim result As New List(Of String)()
            Dim obj = TryCast(token, JObject)
            If obj Is Nothing Then Return result

            For Each prop In obj.Properties()
                result.Add(prop.Name)
            Next

            Return result
        End Function

        Private Shared Function GetJsonStringArray(token As JToken) As List(Of String)
            Dim result As New List(Of String)()
            Dim arr = TryCast(token, JArray)
            If arr Is Nothing Then Return result

            For Each item As JToken In arr
                result.Add(item.ToString())
            Next

            Return result
        End Function

        Private Shared Sub AssertStringSetEqual(expected As IEnumerable(Of String),
                                                actual As IEnumerable(Of String),
                                                message As String)
            Dim expectedSet As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim actualSet As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            If expected IsNot Nothing Then
                For Each item In expected
                    expectedSet.Add(item)
                Next
            End If

            If actual IsNot Nothing Then
                For Each item In actual
                    actualSet.Add(item)
                Next
            End If

            If expectedSet.Count <> actualSet.Count Then
                Throw New InvalidOperationException(message &
                                                    " Expected={" & String.Join(", ", expectedSet) & "}" &
                                                    " Actual={" & String.Join(", ", actualSet) & "}")
            End If

            For Each item In expectedSet
                If Not actualSet.Contains(item) Then
                    Throw New InvalidOperationException(message &
                                                        " Expected={" & String.Join(", ", expectedSet) & "}" &
                                                        " Actual={" & String.Join(", ", actualSet) & "}")
                End If
            Next
        End Sub


        Private Shared Sub TestSkippedFailureDoesNotShrinkRegistrySnapshot()
            Dim host As New MockRepeatedToolScopeHost(
                "Word",
                BuildFakeToolRegistry({"fake_tool_alpha", "fake_tool_beta", "fake_tool_gamma"}))

            Dim ag As New AgentDescriptor() With {
                .Name = "neutral_fake"
            }
            ag.AllowedTools.Add("fake_tool_alpha")
            ag.AllowedTools.Add("fake_tool_beta")
            ag.AllowedTools.Add("fake_tool_gamma")

            Dim firstPayload = SubAgentRunner.InvokeResolvedAsync(
                host,
                ag,
                "before-skipped-failure",
                contextBlob:=Nothing,
                storeResultInMemory:=False,
                cancellationToken:=System.Threading.CancellationToken.None).GetAwaiter().GetResult()

            AssertTrue(JObject.Parse(firstPayload)("error") Is Nothing, "First invocation should succeed.")

            Dim runState As New ToolCallSequencing.ToolingRunState()
            runState.NoteToolFailure("agent_mock",
                                     SubAgentRuntimeHardening.ModelEmptyResponseCode,
                                     "empty",
                                     skippedByPolicy:=True,
                                     returnedToParent:=True)
            AssertTrue(runState.RequiresParentRecovery, "Skipped failure should require parent recovery.")
            runState.NoteRecoveryByLaterToolCall("fake_tool_alpha")
            AssertFalse(runState.HasUnresolvedToolFailure, "Later tool call should recover skipped failure.")

            Dim secondPayload = SubAgentRunner.InvokeResolvedAsync(
                host,
                ag,
                "after-skipped-failure",
                contextBlob:=Nothing,
                storeResultInMemory:=False,
                cancellationToken:=System.Threading.CancellationToken.None).GetAwaiter().GetResult()

            AssertTrue(JObject.Parse(secondPayload)("error") Is Nothing, "Second invocation should still succeed.")

            AssertToolNameSequenceEqual(host.Invocations(0).ResolvedTools, host.Invocations(1).ResolvedTools, "Skipped failure must not shrink resolved tools.")
            AssertEqual(host.Invocations(0).SnapshotToolCount, host.Invocations(1).SnapshotToolCount, "Skipped failure must not shrink snapshot tool count.")
        End Sub

        Private Shared Sub TestPreviousSubAgentCallDoesNotMutateParentRegistrySnapshot()
            Dim host As New MockRepeatedToolScopeHost(
                "Outlook",
                BuildFakeToolRegistry({"fake_tool_alpha", "fake_tool_beta", "fake_tool_gamma"}))

            Dim ag As New AgentDescriptor() With {
                .Name = "neutral_fake"
            }
            ag.AllowedTools.Add("fake_tool_alpha")
            ag.AllowedTools.Add("fake_tool_beta")
            ag.AllowedTools.Add("fake_tool_gamma")

            For i As Integer = 1 To 3
                Dim payload = SubAgentRunner.InvokeResolvedAsync(
                    host,
                    ag,
                    "snapshot-" & i.ToString(),
                    contextBlob:=Nothing,
                    storeResultInMemory:=False,
                    cancellationToken:=System.Threading.CancellationToken.None).GetAwaiter().GetResult()

                AssertTrue(JObject.Parse(payload)("error") Is Nothing, $"Invocation {i} should succeed.")
            Next

            AssertToolNameSequenceEqual(host.Invocations(0).ResolvedTools, host.Invocations(1).ResolvedTools, "Second invocation should not mutate the parent snapshot.")
            AssertToolNameSequenceEqual(host.Invocations(0).ResolvedTools, host.Invocations(2).ResolvedTools, "Third invocation should not mutate the parent snapshot.")
            AssertEqual(host.Invocations(0).SnapshotToolCount, host.Invocations(1).SnapshotToolCount, "Second invocation snapshot count mismatch.")
            AssertEqual(host.Invocations(0).SnapshotToolCount, host.Invocations(2).SnapshotToolCount, "Third invocation snapshot count mismatch.")
        End Sub


    End Class

End Namespace

#End If