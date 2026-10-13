' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ToolingConstants.vb
' Purpose: Central constants for the tooling pipeline shared by Outlook, Word
'          and Excel. These replace the per-host scattered defaults.
'
' Constants:
'  - DefaultMaxToolIterations = 100 (unified across all hosts).
'  - LlmTimeoutBufferSeconds = 60 (added per iteration).
'  - MaxContinuationRetries = 5 (repair attempts for recovery).
'  - SubAgentLargeToolResponseThresholdChars = 30000 (compaction trigger).
'  - SubAgentLargeToolResponseExcerptChars = 8000 (excerpt size when compacted).
'  - MaxLocalizableBlockedFinalChars = 1500 (host post-translation willing threshold).
'  - MaxFallbackToolsListedInGuard = 6 (fallback tools listed in guard prompts).
' =============================================================================


Option Explicit On
Option Strict On

Namespace Agents

    ''' <summary>
    ''' Central tooling-pipeline constants. Per Q7 the user wants a single, unified default
    ''' across Outlook and Word. Hosts may still override via INI_ToolingMaximumIterations
    ''' at runtime, but the default seeded into INI must come from here.
    ''' </summary>
    Public Module ToolingConstants

        ''' <summary>Unified default for INI_ToolingMaximumIterations across all hosts.</summary>
        Public Const DefaultMaxToolIterations As Integer = 100

        ''' <summary>Additional seconds added to the configured LLM timeout per iteration.</summary>
        Public Const LlmTimeoutBufferSeconds As Integer = 60

        ''' <summary>Maximum repair attempts for premature-text / invalid-turn / empty-response recovery.</summary>
        Public Const MaxContinuationRetries As Integer = 5

        ''' <summary>
        ''' Number of full finalization re-planning passes allowed after the ordinary continuation-repair
        ''' budget is exhausted. The restart preserves verified side effects/artifacts and asks the model to
        ''' choose a materially different completion path; it is not permission to replay successful mutations.
        ''' </summary>
        Public Const MaxFinalizationRecoveryRestarts As Integer = 1

        ''' <summary>
        ''' Number of full tool-path recovery re-plans allowed after a concrete tool's own repair/circuit-breaker
        ''' budget becomes terminal. The recovery preserves successful state but may select a different tool or
        ''' a materially different implementation path before the host falls back to no-tool finalization.
        ''' </summary>
        Public Const MaxTerminalToolRecoveryRestarts As Integer = 1

        ''' <summary>Char threshold over which a sub-agent tool response is compacted for model replay.</summary>
        Public Const SubAgentLargeToolResponseThresholdChars As Integer = 30000

        ''' <summary>Excerpt size kept when sub-agent tool responses are compacted.</summary>
        Public Const SubAgentLargeToolResponseExcerptChars As Integer = 8000

        ' =====================================================================
        ' Tool-response payload budget (context "drawer") defaults.
        '
        ' Values are characters, not tokens or a complete model-request budget.
        ' Defaults do not proactively compact history. An empty response enables
        ' recovery; context_compact is an explicit shelf request. Weaker models may
        ' opt into a positive proactive budget with ToolResponseCompactOnlyAfterEmpty=False.
        ' Large-source reference replay is separate from history compaction and keeps
        ' verified full originals; a preview/evidence core is never a semantic summary.
        ' =====================================================================

        ''' <summary>
        ''' Overall character budget for the serialized tool-response payload injected each
        ''' iteration. When exceeded, the host progressively compacts older results by reference
        ''' (still retrievable via context_expand). 0 disables budget-driven compaction.
        ''' No automatic replay budget until an empty response; positive INI values opt into a budget.
        ''' </summary>
        Public Const ToolResponsePayloadBudgetChars As Integer = 0

        ' Configurable context contracts. Defaults keep full history until real recovery pressure.
        Public Const AutoPilotLogDiagnosticsMaxInputChars As System.Int32 = 0
        Public Const TextAnalyzeMaxSourceChars As System.Int32 = 400000
        Public Const TextAnalyzeMaxInstructionChars As System.Int32 = 20000
        Public Const TextAnalyzeMaxAdditionalContextChars As System.Int32 = 120000
        Public Const ToolResponseRecoveryBudgetChars As System.Int32 = 400000
        Public Const ToolResponseCompactOnlyAfterEmpty As System.Boolean = True
        Public Const ToolResponseKeepRecentFullCount As System.Int32 = 2
        Public Const ToolResultReferenceThresholdChars As System.Int32 = 100000
        Public Const ToolResultReferencePreviewChars As System.Int32 = 2000
        Public Const ContextExpandDefaultChars As System.Int32 = 8000
        Public Const ContextExpandMaxChars As System.Int32 = 100000
        ''' <summary>
        ''' Results carrying deliverable/reference identity larger than this are replayed
        ''' to the model through a compact lossless envelope plus result_ref. The full
        ''' ToolResponse remains untouched for artifact registration/verification.
        ''' </summary>
        Public Const DeliverableSafeReplayEnvelopeThresholdChars As Integer = 4000

        ''' <summary>First (milder) threshold for reference-compacting older medium-sized results under budget pressure. Override via INI BudgetMediumCompactionThresholdChars.</summary>
        Public Const BudgetMediumCompactionThresholdChars As Integer = 6000

        ''' <summary>Second (aggressive) threshold for reference-compacting older medium-sized results under budget pressure. Override via INI BudgetAggressiveCompactionThresholdChars.</summary>
        Public Const BudgetAggressiveCompactionThresholdChars As Integer = 2000

        ''' <summary>Preview size kept when older medium-sized results are reference-compacted under budget pressure. Override via INI BudgetCompactionPreviewChars.</summary>
        Public Const BudgetCompactionPreviewChars As Integer = 600

        ''' <summary>Maximum inline size of a deterministic evidence-core replay for a large structured result.</summary>
        Public Const EvidenceReplayCoreMaxChars As Integer = 12000

        ''' <summary>Maximum exact string length retained inside an evidence core. Longer text remains available through context_expand.</summary>
        Public Const EvidenceReplayMaxExactStringChars As Integer = 320

        ''' <summary>Preferred maximum number of exact properties retained per object inside large result arrays before progressively tighter evidence-core projection.</summary>
        Public Const EvidenceReplayPreferredArrayItemProperties As Integer = 12

        ''' <summary>Maximum length of a blocked-final string that the host is willing to translate.</summary>
        Public Const MaxLocalizableBlockedFinalChars As Integer = 1500

        ''' <summary>Maximum number of fallback tools to list in a deliverable-fallback guard prompt.</summary>
        Public Const MaxFallbackToolsListedInGuard As Integer = 6

    End Module

End Namespace
