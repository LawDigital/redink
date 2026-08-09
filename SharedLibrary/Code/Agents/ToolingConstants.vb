' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ToolingConstants.vb
' Purpose: Central constants for the tooling pipeline shared by Outlook, Word
'          and Excel. These replace the per-host scattered defaults.
'
' Constants:
'  - DefaultMaxToolIterations = 50 (unified across all hosts).
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
        Public Const DefaultMaxToolIterations As Integer = 50

        ''' <summary>Additional seconds added to the configured LLM timeout per iteration.</summary>
        Public Const LlmTimeoutBufferSeconds As Integer = 60

        ''' <summary>Maximum repair attempts for premature-text / invalid-turn / empty-response recovery.</summary>
        Public Const MaxContinuationRetries As Integer = 5

        ''' <summary>Char threshold over which a sub-agent tool response is compacted for model replay.</summary>
        Public Const SubAgentLargeToolResponseThresholdChars As Integer = 30000

        ''' <summary>Excerpt size kept when sub-agent tool responses are compacted.</summary>
        Public Const SubAgentLargeToolResponseExcerptChars As Integer = 8000

        ' =====================================================================
        ' Tool-response payload budget (context "drawer") defaults.
        '
        ' These are the DEFAULTS (medium-context models, roughly 128k-200k tokens).
        ' They can be overridden per install via the INI keys:
        '   ToolResponsePayloadBudgetChars
        '   BudgetMediumCompactionThresholdChars
        '   BudgetAggressiveCompactionThresholdChars
        '   BudgetCompactionPreviewChars
        '
        ' Proposed profiles (character counts, not tokens; ~4 chars per token):
        '
        '   STRONG models (very large context, e.g. 1,000,000 tokens):
        '     ToolResponsePayloadBudgetChars           = 800000
        '     BudgetMediumCompactionThresholdChars     = 20000
        '     BudgetAggressiveCompactionThresholdChars = 10000
        '     BudgetCompactionPreviewChars             = 2000
        '     Rationale: keep far more history fully visible; only shelve very large
        '     results, and keep bigger previews so the model rarely needs context_expand.
        '
        '   MEDIUM models (default, ~128k-200k tokens):
        '     ToolResponsePayloadBudgetChars           = 120000
        '     BudgetMediumCompactionThresholdChars     = 6000
        '     BudgetAggressiveCompactionThresholdChars = 2000
        '     BudgetCompactionPreviewChars             = 600
        '
        '   SMALL models (tight context, e.g. 32k-64k tokens):
        '     ToolResponsePayloadBudgetChars           = 40000
        '     BudgetMediumCompactionThresholdChars     = 3000
        '     BudgetAggressiveCompactionThresholdChars = 1200
        '     BudgetCompactionPreviewChars             = 300
        '     Rationale: compact aggressively and early; keep only small previews so the
        '     active window stays within a limited context. The model relies more on
        '     context_expand to page back into shelved results on demand.
        ' =====================================================================

        ''' <summary>
        ''' Overall character budget for the serialized tool-response payload injected each
        ''' iteration. When exceeded, the host progressively compacts older results by reference
        ''' (still retrievable via context_expand). 0 disables budget-driven compaction.
        ''' Default for medium-context models; override via INI ToolResponsePayloadBudgetChars.
        ''' </summary>
        Public Const ToolResponsePayloadBudgetChars As Integer = 120000

        ''' <summary>First (milder) threshold for reference-compacting older medium-sized results under budget pressure. Override via INI BudgetMediumCompactionThresholdChars.</summary>
        Public Const BudgetMediumCompactionThresholdChars As Integer = 6000

        ''' <summary>Second (aggressive) threshold for reference-compacting older medium-sized results under budget pressure. Override via INI BudgetAggressiveCompactionThresholdChars.</summary>
        Public Const BudgetAggressiveCompactionThresholdChars As Integer = 2000

        ''' <summary>Preview size kept when older medium-sized results are reference-compacted under budget pressure. Override via INI BudgetCompactionPreviewChars.</summary>
        Public Const BudgetCompactionPreviewChars As Integer = 600

        ''' <summary>Maximum length of a blocked-final string that the host is willing to translate.</summary>
        Public Const MaxLocalizableBlockedFinalChars As Integer = 1500

        ''' <summary>Maximum number of fallback tools to list in a deliverable-fallback guard prompt.</summary>
        Public Const MaxFallbackToolsListedInGuard As Integer = 6

    End Module

End Namespace
