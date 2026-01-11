' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: TokenOffset.vb
' Purpose: Lightweight data model representing a single token with its character-level
'          position in source text. Used for aligning tokenized text back to original
'          documents for highlighting, extraction, and span-based operations.
'
' Architecture:
' -------------
' TokenOffset is a simple POCO (Plain Old CLR Object) that pairs token text with
' character offsets [Start, End) using C# substring convention (inclusive/exclusive).
'
' Key Use Cases:
'   - Named Entity Recognition: Map token-level tags to character spans
'   - Syntax highlighting: Highlight keywords/tokens in editors
'   - Search results: Highlight matched tokens in original text
'   - Subword alignment: Map SentencePiece/BPE tokens to source text
'
' Offset Semantics:
'   - Start: 0-based inclusive character index where token begins
'   - End: 0-based exclusive character index where token ends
'   - Length: End - Start
'   - Extraction: originalText.Substring(Start, End - Start) == Text
'
' Example:
'   Source: "The quick brown fox"
'   Token: Text="quick", Start=4, End=9
'   Verify: "The quick brown fox".Substring(4, 5) == "quick" ✓
'
' Thread Safety:
'   Read-safe after construction. Property setters NOT thread-safe.
'   Treat as write-once after creation, then read-only.
'
' Performance:
'   Zero overhead; plain data class. Struct could reduce GC pressure for large
'   collections, but current design prioritizes reference semantics.
'
' Dependencies:
'   None (Framework-only).
'
' =============================================================================

Option Strict On
Option Explicit On

Namespace SharedLibrary

    ''' <summary>
    ''' Represents a single token with its character-level position in source text.
    ''' Enables alignment between tokenized text and original document for highlighting, extraction, and span operations.
    ''' </summary>
    ''' <remarks>
    ''' Simple data transfer object (DTO) with no behavior or validation.
    ''' Lightweight and serializable for transfer and persistence.
    ''' 
    ''' Typical lifecycle:
    ''' 1. Tokenizer creates TokenOffset instances with positions
    ''' 2. Used for model input (Text) and output alignment (Start/End)
    ''' 3. Predictions mapped back to character positions
    ''' 4. UI highlights or extracts spans using offsets
    ''' 
    ''' Thread safety: Read-safe after initialization; avoid concurrent writes.
    ''' </remarks>
    Public Class TokenOffset

        ''' <summary>Gets or sets the token text (word, subword, or punctuation).</summary>
        ''' <value>Token string extracted from source text.</value>
        ''' <remarks>
        ''' Contains the surface form as it appears in source document (possibly normalized).
        ''' Token types: words ("quick"), subwords ("token"/"ization"), punctuation (","), special tokens ("[CLS]").
        ''' For unmodified tokenization: originalText.Substring(Start, End - Start) == Text
        ''' </remarks>
        Public Property Text As String

        ''' <summary>Gets or sets the 0-based character index where token begins (inclusive).</summary>
        ''' <value>Inclusive start position (0-based character offset).</value>
        ''' <remarks>
        ''' Offsets follow C# substring convention: [Start, End) inclusive/exclusive.
        ''' Use cases: extracting text, highlighting, cursor positioning, span operations.
        ''' Assumptions: Source text unchanged since tokenization (edits invalidate offsets).
        ''' </remarks>
        Public Property Start As Integer

        ''' <summary>Gets or sets the 0-based character index where token ends (exclusive).</summary>
        ''' <value>Exclusive end position (0-based character offset).</value>
        ''' <remarks>
        ''' Points to character AFTER last token character (exclusive bound).
        ''' Length calculation: tokenLength = End - Start.
        ''' Matches .NET Substring semantics: text.Substring(start, length) where length = end - start.
        ''' </remarks>
        Public Property [End] As Integer

    End Class
End Namespace