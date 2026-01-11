' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SearchResult.vb
' Purpose: Lightweight data model representing a single search result from embedding
'          similarity search. Contains match location, matched text, and similarity score.
'
' Architecture:
' -------------
' SearchResult is a simple POCO (Plain Old CLR Object) returned by embedding search
' engines (EmbeddingStore, EmbeddingStore_BagofWords) after similarity comparison.
'
' Key Properties:
'   - DocId: Unique document identifier (file path, database key, GUID)
'   - Text: Matched text snippet (sentence, paragraph, or chunk)
'   - StartOffset/EndOffset: Character positions in source [Start, End)
'   - Score: Similarity score (typically cosine similarity 0.0-1.0)
'
' Typical Workflow:
'   1. User submits search query
'   2. Search engine computes similarity between query and indexed chunks
'   3. Top N results returned as List(Of SearchResult) sorted by Score
'   4. Consumer uses DocId + offsets to locate/highlight matched text
'
' Offset Semantics:
'   - StartOffset: 0-based inclusive character index
'   - EndOffset: 0-based exclusive character index
'   - Range: [StartOffset, EndOffset) follows C# substring convention
'   - Extraction: docText.Substring(StartOffset, EndOffset - StartOffset)
'
' Score Semantics:
'   - Cosine similarity: [0.0, 1.0] where 1.0 = perfect match, 0.0 = orthogonal
'   - Other metrics (Euclidean, dot product) may use different ranges
'   - Results typically sorted descending (highest similarity first)
'
' Thread Safety:
'   Read-safe after construction. Property setters NOT thread-safe.
'   Treat as write-once after creation, then read-only.
'
' Performance:
'   Zero overhead; plain data class with four simple properties.
'
' Dependencies:
'   None (Framework-only).
'
' =============================================================================

Namespace SharedLibrary

    ''' <summary>
    ''' Represents a single search result from embedding similarity search.
    ''' Contains source document ID, matched text, character offsets, and similarity score.
    ''' </summary>
    ''' <remarks>
    ''' Simple data transfer object (DTO) with no behavior or validation.
    ''' Lightweight and serializable for transfer across application boundaries.
    ''' 
    ''' Typical lifecycle:
    ''' 1. Search engine creates instances during similarity search
    ''' 2. Instances sorted by Score and returned as collection
    ''' 3. UI/business logic consumes for display/processing
    ''' 4. Instances discarded after use (no cleanup required)
    ''' 
    ''' Thread safety: Read-safe after initialization; avoid concurrent writes.
    ''' </remarks>
    Public Class SearchResult

        ''' <summary>Gets or sets the unique identifier of the source document.</summary>
        ''' <value>Document ID (e.g., file path, database key, GUID).</value>
        ''' <remarks>
        ''' Format is application-defined and treated as opaque string by embedding stores.
        ''' Common patterns: file paths, database keys, GUIDs.
        ''' Used for filtering results by source and retrieving full document content.
        ''' </remarks>
        Public Property DocId As String

        ''' <summary>Gets or sets the matched text snippet from the source document.</summary>
        ''' <value>Text content of matched chunk (sentence, paragraph, or custom segment).</value>
        ''' <remarks>
        ''' Contains actual indexed content that was matched. Chunk size determined by indexing strategy.
        ''' Use for: result preview, context snippets, relevance verification.
        ''' May not include full sentence boundaries if chunk splitting cuts mid-sentence.
        ''' </remarks>
        Public Property Text As String

        ''' <summary>Gets or sets the 0-based character index where matched text begins (inclusive).</summary>
        ''' <value>Inclusive start position (0-based character offset).</value>
        ''' <remarks>
        ''' Offsets follow C# substring convention: [StartOffset, EndOffset) inclusive/exclusive.
        ''' Use for: cursor navigation, exact extraction, highlighting.
        ''' Assumes document unchanged since indexing (edits invalidate offsets).
        ''' </remarks>
        Public Property StartOffset As Integer

        ''' <summary>Gets or sets the 0-based character index where matched text ends (exclusive).</summary>
        ''' <value>Exclusive end position (0-based character offset).</value>
        ''' <remarks>
        ''' Points to character AFTER last matched character (exclusive bound).
        ''' Length calculation: matchLength = EndOffset - StartOffset.
        ''' Matches .NET Substring semantics: text.Substring(start, length) where length = end - start.
        ''' </remarks>
        Public Property EndOffset As Integer

        ''' <summary>Gets or sets the similarity score between query and matched text.</summary>
        ''' <value>Similarity score (range depends on search algorithm).</value>
        ''' <remarks>
        ''' Score interpretation by algorithm:
        ''' - Cosine similarity: [0.0, 1.0] where 1.0 = perfect match, 0.0 = orthogonal (typical threshold: > 0.7)
        ''' - Dot product: [-∞, +∞] where higher = more similar (not normalized)
        ''' - Euclidean distance: [0, +∞] where lower = more similar (sometimes inverted: 1/(1+distance))
        ''' Results typically sorted descending by Score (highest first).
        ''' </remarks>
        Public Property Score As Single

    End Class
End Namespace