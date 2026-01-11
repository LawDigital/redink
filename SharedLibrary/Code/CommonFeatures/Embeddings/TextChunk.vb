' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: TextChunk.vb
' Purpose: Lightweight data model representing a text segment with position metadata
'          and optional vector embedding. Fundamental indexing unit for embedding-based
'          search systems and bag-of-words stores.
'
' Architecture:
' -------------
' TextChunk is a simple POCO (Plain Old CLR Object) that combines text content,
' position tracking, and vector embeddings for similarity search.
'
' Key Properties:
'   - Text: Raw text segment (50-1000 chars typical)
'   - Position: Sequential chunk index (0, 1, 2, ...)
'   - StartOffset/EndOffset: Character positions in source document [Start, End)
'   - Vector: Embedding array (384-1536 floats for neural, vocab-size for BoW)
'
' Chunking Strategies:
'   1. Fixed-size tokens (128, 256, 512 tokens with overlap)
'   2. Sentence-based (1-5 sentences per chunk)
'   3. Paragraph-based (semantic boundaries)
'   4. Sliding window (overlapping for better recall)
'
' Offset Semantics:
'   - StartOffset: 0-based inclusive character index
'   - EndOffset: 0-based exclusive character index
'   - Range: [StartOffset, EndOffset) follows C# substring convention
'   - Extraction: doc.Substring(StartOffset, EndOffset - StartOffset) == Text
'
' Vector Types:
'   - Neural embeddings: Dense (e.g., 384 floats for all-MiniLM-L6-v2)
'   - Bag-of-words: Vocab-sized (5000-50000 dims, mostly zeros)
'   - Initially Nothing; populated during indexing
'
' Thread Safety:
'   Read-safe after construction. Property setters NOT thread-safe.
'   Treat as write-once after creation, then read-only.
'
' Performance:
'   Memory: 3 integers + 2 references per chunk.
'   Vector dominates (384 floats = 1.5 KB; 10K chunks = ~15 MB).
'
' Dependencies:
'   None (Framework-only).
'
' =============================================================================

Option Strict On
Option Explicit On

Namespace SharedLibrary

    ''' <summary>
    ''' Represents a text segment with position metadata and optional vector embedding.
    ''' Fundamental indexing unit for embedding-based search and bag-of-words systems.
    ''' </summary>
    ''' <remarks>
    ''' Simple data transfer object (DTO) with no behavior or validation.
    ''' Lightweight and serializable for persistence and transfer.
    ''' 
    ''' Typical lifecycle:
    ''' 1. Document splitter creates chunks with Text and offsets
    ''' 2. Embedding generator computes and assigns Vector
    ''' 3. Chunks stored in EmbeddingStore for indexing
    ''' 4. Search compares query vector against chunk Vectors
    ''' 5. Top-N similar chunks returned as SearchResults
    ''' 
    ''' Thread safety: Read-safe after initialization; avoid concurrent writes.
    ''' </remarks>
    Public Class TextChunk

        ''' <summary>Gets or sets the text content of this chunk.</summary>
        ''' <value>Text segment extracted from source document.</value>
        ''' <remarks>
        ''' Contains raw text for vectorization. Common sizes: 1-5 sentences (50-300 words),
        ''' 128-512 tokens, or 200-1000 characters. Include complete sentences when possible.
        ''' Use offsets to retrieve surrounding context from original document.
        ''' </remarks>
        Public Property Text As String

        ''' <summary>Gets or sets the sequential index of this chunk within the document.</summary>
        ''' <value>0-based sequential position (0 = first chunk, 1 = second chunk, etc.).</value>
        ''' <remarks>
        ''' Ordinal index for iterating chunks in document order. NOT a character offset.
        ''' Used for: ordering (chunks.OrderBy), logging, determining relative location.
        ''' Independent of character offsets (Position=5 can start at any character position).
        ''' </remarks>
        Public Property Position As Integer

        ''' <summary>Gets or sets the 0-based character index where chunk begins (inclusive).</summary>
        ''' <value>Inclusive start position (0-based character offset).</value>
        ''' <remarks>
        ''' Offsets follow C# substring convention: [StartOffset, EndOffset) inclusive/exclusive.
        ''' Use cases: extracting text, highlighting, cursor navigation.
        ''' Assumptions: Source text unchanged since chunking (edits invalidate offsets).
        ''' </remarks>
        Public Property StartOffset As Integer

        ''' <summary>Gets or sets the 0-based character index where chunk ends (exclusive).</summary>
        ''' <value>Exclusive end position (0-based character offset).</value>
        ''' <remarks>
        ''' Points to character AFTER last chunk character (exclusive bound).
        ''' Length calculation: chunkLength = EndOffset - StartOffset.
        ''' Matches .NET Substring semantics: text.Substring(start, length) where length = end - start.
        ''' </remarks>
        Public Property EndOffset As Integer

        ''' <summary>Gets or sets the vector embedding representation of chunk text.</summary>
        ''' <value>Dense/sparse vector as Single() array. Nothing before embedding computation.</value>
        ''' <remarks>
        ''' Numerical representation in vector space for similarity search.
        ''' Common dimensions: 384 (all-MiniLM-L6-v2), 1536 (text-embedding-ada-002), 5000-50000 (BoW).
        ''' Lifecycle: Initially Nothing → Populated during indexing → Used for similarity comparison.
        ''' Memory: 384-dim = 1.5 KB/chunk; 10K chunks = ~15 MB. Avoid modifying after indexing.
        ''' </remarks>
        Public Property Vector As Single()

    End Class
End Namespace