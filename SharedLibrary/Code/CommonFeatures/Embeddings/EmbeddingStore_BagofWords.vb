' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: EmbeddingStore_BagofWords.vb
' Purpose: Lightweight, offline embedding and similarity search using classical 
'          bag-of-words vectorization with term frequency (TF) and cosine similarity.
'          Serves as a fallback when neural embeddings are unavailable.
'
' Architecture:
'   - In-memory document store mapping docId -> chunked text segments
'   - Global vocabulary rebuilt on each indexing operation (token -> dimension index)
'   - TF vectorization: term counts in dense vectors sized by vocabulary
'   - Brute-force cosine similarity search across all indexed chunks
'   - Tokenization: regex-based Unicode word/number extraction, stopword filtering, bigrams
'
' Key Limitations:
'   - NOT thread-safe (shared state in store/vocab)
'   - Vocabulary rebuild is O(D*C*T) on every IndexDocument call
'   - No TF-IDF, stemming, or lemmatization
'   - Brute-force O(N) search; suitable for <10K chunks
'   - RAM-only, no persistence
'
' Dependencies: System.Text.RegularExpressions, TextChunk, SearchResult
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Text.RegularExpressions

Namespace SharedLibrary

    ''' <summary>
    ''' In-memory bag-of-words embedding store with cosine similarity search.
    ''' Provides a lightweight alternative to neural embeddings for semantic search
    ''' when offline operation or simplicity is required.
    ''' </summary>
    ''' <remarks>
    ''' Uses term frequency (TF) vectorization without IDF weighting.
    ''' Vocabulary is rebuilt on each document indexing operation, making it
    ''' suitable for small to medium-sized document sets (typically &lt;10,000 chunks).
    ''' For larger corpora, consider neural embeddings (ONNX models or cloud APIs).
    ''' </remarks>
    Public Class EmbeddingStore_BagofWords

        ''' <summary>
        ''' Maps document IDs to their chunked text segments with computed vectors.
        ''' </summary>
        Private store As Dictionary(Of String, List(Of TextChunk))

        ''' <summary>
        ''' Global vocabulary mapping tokens to vector dimension indices.
        ''' Rebuilt on each IndexDocument call to incorporate new terms.
        ''' </summary>
        Private vocab As Dictionary(Of String, Integer)

        ''' <summary>
        ''' Initializes a new instance of the EmbeddingStore_BagofWords class.
        ''' </summary>
        Public Sub New()
            store = New Dictionary(Of String, List(Of TextChunk))()
            vocab = New Dictionary(Of String, Integer)()
        End Sub

        ''' <summary>
        ''' Indexes a document by storing its chunks and rebuilding all vectors across the entire corpus.
        ''' </summary>
        ''' <param name="docId">Unique identifier for the document.</param>
        ''' <param name="chunks">Pre-split text chunks with StartOffset/EndOffset metadata.</param>
        ''' <remarks>
        ''' WARNING: Rebuilds vocabulary and all vectors for ALL documents (O(total_chunks * avg_tokens)).
        ''' For incremental indexing, consider batching multiple documents or implementing differential updates.
        ''' </remarks>
        Public Sub IndexDocument(docId As String, chunks As List(Of TextChunk))
            store(docId) = chunks
            RebuildVectors()
        End Sub

        ''' <summary>
        ''' Rebuilds the global vocabulary and recalculates term frequency vectors for all chunks.
        ''' </summary>
        ''' <remarks>
        ''' Two-pass algorithm:
        ''' Pass 1: Scan all chunks to build vocabulary (token -> dimension index mapping).
        ''' Pass 2: For each chunk, count token occurrences and populate dense vector.
        ''' Vocabulary size determines vector dimensionality; grows unbounded with unique tokens.
        ''' </remarks>
        Private Sub RebuildVectors()
            ' Build vocabulary across all chunks
            vocab.Clear()
            For Each chunks In store.Values
                For Each chunk In chunks
                    For Each token In SimpleTokenizer(chunk.Text)
                        If Not vocab.ContainsKey(token) Then
                            vocab(token) = vocab.Count
                        End If
                    Next
                Next
            Next

            ' Assign vector for each chunk
            For Each kvp In store
                For Each chunk In kvp.Value
                    Dim counts As New Dictionary(Of Integer, Single)()
                    For Each token In SimpleTokenizer(chunk.Text)
                        Dim idx = vocab(token)
                        If counts.ContainsKey(idx) Then
                            counts(idx) += 1
                        Else
                            counts(idx) = 1
                        End If
                    Next
                    Dim vector(vocab.Count - 1) As Single
                    For Each c In counts
                        vector(c.Key) = c.Value
                    Next
                    chunk.Vector = vector
                Next
            Next
        End Sub

        ''' <summary>
        ''' Searches indexed chunks for semantic similarity to the query text.
        ''' </summary>
        ''' <param name="query">The search query text to vectorize and compare.</param>
        ''' <param name="allDocs">If True, search across all documents; if False, restrict to currentDocId.</param>
        ''' <param name="findAll">If True, return all results sorted by score; if False, return only top result.</param>
        ''' <param name="currentDocId">Document ID to restrict search to when allDocs=False.</param>
        ''' <param name="currentPosition">Ignore chunks before this offset when searching within currentDocId and findAll=False.</param>
        ''' <returns>List of SearchResult objects sorted by descending cosine similarity score.</returns>
        ''' <remarks>
        ''' Performs brute-force comparison against all candidate chunks (O(N) similarity calculations).
        ''' Score range: [0.0, 1.0] where 1.0 is perfect match, 0.0 is orthogonal (no common terms).
        ''' </remarks>
        Public Function Search(query As String,
                               allDocs As Boolean,
                               findAll As Boolean,
                               currentDocId As String,
                               currentPosition As Integer) As List(Of SearchResult)
            Dim qVec = GetVectorForText(query)
            Dim results As New List(Of SearchResult)()
            For Each kvp In store
                Dim docId = kvp.Key
                If Not allDocs AndAlso docId <> currentDocId Then Continue For
                For Each chunk In kvp.Value
                    If Not findAll AndAlso docId = currentDocId AndAlso chunk.StartOffset < currentPosition Then Continue For
                    Dim score = CosineSimilarity(qVec, chunk.Vector)
                    results.Add(New SearchResult With {
                        .DocId = docId,
                        .Text = chunk.Text,
                        .StartOffset = chunk.StartOffset,
                        .EndOffset = chunk.EndOffset,
                        .Score = score
                    })
                Next
            Next
            results.Sort(Function(a, b) b.Score.CompareTo(a.Score))
            If Not findAll AndAlso results.Count > 0 Then
                Return New List(Of SearchResult) From {results(0)}
            End If
            Return results
        End Function

        ''' <summary>
        ''' Converts arbitrary text into a term frequency vector using the current vocabulary.
        ''' </summary>
        ''' <param name="text">Text to vectorize.</param>
        ''' <returns>Dense vector of size vocab.Count with term frequencies.</returns>
        ''' <remarks>
        ''' Tokens not in vocabulary are silently ignored (zero contribution).
        ''' This is expected for query terms not seen during indexing.
        ''' </remarks>
        Private Function GetVectorForText(text As String) As Single()
            Dim counts As New Dictionary(Of Integer, Single)()
            For Each token In SimpleTokenizer(text)
                If vocab.ContainsKey(token) Then
                    Dim idx = vocab(token)
                    If counts.ContainsKey(idx) Then
                        counts(idx) += 1
                    Else
                        counts(idx) = 1
                    End If
                End If
            Next
            Dim vector(vocab.Count - 1) As Single
            For Each c In counts
                vector(c.Key) = c.Value
            Next
            Return vector
        End Function

        ''' <summary>
        ''' Computes cosine similarity between two vectors.
        ''' </summary>
        ''' <param name="vec1">First vector.</param>
        ''' <param name="vec2">Second vector.</param>
        ''' <returns>Cosine similarity in range [0.0, 1.0]; 0.0 if either vector has zero norm.</returns>
        ''' <remarks>
        ''' Formula: cos(θ) = (A · B) / (||A|| * ||B||).
        ''' Handles mismatched vector lengths by using the minimum dimension (defensive coding).
        ''' Returns 0.0 for zero-magnitude vectors to avoid division by zero.
        ''' </remarks>
        Private Function CosineSimilarity(vec1 As Single(), vec2 As Single()) As Single
            Dim dot As Single = 0, normA As Single = 0, normB As Single = 0
            For i = 0 To Math.Min(vec1.Length, vec2.Length) - 1
                dot += vec1(i) * vec2(i)
                normA += vec1(i) * vec1(i)
                normB += vec2(i) * vec2(i)
            Next
            Return If(normA > 0 AndAlso normB > 0, CSng(dot / (Math.Sqrt(normA) * Math.Sqrt(normB))), 0)
        End Function

        ''' <summary>
        ''' Splits text into tokens, filters stopwords, and generates bigrams for richer context.
        ''' </summary>
        ''' <param name="text">Input text to tokenize.</param>
        ''' <returns>Sequence of unigram tokens followed by bigram tokens (e.g., "word1_word2").</returns>
        ''' <remarks>
        ''' Tokenization pipeline:
        ''' 1. Extract Unicode letters and numbers via regex: [\p{L}\p{N}]+
        ''' 2. Convert to lowercase for case-insensitive matching.
        ''' 3. Filter minimal bilingual (DE/EN) stopword list.
        ''' 4. Emit unigrams, then emit bigrams (consecutive token pairs with underscore separator).
        ''' Bigrams improve phrase matching (e.g., "machine_learning" vs just "machine" + "learning").
        ''' Stopword list should be expanded for production use; current set is illustrative.
        ''' </remarks>
        Private Iterator Function SimpleTokenizer(text As String) As IEnumerable(Of String)
            ' Basic tokenization: words and numbers
            Dim tokens = Regex.Matches(text.ToLowerInvariant(), "[\p{L}\p{N}]+") _
                      .Cast(Of Match)() _
                      .Select(Function(m) m.Value) _
                      .ToList()

            ' Stopword list (bilingual German/English example)
            Dim stopwords As New HashSet(Of String) From {
                    "und", "oder", "der", "die", "das", "ist", "zu", "in", "im",
                    "on", "the", "and", "a", "an", "of", "for"
                }

            ' Filter stopwords
            Dim filtered = tokens.Where(Function(t) Not stopwords.Contains(t)).ToList()

            ' Yield individual tokens
            For Each token In filtered
                Yield token
            Next

            ' Yield bigrams
            For i As Integer = 0 To filtered.Count - 2
                Yield filtered(i) & "_" & filtered(i + 1)
            Next
        End Function

    End Class

End Namespace