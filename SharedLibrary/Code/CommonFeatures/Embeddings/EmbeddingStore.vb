' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: EmbeddingStore.vb
' Purpose: ONNX-based neural embedding store for semantic similarity search.
'          Loads Sentence-Transformer models (e.g., all-MiniLM-L6-v2) to compute
'          dense vector embeddings and perform cosine similarity ranking.
'
' Architecture:
'   - ONNX inference session with pre-trained Sentence-Transformer model
'   - WordPieceTokenizer for text preprocessing (uses vocab.txt)
'   - In-memory document store mapping docId -> text chunks with embeddings
'   - Cosine similarity search with configurable scope (all docs vs. current doc)
'
' Model Requirements:
'   - Input tensors (Int64, shape [1, seqLen]): input_ids, attention_mask, token_type_ids
'   - Output tensor: Float32 embedding vector (dimension depends on model)
'   - Recommended: all-MiniLM-L6-v2-onnx (384-dim, 256 max tokens, ~80 MB)
'   - Alternative models supported (see model list below)
'
' Key Limitations:
'   - NOT thread-safe (shared session and store state)
'   - Requires ONNX model file and matching vocab.txt at runtime
'   - No persistence; index is RAM-only
'   - Maximum sequence length is model-dependent (typically 128-384 tokens)
'
' Dependencies: Microsoft.ML.OnnxRuntime, Microsoft.ML.Tokenizers, TextChunk, SearchResult
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.ML.OnnxRuntime
Imports Microsoft.ML.OnnxRuntime.Tensors
Imports Microsoft.ML.Tokenizers

Namespace SharedLibrary

    ''' <summary>
    ''' ONNX-based embedding store for semantic document search using Sentence-Transformer models.
    ''' </summary>
    ''' <remarks>
    ''' <para><strong>Supported ONNX Models (no code changes required):</strong></para>
    ''' <list type="number">
    ''' <item>
    ''' <term>all-MiniLM-L6-v2-onnx</term>
    ''' <description>
    ''' URL: https://huggingface.co/onnx-models/all-MiniLM-L6-v2-onnx
    ''' <br/>Dimension: 384, Max Seq Length: 256, Size: ~80 MB
    ''' </description>
    ''' </item>
    ''' <item>
    ''' <term>all-mpnet-base-v2-onnx</term>
    ''' <description>
    ''' URL: https://huggingface.co/onnx-models/all-mpnet-base-v2-onnx
    ''' <br/>Dimension: 768, Max Seq Length: 384
    ''' </description>
    ''' </item>
    ''' <item>
    ''' <term>all-MiniLM-L12-v2-onnx</term>
    ''' <description>
    ''' URL: https://huggingface.co/onnx-models/all-MiniLM-L12-v2-onnx
    ''' <br/>Dimension: 384, Max Seq Length: 128
    ''' </description>
    ''' </item>
    ''' <item>
    ''' <term>all-MiniLM-L6-v2-fine-tuned-epochs-8-onnx</term>
    ''' <description>
    ''' URL: https://huggingface.co/onnx-models/all-MiniLM-L6-v2-fine-tuned-epochs-8-onnx
    ''' <br/>Dimension: 384, Max Seq Length: 256
    ''' </description>
    ''' </item>
    ''' <item>
    ''' <term>LightEmbed/sbert-all-MiniLM-L6-v2-onnx</term>
    ''' <description>
    ''' URL: https://huggingface.co/LightEmbed/sbert-all-MiniLM-L6-v2-onnx
    ''' <br/>Dimension: 384, Max Seq Length: 256
    ''' </description>
    ''' </item>
    ''' </list>
    ''' <para><strong>Important:</strong> Ensure input tensor names match exactly: 
    ''' "input_ids", "attention_mask", "token_type_ids" (all Int64).</para>
    ''' <para>ONNX Opset: ≥11, compatible with Microsoft.ML.OnnxRuntime v1.15.0+</para>
    ''' </remarks>
    Public Class EmbeddingStore

        ''' <summary>
        ''' Maps document IDs to their chunked text segments with computed embeddings.
        ''' </summary>
        Private ReadOnly store As Dictionary(Of String, List(Of TextChunk))

        ''' <summary>
        ''' ONNX inference session for the loaded Sentence-Transformer model.
        ''' </summary>
        Private ReadOnly session As InferenceSession

        ''' <summary>
        ''' WordPiece tokenizer using the model's vocabulary file.
        ''' </summary>
        Private ReadOnly tokenizer As WordPieceTokenizer

        ''' <summary>
        ''' Initializes a new instance with default model paths.
        ''' </summary>
        ''' <remarks>
        ''' Looks for "model.onnx" and "vocab.txt" in the application directory.
        ''' </remarks>
        Public Sub New()
            Me.New(
                modelPath:="model.onnx",
                vocabPath:="vocab.txt"
            )
        End Sub

        ''' <summary>
        ''' Initializes a new instance with custom model and vocabulary paths.
        ''' </summary>
        ''' <param name="modelPath">Path to the ONNX model file.</param>
        ''' <param name="vocabPath">Path to the vocab.txt file (WordPiece vocabulary).</param>
        ''' <remarks>
        ''' If files are missing or initialization fails, store is set to Nothing
        ''' and an error message is displayed. Check store for Nothing before use.
        ''' </remarks>
        Public Sub New(modelPath As String, vocabPath As String)
            ' Initialize store before any potential early return
            Me.store = New Dictionary(Of String, List(Of TextChunk))()

            ' Load ONNX model
            If Not File.Exists(modelPath) Then
                MessageBox.Show($"Error in EmbeddingStore: Embedding model not found: {modelPath}")
                Me.store = Nothing
                Return
            End If
            Me.session = New InferenceSession(modelPath)

            ' Load vocabulary and initialize tokenizer
            If Not File.Exists(vocabPath) Then
                MessageBox.Show($"Error in EmbeddingStore: Embedding vocabulary not found: {vocabPath}")
                Me.store = Nothing
                Return
            End If

            Dim options As New WordPieceOptions() With {
                .Normalizer = New LowerCaseNormalizer()
            }
            Me.tokenizer = WordPieceTokenizer.Create(vocabPath, options)

            ' Additional safeguards: ensure critical components initialized
            If Me.tokenizer Is Nothing Then
                MessageBox.Show("Error in EmbeddingStore: Failed to initialize tokenizer")
                Me.store = Nothing
                Return
            End If
            If Me.session Is Nothing Then
                MessageBox.Show("Error in EmbeddingStore: Failed to initialize ONNX session")
                Me.store = Nothing
                Return
            End If
        End Sub

        ''' <summary>
        ''' Computes the embedding vector for the given text using the ONNX model.
        ''' </summary>
        ''' <param name="text">Text to embed.</param>
        ''' <returns>Float32 embedding vector (dimension depends on model, typically 384 or 768).</returns>
        ''' <remarks>
        ''' Tokenizes text using WordPiece, truncates to maxLen (256), and runs ONNX inference.
        ''' Returns dense vector suitable for cosine similarity comparison.
        ''' </remarks>
        Private Function GetEmbedding(text As String) As Single()
            ' Safeguard: ensure components are initialized
            If Me.tokenizer Is Nothing Then
                Debug.WriteLine("Tokenizer not initialized")
            End If
            If Me.session Is Nothing Then
                Debug.WriteLine("ONNX session not initialized")
            End If

            ' Tokenize text
            Const maxLen As Integer = 256
            Dim normalized As String = Nothing
            Dim charsUsed As Integer = 0
            Dim ids As IReadOnlyList(Of Integer) =
                Me.tokenizer.EncodeToIds(
                    text,
                    maxLen,
                    normalized,
                    charsUsed,
                    considerPreTokenization:=True,
                    considerNormalization:=True)

            If ids Is Nothing OrElse ids.Count = 0 Then
                Debug.WriteLine("No token IDs returned")
            End If

            ' Build tensors: [1, seqLen]
            Dim seqLen = ids.Count
            Dim inputIds = New DenseTensor(Of Int64)(New Integer() {1, seqLen})
            Dim attentionMask = New DenseTensor(Of Int64)(New Integer() {1, seqLen})
            Dim tokenTypeIds = New DenseTensor(Of Int64)(New Integer() {1, seqLen})

            For i As Integer = 0 To seqLen - 1
                inputIds(0, i) = ids(i)
                attentionMask(0, i) = 1L
                tokenTypeIds(0, i) = 0L
            Next

            ' Create NamedOnnxValue inputs (explicit Int64 type)
            Dim inputs = New List(Of NamedOnnxValue) From {
                NamedOnnxValue.CreateFromTensor(Of Int64)("input_ids", inputIds),
                NamedOnnxValue.CreateFromTensor(Of Int64)("attention_mask", attentionMask),
                NamedOnnxValue.CreateFromTensor(Of Int64)("token_type_ids", tokenTypeIds)
            }

            ' Run inference
            Using results = Me.session.Run(inputs)
                If results Is Nothing OrElse results.Count = 0 Then
                    Debug.WriteLine("ONNX runtime returned no results")
                End If

                Dim outTensor = results.First().AsTensor(Of Single)()
                If outTensor Is Nothing Then
                    Debug.WriteLine("Result could not be read as Tensor(Of Single)")
                End If

                Return outTensor.ToArray()
            End Using
        End Function

        ''' <summary>
        ''' Indexes a document by computing embeddings for all chunks and storing them.
        ''' </summary>
        ''' <param name="docId">Unique identifier for the document.</param>
        ''' <param name="chunks">Pre-split text chunks with StartOffset/EndOffset metadata.</param>
        ''' <remarks>
        ''' Each chunk's Vector property is populated with its embedding.
        ''' Existing document with same docId is replaced.
        ''' </remarks>
        Public Sub IndexDocument(docId As String, chunks As List(Of TextChunk))
            For Each chunk In chunks
                chunk.Vector = GetEmbedding(chunk.Text)
            Next
            store(docId) = chunks
        End Sub

        ''' <summary>
        ''' Searches indexed chunks for semantic similarity to the query text.
        ''' </summary>
        ''' <param name="query">The search query text to vectorize and compare.</param>
        ''' <param name="allDocs">If True, search across all documents; if False, restrict to currentDocId.</param>
        ''' <param name="findAll">If True, return all results; if False, return only top result.</param>
        ''' <param name="currentDocId">Document ID to restrict search to when allDocs=False.</param>
        ''' <param name="currentPosition">Ignore chunks before this offset when searching within currentDocId and findAll=False.</param>
        ''' <returns>List of SearchResult objects sorted by descending score, then by docId and offset.</returns>
        ''' <remarks>
        ''' Results are sorted by: (1) score descending, (2) docId ascending, (3) startOffset ascending.
        ''' Only chunks with score > 0 are included.
        ''' </remarks>
        Public Function Search(query As String,
                               allDocs As Boolean,
                               findAll As Boolean,
                               currentDocId As String,
                               currentPosition As Integer) As List(Of SearchResult)

            Dim qVec = GetEmbedding(query)
            Dim results As New List(Of SearchResult)()

            ' Iterate documents alphabetically
            For Each docId In store.Keys.OrderBy(Function(k) k)
                If Not allDocs AndAlso docId <> currentDocId Then Continue For

                ' Iterate chunks by ascending StartOffset
                For Each chunk In store(docId).OrderBy(Function(c) c.StartOffset)
                    If Not findAll AndAlso docId = currentDocId AndAlso chunk.StartOffset < currentPosition Then
                        Continue For
                    End If

                    Dim score = CosineSimilarity(qVec, chunk.Vector)
                    If score > 0 Then
                        results.Add(New SearchResult With {
                            .DocId = docId,
                            .Text = chunk.Text,
                            .StartOffset = chunk.StartOffset,
                            .EndOffset = chunk.EndOffset,
                            .Score = score
                        })
                    End If
                Next
            Next

            ' Final sort: score descending, docId ascending, startOffset ascending
            results.Sort(Function(a, b)
                             Dim cmp = b.Score.CompareTo(a.Score)
                             If cmp <> 0 Then Return cmp
                             cmp = a.DocId.CompareTo(b.DocId)
                             If cmp <> 0 Then Return cmp
                             Return a.StartOffset.CompareTo(b.StartOffset)
                         End Function)

            ' Return only top result if requested
            If Not findAll AndAlso results.Count > 0 Then
                Return New List(Of SearchResult) From {results(0)}
            End If

            Return results
        End Function

        ''' <summary>
        ''' Computes cosine similarity between two vectors.
        ''' </summary>
        ''' <param name="vec1">First embedding vector.</param>
        ''' <param name="vec2">Second embedding vector.</param>
        ''' <returns>Cosine similarity in range [0.0, 1.0]; 0.0 if either vector has zero norm.</returns>
        ''' <remarks>
        ''' Formula: cos(θ) = (A · B) / (||A|| * ||B||).
        ''' Handles mismatched vector lengths by using the minimum dimension (defensive coding).
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
    End Class

End Namespace