' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: MlNetTokenizer.vb
' Purpose: SentencePiece-based subword tokenization using Microsoft.ML.Tokenizers
'          for embedding model input preparation. Converts text to token IDs with
'          padding/truncation and tracks token offsets for alignment with source text.
'
' Architecture:
'   - MlNetTokenizer (Module): Stateful wrapper around ML.NET LlamaTokenizer.
'     Maintains a single shared tokenizer instance (_tokenizer) loaded from a
'     SentencePiece model file. NOT thread-safe; callers must synchronize access.
'   
'   - TokenOffset (Class): Data model representing a token with its character-level
'     start/end positions in the original text for alignment purposes.
'
'   - Key Methods:
'     * LoadModel: One-time initialization; reads SentencePiece model from disk
'     * TokenizeToIds: Converts text to padded/truncated integer array for models
'     * Tokenize/TokenizeWithOffsets: PLACEHOLDER methods using whitespace splitting
'       (TODO: replace with proper SentencePiece subword segmentation)
'
' Dependencies:
'   - Microsoft.ML.Tokenizers.LlamaTokenizer (NuGet: Microsoft.ML.Tokenizers)
'   - System.Text.RegularExpressions (fallback tokenization)
'   - System.IO (model file reading)
'
' Thread Safety:
'   NOT thread-safe. Module-level _tokenizer is shared state.
'   Use ThreadStatic or AsyncLocal for multi-threaded scenarios.
'
' Model Requirements:
'   - SentencePiece model file (.model format, e.g., spm.model)
'   - Default unkId=3 (LLaMA standard); verify against your model
'   - PadId fixed at 0; ensure compatibility with embedding model
'
' Limitations:
'   - Tokenize/TokenizeWithOffsets use naive whitespace splitting instead of
'     SentencePiece subword tokenization (placeholder implementation)
'   - No vocabulary size validation
'   - No batch tokenization support
'
' Usage:
'   MlNetTokenizer.LoadModel("path\to\spm.model", unkId:=3)
'   Dim tokenIds As Integer() = MlNetTokenizer.TokenizeToIds("Text", maxLen:=512)
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports System.Text.RegularExpressions
Imports Microsoft.ML.Tokenizers

Namespace SharedLibrary

    ''' <summary>
    ''' Module providing SentencePiece tokenization capabilities via Microsoft.ML.Tokenizers.
    ''' Converts text into token IDs for embedding models and tracks character offsets.
    ''' </summary>
    ''' <remarks>
    ''' This module wraps the ML.NET LlamaTokenizer to provide a simplified interface for
    ''' tokenizing text for embedding generation. It maintains stateful tokenizer instances
    ''' and is NOT thread-safe. Call LoadModel once during initialization before using
    ''' tokenization methods.
    ''' 
    ''' IMPORTANT: Tokenize() and TokenizeWithOffsets() currently use whitespace splitting
    ''' as a placeholder. For production use, implement proper SentencePiece subword segmentation
    ''' using the _tokenizer.Encode() method and its offset tracking capabilities.
    ''' </remarks>
    Public Module MlNetTokenizer

        ''' <summary>
        ''' The loaded SentencePiece tokenizer instance. Null until LoadModel is called.
        ''' </summary>
        Private _tokenizer As LlamaTokenizer

        ''' <summary>
        ''' Token ID used for padding sequences to fixed length. Fixed at 0.
        ''' </summary>
        Private _padId As Integer = 0

        ''' <summary>
        ''' Token ID representing unknown/out-of-vocabulary tokens.
        ''' Set via LoadModel; typically 3 for LLaMA models.
        ''' </summary>
        Private _unkId As Integer

        ''' <summary>
        ''' Loads a SentencePiece model from disk and initializes the tokenizer.
        ''' </summary>
        ''' <param name="spmModelPath">File path to the SentencePiece model file (e.g., spm.model).</param>
        ''' <param name="unkId">Token ID for unknown/OOV tokens. Default is 3 (LLaMA standard).</param>
        ''' <exception cref="FileNotFoundException">Thrown if spmModelPath does not exist.</exception>
        ''' <exception cref="IOException">Thrown if the model file cannot be read.</exception>
        ''' <remarks>
        ''' This method must be called once before any tokenization operations.
        ''' The model file is typically distributed with the embedding model and contains
        ''' the vocabulary and subword segmentation rules.
        ''' 
        ''' Thread Safety: Not safe to call concurrently with tokenization methods.
        ''' Call during single-threaded initialization phase.
        ''' </remarks>
        Public Sub LoadModel(spmModelPath As String, Optional unkId As Integer = 3)
            _unkId = unkId
            Using fs As FileStream = File.OpenRead(spmModelPath)
                _tokenizer = LlamaTokenizer.Create(fs)
            End Using
        End Sub

        ''' <summary>
        ''' Gets the token ID used for padding sequences.
        ''' </summary>
        ''' <value>The padding token ID (always 0 in current implementation).</value>
        ''' <remarks>
        ''' Used to construct attention masks when preparing batched inputs for models.
        ''' Most transformer-based embedding models expect padding tokens to be masked out.
        ''' </remarks>
        Public ReadOnly Property PadId As Integer
            Get
                Return _padId
            End Get
        End Property

        ''' <summary>
        ''' Tokenizes text into token IDs, padding or truncating to the specified maximum length.
        ''' </summary>
        ''' <param name="text">The input text to tokenize.</param>
        ''' <param name="maxLen">Maximum sequence length. Shorter sequences are padded; longer are truncated.</param>
        ''' <returns>Array of token IDs with exactly maxLen elements.</returns>
        ''' <exception cref="InvalidOperationException">Thrown if LoadModel has not been called.</exception>
        ''' <remarks>
        ''' This method:
        ''' 1. Encodes text using SentencePiece algorithm (no BOS/EOS tokens added).
        ''' 2. Copies token IDs to avoid ReadOnlySpan indexing issues.
        ''' 3. Pads with PadId if result is shorter than maxLen.
        ''' 4. Truncates if result is longer than maxLen.
        ''' 
        ''' The resulting array is suitable for direct use as model input (e.g., ONNX Runtime input tensors).
        ''' Caller should construct corresponding attention masks (1 for real tokens, 0 for padding).
        ''' </remarks>
        Public Function TokenizeToIds(text As String, maxLen As Integer) As Integer()
            If _tokenizer Is Nothing Then
                Throw New System.Exception("Tokenizer not initialized. Call MlNetTokenizer.LoadModel first.")
            End If

            ' Copy result to array to avoid ReadOnlySpan indexing issues
            Dim rawIdsArr As Integer() = _tokenizer _
            .EncodeToIds(text, addBeginningOfSentence:=False, addEndOfSentence:=False) _
            .ToArray()

            Dim ids(maxLen - 1) As Integer
            For i As Integer = 0 To maxLen - 1
                ids(i) = If(i < rawIdsArr.Length, rawIdsArr(i), _padId)
            Next

            Return ids
        End Function

        ''' <summary>
        ''' Splits text into tokens using whitespace boundaries.
        ''' </summary>
        ''' <param name="text">The input text to tokenize.</param>
        ''' <returns>List of token strings.</returns>
        ''' <exception cref="InvalidOperationException">Thrown if LoadModel has not been called.</exception>
        ''' <remarks>
        ''' PLACEHOLDER IMPLEMENTATION: This method currently uses simple whitespace splitting
        ''' via regex instead of true SentencePiece subword tokenization.
        ''' 
        ''' TODO: Replace with proper subword tokenization using _tokenizer.Encode() to get
        ''' actual subword segments (e.g., "tokenization" -> ["token", "ization"]).
        ''' 
        ''' Current behavior splits on any whitespace (\s+), which does not match the
        ''' granularity of SentencePiece subword units returned by TokenizeToIds.
        ''' </remarks>
        Public Function Tokenize(text As String) As List(Of String)
            If _tokenizer Is Nothing Then
                Throw New System.Exception("Tokenizer not initialized.")
            End If
            Return Regex.Split(text, "\s+").ToList()
        End Function

        ''' <summary>
        ''' Splits text into tokens with character-level offset tracking.
        ''' </summary>
        ''' <param name="text">The input text to tokenize.</param>
        ''' <returns>List of TokenOffset objects containing token text and start/end character positions.</returns>
        ''' <remarks>
        ''' PLACEHOLDER IMPLEMENTATION: This method currently uses simple whitespace splitting
        ''' with regex Match positions. Does not reflect actual SentencePiece subword boundaries.
        ''' 
        ''' TODO: Use _tokenizer.Encode() with offset tracking to align subword tokens
        ''' with their original character positions in the input text. This is critical
        ''' for highlighting or extracting matched spans in the original document.
        ''' 
        ''' Current behavior: Finds non-whitespace sequences (\S+) and records their positions.
        ''' Offsets are 0-based character indices into the original text string.
        ''' </remarks>
        Public Function TokenizeWithOffsets(text As String) As List(Of TokenOffset)
            Dim list As New List(Of TokenOffset)
            For Each m As Match In Regex.Matches(text, "\S+")
                list.Add(New TokenOffset With {
                .Text = m.Value,
                .Start = m.Index,
                .End = m.Index + m.Length
            })
            Next
            Return list
        End Function

        ''' <summary>
        ''' Represents a single token with its character-level position in the source text.
        ''' </summary>
        ''' <remarks>
        ''' Used to align tokenized text back to the original document for highlighting,
        ''' extraction, or downstream NLP processing that requires character-level spans.
        ''' Start and End are 0-based character indices (End is exclusive).
        ''' </remarks>
        Public Class TokenOffset
            ''' <summary>
            ''' The token text (substring from original input).
            ''' </summary>
            Public Property Text As String

            ''' <summary>
            ''' 0-based character index where this token starts in the original text.
            ''' </summary>
            Public Property Start As Integer

            ''' <summary>
            ''' 0-based character index where this token ends in the original text (exclusive).
            ''' </summary>
            Public Property [End] As Integer
        End Class

    End Module

End Namespace