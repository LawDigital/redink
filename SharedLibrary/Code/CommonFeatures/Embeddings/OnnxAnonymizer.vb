' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: OnnxAnonymizer.vb
' Purpose: ONNX-based Named Entity Recognition (NER) with selective anonymization
'          and de-anonymization. Detects and replaces person names (PER), organizations
'          (ORG), and other entity types with reversible placeholders for privacy.
'
' Architecture:
' -------------
' OnnxAnonymizer is a stateful NER module that wraps ONNX Runtime for token classification.
' Uses BIO tagging scheme (B-PER, I-PER, O) to detect entity spans and replaces whitelisted
' types with unique placeholders (<PER1>, <ORG2>) for round-trip anonymization.
'
' Algorithm:
'   1. Tokenization: SentencePiece subword tokenization with padding to maxLen
'   2. Inference: ONNX Runtime executes token classification model
'   3. Decoding: Argmax over logits → BIO labels → entity span merging
'   4. Filtering: Only whitelisted entity types replaced
'   5. Replacement: Generate unique placeholders, track bidirectional mapping
'   6. Reversal: String replacement using reverse map
'
' BIO Tagging:
'   - B-{TYPE}: Beginning of entity (e.g., B-PER for "John" in "John Smith")
'   - I-{TYPE}: Inside/continuation (e.g., I-PER for "Smith")
'   - O: Outside any entity (not anonymized)
'
' Model Requirements:
'   - ONNX model: Logits shape [batch_size, sequence_length, num_labels]
'   - Label map: Newline-separated BIO tags (O\nB-PER\nI-PER\nB-ORG\n...)
'   - Input names: "input_ids" (required); "attention_mask", "token_type_ids" (optional)
'   - Compatible with BERT-style token classification (e.g., dslim/bert-base-NER)
'
' Thread Safety:
'   NOT thread-safe. Module-level state shared.
'   Use external synchronization or refactor to class for concurrent use.
'
' Performance:
'   - Initialize: I/O-bound (load ONNX model, SentencePiece, label map); call once
'   - Anonymize: O(sequence_length) inference + O(entities) replacement
'   - Reverse: O(placeholders * text_length) string replacement
'
' Limitations:
'   - Whitelist-based: Only configured types anonymized
'   - No coreference resolution: Repeated mentions get separate placeholders
'   - Stateful: Single Anonymize/Reverse cycle per session
'   - Truncation: Entities beyond maxLen not detected
'
' Dependencies:
'   - Microsoft.ML.OnnxRuntime (NuGet)
'   - Microsoft.ML.OnnxRuntime.Tensors (NuGet)
'   - MlNetTokenizer (internal)
'
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports Microsoft.ML.OnnxRuntime
Imports Microsoft.ML.OnnxRuntime.Tensors

Namespace SharedLibrary

    ''' <summary>
    ''' ONNX-based Named Entity Recognition (NER) module with selective anonymization.
    ''' Detects entities using token classification and replaces whitelisted types with reversible placeholders.
    ''' </summary>
    ''' <remarks>
    ''' Wraps ONNX Runtime for BIO-tagged NER inference. Maintains stateful mappings between
    ''' original entities and placeholders for round-trip anonymization.
    ''' 
    ''' Thread Safety: NOT thread-safe. Use external synchronization or refactor to class.
    ''' Privacy Note: Mappings contain sensitive data; clear or encrypt after use.
    ''' </remarks>
    Public Module OnnxAnonymizer

        ''' <summary>Whitelist of entity types to anonymize (e.g., "PER", "ORG"). Configure via SetEntityTypesToAnonymize.</summary>
        Private _entityTypesToAnonymize As HashSet(Of String) =
        New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {"PER", "ORG"}

        ''' <summary>ONNX Runtime inference session for the NER model.</summary>
        Private _session As InferenceSession

        ''' <summary>Maximum token sequence length for model input (padding/truncation threshold).</summary>
        Private _maxLen As Integer

        ''' <summary>Maps model output label IDs to BIO tag strings (e.g., 0 → "O", 1 → "B-PER").</summary>
        Private _id2Label As Dictionary(Of Integer, String)

        ''' <summary>Bidirectional mapping: original entity text → placeholder (e.g., "John Smith" → "&lt;PER1&gt;").</summary>
        Private _mapping As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        ''' <summary>Reverse mapping: placeholder → original entity text (e.g., "&lt;PER1&gt;" → "John Smith").</summary>
        Private _reverseMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        ''' <summary>Per-entity-type counters for generating unique placeholder IDs (e.g., PER1, PER2).</summary>
        Private _counters As Dictionary(Of String, Integer)

        ''' <summary>Configures which entity types should be anonymized.</summary>
        ''' <param name="types">Collection of entity type labels (e.g., "PER", "ORG", "LOC").</param>
        ''' <remarks>
        ''' Types are case-insensitive and match base labels in BIO scheme (without "B-"/"I-" prefixes).
        ''' Default: {"PER", "ORG"}. Call before Anonymize() to take effect.
        ''' </remarks>
        Public Sub SetEntityTypesToAnonymize(types As IEnumerable(Of String))
            _entityTypesToAnonymize = New HashSet(Of String)(types, StringComparer.OrdinalIgnoreCase)
        End Sub

        ''' <summary>Gets the entity-to-placeholder mapping from the last Anonymize() call.</summary>
        ''' <value>Read-only dictionary mapping original entity text to placeholders.</value>
        ''' <remarks>Use to inspect or persist mappings for later de-anonymization. Cleared on each Anonymize() call.</remarks>
        Public ReadOnly Property Mapping As IReadOnlyDictionary(Of String, String)
            Get
                Return _mapping
            End Get
        End Property

        ''' <summary>Initializes the ONNX NER model, tokenizer, and label mapping.</summary>
        ''' <param name="modelPath">File path to ONNX model (.onnx file).</param>
        ''' <param name="spmModelPath">File path to SentencePiece tokenizer model (e.g., spm.model).</param>
        ''' <param name="labelMapPath">File path to label map (newline-separated BIO tags).</param>
        ''' <param name="maxSequenceLength">Maximum token sequence length. Default is 128.</param>
        ''' <remarks>
        ''' Must be called once before Anonymize(). Label map format: one BIO tag per line (0: O, 1: B-PER, 2: I-PER, ...).
        ''' Model inputs: "input_ids" (required); "attention_mask", "token_type_ids" (optional, auto-detected).
        ''' Not thread-safe with Anonymize(); call during single-threaded initialization.
        ''' </remarks>
        Public Sub Initialize(
        modelPath As String,
        spmModelPath As String,
        labelMapPath As String,
        Optional maxSequenceLength As Integer = 128
    )
            ' Load label map (ID → BIO tag string)
            Dim lines = File.ReadAllLines(labelMapPath)
            _id2Label = New Dictionary(Of Integer, String)(lines.Length)
            For i As Integer = 0 To lines.Length - 1
                _id2Label(i) = lines(i).Trim()
            Next

            ' Initialize ONNX Runtime session
            _session = New InferenceSession(modelPath)

            ' Initialize tokenizer
            MlNetTokenizer.LoadModel(spmModelPath, unkId:=3)
            _maxLen = maxSequenceLength
        End Sub

        ''' <summary>Anonymizes text by replacing whitelisted entity types with unique placeholders.</summary>
        ''' <param name="text">Input text to anonymize.</param>
        ''' <returns>Anonymized text with entities replaced by placeholders (e.g., "&lt;PER1&gt;", "&lt;ORG2&gt;").</returns>
        ''' <remarks>
        ''' Workflow: Tokenize → Inference → Decode BIO labels → Filter whitelist → Replace with placeholders.
        ''' Placeholder format: &lt;{TYPE}{COUNTER}&gt;. Populates Mapping property.
        ''' Performance: O(sequence_length) inference + O(entities) replacement. Entities beyond maxLen not detected.
        ''' </remarks>
        Public Function Anonymize(text As String) As String
            If _session Is Nothing Then
                Throw New InvalidOperationException("Please call OnnxAnonymizer.Initialize first.")
            End If

            ' Reset counters and mappings for new anonymization session
            _mapping.Clear()
            _reverseMap.Clear()
            _counters = _entityTypesToAnonymize.ToDictionary(Function(lbl) lbl, Function(lbl) 1)

            ' Tokenization and tensor construction
            Dim ids = MlNetTokenizer.TokenizeToIds(text, _maxLen)
            Dim seqLen = ids.Length
            Dim inputIds = New DenseTensor(Of Int64)(New Integer() {1, seqLen})
            Dim attention = New DenseTensor(Of Int64)(New Integer() {1, seqLen})
            For i As Integer = 0 To seqLen - 1
                inputIds(0, i) = CType(ids(i), Int64)
                attention(0, i) = If(ids(i) = MlNetTokenizer.PadId, 0L, 1L)
            Next

            ' Build ONNX input tensors based on model metadata
            Dim inputs As New List(Of NamedOnnxValue) From {
        NamedOnnxValue.CreateFromTensor("input_ids", inputIds)
    }
            Dim meta = _session.InputMetadata
            If meta.ContainsKey("attention_mask") Then
                inputs.Add(NamedOnnxValue.CreateFromTensor("attention_mask", attention))
            End If
            If meta.ContainsKey("token_type_ids") Then
                Dim tokenType = New DenseTensor(Of Int64)(New Integer() {1, seqLen})
                inputs.Add(NamedOnnxValue.CreateFromTensor("token_type_ids", tokenType))
            End If

            ' Run inference and decode BIO labels to entity spans
            Dim entities As List(Of (Start As Integer, [End] As Integer, Label As String, Text As String))
            Using results = _session.Run(inputs)
                entities = DecodePredictedLabels(text, results)
            End Using

            ' Replace only whitelisted entity types (descending order to preserve offsets)
            Dim sb As New System.Text.StringBuilder(text)
            Dim toReplace = entities _
        .Where(Function(entity) _entityTypesToAnonymize.Contains(entity.Label)) _
        .OrderByDescending(Function(entity) entity.Start)

            For Each match In toReplace
                ' Generate placeholder if not already mapped
                If Not _mapping.ContainsKey(match.Text) Then
                    Dim cnt = _counters(match.Label)
                    Dim ph = $"<{match.Label}{cnt}>"
                    _mapping(match.Text) = ph
                    _reverseMap(ph) = match.Text
                    _counters(match.Label) = cnt + 1
                End If

                ' Replace entity span with placeholder using character offsets
                sb.Remove(match.Start, match.End - match.Start) _
          .Insert(match.Start, _mapping(match.Text))
            Next

            Return sb.ToString()
        End Function

        ''' <summary>Reverses anonymization by replacing placeholders with original entity text.</summary>
        ''' <param name="anonymized">Anonymized text containing placeholders.</param>
        ''' <returns>De-anonymized text with original entities restored.</returns>
        ''' <remarks>
        ''' Uses reverse map from last Anonymize() call. Simple string replacement per placeholder.
        ''' Performance: O(placeholders * text_length). Only works with current session's mappings.
        ''' </remarks>
        Public Function Reverse(anonymized As String) As String
            Dim s = anonymized
            For Each kv In _reverseMap
                s = s.Replace(kv.Key, kv.Value)
            Next
            Return s
        End Function

        ''' <summary>Decodes ONNX model logits into entity spans using BIO tagging scheme.</summary>
        ''' <param name="originalText">Original input text for offset calculation.</param>
        ''' <param name="results">ONNX inference output containing logits tensor.</param>
        ''' <returns>List of detected entities with character-level offsets, labels, and text.</returns>
        ''' <remarks>
        ''' Algorithm: Locate logits → Argmax per token → Merge using BIO scheme (B-{TYPE} starts, I-{TYPE} continues, O closes).
        ''' Uses MlNetTokenizer.TokenizeWithOffsets for character alignment. Entities beyond maxLen not detected.
        ''' </remarks>
        Private Function DecodePredictedLabels(
        originalText As String,
        results As IDisposableReadOnlyCollection(Of DisposableNamedOnnxValue)
    ) As List(Of (Start As Integer, [End] As Integer, Label As String, Text As String))

            ' Find logits tensor (output node containing "logit" in name)
            Dim logitsNode = results _
            .FirstOrDefault(Function(x) x.Name.ToLower().Contains("logit"))
            If logitsNode Is Nothing Then
                Throw New InvalidOperationException("Logits tensor not found.")
            End If

            Dim logits = logitsNode.AsTensor(Of Single)()
            Dim dims = logits.Dimensions.ToArray()
            Dim seqLen = dims(1)
            Dim numLabels = dims(2)

            ' Get token offsets for character-level span extraction
            Dim offsets = MlNetTokenizer.TokenizeWithOffsets(originalText)

            Dim list = New List(Of (Integer, Integer, String, String))
            Dim curLabel As String = Nothing
            Dim spanStart As Integer = 0, spanEnd As Integer = 0

            ' BIO span merging loop
            For i As Integer = 0 To Math.Min(offsets.Count - 1, seqLen - 1)
                ' Argmax over label dimension to get predicted BIO tag
                Dim bestIdx = 0
                Dim bestVal = Single.MinValue
                For j As Integer = 0 To numLabels - 1
                    Dim v = logits(0, i, j)
                    If v > bestVal Then
                        bestVal = v
                        bestIdx = j
                    End If
                Next

                Dim fullLabel = If(_id2Label.ContainsKey(bestIdx), _id2Label(bestIdx), "O")
                Dim off = offsets(i)

                If fullLabel.StartsWith("B-") Then
                    ' Close previous span if open
                    If curLabel IsNot Nothing Then
                        list.Add((spanStart, spanEnd, curLabel,
                              originalText.Substring(spanStart, spanEnd - spanStart)))
                    End If
                    ' Start new span
                    curLabel = fullLabel.Substring(2)
                    spanStart = off.Start
                    spanEnd = off.End

                ElseIf fullLabel.StartsWith("I-") AndAlso curLabel = fullLabel.Substring(2) Then
                    ' Continue current span (only if type matches)
                    spanEnd = off.End

                Else
                    ' Outside or type mismatch → close current span
                    If curLabel IsNot Nothing Then
                        list.Add((spanStart, spanEnd, curLabel,
                              originalText.Substring(spanStart, spanEnd - spanStart)))
                        curLabel = Nothing
                    End If
                End If
            Next

            ' Add final open span if exists
            If curLabel IsNot Nothing Then
                list.Add((spanStart, spanEnd, curLabel,
                      originalText.Substring(spanStart, spanEnd - spanStart)))
            End If

            Return list
        End Function

        ''' <summary>Releases ONNX Runtime resources.</summary>
        ''' <remarks>
        ''' Frees native ONNX model resources. Initialize() must be called again after disposal.
        ''' Does not clear mappings; manually clear if they contain sensitive data.
        ''' </remarks>
        Public Sub Dispose()
            If _session IsNot Nothing Then
                _session.Dispose()
                _session = Nothing
            End If
        End Sub

    End Module
End Namespace