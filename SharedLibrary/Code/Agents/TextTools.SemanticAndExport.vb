' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: TextTools.SemanticAndExport.vb
' Purpose: Host-agnostic text_* and semantic-index tools for the agent layer.
'
' Tools:
'   - text_export_to_text
'   - semantic_index_create_from_file
'   - semantic_index_create_from_text
'   - semantic_index_validate
'   - semantic_index_search
'   - semantic_index_search_continuation
'   - semantic_index_load_entries
'   - semantic_index_verify_answer
'   - semantic_index_retrieve_after_verification
'   - semantic_index_reset_conversation
'   - semantic_index_invalidate_cache
'
' Notes:
'   - All semantic-search logic uses SharedMethods.SemanticSearch.* shared helpers.
'   - No WinForms interactive APIs are used here.
'   - LLM-dependent operations require an ISharedContext supplied by the host.
' =============================================================================

Option Strict On
Option Explicit On
Option Infer On

Imports System.Collections.Concurrent
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace Agents

    Partial Public NotInheritable Class TextTools

        Public Const ToolExportToText As String = "text_export_to_text"

        Public Const ToolSemanticIndexCreateFromFile As String = "semantic_index_create_from_file"
        Public Const ToolSemanticIndexCreateFromText As String = "semantic_index_create_from_text"
        Public Const ToolSemanticIndexValidate As String = "semantic_index_validate"
        Public Const ToolSemanticIndexSearch As String = "semantic_index_search"
        Public Const ToolSemanticIndexSearchContinuation As String = "semantic_index_search_continuation"
        Public Const ToolSemanticIndexLoadEntries As String = "semantic_index_load_entries"
        Public Const ToolSemanticIndexVerifyAnswer As String = "semantic_index_verify_answer"
        Public Const ToolSemanticIndexRetrieveAfterVerification As String = "semantic_index_retrieve_after_verification"
        Public Const ToolSemanticIndexResetConversation As String = "semantic_index_reset_conversation"
        Public Const ToolSemanticIndexInvalidateCache As String = "semantic_index_invalidate_cache"

        Private Const TextExportAdapterVersion As System.String = "text-export-v3"

        Private Const SemanticSearchTaskName As String = "SemanticSearch"
        Private Const SemanticIndexGenerationTaskName As String = "SemanticSearchIndex"
        Private Const DefaultTextExportDirectoryName As String = "extracted_text"

        Private Shared ReadOnly SemanticConversationStore As New ConcurrentDictionary(
            Of String,
            SemanticConversationStateItem)(StringComparer.OrdinalIgnoreCase)

        Private Shared ReadOnly SemanticRetrievalStore As New ConcurrentDictionary(
            Of String,
            SemanticRetrievalStateItem)(StringComparer.OrdinalIgnoreCase)

        Private Shared ReadOnly SemanticVerificationStore As New ConcurrentDictionary(
            Of String,
            SemanticVerificationStateItem)(StringComparer.OrdinalIgnoreCase)

        Private Shared ReadOnly SupportedTextExportExtensions As New HashSet(Of String)(
            StringComparer.OrdinalIgnoreCase) From {
                ".txt", ".rtf", ".doc", ".docx", ".docm", ".pdf",
                ".xlsx", ".xlsm", ".pptx", ".pptm",
                ".ini", ".csv", ".log", ".json", ".xml", ".html", ".htm", ".md", ".yaml", ".yml",
                ".vb", ".cs", ".js", ".ts", ".py", ".java", ".cpp", ".c", ".h", ".sql",
                ".eml", ".msg",
                ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff", ".tif", ".webp", ".svg",
                ".mp3", ".wav", ".ogg", ".flac", ".m4a", ".aac", ".wma", ".opus", ".webm",
                ".mp4", ".avi", ".mkv", ".mov", ".wmv"
            }

        Private Shared ReadOnly MetadataProfileMap As New Dictionary(
            Of String,
            SharedMethods.SemanticSearchMetadataProfile)(StringComparer.OrdinalIgnoreCase) From {
                {"generic", SharedMethods.SemanticSearchMetadataProfile.Generic},
                {"technical_manual", SharedMethods.SemanticSearchMetadataProfile.TechnicalManual},
                {"legal", SharedMethods.SemanticSearchMetadataProfile.Legal},
                {"contract", SharedMethods.SemanticSearchMetadataProfile.Contract},
                {"investigation", SharedMethods.SemanticSearchMetadataProfile.Investigation},
                {"compliance", SharedMethods.SemanticSearchMetadataProfile.Compliance},
                {"narrative", SharedMethods.SemanticSearchMetadataProfile.Narrative},
                {"corporate_transaction", SharedMethods.SemanticSearchMetadataProfile.CorporateTransaction},
                {"dispute", SharedMethods.SemanticSearchMetadataProfile.Dispute},
                {"regulatory", SharedMethods.SemanticSearchMetadataProfile.Regulatory},
                {"data_protection_and_privacy", SharedMethods.SemanticSearchMetadataProfile.DataProtectionAndPrivacy},
                {"corporate_governance", SharedMethods.SemanticSearchMetadataProfile.CorporateGovernance},
                {"employment_and_hr", SharedMethods.SemanticSearchMetadataProfile.EmploymentAndHR},
                {"finance_and_accounting", SharedMethods.SemanticSearchMetadataProfile.FinanceAndAccounting},
                {"tax", SharedMethods.SemanticSearchMetadataProfile.Tax},
                {"risk_management", SharedMethods.SemanticSearchMetadataProfile.RiskManagement},
                {"operations_and_projects", SharedMethods.SemanticSearchMetadataProfile.OperationsAndProjects},
                {"procurement_and_supply", SharedMethods.SemanticSearchMetadataProfile.ProcurementAndSupply},
                {"sales_and_commercial", SharedMethods.SemanticSearchMetadataProfile.SalesAndCommercial},
                {"insurance", SharedMethods.SemanticSearchMetadataProfile.Insurance},
                {"real_estate", SharedMethods.SemanticSearchMetadataProfile.RealEstate},
                {"intellectual_property", SharedMethods.SemanticSearchMetadataProfile.IntellectualProperty},
                {"business_records", SharedMethods.SemanticSearchMetadataProfile.BusinessRecords}
            }

        Private NotInheritable Class SemanticConversationStateItem
            Public Property Handle As String = ""
            Public Property Path As String = ""
            Public Property State As SharedMethods.SemanticSearchConversationState =
                New SharedMethods.SemanticSearchConversationState()
            Public Property Options As SharedMethods.SemanticSearchRetrievalOptions = Nothing
            Public Property UpdatedUtc As DateTime = DateTime.UtcNow
        End Class

        Private NotInheritable Class SemanticRetrievalStateItem
            Public Property Handle As String = ""
            Public Property Path As String = ""
            Public Property ConversationHandle As String = ""
            Public Property Retrieval As SharedMethods.SemanticSearchRetrievalResult = Nothing
            Public Property Options As SharedMethods.SemanticSearchRetrievalOptions = Nothing
            Public Property UpdatedUtc As DateTime = DateTime.UtcNow
        End Class

        Private NotInheritable Class SemanticVerificationStateItem
            Public Property Handle As String = ""
            Public Property Path As String = ""
            Public Property RetrievalHandle As String = ""
            Public Property Verification As SharedMethods.SemanticSearchResponseVerificationResult = Nothing
            Public Property UpdatedUtc As DateTime = DateTime.UtcNow
        End Class

        Private NotInheritable Class TextExtractionOutcome
            Public Property Success As Boolean
            Public Property Content As String = ""
            Public Property ErrorCode As String = ""
            Public Property Message As String = ""
            Public Property PageCount As System.Nullable(Of System.Int32) = Nothing
            Public Property OcrUsed As System.Nullable(Of System.Boolean) = Nothing
            Public Property OcrAttempted As System.Nullable(Of System.Boolean) = Nothing
            Public Property OcrSkipped As System.Nullable(Of System.Boolean) = Nothing
            Public Property OcrDurationMilliseconds As System.Nullable(Of System.Int64) = Nothing
            Public Property ExtractionComplete As System.Nullable(Of System.Boolean) = Nothing
            Public Property ExtractionCoverageBasis As System.String = "unverified"
            Public Property ExtractionWarnings As New System.Collections.Generic.List(Of System.String)()
            Public Property ContentFormat As System.String = "unknown"
            Public Property ProcessedRanges As New System.Collections.Generic.List(Of TextExtractionProcessedRange)()
        End Class

        Friend Shared Function IsExtendedTextTool(name As String) As Boolean
            If String.IsNullOrWhiteSpace(name) Then
                Return False
            End If

            Select Case name.Trim()
                Case ToolExportToText,
                     ToolSemanticIndexCreateFromFile,
                     ToolSemanticIndexCreateFromText,
                     ToolSemanticIndexValidate,
                     ToolSemanticIndexSearch,
                     ToolSemanticIndexSearchContinuation,
                     ToolSemanticIndexLoadEntries,
                     ToolSemanticIndexVerifyAnswer,
                     ToolSemanticIndexRetrieveAfterVerification,
                     ToolSemanticIndexResetConversation,
                     ToolSemanticIndexInvalidateCache
                    Return True

                Case Else
                    Return False
            End Select
        End Function

        Friend Shared Function BuildExtendedTools() As List(Of ModelConfig)
            Return New List(Of ModelConfig) From {
                BuildToolConfig(
                    ToolExportToText,
                    "Silently extracts readable text to UTF-8 .txt files. Items retain source_path/output_path/status and add char_count (UTF-16), byte_count, snapshot_sha256 (exact output bytes/BOM), source_sha256, resource_id/resource_reuse, secret-free adapter/config fingerprints and observed PDF page/OCR fields. Successful extraction uses an immutable captured source and a session-scoped single-flight resource: identical source bytes plus identical effective adapter/options reuse the same successful extraction instead of rerunning OCR. Failed/cancelled extraction is never cached. overwrite=false still checks an existing output before extraction; an output published by this session is provenance-verified against its resource, while an unrelated pre-existing .txt remains reuse_validation=not_verified. processed_ranges are host-known ranges only and do not claim paragraph-to-page mapping. Null metadata means unknown, not false or complete. output_path can feed create_word_document.markdown_path when explicitly interpreted as Markdown. Newly written exports are published atomically. Hashes prove byte identity, not OCR accuracy. For directories, structure is preserved.",
                    "{""type"":""object"",""properties"":{" &
                        """input_path"":{""type"":""string"",""description"":""Required file or directory path.""}," &
                        """output_directory"":{""type"":""string"",""description"":""Optional output directory. For directory input, relative paths are preserved under this root.""}," &
                        """recursive"":{""type"":""boolean"",""description"":""For directory input, include subdirectories. Default true.""}," &
                        """overwrite"":{""type"":""boolean"",""description"":""Overwrite existing .txt outputs only when true. Default false.""}," &
                        """ocr_pdf"":{""type"":""boolean"",""description"":""Enable silent OCR heuristics for PDFs when a suitable model is configured. Default false.""}}," &
                        """required"":[""input_path""]}",
                    923,
                    "Text (export to text)",
                    allowRepeatedIdenticalCalls:=True),
                BuildToolConfig(
                    ToolSemanticIndexCreateFromFile,
                    "Create a self-indexed semantic-search UTF-8 text file from an existing source text file.",
                    "{""type"":""object"",""properties"":{" &
                        """input_path"":{""type"":""string"",""description"":""Required source text file path.""}," &
                        """output_path"":{""type"":""string"",""description"":""Required destination indexed text file path.""}," &
                        """metadata_profile"":{""type"":""string"",""description"":""Profile key such as generic, technical_manual, contract, legal, compliance, business_records.""}," &
                        """overwrite"":{""type"":""boolean"",""description"":""Never overwrite unless true. Default false.""}," &
                        """target_bytes"":{""type"":""integer"",""description"":""Preferred segment size in UTF-8 bytes. Default 32768.""}," &
                        """minimum_bytes"":{""type"":""integer"",""description"":""Minimum segment size in UTF-8 bytes. Default 16384.""}," &
                        """maximum_bytes"":{""type"":""integer"",""description"":""Maximum segment size in UTF-8 bytes. Default 49152.""}}," &
                        """required"":[""input_path"",""output_path""]}",
                    924,
                    "Semantic index (create from file)"),
                BuildToolConfig(
                    ToolSemanticIndexCreateFromText,
                    "Create a self-indexed semantic-search UTF-8 text file from supplied in-memory text.",
                    "{""type"":""object"",""properties"":{" &
                        """text"":{""type"":""string"",""description"":""Required source text.""}," &
                        """output_path"":{""type"":""string"",""description"":""Required destination indexed text file path.""}," &
                        """metadata_profile"":{""type"":""string"",""description"":""Profile key such as generic, technical_manual, contract, legal, compliance, business_records.""}," &
                        """overwrite"":{""type"":""boolean"",""description"":""Never overwrite unless true. Default false.""}}," &
                        """required"":[""text"",""output_path""]}",
                    925,
                    "Semantic index (create from text)"),
                BuildToolConfig(
                    ToolSemanticIndexValidate,
                    "Validate whether a file is a readable semantic-search index and return basic counts.",
                    "{""type"":""object"",""properties"":{" &
                        """path"":{""type"":""string"",""description"":""Required indexed text file path.""}}," &
                        """required"":[""path""]}",
                    926,
                    "Semantic index (validate)"),
                BuildToolConfig(
                    ToolSemanticIndexSearch,
                    "Run an initial semantic search against an indexed text file and return grounded source excerpts plus internal handles for later continuation and verification.",
                    "{""type"":""object"",""properties"":{" &
                        """path"":{""type"":""string"",""description"":""Required indexed text file path.""}," &
                        """question"":{""type"":""string"",""description"":""Required current question.""}," &
                        """conversation"":{""type"":""string"",""description"":""Optional conversation context. Default empty.""}," &
                        """previous_entry_ids"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Optional previously used entry ids.""}," &
                        """minimum_selected_segments"":{""type"":""integer"",""description"":""Default 1.""}," &
                        """maximum_selected_segments"":{""type"":""integer"",""description"":""Default 8.""}," &
                        """maximum_total_segments"":{""type"":""integer"",""description"":""Default 24.""}," &
                        """enable_full_scan_fallback"":{""type"":""boolean"",""description"":""Default true.""}," &
                        """force_full_scan"":{""type"":""boolean"",""description"":""Default false.""}}," &
                        """required"":[""path"",""question""]}",
                    927,
                    "Semantic index (search)"),
                BuildToolConfig(
                    ToolSemanticIndexSearchContinuation,
                    "Continue a semantic-search conversation using a prior conversation_handle and return additional grounded source excerpts.",
                    "{""type"":""object"",""properties"":{" &
                        """path"":{""type"":""string"",""description"":""Required indexed text file path.""}," &
                        """question"":{""type"":""string"",""description"":""Required current follow-up question.""}," &
                        """conversation"":{""type"":""string"",""description"":""Optional conversation context. Default empty.""}," &
                        """conversation_handle"":{""type"":""string"",""description"":""Required handle returned by semantic_index_search.""}," &
                        """minimum_selected_segments"":{""type"":""integer"",""description"":""Default 1.""}," &
                        """maximum_selected_segments"":{""type"":""integer"",""description"":""Default 8.""}," &
                        """maximum_total_segments"":{""type"":""integer"",""description"":""Default 24.""}," &
                        """enable_full_scan_fallback"":{""type"":""boolean"",""description"":""Default true.""}," &
                        """force_full_scan"":{""type"":""boolean"",""description"":""Default false.""}}," &
                        """required"":[""path"",""question"",""conversation_handle""]}",
                    928,
                    "Semantic index (continuation)"),
                BuildToolConfig(
                    ToolSemanticIndexLoadEntries,
                    "Load exact indexed source ranges for trusted entry ids without running semantic selection.",
                    "{""type"":""object"",""properties"":{" &
                        """path"":{""type"":""string"",""description"":""Required indexed text file path.""}," &
                        """entry_ids"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Required known entry ids.""}," &
                        """maximum_total_segments"":{""type"":""integer"",""description"":""Default 24.""}," &
                        """context_bytes_before"":{""type"":""integer"",""description"":""Default 2048.""}," &
                        """context_bytes_after"":{""type"":""integer"",""description"":""Default 2048.""}}," &
                        """required"":[""path"",""entry_ids""]}",
                    929,
                    "Semantic index (load entries)"),
                BuildToolConfig(
                    ToolSemanticIndexVerifyAnswer,
                    "Verify whether a drafted answer is supported by the exact source excerpts returned by a previous retrieval_handle.",
                    "{""type"":""object"",""properties"":{" &
                        """path"":{""type"":""string"",""description"":""Required indexed text file path.""}," &
                        """question"":{""type"":""string"",""description"":""Required current question.""}," &
                        """conversation"":{""type"":""string"",""description"":""Optional conversation context. Default empty.""}," &
                        """retrieval_handle"":{""type"":""string"",""description"":""Required handle returned by semantic_index_search or related tools.""}," &
                        """answer"":{""type"":""string"",""description"":""Required drafted answer to verify.""}," &
                        """special_task_name"":{""type"":""string"",""description"":""Optional verification task name. Default SemanticSearch.""}," &
                        """maximum_llm_attempts"":{""type"":""integer"",""description"":""Default 2.""}," &
                        """maximum_conversation_characters"":{""type"":""integer"",""description"":""Default 12000.""}}," &
                        """required"":[""path"",""question"",""retrieval_handle"",""answer""]}",
                    930,
                    "Semantic index (verify answer)"),
                BuildToolConfig(
                    ToolSemanticIndexRetrieveAfterVerification,
                    "Retrieve more semantic-search sources after semantic_index_verify_answer indicates that more evidence is required.",
                    "{""type"":""object"",""properties"":{" &
                        """path"":{""type"":""string"",""description"":""Required indexed text file path.""}," &
                        """question"":{""type"":""string"",""description"":""Required current question.""}," &
                        """conversation"":{""type"":""string"",""description"":""Optional conversation context. Default empty.""}," &
                        """retrieval_handle"":{""type"":""string"",""description"":""Required prior retrieval handle.""}," &
                        """verification_handle"":{""type"":""string"",""description"":""Required verification handle returned by semantic_index_verify_answer.""}," &
                        """minimum_selected_segments"":{""type"":""integer"",""description"":""Default 1.""}," &
                        """maximum_selected_segments"":{""type"":""integer"",""description"":""Default 8.""}," &
                        """maximum_total_segments"":{""type"":""integer"",""description"":""Default 24.""}," &
                        """enable_full_scan_fallback"":{""type"":""boolean"",""description"":""Default true.""}," &
                        """force_full_scan"":{""type"":""boolean"",""description"":""Default false.""}}," &
                        """required"":[""path"",""question"",""retrieval_handle"",""verification_handle""]}",
                    931,
                    "Semantic index (retrieve after verification)"),
                BuildToolConfig(
                    ToolSemanticIndexResetConversation,
                    "Reset and remove a stored semantic-search conversation handle.",
                    "{""type"":""object"",""properties"":{" &
                        """conversation_handle"":{""type"":""string"",""description"":""Required conversation handle.""}}," &
                        """required"":[""conversation_handle""]}",
                    932,
                    "Semantic index (reset conversation)"),
                BuildToolConfig(
                    ToolSemanticIndexInvalidateCache,
                    "Invalidate one indexed-file cache entry or the full semantic-search cache.",
                    "{""type"":""object"",""properties"":{" &
                        """path"":{""type"":""string"",""description"":""Optional indexed text file path. Omit or pass empty to clear the full semantic cache.""}}}",
                    933,
                    "Semantic index (invalidate cache)")
            }
        End Function

        Private Shared Function BuildToolConfig(toolName As String,
                                                description As String,
                                                parametersJson As String,
                                                priority As Integer,
                                                modelDescription As String,
                                                Optional allowRepeatedIdenticalCalls As System.Boolean = False) As ModelConfig
            Dim def As String =
                "{""name"":""" & toolName & """," &
                """description"":""" & EscapeToolJson(description) & """," &
                """parameters"":" & parametersJson & "}"

            Return New ModelConfig() With {
                .ToolName = toolName,
                .ToolDefinition = def,
                .ToolInstructionsPrompt = toolName & ": " & description,
                .ModelDescription = modelDescription,
                .Tool = True,
                .ToolPriority = priority,
                .ToolErrorHandling = "skip",
                .AllowRepeatedIdenticalCalls = allowRepeatedIdenticalCalls
            }
        End Function

        Private Shared Function EscapeToolJson(value As String) As String
            Dim jsonValue As String = JsonConvert.SerializeObject(If(value, ""))
            If jsonValue.Length >= 2 AndAlso
               jsonValue(0) = """"c AndAlso
               jsonValue(jsonValue.Length - 1) = """"c Then
                Return jsonValue.Substring(1, jsonValue.Length - 2)
            End If

            Return jsonValue
        End Function

        Friend Shared Async Function ExecuteExtendedAsync(toolName As String,
                                                          arguments As IDictionary(Of String, Object),
                                                          context As ISharedContext,
                                                          cancellationToken As CancellationToken) As Task(Of String)
            Select Case toolName
                Case ToolExportToText
                    Return Await ExecuteExportToTextAsync(arguments, context, cancellationToken).ConfigureAwait(False)

                Case ToolSemanticIndexCreateFromFile
                    Return Await ExecuteSemanticIndexCreateFromFileAsync(arguments, context, cancellationToken).ConfigureAwait(False)

                Case ToolSemanticIndexCreateFromText
                    Return Await ExecuteSemanticIndexCreateFromTextAsync(arguments, context, cancellationToken).ConfigureAwait(False)

                Case ToolSemanticIndexValidate
                    Return Await ExecuteSemanticIndexValidateAsync(arguments, cancellationToken).ConfigureAwait(False)

                Case ToolSemanticIndexSearch
                    Return Await ExecuteSemanticIndexSearchAsync(arguments, context, cancellationToken).ConfigureAwait(False)

                Case ToolSemanticIndexSearchContinuation
                    Return Await ExecuteSemanticIndexSearchContinuationAsync(arguments, context, cancellationToken).ConfigureAwait(False)

                Case ToolSemanticIndexLoadEntries
                    Return Await ExecuteSemanticIndexLoadEntriesAsync(arguments, cancellationToken).ConfigureAwait(False)

                Case ToolSemanticIndexVerifyAnswer
                    Return Await ExecuteSemanticIndexVerifyAnswerAsync(arguments, context, cancellationToken).ConfigureAwait(False)

                Case ToolSemanticIndexRetrieveAfterVerification
                    Return Await ExecuteSemanticIndexRetrieveAfterVerificationAsync(arguments, context, cancellationToken).ConfigureAwait(False)

                Case ToolSemanticIndexResetConversation
                    Return ExecuteSemanticIndexResetConversation(arguments)

                Case ToolSemanticIndexInvalidateCache
                    Return ExecuteSemanticIndexInvalidateCache(arguments)

                Case Else
                    Return Nothing
            End Select
        End Function

        Private Shared Async Function ExecuteSemanticIndexCreateFromFileAsync(args As IDictionary(Of String, Object),
                                                                             context As ISharedContext,
                                                                             cancellationToken As CancellationToken) As Task(Of String)
            If context Is Nothing Then
                Return BuildError("missing_context", "semantic_index_create_from_file requires a shared LLM context.")
            End If

            Dim inputPath As String = PathPolicy.Resolve(GetStr(args, "input_path"), PathAccess.Read)
            Dim outputPath As String = PathPolicy.Resolve(GetStr(args, "output_path"), PathAccess.Write)

            If String.IsNullOrWhiteSpace(inputPath) OrElse Not File.Exists(inputPath) Then
                Return BuildError("not_found", "The input file was not found.", inputPath)
            End If

            If String.Equals(Path.GetFullPath(inputPath), Path.GetFullPath(outputPath), StringComparison.OrdinalIgnoreCase) Then
                Return BuildError("invalid_argument", "input_path and output_path must be different files.")
            End If

            Dim options As New SharedMethods.SemanticSearchIndexGeneratorOptions() With {
                .TargetBytes = GetInt(args, "target_bytes", SharedMethods.SemanticSearchDefaultTargetBytes),
                .MinimumBytes = GetInt(args, "minimum_bytes", SharedMethods.SemanticSearchDefaultMinimumBytes),
                .MaximumBytes = GetInt(args, "maximum_bytes", SharedMethods.SemanticSearchDefaultMaximumBytes),
                .SpecialTaskName = SemanticIndexGenerationTaskName,
                .MetadataProfile = ResolveMetadataProfile(GetStr(args, "metadata_profile")),
                .OverwriteOutput = GetBool(args, "overwrite", False)
            }

            Dim result As SharedMethods.SemanticSearchIndexGenerationResult =
                Await SharedMethods.CreateSemanticSearchIndexedTextFileAsync(
                    inputPath:=inputPath,
                    outputPath:=outputPath,
                    context:=context,
                    options:=options,
                    cancellationToken:=cancellationToken).ConfigureAwait(False)

            SharedMethods.InvalidateSemanticSearchIndexCache(outputPath)

            Return JsonConvert.SerializeObject(New With {
                Key .output_path = result.OutputPath,
                Key .content_byte_length = result.ContentByteLength,
                Key .document_count = result.DocumentCount,
                Key .segment_count = result.SegmentCount,
                Key .content_sha256 = result.ContentSha256
            })
        End Function

        Private Shared Async Function ExecuteSemanticIndexCreateFromTextAsync(args As IDictionary(Of String, Object),
                                                                             context As ISharedContext,
                                                                             cancellationToken As CancellationToken) As Task(Of String)
            If context Is Nothing Then
                Return BuildError("missing_context", "semantic_index_create_from_text requires a shared LLM context.")
            End If

            Dim text As String = GetStr(args, "text")
            Dim outputPath As String = PathPolicy.Resolve(GetStr(args, "output_path"), PathAccess.Write)

            If text Is Nothing Then
                Return BuildError("missing_text", "text is required.")
            End If

            Dim options As New SharedMethods.SemanticSearchIndexGeneratorOptions() With {
                .SpecialTaskName = SemanticIndexGenerationTaskName,
                .MetadataProfile = ResolveMetadataProfile(GetStr(args, "metadata_profile")),
                .OverwriteOutput = GetBool(args, "overwrite", False)
            }

            Dim result As SharedMethods.SemanticSearchIndexGenerationResult =
                Await SharedMethods.CreateSemanticSearchIndexFromTextAsync(
                    text:=text,
                    outputPath:=outputPath,
                    context:=context,
                    options:=options,
                    cancellationToken:=cancellationToken).ConfigureAwait(False)

            SharedMethods.InvalidateSemanticSearchIndexCache(outputPath)

            Return JsonConvert.SerializeObject(New With {
                Key .output_path = result.OutputPath,
                Key .content_byte_length = result.ContentByteLength,
                Key .document_count = result.DocumentCount,
                Key .segment_count = result.SegmentCount,
                Key .content_sha256 = result.ContentSha256
            })
        End Function

        Private Shared Async Function ExecuteSemanticIndexValidateAsync(args As IDictionary(Of String, Object),
                                                                       cancellationToken As CancellationToken) As Task(Of String)
            Dim path As String = PathPolicy.Resolve(GetStr(args, "path"), PathAccess.Read)
            Dim item As SharedMethods.SemanticSearchIndexCacheItem =
                Await SharedMethods.TryGetSemanticSearchIndexAsync(path, cancellationToken).ConfigureAwait(False)

            If item Is Nothing Then
                Return JsonConvert.SerializeObject(New With {
                    Key .is_valid_index = False,
                    Key .path = path,
                    Key .segment_count = 0,
                    Key .document_count = 0
                })
            End If

            Return JsonConvert.SerializeObject(New With {
                Key .is_valid_index = True,
                Key .path = path,
                Key .segment_count = item.OrderedEntries.Count,
                Key .document_count = item.IndexDocument.Documents.Count
            })
        End Function

        Private Shared Async Function ExecuteSemanticIndexSearchAsync(args As IDictionary(Of String, Object),
                                                                     context As ISharedContext,
                                                                     cancellationToken As CancellationToken) As Task(Of String)
            If context Is Nothing Then
                Return BuildError("missing_context", "semantic_index_search requires a shared LLM context.")
            End If

            Dim path As String = PathPolicy.Resolve(GetStr(args, "path"), PathAccess.Read)
            Dim question As String = GetStr(args, "question")
            Dim conversation As String = GetStr(args, "conversation")

            If String.IsNullOrWhiteSpace(question) Then
                Return BuildError("missing_question", "question is required.")
            End If

            Dim options As SharedMethods.SemanticSearchRetrievalOptions = BuildRetrievalOptions(args, Nothing)
            Dim state As New SharedMethods.SemanticSearchConversationState()

            Dim previousIds As List(Of String) = GetStringList(args, "previous_entry_ids")
            If previousIds.Count > 0 Then
                state.LastUsedEntryIds = previousIds
            End If

            Dim retrieval As SharedMethods.SemanticSearchRetrievalResult =
                Await SharedMethods.RetrieveSemanticSearchAsync(
                    path:=path,
                    context:=context,
                    currentQuestion:=question,
                    conversation:=conversation,
                    conversationState:=state,
                    options:=options,
                    cancellationToken:=cancellationToken).ConfigureAwait(False)

            If Not retrieval.IsIndexed Then
                Return BuildRetrievalResponse(path, retrieval, Nothing, Nothing)
            End If

            Dim conversationHandle As String = StoreConversationState(path, state, options)
            Dim retrievalHandle As String = StoreRetrievalState(path, retrieval, options, conversationHandle)

            Return BuildRetrievalResponse(path, retrieval, retrievalHandle, conversationHandle)
        End Function

        Private Shared Async Function ExecuteSemanticIndexSearchContinuationAsync(args As IDictionary(Of String, Object),
                                                                                 context As ISharedContext,
                                                                                 cancellationToken As CancellationToken) As Task(Of String)
            If context Is Nothing Then
                Return BuildError("missing_context", "semantic_index_search_continuation requires a shared LLM context.")
            End If

            Dim path As String = PathPolicy.Resolve(GetStr(args, "path"), PathAccess.Read)
            Dim question As String = GetStr(args, "question")
            Dim conversation As String = GetStr(args, "conversation")
            Dim conversationHandle As String = GetStr(args, "conversation_handle")

            If String.IsNullOrWhiteSpace(question) Then
                Return BuildError("missing_question", "question is required.")
            End If

            If String.IsNullOrWhiteSpace(conversationHandle) Then
                Return BuildError("missing_conversation_handle", "conversation_handle is required.")
            End If

            Dim stateItem As SemanticConversationStateItem = Nothing
            If Not SemanticConversationStore.TryGetValue(conversationHandle, stateItem) OrElse stateItem Is Nothing Then
                Return BuildError("conversation_not_found", "The conversation_handle was not found.")
            End If

            If Not String.Equals(stateItem.Path, path, StringComparison.OrdinalIgnoreCase) Then
                Return BuildError("path_mismatch", "The conversation_handle belongs to a different index path.")
            End If

            Dim options As SharedMethods.SemanticSearchRetrievalOptions =
                BuildRetrievalOptions(args, stateItem.Options)

            Dim retrieval As SharedMethods.SemanticSearchRetrievalResult =
                Await SharedMethods.RetrieveSemanticSearchAsync(
                    path:=path,
                    context:=context,
                    currentQuestion:=question,
                    conversation:=conversation,
                    conversationState:=stateItem.State,
                    options:=options,
                    cancellationToken:=cancellationToken).ConfigureAwait(False)

            stateItem.Options = options
            stateItem.UpdatedUtc = DateTime.UtcNow

            Dim retrievalHandle As String = StoreRetrievalState(path, retrieval, options, conversationHandle)
            Return BuildRetrievalResponse(path, retrieval, retrievalHandle, conversationHandle)
        End Function

        Private Shared Async Function ExecuteSemanticIndexLoadEntriesAsync(args As IDictionary(Of String, Object),
                                                                          cancellationToken As CancellationToken) As Task(Of String)
            Dim path As String = PathPolicy.Resolve(GetStr(args, "path"), PathAccess.Read)
            Dim entryIds As List(Of String) = GetStringList(args, "entry_ids")

            If entryIds.Count = 0 Then
                Return BuildError("missing_entry_ids", "entry_ids is required.")
            End If

            Dim options As SharedMethods.SemanticSearchRetrievalOptions = BuildRetrievalOptions(args, Nothing)

            Dim retrieval As SharedMethods.SemanticSearchRetrievalResult =
                Await SharedMethods.LoadAdditionalSemanticSearchSourcesAsync(
                    path:=path,
                    ids:=entryIds,
                    options:=options,
                    cancellationToken:=cancellationToken).ConfigureAwait(False)

            Dim retrievalHandle As String = Nothing
            If retrieval IsNot Nothing AndAlso retrieval.IsIndexed Then
                retrievalHandle = StoreRetrievalState(path, retrieval, options, Nothing)
            End If

            Return BuildRetrievalResponse(path, retrieval, retrievalHandle, Nothing)
        End Function

        Private Shared Async Function ExecuteSemanticIndexVerifyAnswerAsync(args As IDictionary(Of String, Object),
                                                                           context As ISharedContext,
                                                                           cancellationToken As CancellationToken) As Task(Of String)
            If context Is Nothing Then
                Return BuildError("missing_context", "semantic_index_verify_answer requires a shared LLM context.")
            End If

            Dim path As String = PathPolicy.Resolve(GetStr(args, "path"), PathAccess.Read)
            Dim question As String = GetStr(args, "question")
            Dim conversation As String = GetStr(args, "conversation")
            Dim retrievalHandle As String = GetStr(args, "retrieval_handle")
            Dim answer As String = GetStr(args, "answer")
            Dim specialTaskName As String = GetStr(args, "special_task_name")
            Dim maximumLlmAttempts As Integer =
                GetInt(args, "maximum_llm_attempts", SharedMethods.SemanticSearchDefaultMaximumLlmAttempts)
            Dim maximumConversationCharacters As Integer =
                GetInt(args, "maximum_conversation_characters", SharedMethods.SemanticSearchDefaultMaximumConversationCharacters)

            If String.IsNullOrWhiteSpace(question) Then
                Return BuildError("missing_question", "question is required.")
            End If

            If String.IsNullOrWhiteSpace(retrievalHandle) Then
                Return BuildError("missing_retrieval_handle", "retrieval_handle is required.")
            End If

            If String.IsNullOrWhiteSpace(answer) Then
                Return BuildError("missing_answer", "answer is required.")
            End If

            Dim retrievalItem As SemanticRetrievalStateItem = Nothing
            If Not SemanticRetrievalStore.TryGetValue(retrievalHandle, retrievalItem) OrElse retrievalItem Is Nothing Then
                Return BuildError("retrieval_not_found", "The retrieval_handle was not found.")
            End If

            If Not String.Equals(retrievalItem.Path, path, StringComparison.OrdinalIgnoreCase) Then
                Return BuildError("path_mismatch", "The retrieval_handle belongs to a different index path.")
            End If

            Dim verification As SharedMethods.SemanticSearchResponseVerificationResult =
                Await SharedMethods.VerifySemanticSearchResponseAsync(
                    path:=path,
                    context:=context,
                    specialTaskName:=specialTaskName,
                    currentQuestion:=question,
                    conversation:=conversation,
                    retrieval:=retrievalItem.Retrieval,
                    responseText:=answer,
                    cancellationToken:=cancellationToken,
                    maximumLlmAttempts:=maximumLlmAttempts,
                    maximumConversationCharacters:=maximumConversationCharacters).ConfigureAwait(False)

            Dim verificationHandle As String = StoreVerificationState(path, retrievalHandle, verification)

            Return JsonConvert.SerializeObject(New With {
                Key .path = path,
                Key .retrieval_handle = retrievalHandle,
                Key .verification_handle = verificationHandle,
                Key .supported = verification.Supported,
                Key .unsupported_claims = verification.UnsupportedClaims,
                Key .missing_details = verification.MissingDetails,
                Key .requires_more_sources = verification.RequiresMoreSources,
                Key .additional_entry_ids = verification.AdditionalEntryIds,
                Key .revised_search_intent = verification.RevisedSearchIntent
            })
        End Function

        Private Shared Async Function ExecuteSemanticIndexRetrieveAfterVerificationAsync(args As IDictionary(Of String, Object),
                                                                                        context As ISharedContext,
                                                                                        cancellationToken As CancellationToken) As Task(Of String)
            If context Is Nothing Then
                Return BuildError("missing_context", "semantic_index_retrieve_after_verification requires a shared LLM context.")
            End If

            Dim path As String = PathPolicy.Resolve(GetStr(args, "path"), PathAccess.Read)
            Dim question As String = GetStr(args, "question")
            Dim conversation As String = GetStr(args, "conversation")
            Dim retrievalHandle As String = GetStr(args, "retrieval_handle")
            Dim verificationHandle As String = GetStr(args, "verification_handle")

            If String.IsNullOrWhiteSpace(question) Then
                Return BuildError("missing_question", "question is required.")
            End If

            If String.IsNullOrWhiteSpace(retrievalHandle) Then
                Return BuildError("missing_retrieval_handle", "retrieval_handle is required.")
            End If

            If String.IsNullOrWhiteSpace(verificationHandle) Then
                Return BuildError("missing_verification_handle", "verification_handle is required.")
            End If

            Dim retrievalItem As SemanticRetrievalStateItem = Nothing
            If Not SemanticRetrievalStore.TryGetValue(retrievalHandle, retrievalItem) OrElse retrievalItem Is Nothing Then
                Return BuildError("retrieval_not_found", "The retrieval_handle was not found.")
            End If

            Dim verificationItem As SemanticVerificationStateItem = Nothing
            If Not SemanticVerificationStore.TryGetValue(verificationHandle, verificationItem) OrElse verificationItem Is Nothing Then
                Return BuildError("verification_not_found", "The verification_handle was not found.")
            End If

            If Not String.Equals(retrievalItem.Path, path, StringComparison.OrdinalIgnoreCase) OrElse
               Not String.Equals(verificationItem.Path, path, StringComparison.OrdinalIgnoreCase) Then
                Return BuildError("path_mismatch", "The supplied handles belong to a different index path.")
            End If

            If Not String.Equals(verificationItem.RetrievalHandle, retrievalHandle, StringComparison.OrdinalIgnoreCase) Then
                Return BuildError("handle_mismatch", "The verification_handle does not belong to the supplied retrieval_handle.")
            End If

            Dim options As SharedMethods.SemanticSearchRetrievalOptions =
                BuildRetrievalOptions(args, retrievalItem.Options)

            Dim additional As SharedMethods.SemanticSearchRetrievalResult =
                Await SharedMethods.RetrieveAdditionalSemanticSearchSourcesAsync(
                    path:=path,
                    context:=context,
                    currentQuestion:=question,
                    conversation:=conversation,
                    previousRetrieval:=retrievalItem.Retrieval,
                    verification:=verificationItem.Verification,
                    options:=options,
                    cancellationToken:=cancellationToken).ConfigureAwait(False)

            Dim merged As SharedMethods.SemanticSearchRetrievalResult =
                MergeRetrievalResults(retrievalItem.Retrieval, additional)

            Dim conversationHandle As String = retrievalItem.ConversationHandle
            If Not String.IsNullOrWhiteSpace(conversationHandle) Then
                Dim conversationItem As SemanticConversationStateItem = Nothing
                If SemanticConversationStore.TryGetValue(conversationHandle, conversationItem) AndAlso conversationItem IsNot Nothing Then
                    SharedMethods.UpdateSemanticSearchConversationState(conversationItem.State, merged)
                    conversationItem.Options = options
                    conversationItem.UpdatedUtc = DateTime.UtcNow
                End If
            End If

            Dim mergedHandle As String = StoreRetrievalState(path, merged, options, conversationHandle)
            Return BuildRetrievalResponse(path, merged, mergedHandle, conversationHandle)
        End Function

        Private Shared Function ExecuteSemanticIndexResetConversation(args As IDictionary(Of String, Object)) As String
            Dim conversationHandle As String = GetStr(args, "conversation_handle")
            If String.IsNullOrWhiteSpace(conversationHandle) Then
                Return BuildError("missing_conversation_handle", "conversation_handle is required.")
            End If

            Dim item As SemanticConversationStateItem = Nothing
            If SemanticConversationStore.TryRemove(conversationHandle, item) AndAlso item IsNot Nothing Then
                SharedMethods.ResetSemanticSearchConversationState(item.State)
                Return JsonConvert.SerializeObject(New With {
                    Key .conversation_handle = conversationHandle,
                    Key .reset = True
                })
            End If

            Return BuildError("conversation_not_found", "The conversation_handle was not found.")
        End Function

        Private Shared Function ExecuteSemanticIndexInvalidateCache(args As IDictionary(Of String, Object)) As String
            Dim rawPath As String = GetStr(args, "path")

            If String.IsNullOrWhiteSpace(rawPath) Then
                SharedMethods.InvalidateSemanticSearchIndexCache()
                SemanticConversationStore.Clear()
                SemanticRetrievalStore.Clear()
                SemanticVerificationStore.Clear()

                Return JsonConvert.SerializeObject(New With {
                    Key .invalidated = "all"
                })
            End If

            Dim normalizedPath As String = PathPolicy.Resolve(rawPath, PathAccess.Read)
            SharedMethods.InvalidateSemanticSearchIndexCache(normalizedPath)
            RemoveSemanticHandlesForPath(normalizedPath)

            Return JsonConvert.SerializeObject(New With {
                Key .invalidated = normalizedPath
            })
        End Function

        Private Shared Async Function ExecuteExportToTextAsync(args As IDictionary(Of String, Object),
                                                               context As ISharedContext,
                                                               cancellationToken As CancellationToken) As Task(Of String)
            Dim requestedPath As String = GetStr(args, "input_path")
            If String.IsNullOrWhiteSpace(requestedPath) Then
                requestedPath = GetStr(args, "path")
            End If

            If String.IsNullOrWhiteSpace(requestedPath) Then
                Return BuildError("missing_input_path", "input_path is required.")
            End If

            Dim inputPath As String = PathPolicy.Resolve(requestedPath, PathAccess.Read)
            Dim recursive As Boolean = GetBool(args, "recursive", True)
            Dim overwrite As Boolean = GetBool(args, "overwrite", False)
            Dim ocrPdf As Boolean = GetBool(args, "ocr_pdf", False)
            Dim outputDirectoryArg As String = GetStr(args, "output_directory")

            Dim inputIsFile As Boolean = File.Exists(inputPath)
            Dim inputIsDirectory As Boolean = Directory.Exists(inputPath)

            If Not inputIsFile AndAlso Not inputIsDirectory Then
                Return BuildError("not_found", "The input path was not found.", inputPath)
            End If

            Dim items As New List(Of Object)()
            Dim convertedCount As Integer = 0
            Dim skippedCount As Integer = 0
            Dim failedCount As Integer = 0

            If inputIsFile Then
                cancellationToken.ThrowIfCancellationRequested()
                Dim outputPath As System.String = ResolveSingleFileTextOutputPath(inputPath, outputDirectoryArg)
                Dim item As Newtonsoft.Json.Linq.JObject = Await ExportSingleTextFileAsync(
                    inputPath, outputPath, overwrite, ocrPdf, context, cancellationToken).ConfigureAwait(False)
                items.Add(item)
                CountTextExportItem(item, convertedCount, skippedCount, failedCount)
                Return JsonConvert.SerializeObject(New With {
                    Key .input_path = inputPath,
                    Key .output_root = Path.GetDirectoryName(outputPath),
                    Key .converted_count = convertedCount,
                    Key .skipped_count = skippedCount,
                    Key .failed_count = failedCount,
                    Key .items = items
                })
            End If

            Dim outputRoot As String = ResolveDirectoryTextOutputRoot(inputPath, outputDirectoryArg)
            Dim searchOption As SearchOption = If(recursive, SearchOption.AllDirectories, SearchOption.TopDirectoryOnly)

            For Each sourcePath As String In Directory.GetFiles(inputPath, "*", searchOption).OrderBy(Function(p) p)
                cancellationToken.ThrowIfCancellationRequested()

                If IsUnderPath(sourcePath, outputRoot) OrElse Global.SharedLibrary.SharedLibrary.GeneratedOutputRegistry.IsGeneratedPath(sourcePath) Then
                    Continue For
                End If

                Dim ext As String = Path.GetExtension(sourcePath)
                If Not IsSupportedTextExportExtension(ext) Then
                    skippedCount += 1
                    items.Add(New With {
                        Key .source_path = sourcePath,
                        Key .status = "skipped_unsupported"
                    })
                    Continue For
                End If

                Dim relativePath As String = sourcePath.Substring(inputPath.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                Dim outputPath As String = PathPolicy.Resolve(Path.Combine(outputRoot, relativePath & ".txt"), PathAccess.Write)

                Dim item As Newtonsoft.Json.Linq.JObject = Await ExportSingleTextFileAsync(
                    sourcePath, outputPath, overwrite, ocrPdf, context, cancellationToken).ConfigureAwait(False)
                items.Add(item)
                CountTextExportItem(item, convertedCount, skippedCount, failedCount)
            Next

            Return JsonConvert.SerializeObject(New With {
                Key .input_path = inputPath,
                Key .output_root = outputRoot,
                Key .converted_count = convertedCount,
                Key .skipped_count = skippedCount,
                Key .failed_count = failedCount,
                Key .items = items
            })
        End Function

        Private Shared Sub CountTextExportItem(item As Newtonsoft.Json.Linq.JObject,
                                               ByRef converted As System.Int32, ByRef skipped As System.Int32,
                                               ByRef failed As System.Int32)
            Select Case item.Value(Of System.String)("status")
                Case "converted" : converted += 1
                Case "skipped_existing", "skipped_unsupported" : skipped += 1
                Case Else : failed += 1
            End Select
        End Sub

        ' Both directory and single-file exports use this one side-effect boundary.
        ' Existing-file reuse is provenance-verified only when this session published the exact output from the same extraction resource.
        Private Shared Async Function ExportSingleTextFileAsync(sourcePath As System.String,
                                                                 outputPath As System.String,
                                                                 overwrite As System.Boolean,
                                                                 ocrPdf As System.Boolean,
                                                                 context As ISharedContext,
                                                                 cancellationToken As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of Newtonsoft.Json.Linq.JObject)
            Dim result As TextExportResult = Await TextExportService.ExportFileAsync(
                context, sourcePath, outputPath, New TextExportOptions With {.Overwrite = overwrite, .OcrPdf = ocrPdf},
                cancellationToken).ConfigureAwait(False)
            Return result.LegacyItem
        End Function

        Friend Shared Function GetTextExportSupportedExtensions() As System.Collections.Generic.IReadOnlyCollection(Of System.String)
            Return New System.Collections.ObjectModel.ReadOnlyCollection(Of System.String)(
                SupportedTextExportExtensions.OrderBy(Function(value As System.String) value, System.StringComparer.OrdinalIgnoreCase).ToList())
        End Function

        Friend Shared Function CreateTextExportCallContext(sourcePath As System.String, context As ISharedContext,
                                                          options As TextExportOptions) As ISharedContext
            If context Is Nothing Then Return Nothing
            Dim isolated As ISharedContext = SharedMethods.CreateIsolatedModelCallContext(context)
            Dim extension As System.String = System.IO.Path.GetExtension(sourcePath).ToLowerInvariant()
            Dim taskName As System.String = System.String.Empty
            If extension = ".pdf" AndAlso options.OcrPdf Then taskName = "OCR"
            If SharedMethods.IsBinaryMediaExtension(extension) Then taskName = SharedMethods.TaskFlagForExtension(extension)
            If Not System.String.IsNullOrEmpty(taskName) Then
                isolated = SharedMethods.CreateIsolatedModelCallContext(SharedMethods.ResolveIsolatedSpecialTaskModel(isolated, taskName).Context)
            End If
            Return isolated
        End Function

        Friend Shared Function GetTextExportProcessingSignature(sourcePath As System.String, context As ISharedContext,
                                                                 options As TextExportOptions) As System.String
            Dim isolatedContext As ISharedContext = CreateTextExportCallContext(sourcePath, context, options)
            Return BuildTextExportProcessingSignature(GetExtractionAdapterId(sourcePath), TextExportAdapterVersion,
                GetExtractionConfigurationFingerprint(sourcePath, isolatedContext, options.OcrPdf),
                TextExtractionResourceRegistry.HashString("ocr_pdf=" & options.OcrPdf.ToString(System.Globalization.CultureInfo.InvariantCulture) & "|ocr_batch_pages=" & options.OcrBatchPages.ToString(System.Globalization.CultureInfo.InvariantCulture)))
        End Function

        Friend Shared Function BuildTextExportProcessingSignature(adapterId As System.String, adapterVersion As System.String,
                                                                  configurationFingerprint As System.String, optionsFingerprint As System.String) As System.String
            If System.String.IsNullOrWhiteSpace(adapterVersion) OrElse System.String.IsNullOrWhiteSpace(configurationFingerprint) OrElse
                System.String.IsNullOrWhiteSpace(optionsFingerprint) Then Return System.String.Empty
            Return TextExtractionResourceRegistry.HashString(adapterVersion & "|typed-adapter-v1|" & adapterId & "|" &
                configurationFingerprint & "|text-export-options-v1|" & optionsFingerprint)
        End Function

        Friend Shared Async Function ExportSingleTextFileCoreAsync(sourcePath As System.String,
                                                                 outputPath As System.String,
                                                                 overwrite As System.Boolean,
                                                                 ocrPdf As System.Boolean,
                                                                 context As ISharedContext,
                                                                 cancellationToken As System.Threading.CancellationToken,
                                                                 Optional executionOptions As TextExportOptions = Nothing) As System.Threading.Tasks.Task(Of Newtonsoft.Json.Linq.JObject)
            Dim item As New Newtonsoft.Json.Linq.JObject From {
                {"source_path", sourcePath}, {"output_path", outputPath}, {"status", "failed"},
                {"char_count", Newtonsoft.Json.Linq.JValue.CreateNull()},
                {"byte_count", Newtonsoft.Json.Linq.JValue.CreateNull()},
                {"page_count", Newtonsoft.Json.Linq.JValue.CreateNull()},
                {"ocr_used", Newtonsoft.Json.Linq.JValue.CreateNull()},
                {"ocr_attempted", Newtonsoft.Json.Linq.JValue.CreateNull()},
                {"ocr_skipped_due_to_heuristics", Newtonsoft.Json.Linq.JValue.CreateNull()},
                {"ocr_duration_ms", Newtonsoft.Json.Linq.JValue.CreateNull()},
                {"extraction_complete", Newtonsoft.Json.Linq.JValue.CreateNull()},
                {"completeness_status", "not_verified"}, {"content_format", "unknown"},
                {"source_association_verified", False}, {"extraction_duration_ms", 0}
            }
            Try
                sourcePath = PathPolicy.Resolve(sourcePath, PathAccess.Read)
                outputPath = PathPolicy.Resolve(outputPath, PathAccess.Write)
                item("source_path") = sourcePath
                item("output_path") = outputPath
                cancellationToken.ThrowIfCancellationRequested()

                If System.IO.File.Exists(outputPath) AndAlso Not overwrite Then
                    item("status") = "skipped_existing"
                    Try
                        Dim existing As TextFileSnapshot = TextFileSnapshot.Read(outputPath)
                        AddTextSnapshotMetadata(item, existing)

                        Dim knownResource As TextExtractionResource = Nothing
                        If TextExtractionResourceRegistry.HasPublishedOutput(outputPath) Then
                            Dim currentSourceHash As System.String = TextFileSnapshot.ComputeFileHash(sourcePath)
                            Dim knownAdapterId As System.String = GetExtractionAdapterId(sourcePath)
                            Dim knownAdapterVersion As System.String = TextExportAdapterVersion
                            Dim knownConfigurationFingerprint As System.String = GetExtractionConfigurationFingerprint(sourcePath, context, ocrPdf)
                            Dim knownOptionsFingerprint As System.String = TextExtractionResourceRegistry.HashString(
                                "ocr_pdf=" & ocrPdf.ToString(System.Globalization.CultureInfo.InvariantCulture) & "|ocr_batch_pages=" & If(executionOptions, New TextExportOptions()).OcrBatchPages.ToString(System.Globalization.CultureInfo.InvariantCulture))
                            If TextExtractionResourceRegistry.TryGetVerifiedPublishedOutput(
                                outputPath, currentSourceHash, knownAdapterId, knownAdapterVersion,
                                knownConfigurationFingerprint, knownOptionsFingerprint, knownResource) Then
                            item("reuse_validation") = "session_resource_verified"
                            item("source_association_verified") = True
                            item("source_sha256") = knownResource.SourceSha256
                            item("source_byte_count") = knownResource.SourceByteCount
                            item("resource_id") = knownResource.ResourceId
                            item("resource_reuse") = knownResource.ReuseStatus
                            item("adapter_id") = knownResource.AdapterId
                            item("adapter_version") = knownResource.AdapterVersion
                            item("configuration_fingerprint") = knownResource.ConfigurationFingerprint
                            item("options_fingerprint") = knownResource.OptionsFingerprint
                            If knownResource.Payload IsNot Nothing Then
                                item("page_count") = If(knownResource.Payload.PageCount.HasValue, New Newtonsoft.Json.Linq.JValue(knownResource.Payload.PageCount.Value), Newtonsoft.Json.Linq.JValue.CreateNull())
                                item("ocr_used") = If(knownResource.Payload.OcrUsed.HasValue, New Newtonsoft.Json.Linq.JValue(knownResource.Payload.OcrUsed.Value), Newtonsoft.Json.Linq.JValue.CreateNull())
                                item("ocr_attempted") = If(knownResource.Payload.OcrAttempted.HasValue, New Newtonsoft.Json.Linq.JValue(knownResource.Payload.OcrAttempted.Value), Newtonsoft.Json.Linq.JValue.CreateNull())
                                item("ocr_skipped_due_to_heuristics") = If(knownResource.Payload.OcrSkipped.HasValue, New Newtonsoft.Json.Linq.JValue(knownResource.Payload.OcrSkipped.Value), Newtonsoft.Json.Linq.JValue.CreateNull())
                                item("ocr_duration_ms") = If(knownResource.Payload.OcrDurationMilliseconds.HasValue, New Newtonsoft.Json.Linq.JValue(knownResource.Payload.OcrDurationMilliseconds.Value), Newtonsoft.Json.Linq.JValue.CreateNull())
                                item("content_format") = knownResource.Payload.ContentFormat
                                item("extraction_coverage_basis") = knownResource.Payload.ExtractionCoverageBasis
                                item("extraction_warnings") = Newtonsoft.Json.Linq.JArray.FromObject(TextExtractionDiagnostics.CopyWarnings(knownResource.Payload.ExtractionWarnings))
                                item("extraction_complete") = If(knownResource.Payload.ExtractionComplete.HasValue,
                                    New Newtonsoft.Json.Linq.JValue(knownResource.Payload.ExtractionComplete.Value), Newtonsoft.Json.Linq.JValue.CreateNull())
                                item("completeness_status") = If(knownResource.Payload.ExtractionComplete.HasValue,
                                    If(knownResource.Payload.ExtractionComplete.GetValueOrDefault(), "complete", "incomplete"), "not_verified")
                                AddProcessedRangesMetadata(item, knownResource.Payload.ProcessedRanges)
                            End If
                            Else
                                item("reuse_validation") = "not_verified"
                                item("warning") = "Existing output left unchanged without extraction. Its association with the current source/options has not been verified."
                            End If
                        Else
                            item("reuse_validation") = "not_verified"
                            item("warning") = "Existing output left unchanged without extraction. Its association with the current source/options has not been verified."
                        End If
                    Catch ex As System.Exception
                        ' Metadata must not break legacy skip semantics, e.g. a write-only
                        ' workspace or an output exceeding the configured text-read limit.
                        item("reuse_validation") = "not_verified"
                        item("metadata_warning") = "Existing output metadata unavailable: " & ex.GetType().Name
                        System.Diagnostics.Debug.WriteLine("Text export skipped-existing metadata: " & ex.GetType().FullName)
                    End Try
                    Return item
                End If

                ' Capture the effective extractor model after the legacy overwrite=false
                ' precheck, then keep its fingerprint and all nested reader calls pinned together.
                context = CreateTextExportCallContext(sourcePath, context, If(executionOptions, New TextExportOptions With {.OcrPdf = ocrPdf}))
                Dim adapterId As System.String = GetExtractionAdapterId(sourcePath)
                Dim adapterVersion As System.String = TextExportAdapterVersion
                Dim configurationFingerprint As System.String = GetExtractionConfigurationFingerprint(sourcePath, context, ocrPdf)
                Dim optionsFingerprint As System.String = TextExtractionResourceRegistry.HashString(
                    "ocr_pdf=" & ocrPdf.ToString(System.Globalization.CultureInfo.InvariantCulture) & "|ocr_batch_pages=" & If(executionOptions, New TextExportOptions()).OcrBatchPages.ToString(System.Globalization.CultureInfo.InvariantCulture))

                Dim timer As System.Diagnostics.Stopwatch = System.Diagnostics.Stopwatch.StartNew()
                Dim result As TextExtractionOutcome = Nothing
                Dim resource As TextExtractionResource = Nothing
                Try
                    resource = Await TextExtractionResourceRegistry.ResolveAsync(
                        sourcePath,
                        System.IO.Path.GetDirectoryName(outputPath),
                        adapterId,
                        adapterVersion,
                        configurationFingerprint,
                        optionsFingerprint,
                        Async Function(snapshotPath As System.String, token As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of TextExtractionPayload)
                            Dim localResult As TextExtractionOutcome =
                                Await TryExtractTextForExportAsync(snapshotPath, context, ocrPdf, token, executionOptions).ConfigureAwait(False)
                            Return ConvertToExtractionPayload(localResult)
                        End Function,
                        cancellationToken).ConfigureAwait(False)

                    result = ConvertFromExtractionPayload(If(resource Is Nothing, Nothing, resource.Payload))
                Finally
                    timer.Stop()
                    item("extraction_duration_ms") = timer.ElapsedMilliseconds
                End Try
                cancellationToken.ThrowIfCancellationRequested()
                item("page_count") = If(result.PageCount.HasValue, New Newtonsoft.Json.Linq.JValue(result.PageCount.Value), Newtonsoft.Json.Linq.JValue.CreateNull())
                item("ocr_used") = If(result.OcrUsed.HasValue, New Newtonsoft.Json.Linq.JValue(result.OcrUsed.Value), Newtonsoft.Json.Linq.JValue.CreateNull())
                item("ocr_attempted") = If(result.OcrAttempted.HasValue, New Newtonsoft.Json.Linq.JValue(result.OcrAttempted.Value), Newtonsoft.Json.Linq.JValue.CreateNull())
                item("ocr_skipped_due_to_heuristics") = If(result.OcrSkipped.HasValue, New Newtonsoft.Json.Linq.JValue(result.OcrSkipped.Value), Newtonsoft.Json.Linq.JValue.CreateNull())
                item("ocr_duration_ms") = If(result.OcrDurationMilliseconds.HasValue, New Newtonsoft.Json.Linq.JValue(result.OcrDurationMilliseconds.Value), Newtonsoft.Json.Linq.JValue.CreateNull())
                item("extraction_coverage_basis") = result.ExtractionCoverageBasis
                item("extraction_warnings") = Newtonsoft.Json.Linq.JArray.FromObject(TextExtractionDiagnostics.CopyWarnings(result.ExtractionWarnings))
                item("extraction_complete") = If(result.ExtractionComplete.HasValue, New Newtonsoft.Json.Linq.JValue(result.ExtractionComplete.Value), Newtonsoft.Json.Linq.JValue.CreateNull())
                item("completeness_status") = If(result.ExtractionComplete.HasValue,
                    If(result.ExtractionComplete.GetValueOrDefault(), "complete", "incomplete"), "not_verified")
                item("content_format") = result.ContentFormat
                If resource IsNot Nothing Then
                    item("resource_id") = resource.ResourceId
                    item("resource_reuse") = resource.ReuseStatus
                    item("adapter_id") = resource.AdapterId
                    item("adapter_version") = resource.AdapterVersion
                    item("configuration_fingerprint") = resource.ConfigurationFingerprint
                    item("options_fingerprint") = resource.OptionsFingerprint
                    item("source_sha256") = resource.SourceSha256
                    item("source_byte_count") = resource.SourceByteCount
                End If
                AddProcessedRangesMetadata(item, result.ProcessedRanges)
                If Not result.Success Then
                    item("error") = result.ErrorCode
                    item("message") = result.Message
                    Return item
                End If
                cancellationToken.ThrowIfCancellationRequested()
                Dim snapshot As TextFileSnapshot = TextFileSnapshot.WriteUtf8Atomic(outputPath, result.Content, overwrite)
                AddTextSnapshotMetadata(item, snapshot)
                If resource IsNot Nothing Then
                    TextExtractionResourceRegistry.RegisterPublishedOutput(outputPath, resource, snapshot.Sha256)
                End If
                item("status") = "converted"
                If resource IsNot Nothing Then item("source_sha256") = resource.SourceSha256
                item("source_association_verified") = resource IsNot Nothing
                item("publication") = "atomic"
                If result.ExtractionComplete.HasValue AndAlso Not result.ExtractionComplete.Value Then
                    item("warning") = "The importer reported a potentially incomplete extraction. Do not treat this as a complete document basis."
                End If
                System.Diagnostics.Debug.WriteLine("Text export: status=converted; chars=" & snapshot.Content.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                    "; duration_ms=" & timer.ElapsedMilliseconds.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                    "; snapshot_sha256=" & snapshot.Sha256)
                Return item
            Catch ex As System.OperationCanceledException
                Throw
            Catch ex As TextFileInputException
                item("error") = ex.ErrorCode
                item("message") = ex.Message
                Return item
            Catch ex As System.Exception
                item("error") = "text_export_failed"
                item("message") = ex.Message
                System.Diagnostics.Debug.WriteLine("Text export failed: " & ex.GetType().FullName)
                Return item
            End Try
        End Function

        Private Shared Sub AddTextSnapshotMetadata(item As Newtonsoft.Json.Linq.JObject, snapshot As TextFileSnapshot)
            item("char_count") = snapshot.Content.Length
            item("byte_count") = snapshot.SizeBytes
            item("snapshot_sha256") = snapshot.Sha256
            item("char_count_unit") = "utf16_code_units"
            item("within_text_read_limit") = snapshot.SizeBytes <= PathPolicy.MaxFileSizeBytes
        End Sub

        Private Shared Function ConvertToExtractionPayload(result As TextExtractionOutcome) As TextExtractionPayload
            If result Is Nothing Then
                Return New TextExtractionPayload With {
                    .Success = False,
                    .ErrorCode = "empty_extraction_result",
                    .Message = "The extraction adapter returned no result."
                }
            End If
            Return New TextExtractionPayload With {
                .Success = result.Success,
                .Content = If(result.Content, System.String.Empty),
                .ErrorCode = If(result.ErrorCode, System.String.Empty),
                .Message = If(result.Message, System.String.Empty),
                .PageCount = result.PageCount,
                .OcrUsed = result.OcrUsed,
                .OcrAttempted = result.OcrAttempted,
                .OcrSkipped = result.OcrSkipped,
                .OcrDurationMilliseconds = result.OcrDurationMilliseconds,
                .ExtractionComplete = result.ExtractionComplete,
                .ExtractionCoverageBasis = result.ExtractionCoverageBasis,
                .ExtractionWarnings = TextExtractionDiagnostics.CopyWarnings(result.ExtractionWarnings),
                .ContentFormat = If(result.ContentFormat, "unknown"),
                .ProcessedRanges = If(result.ProcessedRanges, New System.Collections.Generic.List(Of TextExtractionProcessedRange)())
            }
        End Function

        Private Shared Function ConvertFromExtractionPayload(payload As TextExtractionPayload) As TextExtractionOutcome
            If payload Is Nothing Then
                Return New TextExtractionOutcome With {
                    .Success = False,
                    .ErrorCode = "empty_extraction_resource",
                    .Message = "The extraction resource returned no payload."
                }
            End If
            Return New TextExtractionOutcome With {
                .Success = payload.Success,
                .Content = If(payload.Content, System.String.Empty),
                .ErrorCode = If(payload.ErrorCode, System.String.Empty),
                .Message = If(payload.Message, System.String.Empty),
                .PageCount = payload.PageCount,
                .OcrUsed = payload.OcrUsed,
                .OcrAttempted = payload.OcrAttempted,
                .OcrSkipped = payload.OcrSkipped,
                .OcrDurationMilliseconds = payload.OcrDurationMilliseconds,
                .ExtractionComplete = payload.ExtractionComplete,
                .ExtractionCoverageBasis = payload.ExtractionCoverageBasis,
                .ExtractionWarnings = TextExtractionDiagnostics.CopyWarnings(payload.ExtractionWarnings),
                .ContentFormat = If(payload.ContentFormat, "unknown"),
                .ProcessedRanges = If(payload.ProcessedRanges, New System.Collections.Generic.List(Of TextExtractionProcessedRange)())
            }
        End Function

        Private Shared Function GetExtractionAdapterId(sourcePath As System.String) As System.String
            Dim extension As System.String = System.IO.Path.GetExtension(sourcePath).ToLowerInvariant()
            Select Case extension
                Case ".pdf" : Return "pdf"
                Case ".doc", ".docx", ".docm" : Return "word"
                Case ".xlsx", ".xlsm" : Return "excel"
                Case ".pptx", ".pptm" : Return "powerpoint"
                Case ".eml", ".msg" : Return "mail"
                Case Else : Return "text-or-binary"
            End Select
        End Function

        Private Shared Function GetExtractionConfigurationFingerprint(sourcePath As System.String,
                                                                      context As ISharedContext,
                                                                      ocrPdf As System.Boolean) As System.String
            Dim extension As System.String = System.IO.Path.GetExtension(sourcePath).ToLowerInvariant()
            If extension = ".pdf" AndAlso ocrPdf Then
                If context IsNot Nothing Then SharedMethods.ResolveIsolatedSpecialTaskModel(context, "OCR")
                Return SharedMethods.GetOcrConfigurationFingerprint(context)
            End If
            If SharedMethods.IsBinaryMediaExtension(extension) AndAlso context IsNot Nothing Then
                Dim resolved As SharedMethods.IsolatedSpecialTaskModel = SharedMethods.ResolveIsolatedSpecialTaskModel(
                    context, SharedMethods.TaskFlagForExtension(extension))
                Return TextExtractionResourceRegistry.HashString("media-adapter-v1|" & resolved.Signature & "|" & If(context.SP_InsertClipboard, System.String.Empty))
            End If
            If extension = ".doc" Then
                Return TextExtractionResourceRegistry.HashString("legacy-word|enabled=" &
                    (context IsNot Nothing AndAlso context.INI_AllowLegacyDocFiles).ToString(System.Globalization.CultureInfo.InvariantCulture))
            End If
            Return TextExtractionResourceRegistry.HashString(
                "adapter=" & GetExtractionAdapterId(sourcePath) & "|ocr_pdf=" &
                ocrPdf.ToString(System.Globalization.CultureInfo.InvariantCulture))
        End Function

        Private Shared Sub AddProcessedRangesMetadata(item As Newtonsoft.Json.Linq.JObject,
                                                      ranges As System.Collections.Generic.List(Of TextExtractionProcessedRange))
            Dim array As New Newtonsoft.Json.Linq.JArray()
            If ranges IsNot Nothing Then
                For Each range As TextExtractionProcessedRange In ranges
                    If range Is Nothing Then Continue For
                    array.Add(New Newtonsoft.Json.Linq.JObject From {
                        {"start_page", range.StartPage},
                        {"end_page", range.EndPage},
                        {"association", If(range.Association, "range_only")}
                    })
                Next
            End If
            item("processed_ranges") = array
            item("page_mapping_verified") = False
        End Sub

        Private Shared Async Function TryExtractTextForExportAsync(filePath As String,
                                                                   context As ISharedContext,
                                                                   ocrPdf As Boolean,
                                                                   Optional cancellationToken As System.Threading.CancellationToken = Nothing,
                                                                   Optional executionOptions As TextExportOptions = Nothing) As System.Threading.Tasks.Task(Of TextExtractionOutcome)
            Dim ext As String = Path.GetExtension(filePath).ToLowerInvariant()

            Select Case ext
                Case ".txt", ".ini", ".csv", ".log", ".json", ".xml", ".html", ".htm",
                     ".md", ".yaml", ".yml",
                     ".vb", ".cs", ".js", ".ts", ".py", ".java", ".cpp", ".c", ".h", ".sql"

                    Dim plainError As System.String = Nothing
                    Dim plainText As System.String = SharedMethods.ReadTextFile(filePath, False, plainError)
                    Dim plainResult As TextExtractionOutcome = NormalizeLegacyExtractionResult(plainText, "text", plainError)
                    plainResult.ContentFormat = If(ext = ".md", "markdown", "plain_text")
                    plainResult.OcrUsed = False
                    plainResult.OcrAttempted = False
                    If plainResult.Success Then
                        plainResult.ExtractionComplete = True
                        plainResult.ExtractionCoverageBasis = "plain_text_full_read"
                    End If
                    Return plainResult

                Case ".rtf"
                    Dim rtfText As System.String
                    Dim rtfError As System.String = Nothing
                    If System.Threading.Thread.CurrentThread.GetApartmentState() = System.Threading.ApartmentState.STA Then
                        rtfText = SharedMethods.ReadRtfAsText(filePath, True, rtfError)
                    Else
                        rtfText = Await TextExportService.RunStaReaderAsync(
                            Function() SharedMethods.ReadRtfAsText(filePath, True, rtfError), cancellationToken).ConfigureAwait(False)
                    End If
                    Return NormalizeLegacyExtractionResult(rtfText, "rtf", rtfError)

                Case ".doc"
                    If context Is Nothing OrElse Not context.INI_AllowLegacyDocFiles Then
                        Return New TextExtractionOutcome() With {
                            .Success = False,
                            .ErrorCode = "legacy_doc_disabled",
                            .Message = ".doc extraction is disabled unless AllowLegacyDocFiles is enabled."
                        }
                    End If

                    Dim legacyError As System.String = Nothing
                    If executionOptions IsNot Nothing AndAlso executionOptions.HostReaderDispatcher IsNot Nothing Then
                        Dim legacyText As System.String = Await executionOptions.HostReaderDispatcher.Invoke(
                            Function()
                                If System.Threading.Thread.CurrentThread.GetApartmentState() <> System.Threading.ApartmentState.STA Then
                                    Throw New System.InvalidOperationException("The configured host reader dispatcher did not enter an STA thread.")
                                End If
                                Return SharedMethods.ReadWordDocument(filePath, True, legacyError)
                            End Function, cancellationToken).ConfigureAwait(False)
                        Return NormalizeLegacyExtractionResult(legacyText, "word", legacyError)
                    End If
                    If System.Threading.Thread.CurrentThread.GetApartmentState() <> System.Threading.ApartmentState.STA Then
                        Return New TextExtractionOutcome With {.ErrorCode = "requires_sta", .Message = "Legacy Word extraction requires the foreground host reader dispatcher; extraction was deferred."}
                    End If
                    Return NormalizeLegacyExtractionResult(SharedMethods.ReadWordDocument(filePath, True, legacyError), "word", legacyError)

                Case ".docx", ".docm"
                    Dim docxError As System.String = Nothing
                    Dim docxText As System.String = SharedMethods.ReadDocxSandboxed(filePath, readError:=docxError)
                    Return NormalizeLegacyExtractionResult(docxText, "word", docxError)

                Case ".xlsx", ".xlsm"
                    Dim xlsxError As System.String = Nothing
                    Dim xlsxText As System.String = SharedMethods.ReadXlsxSandboxed(filePath, silent:=True, askWorksheetSelection:=False, readError:=xlsxError)
                    Return NormalizeLegacyExtractionResult(xlsxText, "excel", xlsxError)

                Case ".pptx", ".pptm"
                    Dim pptxError As System.String = Nothing
                    Dim pptxText As System.String = SharedMethods.ReadPptxSandboxed(filePath, readError:=pptxError)
                    Return NormalizeLegacyExtractionResult(pptxText, "powerpoint", pptxError)

                Case ".pdf"
                    Dim pdf As SharedMethods.PdfReadResult = Await SharedMethods.ReadPdfAsTextEx(
                        pdfPath:=filePath,
                        ReturnErrorInsteadOfEmpty:=False,
                        DoOCR:=ocrPdf AndAlso context IsNot Nothing,
                        AskUser:=False,
                        context:=context,
                        ShowOcrProgressWindow:=False,
                        OcrBatchPages:=If(executionOptions, New TextExportOptions()).OcrBatchPages,
                        CancellationToken:=cancellationToken).ConfigureAwait(False)
                    Dim succeeded As System.Boolean = System.String.IsNullOrWhiteSpace(pdf.ErrorCode) AndAlso
                        Not System.String.IsNullOrWhiteSpace(pdf.Content)
                    Return New TextExtractionOutcome() With {
                        .Success = succeeded,
                        .Content = pdf.Content,
                        .ErrorCode = If(succeeded, System.String.Empty, If(System.String.IsNullOrWhiteSpace(pdf.ErrorCode), "empty_pdf_extraction", pdf.ErrorCode)),
                        .Message = If(succeeded, System.String.Empty, If(System.String.IsNullOrWhiteSpace(pdf.ErrorMessage), "No readable PDF content was extracted.", pdf.ErrorMessage)),
                        .PageCount = pdf.PageCount,
                        .OcrUsed = pdf.OcrUsed,
                        .OcrAttempted = pdf.OcrAttempted,
                        .OcrSkipped = pdf.OcrWasSkippedDueToHeuristics,
                        .OcrDurationMilliseconds = pdf.OcrDurationMilliseconds,
                        .ExtractionComplete = If(Not succeeded, CType(False, System.Nullable(Of System.Boolean)), pdf.ExtractionComplete),
                        .ExtractionCoverageBasis = pdf.ExtractionCoverageBasis,
                        .ExtractionWarnings = TextExtractionDiagnostics.CopyWarnings(pdf.ExtractionWarnings),
                        .ContentFormat = If(pdf.OcrUsed, "ocr_model_text", "plain_text"),
                        .ProcessedRanges = BuildPdfProcessedRanges(pdf)
                    }

                Case ".eml"
                    Dim emlError As System.String = Nothing
                    Dim emlText As System.String = SharedMethods.ReadEmlSandboxed(filePath, readError:=emlError)
                    Return NormalizeLegacyExtractionResult(emlText, "mail", emlError)

                Case ".msg"
                    Dim msgError As System.String = Nothing
                    Dim msgText As System.String = SharedMethods.ReadMsgSandboxed(filePath, readError:=msgError)
                    Return NormalizeLegacyExtractionResult(msgText, "mail", msgError)

                Case Else
                    If SharedMethods.IsBinaryMediaExtension(ext) Then
                        If context Is Nothing OrElse Not SharedMethods.IsModelCapableForExtension(context, ext) Then
                            Return New TextExtractionOutcome() With {
                                .Success = False,
                                .ErrorCode = "unsupported_binary_media",
                                .Message = "No suitable model configuration is available for this binary media type."
                            }
                        End If

                        Dim taskFlag As String = SharedMethods.TaskFlagForExtension(ext)
                        Dim content As String =
                            Await SharedMethods.ReadBinaryFileViaLLM(
                                filePath:=filePath,
                                context:=context,
                                askUser:=False,
                                taskFlag:=taskFlag,
                                cancellationToken:=cancellationToken).ConfigureAwait(False)

                        Dim mediaFailed As System.Boolean = System.String.IsNullOrWhiteSpace(content) OrElse
                            content.StartsWith("HTTP Error ", System.StringComparison.OrdinalIgnoreCase) OrElse
                            content.StartsWith("An unexpected error occurred when accessing the LLM endpoint:", System.StringComparison.OrdinalIgnoreCase) OrElse
                            content = "Aborted by user."
                        Return New TextExtractionOutcome() With {
                            .Success = Not mediaFailed,
                            .Content = If(mediaFailed, System.String.Empty, content),
                            .ErrorCode = If(mediaFailed, "media_extraction_failed", System.String.Empty),
                            .Message = If(mediaFailed, "The media adapter returned no usable extraction.", System.String.Empty)
                        }
                    End If

                    Return New TextExtractionOutcome() With {
                        .Success = False,
                        .ErrorCode = "unsupported_extension",
                        .Message = "The file type is not supported for text export."
                    }
            End Select
        End Function

        ' Consume reader-reported status at the adapter boundary. Document contents that
        ' literally start with Error: or Error reading remain ordinary source text.
        Private Shared Function NormalizeLegacyExtractionResult(content As System.String, adapterId As System.String,
                                                                readError As System.String) As TextExtractionOutcome
            Dim value As System.String = If(content, System.String.Empty)
            If Not System.String.IsNullOrEmpty(readError) Then
                Dim empty As System.Boolean = readError = "Error: No data found in .xlsx." OrElse
                    readError = "Error: No text content found in .pptx." OrElse readError = "Error: No text content found in .docx." OrElse
                    readError = "Error: Empty .eml file."
                Return New TextExtractionOutcome With {
                    .Success = False, .ErrorCode = If(empty, "empty_extraction", "reader_failed"),
                    .Message = "The " & adapterId & " reader reported: " & readError,
                    .ExtractionComplete = False
                }
            End If
            If readError Is Nothing Then
                Return New TextExtractionOutcome With {.Success = False, .ErrorCode = "reader_status_unknown",
                    .Message = "The reader returned no structured completion status."}
            End If
            Dim incomplete As System.Boolean = adapterId = "mail" AndAlso
                System.Text.RegularExpressions.Regex.IsMatch(value, "(?im)^\[Skipped: (?:failed to extract|PDF extraction failed|max nesting depth reached)")
            Return New TextExtractionOutcome With {
                .Success = True, .Content = value,
                .ExtractionComplete = If(incomplete, CType(False, System.Nullable(Of System.Boolean)), Nothing),
                .ExtractionCoverageBasis = If(incomplete, "reader_reported_gaps", "reader_coverage_unverified")
            }
        End Function

        Private Shared Function BuildPdfProcessedRanges(pdf As SharedMethods.PdfReadResult) As System.Collections.Generic.List(Of TextExtractionProcessedRange)
            Dim ranges As New System.Collections.Generic.List(Of TextExtractionProcessedRange)()
            If pdf Is Nothing Then Return ranges
            If pdf.TextPageSequenceComplete AndAlso pdf.TextProcessedPageCount > 0 AndAlso System.String.IsNullOrWhiteSpace(pdf.ErrorCode) Then
                ranges.Add(New TextExtractionProcessedRange With {
                    .StartPage = 1,
                    .EndPage = pdf.TextProcessedPageCount,
                    .Association = "text_extraction_range"
                })
            End If
            If pdf.OcrProcessedRanges IsNot Nothing AndAlso pdf.OcrProcessedRanges.Count > 0 Then
                ranges.AddRange(pdf.OcrProcessedRanges)
            End If
            Return ranges
        End Function

        Private Shared Function ResolveMetadataProfile(value As String) As SharedMethods.SemanticSearchMetadataProfile
            If String.IsNullOrWhiteSpace(value) Then
                Return SharedMethods.SemanticSearchMetadataProfile.Generic
            End If

            Dim resolved As SharedMethods.SemanticSearchMetadataProfile
            If MetadataProfileMap.TryGetValue(value.Trim(), resolved) Then
                Return resolved
            End If

            Return SharedMethods.SemanticSearchMetadataProfile.Generic
        End Function

        Private Shared Function BuildRetrievalOptions(args As IDictionary(Of String, Object),
                                                      defaults As SharedMethods.SemanticSearchRetrievalOptions) As SharedMethods.SemanticSearchRetrievalOptions
            Dim options As SharedMethods.SemanticSearchRetrievalOptions =
                If(defaults, New SharedMethods.SemanticSearchRetrievalOptions())

            options.MinimumSelectedSegments =
                GetInt(args, "minimum_selected_segments", options.MinimumSelectedSegments)
            options.MaximumSelectedSegments =
                GetInt(args, "maximum_selected_segments", options.MaximumSelectedSegments)
            options.MaximumTotalSegments =
                GetInt(args, "maximum_total_segments", options.MaximumTotalSegments)
            options.ContextBytesBefore =
                GetInt(args, "context_bytes_before", options.ContextBytesBefore)
            options.ContextBytesAfter =
                GetInt(args, "context_bytes_after", options.ContextBytesAfter)
            options.EnableFullScanFallback =
                GetBool(args, "enable_full_scan_fallback", options.EnableFullScanFallback)
            options.ForceFullScan =
                GetBool(args, "force_full_scan", options.ForceFullScan)
            options.SpecialTaskName = SemanticSearchTaskName

            Return options
        End Function

        Private Shared Function BuildRetrievalResponse(path As String,
                                                       retrieval As SharedMethods.SemanticSearchRetrievalResult,
                                                       retrievalHandle As String,
                                                       conversationHandle As String) As String
            If retrieval Is Nothing Then
                Return JsonConvert.SerializeObject(New With {
                    Key .path = path,
                    Key .is_indexed = False,
                    Key .diagnostic_message = "No retrieval result was returned."
                })
            End If

            Dim loadedSources As List(Of Object) =
                retrieval.LoadedSources.Select(
                    Function(item As SharedMethods.SemanticSearchLoadedSourceSegment) New With {
                        Key .entry_ids = item.EntryIds,
                        Key .document_id = item.DocumentId,
                        Key .document_stable_id = item.DocumentStableId,
                        Key .document_name = item.DocumentName,
                        Key .absolute_start_byte = item.AbsoluteStartByte,
                        Key .relative_start_byte = item.RelativeStartByte,
                        Key .document_relative_start_byte = item.DocumentRelativeStartByte,
                        Key .length_bytes = item.LengthBytes,
                        Key .text = item.Text
                    }).Cast(Of Object)().ToList()

            Return JsonConvert.SerializeObject(New With {
                Key .path = path,
                Key .is_indexed = retrieval.IsIndexed,
                Key .retrieval_handle = retrievalHandle,
                Key .conversation_handle = conversationHandle,
                Key .selected_entry_ids = retrieval.SelectedEntryIds,
                Key .loaded_sources = loadedSources,
                Key .reduced_source_text = retrieval.ReducedSourceText,
                Key .used_fallback = retrieval.UsedFallback,
                Key .diagnostic_message = retrieval.DiagnosticMessage
            })
        End Function

        Private Shared Function MergeRetrievalResults(previousRetrieval As SharedMethods.SemanticSearchRetrievalResult,
                                                      additionalRetrieval As SharedMethods.SemanticSearchRetrievalResult) As SharedMethods.SemanticSearchRetrievalResult
            If previousRetrieval Is Nothing Then
                Return additionalRetrieval
            End If

            If additionalRetrieval Is Nothing Then
                Return previousRetrieval
            End If

            Dim merged As New SharedMethods.SemanticSearchRetrievalResult() With {
                .IsIndexed = previousRetrieval.IsIndexed OrElse additionalRetrieval.IsIndexed,
                .SearchPreparation = If(previousRetrieval.SearchPreparation, additionalRetrieval.SearchPreparation),
                .Selection = If(additionalRetrieval.Selection, previousRetrieval.Selection),
                .UsedFallback = previousRetrieval.UsedFallback OrElse additionalRetrieval.UsedFallback,
                .DiagnosticMessage = If(
                    String.IsNullOrWhiteSpace(additionalRetrieval.DiagnosticMessage),
                    previousRetrieval.DiagnosticMessage,
                    additionalRetrieval.DiagnosticMessage)
            }

            merged.SelectedEntryIds =
                previousRetrieval.SelectedEntryIds.
                    Concat(additionalRetrieval.SelectedEntryIds).
                    Where(Function(id As String) Not String.IsNullOrWhiteSpace(id)).
                    Distinct(StringComparer.OrdinalIgnoreCase).
                    ToList()

            Dim sourceMap As New Dictionary(Of String, SharedMethods.SemanticSearchLoadedSourceSegment)(StringComparer.Ordinal)
            For Each source As SharedMethods.SemanticSearchLoadedSourceSegment In previousRetrieval.LoadedSources.Concat(additionalRetrieval.LoadedSources)
                Dim key As String =
                    source.AbsoluteStartByte.ToString(System.Globalization.CultureInfo.InvariantCulture) & "|" &
                    source.LengthBytes.ToString(System.Globalization.CultureInfo.InvariantCulture) & "|" &
                    If(source.DocumentId, "")

                If Not sourceMap.ContainsKey(key) Then
                    sourceMap.Add(key, source)
                End If
            Next

            merged.LoadedSources = sourceMap.Values.
                OrderBy(Function(item As SharedMethods.SemanticSearchLoadedSourceSegment) item.AbsoluteStartByte).
                ToList()

            merged.FullScanResults =
                previousRetrieval.FullScanResults.
                    Concat(additionalRetrieval.FullScanResults).
                    GroupBy(Function(item As SharedMethods.SemanticSearchSegmentScanResult) item.Id, StringComparer.OrdinalIgnoreCase).
                    Select(Function(group) group.OrderByDescending(Function(item) item.Relevance).First()).
                    ToList()

            merged.ReducedSourceText = MergeSourceText(
                previousRetrieval.ReducedSourceText,
                additionalRetrieval.ReducedSourceText)

            Return merged
        End Function

        Private Shared Function MergeSourceText(previousText As String, additionalText As String) As String
            Dim leftText As String = If(previousText, "")
            Dim rightText As String = If(additionalText, "")

            If String.IsNullOrWhiteSpace(leftText) Then
                Return rightText
            End If

            If String.IsNullOrWhiteSpace(rightText) Then
                Return leftText
            End If

            If leftText.IndexOf(rightText, StringComparison.Ordinal) >= 0 Then
                Return leftText
            End If

            If rightText.IndexOf(leftText, StringComparison.Ordinal) >= 0 Then
                Return rightText
            End If

            Return leftText.TrimEnd() & vbCrLf & vbCrLf & rightText.TrimStart()
        End Function

        Private Shared Function StoreConversationState(path As String,
                                                       state As SharedMethods.SemanticSearchConversationState,
                                                       options As SharedMethods.SemanticSearchRetrievalOptions) As String
            Dim handle As String = "ssc_" & Guid.NewGuid().ToString("N")
            SemanticConversationStore(handle) = New SemanticConversationStateItem() With {
                .Handle = handle,
                .Path = path,
                .State = state,
                .Options = options,
                .UpdatedUtc = DateTime.UtcNow
            }
            Return handle
        End Function

        Private Shared Function StoreRetrievalState(path As String,
                                                    retrieval As SharedMethods.SemanticSearchRetrievalResult,
                                                    options As SharedMethods.SemanticSearchRetrievalOptions,
                                                    conversationHandle As String) As String
            Dim handle As String = "ssr_" & Guid.NewGuid().ToString("N")
            SemanticRetrievalStore(handle) = New SemanticRetrievalStateItem() With {
                .Handle = handle,
                .Path = path,
                .ConversationHandle = If(conversationHandle, ""),
                .Retrieval = retrieval,
                .Options = options,
                .UpdatedUtc = DateTime.UtcNow
            }
            Return handle
        End Function

        Private Shared Function StoreVerificationState(path As String,
                                                       retrievalHandle As String,
                                                       verification As SharedMethods.SemanticSearchResponseVerificationResult) As String
            Dim handle As String = "ssv_" & Guid.NewGuid().ToString("N")
            SemanticVerificationStore(handle) = New SemanticVerificationStateItem() With {
                .Handle = handle,
                .Path = path,
                .RetrievalHandle = retrievalHandle,
                .Verification = verification,
                .UpdatedUtc = DateTime.UtcNow
            }
            Return handle
        End Function

        Private Shared Sub RemoveSemanticHandlesForPath(path As String)
            For Each kvp In SemanticConversationStore
                If String.Equals(kvp.Value.Path, path, StringComparison.OrdinalIgnoreCase) Then
                    Dim removed As SemanticConversationStateItem = Nothing
                    SemanticConversationStore.TryRemove(kvp.Key, removed)
                End If
            Next

            For Each kvp In SemanticRetrievalStore
                If String.Equals(kvp.Value.Path, path, StringComparison.OrdinalIgnoreCase) Then
                    Dim removed As SemanticRetrievalStateItem = Nothing
                    SemanticRetrievalStore.TryRemove(kvp.Key, removed)
                End If
            Next

            For Each kvp In SemanticVerificationStore
                If String.Equals(kvp.Value.Path, path, StringComparison.OrdinalIgnoreCase) Then
                    Dim removed As SemanticVerificationStateItem = Nothing
                    SemanticVerificationStore.TryRemove(kvp.Key, removed)
                End If
            Next
        End Sub

        Private Shared Function ResolveSingleFileTextOutputPath(inputPath As String,
                                                                outputDirectoryArg As String) As String
            If String.IsNullOrWhiteSpace(outputDirectoryArg) Then
                Return PathPolicy.Resolve(inputPath & ".txt", PathAccess.Write)
            End If

            Dim outputRoot As String = ResolveRelativeOrAbsoluteOutputDirectory(inputPath, outputDirectoryArg, False)
            Return PathPolicy.Resolve(Path.Combine(outputRoot, Path.GetFileName(inputPath) & ".txt"), PathAccess.Write)
        End Function

        Private Shared Function ResolveDirectoryTextOutputRoot(inputDirectory As String,
                                                               outputDirectoryArg As String) As String
            If String.IsNullOrWhiteSpace(outputDirectoryArg) Then
                Return PathPolicy.Resolve(Path.Combine(inputDirectory, DefaultTextExportDirectoryName), PathAccess.Write)
            End If

            Return ResolveRelativeOrAbsoluteOutputDirectory(inputDirectory, outputDirectoryArg, True)
        End Function

        Private Shared Function ResolveRelativeOrAbsoluteOutputDirectory(inputPath As String,
                                                                        outputDirectoryArg As String,
                                                                        inputIsDirectory As Boolean) As String
            Dim candidate As String = outputDirectoryArg.Trim()

            If Not Path.IsPathRooted(candidate) Then
                Dim baseDirectory As String =
                    If(inputIsDirectory, inputPath, Path.GetDirectoryName(inputPath))
                candidate = Path.Combine(baseDirectory, candidate)
            End If

            Return PathPolicy.Resolve(candidate, PathAccess.Write)
        End Function

        Private Shared Function IsSupportedTextExportExtension(extension As String) As Boolean
            Return SupportedTextExportExtensions.Contains(If(extension, ""))
        End Function

        Private Shared Function IsUnderPath(candidatePath As String, rootPath As String) As Boolean
            If String.IsNullOrWhiteSpace(candidatePath) OrElse String.IsNullOrWhiteSpace(rootPath) Then
                Return False
            End If

            Dim fullCandidate As String = Path.GetFullPath(candidatePath).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            Dim fullRoot As String = Path.GetFullPath(rootPath).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)

            If String.Equals(fullCandidate, fullRoot, StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If

            Return fullCandidate.StartsWith(fullRoot & Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) OrElse
                   fullCandidate.StartsWith(fullRoot & Path.AltDirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function GetStringList(args As IDictionary(Of String, Object), name As String) As List(Of String)
            Dim result As New List(Of String)()
            If args Is Nothing Then
                Return result
            End If

            Dim value As Object = Nothing
            If Not args.TryGetValue(name, value) OrElse value Is Nothing Then
                Return result
            End If

            If TypeOf value Is String Then
                Dim textValue As String = CStr(value)
                For Each part As String In textValue.Split(New Char() {","c, ControlChars.Cr, ControlChars.Lf}, StringSplitOptions.RemoveEmptyEntries)
                    Dim trimmed As String = part.Trim()
                    If trimmed <> "" Then
                        result.Add(trimmed)
                    End If
                Next
                Return result
            End If

            If TypeOf value Is Newtonsoft.Json.Linq.JArray Then
                For Each token In DirectCast(value, Newtonsoft.Json.Linq.JArray)
                    Dim textValue As String = token.ToString().Trim()
                    If textValue <> "" Then
                        result.Add(textValue)
                    End If
                Next
                Return result
            End If

            If TypeOf value Is System.Collections.IEnumerable Then
                For Each item As Object In DirectCast(value, System.Collections.IEnumerable)
                    If item Is Nothing Then Continue For
                    Dim textValue As String = item.ToString().Trim()
                    If textValue <> "" Then
                        result.Add(textValue)
                    End If
                Next
            End If

            Return result.
                Distinct(StringComparer.OrdinalIgnoreCase).
                ToList()
        End Function

        Private Shared Function BuildError(code As String,
                                           message As String,
                                           Optional path As String = Nothing) As String
            Return JsonConvert.SerializeObject(New With {
                Key .error = code,
                Key .message = message,
                Key .path = path
            })
        End Function

    End Class

End Namespace
