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
                    "Silently extracts readable text from a supported file or from all supported files in a directory and saves UTF-8 .txt files without using host UI prompts. For directory input, subdirectories can be included and structure is preserved under the output directory.",
                    "{""type"":""object"",""properties"":{" &
                        """input_path"":{""type"":""string"",""description"":""Required file or directory path.""}," &
                        """output_directory"":{""type"":""string"",""description"":""Optional output directory. For directory input, relative paths are preserved under this root.""}," &
                        """recursive"":{""type"":""boolean"",""description"":""For directory input, include subdirectories. Default true.""}," &
                        """overwrite"":{""type"":""boolean"",""description"":""Overwrite existing .txt outputs only when true. Default false.""}," &
                        """ocr_pdf"":{""type"":""boolean"",""description"":""Enable silent OCR heuristics for PDFs when a suitable model is configured. Default false.""}}," &
                        """required"":[""input_path""]}",
                    923,
                    "Text (export to text)"),
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
                                                modelDescription As String) As ModelConfig
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
                .ToolErrorHandling = "skip"
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

                Dim outputPath As String =
                    ResolveSingleFileTextOutputPath(inputPath, outputDirectoryArg)

                Dim result As TextExtractionOutcome =
                    Await TryExtractTextForExportAsync(inputPath, context, ocrPdf).ConfigureAwait(False)

                If File.Exists(outputPath) AndAlso Not overwrite Then
                    skippedCount += 1
                    items.Add(New With {
                        Key .source_path = inputPath,
                        Key .output_path = outputPath,
                        Key .status = "skipped_existing"
                    })
                ElseIf Not result.Success Then
                    failedCount += 1
                    items.Add(New With {
                        Key .source_path = inputPath,
                        Key .output_path = outputPath,
                        Key .status = "failed",
                        Key .error = result.ErrorCode,
                        Key .message = result.Message
                    })
                Else
                    Directory.CreateDirectory(Path.GetDirectoryName(outputPath))
                    File.WriteAllText(outputPath, result.Content, Encoding.UTF8)
                    convertedCount += 1
                    items.Add(New With {
                        Key .source_path = inputPath,
                        Key .output_path = outputPath,
                        Key .status = "converted"
                    })
                End If

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

                If IsUnderPath(sourcePath, outputRoot) Then
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

                If File.Exists(outputPath) AndAlso Not overwrite Then
                    skippedCount += 1
                    items.Add(New With {
                        Key .source_path = sourcePath,
                        Key .output_path = outputPath,
                        Key .status = "skipped_existing"
                    })
                    Continue For
                End If

                Dim result As TextExtractionOutcome =
                    Await TryExtractTextForExportAsync(sourcePath, context, ocrPdf).ConfigureAwait(False)

                If Not result.Success Then
                    failedCount += 1
                    items.Add(New With {
                        Key .source_path = sourcePath,
                        Key .output_path = outputPath,
                        Key .status = "failed",
                        Key .error = result.ErrorCode,
                        Key .message = result.Message
                    })
                    Continue For
                End If

                Directory.CreateDirectory(Path.GetDirectoryName(outputPath))
                File.WriteAllText(outputPath, result.Content, Encoding.UTF8)

                convertedCount += 1
                items.Add(New With {
                    Key .source_path = sourcePath,
                    Key .output_path = outputPath,
                    Key .status = "converted"
                })
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

        Private Shared Async Function TryExtractTextForExportAsync(filePath As String,
                                                                   context As ISharedContext,
                                                                   ocrPdf As Boolean) As Task(Of TextExtractionOutcome)
            Dim ext As String = Path.GetExtension(filePath).ToLowerInvariant()

            Select Case ext
                Case ".txt", ".ini", ".csv", ".log", ".json", ".xml", ".html", ".htm",
                     ".md", ".yaml", ".yml",
                     ".vb", ".cs", ".js", ".ts", ".py", ".java", ".cpp", ".c", ".h", ".sql"

                    Return New TextExtractionOutcome() With {
                        .Success = True,
                        .Content = SharedMethods.ReadTextFile(filePath, False)
                    }

                Case ".rtf"
                    Return New TextExtractionOutcome() With {
                        .Success = True,
                        .Content = SharedMethods.ReadRtfAsText(filePath, False)
                    }

                Case ".doc"
                    If context Is Nothing OrElse Not context.INI_AllowLegacyDocFiles Then
                        Return New TextExtractionOutcome() With {
                            .Success = False,
                            .ErrorCode = "legacy_doc_disabled",
                            .Message = ".doc extraction is disabled unless AllowLegacyDocFiles is enabled."
                        }
                    End If

                    Return New TextExtractionOutcome() With {
                        .Success = True,
                        .Content = SharedMethods.ReadWordDocument(filePath, False)
                    }

                Case ".docx", ".docm"
                    Return New TextExtractionOutcome() With {
                        .Success = True,
                        .Content = SharedMethods.ReadDocxSandboxed(filePath)
                    }

                Case ".xlsx", ".xlsm"
                    Return New TextExtractionOutcome() With {
                        .Success = True,
                        .Content = SharedMethods.ReadXlsxSandboxed(filePath, silent:=True, askWorksheetSelection:=False)
                    }

                Case ".pptx", ".pptm"
                    Return New TextExtractionOutcome() With {
                        .Success = True,
                        .Content = SharedMethods.ReadPptxSandboxed(filePath)
                    }

                Case ".pdf"
                    Return New TextExtractionOutcome() With {
                        .Success = True,
                        .Content = Await SharedMethods.ReadPdfAsText(
                            pdfPath:=filePath,
                            ReturnErrorInsteadOfEmpty:=True,
                            DoOCR:=ocrPdf AndAlso context IsNot Nothing,
                            AskUser:=False,
                            context:=context,
                            ShowOcrProgressWindow:=False).ConfigureAwait(False)
                    }

                Case ".eml"
                    Return New TextExtractionOutcome() With {
                        .Success = True,
                        .Content = SharedMethods.ReadEmlSandboxed(filePath)
                    }

                Case ".msg"
                    Return New TextExtractionOutcome() With {
                        .Success = True,
                        .Content = SharedMethods.ReadMsgSandboxed(filePath)
                    }

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
                                taskFlag:=taskFlag).ConfigureAwait(False)

                        Return New TextExtractionOutcome() With {
                            .Success = True,
                            .Content = content
                        }
                    End If

                    Return New TextExtractionOutcome() With {
                        .Success = False,
                        .ErrorCode = "unsupported_extension",
                        .Message = "The file type is not supported for text export."
                    }
            End Select
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
