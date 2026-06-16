
' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.Tools.DataCollector.vb
' Purpose:
'   Config-driven structured email data collection tool for AutoPilot.
'
' Notes:
'   - The model extracts values; this tool validates, normalizes, duplicate-checks,
'     writes, and audit-logs.
'   - Target paths and formats are configuration-bound only.
'   - XML is intentionally not implemented in v1, but the writer abstraction allows
'     adding it later.
'   - DEBUG-only self-test helpers are included at the bottom of this file and can
'     be removed easily later.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Net.Mail
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    Private Shared AP_DataCollectorConfigPathPlaceholder As String = ""

    Private Const AP_DataCollectorToolVersion As String = "1.0.0"
    Private Const AP_DataCollectorSchemaVersion As String = "1.0"

    Private Const AP_DC_Status_Success As String = "success"
    Private Const AP_DC_Status_Preview As String = "preview"
    Private Const AP_DC_Status_Written As String = "written"
    Private Const AP_DC_Status_Updated As String = "updated"
    Private Const AP_DC_Status_DuplicateDetected As String = "duplicate_detected"
    Private Const AP_DC_Status_DuplicateIgnored As String = "duplicate_ignored"
    Private Const AP_DC_Status_ValidationFailed As String = "validation_failed"
    Private Const AP_DC_Status_NormalizationFailed As String = "normalization_failed"
    Private Const AP_DC_Status_UseCaseNotFound As String = "use_case_not_found"
    Private Const AP_DC_Status_ConfigurationInvalid As String = "configuration_invalid"
    Private Const AP_DC_Status_TargetPathNotAllowed As String = "target_path_not_allowed"
    Private Const AP_DC_Status_TargetFileInvalid As String = "target_file_invalid"
    Private Const AP_DC_Status_TargetFileLocked As String = "target_file_locked"
    Private Const AP_DC_Status_WriteFailed As String = "write_failed"
    Private Const AP_DC_Status_LogWriteFailed As String = "log_write_failed"
    Private Const AP_DC_Status_UnsupportedFormat As String = "unsupported_format"
    Private Const AP_DC_Status_UnsupportedWriteMode As String = "unsupported_write_mode"

    Private Interface IDataCollectorOutputWriter
        Function Preview(record As JObject, useCase As DataCollectorUseCaseConfig) As JObject
        Function Write(targetPath As String,
                       record As JObject,
                       useCase As DataCollectorUseCaseConfig,
                       config As DataCollectorConfiguration,
                       effectiveWriteMode As String,
                       duplicateResult As DataCollectorDuplicateCheckResult) As DataCollectorWriteResult
    End Interface

    Private Class DataCollectorConfiguration
        Public Property schemaVersion As String
        Public Property configVersion As String
        Public Property enabled As Boolean
        Public Property allowedBaseDirectory As String
        Public Property defaults As DataCollectorDefaultsConfig
        Public Property log As DataCollectorLogConfig
        Public Property useCases As List(Of DataCollectorUseCaseConfig)
    End Class

    Private Class DataCollectorDefaultsConfig
        Public Property culture As String = "invariant"
        Public Property timezone As String = ""
        Public Property encoding As String = "utf-8"
        Public Property createDirectories As Boolean = True
        Public Property atomicWrites As Boolean = True
        Public Property backupBeforeStructuredUpdate As Boolean = True
        Public Property maxOriginalRequestLength As Integer = 50000
        Public Property maxRecordSize As Integer = 250000
        Public Property maxFieldLength As Integer = 100000
    End Class

    Private Class DataCollectorLogConfig
        Public Property enabled As Boolean
        Public Property directory As String
        Public Property fileNameTemplate As String
        Public Property format As String = "jsonl"
        Public Property logDryRuns As Boolean = False
    End Class

    Private Class DataCollectorUseCaseConfig
        Public Property id As String
        Public Property enabled As Boolean
        Public Property name As String
        Public Property description As String
        Public Property applicability As DataCollectorApplicabilityConfig
        Public Property extraction As DataCollectorExtractionConfig
        Public Property target As DataCollectorTargetConfig
        Public Property duplicateControl As DataCollectorDuplicateControlConfig
        Public Property csv As DataCollectorCsvConfig
        Public Property jsonl As DataCollectorJsonLinesConfig
        Public Property fields As List(Of DataCollectorFieldConfig)
    End Class

    Private Class DataCollectorApplicabilityConfig
        Public Property modelInstructions As String
        Public Property subjectHints As List(Of String)
        Public Property bodyHints As List(Of String)
    End Class

    Private Class DataCollectorExtractionConfig
        Public Property modelInstructions As String
        Public Property includeOriginalRequest As Boolean
        Public Property originalRequestFieldName As String = "originalRequest"
    End Class

    Private Class DataCollectorTargetConfig
        Public Property directory As String
        Public Property fileNameTemplate As String
        Public Property format As String
        Public Property mode As String
        Public Property encoding As String
        Public Property fileNameMissingFieldPolicy As String = "reject"
        Public Property fileNameMissingFieldFallbackValue As String = "missing"
    End Class

    Private Class DataCollectorDuplicateControlConfig
        Public Property enabled As Boolean
        Public Property policy As String = "allow"
        Public Property recordKey As List(Of String)
        Public Property caseSensitive As Boolean = False
        Public Property trimWhitespace As Boolean = True
        Public Property normalizeBeforeCompare As Boolean = True
        Public Property allowMissingDuplicateKeys As Boolean = False
    End Class

    Private Class DataCollectorCsvConfig
        Public Property delimiter As String = ";"
        Public Property includeHeader As Boolean = True
        Public Property quoteMode As String = "whenNeeded"
        Public Property newline As String = "CRLF"
        Public Property protectAgainstFormulaInjection As Boolean = True
    End Class

    Private Class DataCollectorJsonLinesConfig
        Public Property invalidLinePolicy As String = "reject"
    End Class

    Private Class DataCollectorFieldConfig
        Public Property name As String
        Public Property displayName As String
        Public Property description As String
        Public Property type As String
        Public Property required As Boolean
        Public Property nullable As Boolean = False
        Public Property defaultValue As Object
        Public Property modelExtractionInstructions As String
        Public Property normalization As DataCollectorFieldNormalizationConfig
        Public Property validation As DataCollectorFieldValidationConfig
        Public Property output As JObject
    End Class

    Private Class DataCollectorFieldNormalizationConfig
        Public Property trim As Boolean = False
        Public Property collapseWhitespace As Boolean = False
        Public Property inputFormats As List(Of String)
        Public Property outputFormat As String
        Public Property removeCurrencySymbols As Boolean = False
        Public Property acceptedDecimalSeparators As List(Of String)
        Public Property acceptedThousandsSeparators As List(Of String)
        Public Property outputDecimalSeparator As String = "."
        Public Property uppercase As Boolean = False
        Public Property lowercase As Boolean = False
    End Class

    Private Class DataCollectorFieldValidationConfig
        Public Property min As Decimal?
        Public Property max As Decimal?
        Public Property minDate As String
        Public Property maxDate As String
        Public Property allowFutureDate As Boolean? = Nothing
        Public Property minLength As Integer?
        Public Property maxLength As Integer?
        Public Property regex As String
        Public Property allowedValues As List(Of String)
        Public Property customErrorMessage As String
    End Class

    Private Class DataCollectorCollectionRequest
        Public Property useCaseId As String
        Public Property dryRun As Boolean
        Public Property source As DataCollectorRequestSource
        Public Property values As JObject
        Public Property originalRequest As DataCollectorOriginalRequest
        Public Property modelNotes As List(Of DataCollectorModelNote)
    End Class

    Private Class DataCollectorRequestSource
        Public Property messageId As String
        Public Property threadId As String
        Public Property from As String
        Public Property [to] As List(Of String)
        Public Property cc As List(Of String)
        Public Property subject As String
        Public Property receivedAt As String
    End Class

    Private Class DataCollectorOriginalRequest
        Public Property include As Boolean
        Public Property content As String
    End Class

    Private Class DataCollectorModelNote
        Public Property field As String
        Public Property note As String
    End Class

    Private Class DataCollectorMessageItem
        Public Property code As String
        Public Property message As String
        Public Property field As String
        Public Property originalValue As Object
        Public Property normalizedValue As Object
        Public Property details As Object
        Public Property recordKey As Object
    End Class

    Private Class DataCollectorNormalizationItem
        Public Property field As String
        Public Property originalValue As Object
        Public Property normalizedValue As Object
        Public Property normalizationApplied As Boolean
    End Class

    Private Class DataCollectorFieldEvaluation
        Public Property definition As DataCollectorFieldConfig
        Public Property originalToken As JToken
        Public Property normalizedValue As Object
        Public Property outputToken As JToken
        Public Property normalizationApplied As Boolean
    End Class

    Private Class DataCollectorDuplicateCheckResult
        Public Property enabled As Boolean
        Public Property duplicateFound As Boolean
        Public Property existingRecord As JObject
        Public Property existingIndex As Integer = -1
        Public Property recordKey As JObject = New JObject()
        Public Property warnings As New List(Of DataCollectorMessageItem)
        Public Property blockingErrors As New List(Of DataCollectorMessageItem)
    End Class

    Private Class DataCollectorWriteResult
        Public Property success As Boolean = True
        Public Property status As String = AP_DC_Status_Written
        Public Property warnings As New List(Of DataCollectorMessageItem)
        Public Property errors As New List(Of DataCollectorMessageItem)
    End Class

    Private Class DataCollectorOperationResult
        Public Property success As Boolean
        Public Property responseJson As String
        Public Property errorMessage As String
        Public Property errorCode As String
    End Class

    Private Class DataCollectorReadResult
        Public Property success As Boolean = True
        Public Property records As New List(Of JObject)
        Public Property warnings As New List(Of DataCollectorMessageItem)
        Public Property errors As New List(Of DataCollectorMessageItem)
    End Class

    Private NotInheritable Class CsvOutputWriter
        Implements IDataCollectorOutputWriter

        Private ReadOnly _owner As ThisAddIn

        Public Sub New(owner As ThisAddIn)
            _owner = owner
        End Sub

        Public Function Preview(record As JObject, useCase As DataCollectorUseCaseConfig) As JObject Implements IDataCollectorOutputWriter.Preview
            Return _owner.BuildCsvPreview(record, useCase)
        End Function

        Public Function Write(targetPath As String,
                              record As JObject,
                              useCase As DataCollectorUseCaseConfig,
                              config As DataCollectorConfiguration,
                              effectiveWriteMode As String,
                              duplicateResult As DataCollectorDuplicateCheckResult) As DataCollectorWriteResult Implements IDataCollectorOutputWriter.Write
            Return _owner.WriteCsvRecordCore(targetPath, record, useCase, config, effectiveWriteMode, duplicateResult)
        End Function
    End Class

    Private NotInheritable Class JsonArrayOutputWriter
        Implements IDataCollectorOutputWriter

        Private ReadOnly _owner As ThisAddIn

        Public Sub New(owner As ThisAddIn)
            _owner = owner
        End Sub

        Public Function Preview(record As JObject, useCase As DataCollectorUseCaseConfig) As JObject Implements IDataCollectorOutputWriter.Preview
            Return New JObject(
                New JProperty("format", "json"),
                New JProperty("record", record.DeepClone()))
        End Function

        Public Function Write(targetPath As String,
                              record As JObject,
                              useCase As DataCollectorUseCaseConfig,
                              config As DataCollectorConfiguration,
                              effectiveWriteMode As String,
                              duplicateResult As DataCollectorDuplicateCheckResult) As DataCollectorWriteResult Implements IDataCollectorOutputWriter.Write
            Return _owner.WriteJsonArrayRecordCore(targetPath, record, useCase, config, effectiveWriteMode, duplicateResult)
        End Function
    End Class

    Private NotInheritable Class JsonLinesOutputWriter
        Implements IDataCollectorOutputWriter

        Private ReadOnly _owner As ThisAddIn

        Public Sub New(owner As ThisAddIn)
            _owner = owner
        End Sub

        Public Function Preview(record As JObject, useCase As DataCollectorUseCaseConfig) As JObject Implements IDataCollectorOutputWriter.Preview
            Return New JObject(
                New JProperty("format", "jsonl"),
                New JProperty("record", record.DeepClone()))
        End Function

        Public Function Write(targetPath As String,
                              record As JObject,
                              useCase As DataCollectorUseCaseConfig,
                              config As DataCollectorConfiguration,
                              effectiveWriteMode As String,
                              duplicateResult As DataCollectorDuplicateCheckResult) As DataCollectorWriteResult Implements IDataCollectorOutputWriter.Write
            Return _owner.WriteJsonLinesRecordCore(targetPath, record, useCase, config, effectiveWriteMode, duplicateResult)
        End Function
    End Class

    Private Function BuildListCollectionUseCasesTool() As ModelConfig
        Return New ModelConfig() With {
            .ToolOnly = True,
            .Tool = True,
            .ToolName = AP_Tool_ListCollectionUseCases,
            .ModelDescription = "List configured collection use cases",
            .ToolInstructionsPrompt =
                AP_Tool_ListCollectionUseCases & ": Lists all enabled data-collection use cases from the DataCollector JSON configuration. " &
                "Use this first to discover available use cases, applicability rules, field extraction instructions, required fields, target format, write mode, and duplicate policy.",
            .ToolDefinition = BuildDataCollectorToolDefinition(
                AP_Tool_ListCollectionUseCases,
                "Lists enabled data-collection use cases defined by the DataCollector configuration.",
                New JObject(
                    New JProperty("type", "object"),
                    New JProperty("properties", New JObject()),
                    New JProperty("required", New JArray())))
        }
    End Function

    Private Function BuildCollectDataTool() As ModelConfig
        Return New ModelConfig() With {
            .ToolOnly = True,
            .Tool = True,
            .ToolName = AP_Tool_CollectData,
            .ModelDescription = "Collect structured data into configured target files",
            .ToolInstructionsPrompt =
                AP_Tool_CollectData & ": Validates, normalizes, duplicate-checks, writes, and audit-logs structured values for a configured use case. " &
                "The model must never provide directory, path, filename, or format overrides. Those are resolved exclusively from configuration.",
            .ToolDefinition = BuildDataCollectorToolDefinition(
                AP_Tool_CollectData,
                "Collects validated structured data into configuration-bound CSV, JSON, or JSONL targets.",
                BuildDataCollectorRequestParameters(includeDryRun:=True))
        }
    End Function

    Private Function BuildPreviewCollectionTool() As ModelConfig
        Return New ModelConfig() With {
            .ToolOnly = True,
            .Tool = True,
            .ToolName = AP_Tool_PreviewCollection,
            .ModelDescription = "Preview structured collection without writing",
            .ToolInstructionsPrompt =
                AP_Tool_PreviewCollection & ": Runs the full collection pipeline except the final data write. " &
                "Use this to preview resolved target path, duplicate handling, validation, normalization, and output shape before persisting anything.",
            .ToolDefinition = BuildDataCollectorToolDefinition(
                AP_Tool_PreviewCollection,
                "Previews a configured structured data collection write without writing the target data file.",
                BuildDataCollectorRequestParameters(includeDryRun:=False))
        }
    End Function

    Private Shared Function BuildDataCollectorToolDefinition(toolName As String,
                                                             description As String,
                                                             parameters As JObject) As String
        Return New JObject(
            New JProperty("name", toolName),
            New JProperty("description", description),
            New JProperty("parameters", parameters)).
            ToString(Newtonsoft.Json.Formatting.None)
    End Function

    Private Shared Function BuildDataCollectorRequestParameters(includeDryRun As Boolean) As JObject
        Dim properties As New JObject(
            New JProperty("useCaseId",
                New JObject(
                    New JProperty("type", "string"),
                    New JProperty("description", "Configured use case id."))),
            New JProperty("source",
                New JObject(
                    New JProperty("type", "object"),
                    New JProperty("properties",
                        New JObject(
                            New JProperty("messageId", New JObject(New JProperty("type", "string"))),
                            New JProperty("threadId", New JObject(New JProperty("type", "string"))),
                            New JProperty("from", New JObject(New JProperty("type", "string"))),
                            New JProperty("to",
                                New JObject(
                                    New JProperty("type", "array"),
                                    New JProperty("items", New JObject(New JProperty("type", "string"))))),
                            New JProperty("cc",
                                New JObject(
                                    New JProperty("type", "array"),
                                    New JProperty("items", New JObject(New JProperty("type", "string"))))),
                            New JProperty("subject", New JObject(New JProperty("type", "string"))),
                            New JProperty("receivedAt", New JObject(New JProperty("type", "string"))))))),
            New JProperty("values",
                New JObject(
                    New JProperty("type", "object"),
                    New JProperty("additionalProperties", True),
                    New JProperty("description", "Structured values extracted by the model. Keys must match configured field names."))),
            New JProperty("originalRequest",
                New JObject(
                    New JProperty("type", "object"),
                    New JProperty("properties",
                        New JObject(
                            New JProperty("include", New JObject(New JProperty("type", "boolean"))),
                            New JProperty("content", New JObject(New JProperty("type", "string"))))))),
            New JProperty("modelNotes",
                New JObject(
                    New JProperty("type", "array"),
                    New JProperty("items",
                        New JObject(
                            New JProperty("type", "object"),
                            New JProperty("properties",
                                New JObject(
                                    New JProperty("field", New JObject(New JProperty("type", "string"))),
                                    New JProperty("note", New JObject(New JProperty("type", "string"))))))))))

        If includeDryRun Then
            properties.Add(
                New JProperty("dryRun",
                    New JObject(
                        New JProperty("type", "boolean"),
                        New JProperty("description", "If true, perform preview behavior without writing."))))
        End If

        Return New JObject(
            New JProperty("type", "object"),
            New JProperty("properties", properties),
            New JProperty("required", New JArray("useCaseId", "values")))
    End Function

    Private Function IsDataCollectorToolAvailable() As Boolean
        Dim path As String = GetDataCollectorConfigurationPath()
        If String.IsNullOrWhiteSpace(path) Then
            Return False
        End If

        Dim config As DataCollectorConfiguration = Nothing
        Dim errors As List(Of DataCollectorMessageItem) = Nothing

        If Not TryLoadValidatedDataCollectorConfiguration(config, errors) Then
            Return False
        End If

        Return config IsNot Nothing AndAlso
               config.enabled AndAlso
               config.useCases IsNot Nothing AndAlso
               config.useCases.Any(Function(x) x IsNot Nothing AndAlso x.enabled)
    End Function

    Private Function ExecuteListCollectionUseCasesTool(toolCall As ToolCall,
                                                       context As ToolExecutionContext,
                                                       ct As CancellationToken) As Task(Of ToolResponse)
        Dim result As DataCollectorOperationResult = ExecuteListCollectionUseCasesCore(context)
        Return Task.FromResult(CreateDataCollectorToolResponse(toolCall, result))
    End Function

    Private Function ExecuteCollectDataTool(toolCall As ToolCall,
                                            context As ToolExecutionContext,
                                            ct As CancellationToken) As Task(Of ToolResponse)
        Dim request As DataCollectorCollectionRequest = Nothing
        Dim parseResult As DataCollectorOperationResult = TryParseDataCollectorRequest(toolCall.Arguments, request)

        If Not parseResult.success Then
            Return Task.FromResult(CreateDataCollectorToolResponse(toolCall, parseResult))
        End If

        Dim result As DataCollectorOperationResult =
            ExecuteDataCollectorRequestCore(request, forceDryRun:=False, invokedViaPreviewTool:=False, context:=context)

        Return Task.FromResult(CreateDataCollectorToolResponse(toolCall, result))
    End Function

    Private Function ExecutePreviewCollectionTool(toolCall As ToolCall,
                                                  context As ToolExecutionContext,
                                                  ct As CancellationToken) As Task(Of ToolResponse)
        Dim request As DataCollectorCollectionRequest = Nothing
        Dim parseResult As DataCollectorOperationResult = TryParseDataCollectorRequest(toolCall.Arguments, request)

        If Not parseResult.success Then
            Return Task.FromResult(CreateDataCollectorToolResponse(toolCall, parseResult))
        End If

        Dim result As DataCollectorOperationResult =
            ExecuteDataCollectorRequestCore(request, forceDryRun:=True, invokedViaPreviewTool:=True, context:=context)

        Return Task.FromResult(CreateDataCollectorToolResponse(toolCall, result))
    End Function

    Private Function CreateDataCollectorToolResponse(toolCall As ToolCall,
                                                     result As DataCollectorOperationResult) As ToolResponse
        Return New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName,
            .Success = result.success,
            .Response = result.responseJson,
            .ErrorMessage = result.errorMessage,
            .ErrorCode = result.errorCode,
            .Timestamp = DateTime.UtcNow
        }
    End Function

    Private Function ExecuteListCollectionUseCasesCore(context As ToolExecutionContext) As DataCollectorOperationResult
        Dim config As DataCollectorConfiguration = Nothing
        Dim errors As List(Of DataCollectorMessageItem) = Nothing

        If Not TryLoadValidatedDataCollectorConfiguration(config, errors) Then
            Return BuildOperationResultFailure(AP_DC_Status_ConfigurationInvalid, errors)
        End If

        If context IsNot Nothing Then
            context.Log("DataCollector: listing configured use cases.")
        End If

        Dim payload As New JObject(
            New JProperty("success", True),
            New JProperty("status", AP_DC_Status_Success),
            New JProperty("schemaVersion", If(config.schemaVersion, "")),
            New JProperty("configVersion", If(config.configVersion, "")),
            New JProperty("useCases", BuildModelUseCaseDescriptorArray(config)))

        Return New DataCollectorOperationResult With {
            .success = True,
            .responseJson = SerializeJson(payload)
        }
    End Function

    Private Function ExecuteDataCollectorRequestCore(request As DataCollectorCollectionRequest,
                                                     forceDryRun As Boolean,
                                                     invokedViaPreviewTool As Boolean,
                                                     context As ToolExecutionContext) As DataCollectorOperationResult
        Dim config As DataCollectorConfiguration = Nothing
        Dim configErrors As List(Of DataCollectorMessageItem) = Nothing

        If Not TryLoadValidatedDataCollectorConfiguration(config, configErrors) Then
            Return BuildOperationResultFailure(AP_DC_Status_ConfigurationInvalid, configErrors)
        End If

        If request Is Nothing Then
            Return BuildOperationResultFailure(
                AP_DC_Status_ValidationFailed,
                New List(Of DataCollectorMessageItem) From {
                    CreateMessage("REQUEST_INVALID", "The request payload is missing or invalid.")
                })
        End If

        If String.IsNullOrWhiteSpace(request.useCaseId) Then
            Return BuildOperationResultFailure(
                AP_DC_Status_ValidationFailed,
                New List(Of DataCollectorMessageItem) From {
                    CreateMessage("MISSING_USE_CASE_ID", "The field 'useCaseId' is required.")
                })
        End If

        Dim useCase = config.useCases.
            FirstOrDefault(Function(x) x IsNot Nothing AndAlso
                                       x.enabled AndAlso
                                       String.Equals(x.id, request.useCaseId, StringComparison.OrdinalIgnoreCase))

        If useCase Is Nothing Then
            Return BuildOperationResultFailure(
                AP_DC_Status_UseCaseNotFound,
                New List(Of DataCollectorMessageItem) From {
                    CreateMessage("USE_CASE_NOT_FOUND", $"The use case '{request.useCaseId}' was not found or is disabled.")
                })
        End If

        request.source = MergeRequestSource(request.source)
        If request.values Is Nothing Then
            request.values = New JObject()
        End If
        If request.originalRequest Is Nothing Then
            request.originalRequest = New DataCollectorOriginalRequest()
        End If
        If request.modelNotes Is Nothing Then
            request.modelNotes = New List(Of DataCollectorModelNote)()
        End If

        Dim dryRun As Boolean = forceDryRun OrElse request.dryRun
        Dim warnings As New List(Of DataCollectorMessageItem)()
        Dim errors As New List(Of DataCollectorMessageItem)()
        Dim normalization As New List(Of DataCollectorNormalizationItem)()

        If context IsNot Nothing Then
            context.Log($"DataCollector: useCase={useCase.id}, dryRun={dryRun}")
        End If
        ApDashboardLog($"🗂 DataCollector: {useCase.id}", "step")

        Dim evaluations As Dictionary(Of String, DataCollectorFieldEvaluation) =
            NormalizeAndValidateFields(config, useCase, request, normalization, errors)

        Dim shouldWriteAudit As Boolean =
            config.log IsNot Nothing AndAlso
            config.log.enabled AndAlso
            ((Not dryRun AndAlso Not invokedViaPreviewTool) OrElse
             (dryRun AndAlso config.log.logDryRuns))

        If errors.Count > 0 Then
            Dim failStatus As String =
                If(errors.Any(Function(x) String.Equals(x.code, "NORMALIZATION_FAILED", StringComparison.OrdinalIgnoreCase)),
                   AP_DC_Status_NormalizationFailed,
                   AP_DC_Status_ValidationFailed)

            TryWriteAuditLogSafe(
                config, useCase, request, failStatus, If(useCase.target?.mode, ""),
                Nothing,
                BuildOriginalValuesObject(evaluations),
                BuildNormalizedValuesObject(evaluations, useCase, request),
                Nothing,
                warnings,
                errors,
                dryRun,
                shouldWriteAudit,
                warnings)

            Return BuildOperationResultFailure(failStatus, errors, warnings, normalization, useCase.id)
        End If

        Dim record As JObject = BuildOutputRecord(useCase, request, evaluations, warnings, errors)
        If errors.Count > 0 Then
            TryWriteAuditLogSafe(
                config, useCase, request, AP_DC_Status_ValidationFailed, If(useCase.target?.mode, ""),
                Nothing,
                BuildOriginalValuesObject(evaluations),
                record,
                Nothing,
                warnings,
                errors,
                dryRun,
                shouldWriteAudit,
                warnings)

            Return BuildOperationResultFailure(AP_DC_Status_ValidationFailed, errors, warnings, normalization, useCase.id)
        End If

        Dim targetPath As String = Nothing
        If Not TryResolveTargetPath(config, useCase, request, record, targetPath, errors) Then
            TryWriteAuditLogSafe(
                config, useCase, request, AP_DC_Status_TargetPathNotAllowed, If(useCase.target?.mode, ""),
                Nothing,
                BuildOriginalValuesObject(evaluations),
                record,
                Nothing,
                warnings,
                errors,
                dryRun,
                shouldWriteAudit,
                warnings)

            Return BuildOperationResultFailure(AP_DC_Status_TargetPathNotAllowed, errors, warnings, normalization, useCase.id)
        End If

        Dim duplicateResult As DataCollectorDuplicateCheckResult =
            CheckForDuplicate(config, useCase, targetPath, record)

        For Each warn In duplicateResult.warnings
            warnings.Add(warn)
        Next
        For Each blockErr In duplicateResult.blockingErrors
            errors.Add(blockErr)
        Next

        If errors.Count > 0 Then
            TryWriteAuditLogSafe(
                config, useCase, request, AP_DC_Status_ValidationFailed, If(useCase.target?.mode, ""),
                targetPath,
                BuildOriginalValuesObject(evaluations),
                record,
                duplicateResult,
                warnings,
                errors,
                dryRun,
                shouldWriteAudit,
                warnings)

            Return BuildOperationResultFailure(AP_DC_Status_ValidationFailed, errors, warnings, normalization, useCase.id)
        End If

        Dim effectiveWriteMode As String = If(useCase.target?.mode, "appendOrCreate")
        Dim duplicatePolicy As String = GetDuplicatePolicy(useCase)

        If duplicateResult.enabled AndAlso duplicateResult.duplicateFound Then
            Select Case duplicatePolicy
                Case "reject"
                    errors.Add(CreateMessage(
                        "DUPLICATE_RECORD",
                        "A record with the same duplicate key already exists.",
                        details:=Nothing,
                        recordKey:=TokenToPlainObject(duplicateResult.recordKey)))

                    TryWriteAuditLogSafe(
                        config, useCase, request, AP_DC_Status_DuplicateDetected, effectiveWriteMode,
                        targetPath,
                        BuildOriginalValuesObject(evaluations),
                        record,
                        duplicateResult,
                        warnings,
                        errors,
                        dryRun,
                        shouldWriteAudit,
                        warnings)

                    Return BuildOperationResultFailure(AP_DC_Status_DuplicateDetected, errors, warnings, normalization, useCase.id)

                Case "ignore"
                    Dim obj As New JObject(
                        New JProperty("success", True),
                        New JProperty("status", AP_DC_Status_DuplicateIgnored),
                        New JProperty("useCaseId", useCase.id),
                        New JProperty("operation", effectiveWriteMode),
                        New JProperty("target", New JObject(
                            New JProperty("format", useCase.target.format),
                            New JProperty("path", targetPath))),
                        New JProperty("recordKey", duplicateResult.recordKey.DeepClone()),
                        New JProperty("normalization", JArray.FromObject(normalization)),
                        New JProperty("warnings", JArray.FromObject(warnings)))

                    TryWriteAuditLogSafe(
                        config, useCase, request, AP_DC_Status_DuplicateIgnored, effectiveWriteMode,
                        targetPath,
                        BuildOriginalValuesObject(evaluations),
                        record,
                        duplicateResult,
                        warnings,
                        New List(Of DataCollectorMessageItem)(),
                        dryRun,
                        shouldWriteAudit,
                        warnings)

                    Return New DataCollectorOperationResult With {
                        .success = True,
                        .responseJson = SerializeJson(obj)
                    }

                Case "upsert"
                    effectiveWriteMode = "upsert"

                Case "appendWithWarning"
                    warnings.Add(CreateMessage(
                        "DUPLICATE_APPEND_WITH_WARNING",
                        "A duplicate record exists, but the configured policy allows appending with warning.",
                        recordKey:=TokenToPlainObject(duplicateResult.recordKey)))

                Case Else
                    ' allow
            End Select
        End If

        Dim writer As IDataCollectorOutputWriter = CreateDataCollectorOutputWriter(useCase.target.format)
        If writer Is Nothing Then
            Return BuildOperationResultFailure(
                AP_DC_Status_UnsupportedFormat,
                New List(Of DataCollectorMessageItem) From {
                    CreateMessage("UNSUPPORTED_FORMAT", $"The target format '{useCase.target.format}' is not supported.")
                },
                warnings,
                normalization,
                useCase.id)
        End If

        If dryRun Then
            Dim preview As JObject = writer.Preview(record, useCase)
            preview("operation") = effectiveWriteMode
            preview("record") = record.DeepClone()

            Dim previewObj As New JObject(
                New JProperty("success", True),
                New JProperty("status", AP_DC_Status_Preview),
                New JProperty("useCaseId", useCase.id),
                New JProperty("wouldWrite", True),
                New JProperty("target", New JObject(
                    New JProperty("format", useCase.target.format),
                    New JProperty("path", targetPath))),
                New JProperty("recordKey", duplicateResult.recordKey.DeepClone()),
                New JProperty("preview", preview),
                New JProperty("normalization", JArray.FromObject(normalization)),
                New JProperty("warnings", JArray.FromObject(warnings)))

            TryWriteAuditLogSafe(
                config, useCase, request, AP_DC_Status_Preview, effectiveWriteMode,
                targetPath,
                BuildOriginalValuesObject(evaluations),
                record,
                duplicateResult,
                warnings,
                New List(Of DataCollectorMessageItem)(),
                dryRun,
                shouldWriteAudit,
                warnings)

            Return New DataCollectorOperationResult With {
                .success = True,
                .responseJson = SerializeJson(previewObj)
            }
        End If

        Dim writeResult As DataCollectorWriteResult =
            writer.Write(targetPath, record, useCase, config, effectiveWriteMode, duplicateResult)

        For Each warn In writeResult.warnings
            warnings.Add(warn)
        Next
        For Each errItem In writeResult.errors
            errors.Add(errItem)
        Next

        If Not writeResult.success OrElse errors.Count > 0 Then
            Dim failStatus As String =
                If(errors.Any(Function(x) String.Equals(x.code, "TARGET_FILE_INVALID", StringComparison.OrdinalIgnoreCase)),
                   AP_DC_Status_TargetFileInvalid,
                   If(errors.Any(Function(x) String.Equals(x.code, "TARGET_FILE_LOCKED", StringComparison.OrdinalIgnoreCase)),
                      AP_DC_Status_TargetFileLocked,
                      AP_DC_Status_WriteFailed))

            TryWriteAuditLogSafe(
                config, useCase, request, failStatus, effectiveWriteMode,
                targetPath,
                BuildOriginalValuesObject(evaluations),
                record,
                duplicateResult,
                warnings,
                errors,
                dryRun,
                shouldWriteAudit,
                warnings)

            Return BuildOperationResultFailure(failStatus, errors, warnings, normalization, useCase.id)
        End If

        Dim logWarnings As New List(Of DataCollectorMessageItem)()
        TryWriteAuditLogSafe(
            config, useCase, request, writeResult.status, effectiveWriteMode,
            targetPath,
            BuildOriginalValuesObject(evaluations),
            record,
            duplicateResult,
            warnings,
            New List(Of DataCollectorMessageItem)(),
            dryRun,
            shouldWriteAudit,
            logWarnings)

        For Each warn In logWarnings
            warnings.Add(warn)
        Next

        Dim successObj As New JObject(
            New JProperty("success", True),
            New JProperty("status", writeResult.status),
            New JProperty("useCaseId", useCase.id),
            New JProperty("operation", effectiveWriteMode),
            New JProperty("target", New JObject(
                New JProperty("format", useCase.target.format),
                New JProperty("path", targetPath))),
            New JProperty("recordKey", duplicateResult.recordKey.DeepClone()),
            New JProperty("normalization", JArray.FromObject(normalization)),
            New JProperty("warnings", JArray.FromObject(warnings)))

        Return New DataCollectorOperationResult With {
            .success = True,
            .responseJson = SerializeJson(successObj)
        }
    End Function

    Private Function TryParseDataCollectorRequest(arguments As IDictionary(Of String, Object),
                                                  ByRef request As DataCollectorCollectionRequest) As DataCollectorOperationResult
        Try
            Dim root As JObject = If(arguments IsNot Nothing, JObject.FromObject(arguments), New JObject())
            request = root.ToObject(Of DataCollectorCollectionRequest)()

            If request Is Nothing Then
                request = New DataCollectorCollectionRequest()
            End If
            If request.values Is Nothing Then
                request.values = New JObject()
            End If
            If request.source Is Nothing Then
                request.source = New DataCollectorRequestSource()
            End If
            If request.originalRequest Is Nothing Then
                request.originalRequest = New DataCollectorOriginalRequest()
            End If
            If request.modelNotes Is Nothing Then
                request.modelNotes = New List(Of DataCollectorModelNote)()
            End If

            Return New DataCollectorOperationResult With {.success = True}
        Catch ex As Exception
            Return BuildOperationResultFailure(
                AP_DC_Status_ValidationFailed,
                New List(Of DataCollectorMessageItem) From {
                    CreateMessage("REQUEST_INVALID", $"The request payload could not be parsed: {ex.Message}")
                })
        End Try
    End Function

    Private Function TryLoadValidatedDataCollectorConfiguration(ByRef config As DataCollectorConfiguration,
                                                                ByRef errors As List(Of DataCollectorMessageItem)) As Boolean
        errors = New List(Of DataCollectorMessageItem)()

        Dim path As String = GetDataCollectorConfigurationPath()
        If String.IsNullOrWhiteSpace(path) Then
            errors.Add(CreateMessage("CONFIG_PATH_MISSING", "No DataCollector configuration path is configured."))
            Return False
        End If

        If Not File.Exists(path) Then
            errors.Add(CreateMessage("CONFIG_FILE_NOT_FOUND", $"The DataCollector configuration file was not found: {path}"))
            Return False
        End If

        Try
            Dim json As String = File.ReadAllText(path, DetectEncoding("utf-8"))
            config = JsonConvert.DeserializeObject(Of DataCollectorConfiguration)(json)
        Catch ex As Exception
            errors.Add(CreateMessage("CONFIG_FILE_INVALID", $"The DataCollector configuration file could not be read: {ex.Message}"))
            Return False
        End Try

        ValidateDataCollectorConfiguration(config, errors)
        Return errors.Count = 0
    End Function

    Private Function GetDataCollectorConfigurationPath() As String
        Dim rawPath As String = ExpandEnvironmentVariables(INI_DataCollectorPath)

        If String.IsNullOrWhiteSpace(rawPath) Then
            rawPath = AP_DataCollectorConfigPathPlaceholder
        End If

        Return ExpandEnvironmentVariables(rawPath)
    End Function

    Private Sub ValidateDataCollectorConfiguration(config As DataCollectorConfiguration,
                                                   errors As List(Of DataCollectorMessageItem))
        If config Is Nothing Then
            errors.Add(CreateMessage("CONFIG_INVALID", "The DataCollector configuration is missing or invalid."))
            Exit Sub
        End If

        If Not String.Equals(If(config.schemaVersion, "").Trim(), AP_DataCollectorSchemaVersion, StringComparison.Ordinal) Then
            errors.Add(CreateMessage("SCHEMA_VERSION_INVALID", $"Unsupported schemaVersion '{If(config.schemaVersion, "")}'. Expected '{AP_DataCollectorSchemaVersion}'."))
        End If

        If Not config.enabled Then
            errors.Add(CreateMessage("CONFIG_DISABLED", "The DataCollector configuration is disabled."))
        End If

        Dim baseDirectory As String = NormalizeFullPath(ExpandEnvironmentVariables(If(config.allowedBaseDirectory, "")))
        If String.IsNullOrWhiteSpace(baseDirectory) Then
            errors.Add(CreateMessage("ALLOWED_BASE_DIRECTORY_MISSING", "The field 'allowedBaseDirectory' is required."))
        ElseIf Not Path.IsPathRooted(baseDirectory) Then
            errors.Add(CreateMessage("ALLOWED_BASE_DIRECTORY_INVALID", "The field 'allowedBaseDirectory' must be an absolute path."))
        End If

        If config.defaults Is Nothing Then
            config.defaults = New DataCollectorDefaultsConfig()
        End If
        If config.log Is Nothing Then
            config.log = New DataCollectorLogConfig()
        End If
        If config.useCases Is Nothing Then
            config.useCases = New List(Of DataCollectorUseCaseConfig)()
        End If

        If config.log.enabled Then
            ValidateDirectorySettingUnderBase(baseDirectory, config.log.directory, "log.directory", errors)
            ValidateFileNameTemplate(config.log.fileNameTemplate, "log.fileNameTemplate", Nothing, errors)

            If Not String.Equals(If(config.log.format, "jsonl"), "jsonl", StringComparison.OrdinalIgnoreCase) Then
                errors.Add(CreateMessage("LOG_FORMAT_INVALID", "Only 'jsonl' audit log format is supported."))
            End If
        End If

        Dim seenUseCaseIds As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each useCase In config.useCases
            If useCase Is Nothing Then
                errors.Add(CreateMessage("USE_CASE_INVALID", "A use case entry is null."))
                Continue For
            End If

            If String.IsNullOrWhiteSpace(useCase.id) Then
                errors.Add(CreateMessage("USE_CASE_ID_MISSING", "Each use case must define an id."))
            ElseIf Not seenUseCaseIds.Add(useCase.id.Trim()) Then
                errors.Add(CreateMessage("USE_CASE_ID_DUPLICATE", $"Duplicate use case id '{useCase.id}'."))
            End If

            If useCase.extraction Is Nothing Then
                useCase.extraction = New DataCollectorExtractionConfig()
            End If
            If useCase.duplicateControl Is Nothing Then
                useCase.duplicateControl = New DataCollectorDuplicateControlConfig()
            End If

            If useCase.target Is Nothing Then
                errors.Add(CreateMessage("USE_CASE_TARGET_MISSING", $"Use case '{useCase.id}' is missing a target definition."))
            Else
                ValidateDirectorySettingUnderBase(baseDirectory, useCase.target.directory, $"useCases[{useCase.id}].target.directory", errors)
                ValidateFileNameTemplate(useCase.target.fileNameTemplate, $"useCases[{useCase.id}].target.fileNameTemplate", useCase, errors)

                Dim fmt As String = If(useCase.target.format, "").Trim().ToLowerInvariant()
                If fmt <> "csv" AndAlso fmt <> "json" AndAlso fmt <> "jsonl" Then
                    errors.Add(CreateMessage("TARGET_FORMAT_INVALID", $"Use case '{useCase.id}' has unsupported target format '{useCase.target.format}'."))
                End If

                If Not IsSupportedWriteMode(If(useCase.target.mode, "").Trim()) Then
                    errors.Add(CreateMessage("WRITE_MODE_INVALID", $"Use case '{useCase.id}' has unsupported write mode '{useCase.target.mode}'."))
                End If
            End If

            If useCase.fields Is Nothing OrElse useCase.fields.Count = 0 Then
                errors.Add(CreateMessage("FIELDS_MISSING", $"Use case '{useCase.id}' must define at least one field."))
            Else
                Dim seenFieldNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                For Each field In useCase.fields
                    If field Is Nothing Then
                        errors.Add(CreateMessage("FIELD_INVALID", $"Use case '{useCase.id}' contains a null field definition."))
                        Continue For
                    End If

                    If String.IsNullOrWhiteSpace(field.name) Then
                        errors.Add(CreateMessage("FIELD_NAME_MISSING", $"Use case '{useCase.id}' contains a field without a name."))
                    ElseIf Not seenFieldNames.Add(field.name.Trim()) Then
                        errors.Add(CreateMessage("FIELD_NAME_DUPLICATE", $"Use case '{useCase.id}' has duplicate field name '{field.name}'."))
                    End If

                    Select Case If(field.type, "").Trim().ToLowerInvariant()
                        Case "string", "integer", "decimal", "boolean", "date", "datetime", "email", "uri", "enum"
                        Case Else
                            errors.Add(CreateMessage("FIELD_TYPE_INVALID", $"Use case '{useCase.id}' field '{field.name}' has unsupported type '{field.type}'."))
                    End Select
                Next

                If useCase.extraction.includeOriginalRequest Then
                    Dim originalFieldName As String = If(useCase.extraction.originalRequestFieldName, "originalRequest").Trim()
                    If seenFieldNames.Contains(originalFieldName) Then
                        errors.Add(CreateMessage("FIELD_NAME_DUPLICATE", $"Use case '{useCase.id}' originalRequestFieldName '{originalFieldName}' conflicts with a configured field name."))
                    End If
                End If
            End If

            Dim duplicatePolicy As String = GetDuplicatePolicy(useCase)
            Select Case duplicatePolicy
                Case "allow", "reject", "ignore", "upsert", "appendWithWarning"
                Case Else
                    errors.Add(CreateMessage("DUPLICATE_POLICY_INVALID", $"Use case '{useCase.id}' has unsupported duplicate policy '{useCase.duplicateControl.policy}'."))
            End Select

            If useCase.duplicateControl.recordKey IsNot Nothing AndAlso useCase.fields IsNot Nothing Then
                For Each keyField In useCase.duplicateControl.recordKey
                    If Not useCase.fields.Any(Function(f) f IsNot Nothing AndAlso String.Equals(f.name, keyField, StringComparison.OrdinalIgnoreCase)) Then
                        errors.Add(CreateMessage("RECORD_KEY_INVALID", $"Use case '{useCase.id}' recordKey references unknown field '{keyField}'."))
                    End If
                Next
            End If
        Next
    End Sub

    Private Sub ValidateDirectorySettingUnderBase(baseDirectory As String,
                                                  configuredDirectory As String,
                                                  settingName As String,
                                                  errors As List(Of DataCollectorMessageItem))
        If String.IsNullOrWhiteSpace(configuredDirectory) Then
            errors.Add(CreateMessage("DIRECTORY_MISSING", $"The setting '{settingName}' is required."))
            Exit Sub
        End If

        If Path.IsPathRooted(configuredDirectory) Then
            errors.Add(CreateMessage("UNSAFE_DIRECTORY", $"The setting '{settingName}' must be relative to allowedBaseDirectory."))
            Exit Sub
        End If

        Try
            Dim combined As String = NormalizeFullPath(Path.Combine(baseDirectory, configuredDirectory))
            If Not IsPathWithinBaseDirectory(combined, baseDirectory) Then
                errors.Add(CreateMessage("UNSAFE_DIRECTORY", $"The setting '{settingName}' resolves outside allowedBaseDirectory."))
            End If
        Catch ex As Exception
            errors.Add(CreateMessage("UNSAFE_DIRECTORY", $"The setting '{settingName}' is invalid: {ex.Message}"))
        End Try
    End Sub

    Private Sub ValidateFileNameTemplate(template As String,
                                         settingName As String,
                                         useCase As DataCollectorUseCaseConfig,
                                         errors As List(Of DataCollectorMessageItem))
        If String.IsNullOrWhiteSpace(template) Then
            errors.Add(CreateMessage("FILENAME_TEMPLATE_MISSING", $"The setting '{settingName}' is required."))
            Exit Sub
        End If

        If template.Contains(Path.DirectorySeparatorChar) OrElse template.Contains(Path.AltDirectorySeparatorChar) Then
            errors.Add(CreateMessage("FILENAME_TEMPLATE_INVALID", $"The setting '{settingName}' must not contain directory separators."))
            Exit Sub
        End If

        For Each match As Match In Regex.Matches(template, "\{([^{}]+)\}")
            Dim token As String = match.Groups(1).Value.Trim()

            If Not IsSupportedFileNamePlaceholder(token) Then
                errors.Add(CreateMessage("FILENAME_TEMPLATE_INVALID", $"The setting '{settingName}' contains unsupported placeholder '{token}'."))
            End If

            If token.StartsWith("field:", StringComparison.OrdinalIgnoreCase) AndAlso useCase IsNot Nothing Then
                Dim fieldName As String = token.Substring(6).Trim()
                If useCase.fields Is Nothing OrElse Not useCase.fields.Any(Function(f) f IsNot Nothing AndAlso String.Equals(f.name, fieldName, StringComparison.OrdinalIgnoreCase)) Then
                    errors.Add(CreateMessage("FILENAME_TEMPLATE_INVALID", $"The setting '{settingName}' references unknown field placeholder '{fieldName}'."))
                End If
            End If
        Next
    End Sub

    Private Shared Function IsSupportedFileNamePlaceholder(token As String) As Boolean
        Select Case token
            Case "yyyy", "MM", "dd", "HH", "mm", "ss", "fff", "useCaseId", "messageId", "senderDomain"
                Return True
        End Select

        If token.StartsWith("field:", StringComparison.OrdinalIgnoreCase) Then
            Dim name As String = token.Substring(6).Trim()
            Return Regex.IsMatch(name, "^[A-Za-z0-9_]+$")
        End If

        Return False
    End Function

    Private Shared Function IsSupportedWriteMode(mode As String) As Boolean
        Select Case If(mode, "").Trim()
            Case "appendOrCreate", "append", "create", "overwrite", "upsert", "rejectIfExists"
                Return True
            Case Else
                Return False
        End Select
    End Function

    Private Function BuildModelUseCaseDescriptorArray(config As DataCollectorConfiguration) As JArray
        Dim arr As New JArray()

        If config Is Nothing OrElse config.useCases Is Nothing Then
            Return arr
        End If

        For Each useCase In config.useCases.Where(Function(x) x IsNot Nothing AndAlso x.enabled)
            Dim fields As New JArray()

            If useCase.fields IsNot Nothing Then
                For Each field In useCase.fields
                    fields.Add(New JObject(
                        New JProperty("name", If(field.name, "")),
                        New JProperty("displayName", If(field.displayName, "")),
                        New JProperty("description", If(field.description, "")),
                        New JProperty("type", If(field.type, "")),
                        New JProperty("required", field.required),
                        New JProperty("nullable", field.nullable),
                        New JProperty("modelExtractionInstructions", If(field.modelExtractionInstructions, "")),
                        New JProperty("allowedValues",
                            JArray.FromObject(If(field.validation?.allowedValues, New List(Of String)())))))
                Next
            End If

            arr.Add(New JObject(
                New JProperty("id", If(useCase.id, "")),
                New JProperty("name", If(useCase.name, "")),
                New JProperty("description", If(useCase.description, "")),
                New JProperty("applicabilityInstructions", If(useCase.applicability?.modelInstructions, "")),
                New JProperty("subjectHints", JArray.FromObject(If(useCase.applicability?.subjectHints, New List(Of String)()))),
                New JProperty("bodyHints", JArray.FromObject(If(useCase.applicability?.bodyHints, New List(Of String)()))),
                New JProperty("extractionInstructions", If(useCase.extraction?.modelInstructions, "")),
                New JProperty("includeOriginalRequest", If(useCase.extraction IsNot Nothing AndAlso useCase.extraction.includeOriginalRequest, True, False)),
                New JProperty("originalRequestFieldName", If(useCase.extraction?.originalRequestFieldName, "originalRequest")),
                New JProperty("targetFormat", If(useCase.target?.format, "")),
                New JProperty("writeMode", If(useCase.target?.mode, "")),
                New JProperty("duplicatePolicy", GetDuplicatePolicy(useCase)),
                New JProperty("targetDescription", $"{If(useCase.target?.directory, "")}\{If(useCase.target?.fileNameTemplate, "")}"),
                New JProperty("fields", fields)))
        Next

        Return arr
    End Function

    Private Function MergeRequestSource(source As DataCollectorRequestSource) As DataCollectorRequestSource
        If source Is Nothing Then
            source = New DataCollectorRequestSource()
        End If
        If source.to Is Nothing Then
            source.to = New List(Of String)()
        End If
        If source.cc Is Nothing Then
            source.cc = New List(Of String)()
        End If

        If _apCurrentMailInfo IsNot Nothing Then
            If String.IsNullOrWhiteSpace(source.messageId) Then
                source.messageId = _apCurrentMailInfo.EntryID
            End If
            If String.IsNullOrWhiteSpace(source.from) Then
                source.from = _apCurrentMailInfo.SenderEmail
            End If
            If String.IsNullOrWhiteSpace(source.subject) Then
                source.subject = _apCurrentMailInfo.Subject
            End If
            If String.IsNullOrWhiteSpace(source.receivedAt) Then
                source.receivedAt = _apCurrentMailInfo.ReceivedTime.ToString("o", CultureInfo.InvariantCulture)
            End If
        End If

        Return source
    End Function

    Private Function NormalizeAndValidateFields(config As DataCollectorConfiguration,
                                                useCase As DataCollectorUseCaseConfig,
                                                request As DataCollectorCollectionRequest,
                                                normalization As List(Of DataCollectorNormalizationItem),
                                                errors As List(Of DataCollectorMessageItem)) As Dictionary(Of String, DataCollectorFieldEvaluation)
        Dim result As New Dictionary(Of String, DataCollectorFieldEvaluation)(StringComparer.OrdinalIgnoreCase)
        Dim culture As CultureInfo = ResolveCulture(config.defaults)

        For Each field In useCase.fields
            Dim eval As New DataCollectorFieldEvaluation() With {
                .definition = field,
                .originalToken = If(request.values IsNot Nothing, request.values(field.name), Nothing),
                .normalizedValue = Nothing,
                .outputToken = Nothing,
                .normalizationApplied = False
            }

            If eval.originalToken Is Nothing AndAlso field.defaultValue IsNot Nothing Then
                eval.originalToken = JToken.FromObject(field.defaultValue)
            End If

            NormalizeFieldValue(field, eval, culture, normalization, errors)
            ValidateFieldValue(config.defaults, field, eval, errors)

            result(field.name) = eval
        Next

        If request.originalRequest IsNot Nothing AndAlso request.originalRequest.include Then
            Dim content As String = If(request.originalRequest.content, "")
            If content.Length > config.defaults.maxOriginalRequestLength Then
                errors.Add(CreateMessage(
                    "ORIGINAL_REQUEST_TOO_LARGE",
                    $"The originalRequest content exceeds the configured maximum length of {config.defaults.maxOriginalRequestLength}."))
            End If
        End If

        Return result
    End Function

    Private Sub NormalizeFieldValue(field As DataCollectorFieldConfig,
                                    eval As DataCollectorFieldEvaluation,
                                    culture As CultureInfo,
                                    normalization As List(Of DataCollectorNormalizationItem),
                                    errors As List(Of DataCollectorMessageItem))
        Dim originalValue As Object = TokenToPlainObject(eval.originalToken)
        Dim fieldType As String = If(field.type, "").Trim().ToLowerInvariant()
        Dim norm As DataCollectorFieldNormalizationConfig = If(field.normalization, New DataCollectorFieldNormalizationConfig())

        If eval.originalToken Is Nothing OrElse eval.originalToken.Type = JTokenType.Null Then
            eval.outputToken = JValue.CreateNull()
            normalization.Add(New DataCollectorNormalizationItem With {
                .field = field.name,
                .originalValue = originalValue,
                .normalizedValue = Nothing,
                .normalizationApplied = False
            })
            Exit Sub
        End If

        Dim rawText As String = eval.originalToken.ToString()
        Dim normalizedValue As Object = Nothing
        Dim outputToken As JToken = Nothing
        Dim applied As Boolean = False

        Try
            Select Case fieldType
                Case "string", "enum", "email", "uri"
                    Dim text As String = NormalizeString(rawText, norm, applied)
                    If text = "" Then
                        normalizedValue = Nothing
                        outputToken = JValue.CreateNull()
                    Else
                        normalizedValue = text
                        outputToken = New JValue(text)
                    End If

                Case "boolean"
                    Dim text As String = NormalizeString(rawText, norm, applied)
                    If text = "" Then
                        normalizedValue = Nothing
                        outputToken = JValue.CreateNull()
                    Else
                        Dim parsed As Boolean
                        If Not TryParseBoolean(text, parsed) Then
                            Throw New FormatException($"Could not normalize value '{rawText}' as boolean.")
                        End If
                        normalizedValue = parsed
                        outputToken = New JValue(parsed)
                        applied = True
                    End If

                Case "integer"
                    Dim parsed As Long
                    If Not TryParseInteger(rawText, norm, parsed) Then
                        Throw New FormatException($"Could not normalize value '{rawText}' as integer.")
                    End If
                    normalizedValue = parsed
                    outputToken = New JValue(parsed)
                    applied = True

                Case "decimal"
                    Dim parsed As Decimal
                    If Not TryParseDecimalValue(rawText, norm, parsed) Then
                        Throw New FormatException($"Could not normalize value '{rawText}' as decimal.")
                    End If
                    normalizedValue = parsed
                    outputToken = New JValue(parsed)
                    applied = True

                Case "date"
                    Dim parsed As DateTime
                    If Not TryParseDateValue(rawText, norm, culture, parsed, dateOnly:=True) Then
                        Throw New FormatException($"Could not normalize value '{rawText}' as date.")
                    End If
                    normalizedValue = parsed.Date
                    outputToken = New JValue(parsed.ToString(If(String.IsNullOrWhiteSpace(norm.outputFormat), "yyyy-MM-dd", norm.outputFormat), CultureInfo.InvariantCulture))
                    applied = True

                Case "datetime"
                    Dim parsed As DateTime
                    If Not TryParseDateValue(rawText, norm, culture, parsed, dateOnly:=False) Then
                        Throw New FormatException($"Could not normalize value '{rawText}' as datetime.")
                    End If
                    normalizedValue = parsed
                    outputToken = New JValue(parsed.ToString(If(String.IsNullOrWhiteSpace(norm.outputFormat), "o", norm.outputFormat), CultureInfo.InvariantCulture))
                    applied = True

                Case Else
                    Throw New FormatException($"Unsupported field type '{field.type}'.")
            End Select
        Catch ex As Exception
            errors.Add(CreateMessage(
                "NORMALIZATION_FAILED",
                If(field.validation?.customErrorMessage, ex.Message),
                field.name,
                originalValue))
            outputToken = JValue.CreateNull()
            normalizedValue = Nothing
        End Try

        eval.normalizedValue = normalizedValue
        eval.outputToken = If(outputToken, JValue.CreateNull())
        eval.normalizationApplied = applied

        normalization.Add(New DataCollectorNormalizationItem With {
            .field = field.name,
            .originalValue = originalValue,
            .normalizedValue = TokenToPlainObject(eval.outputToken),
            .normalizationApplied = applied
        })
    End Sub

    Private Sub ValidateFieldValue(defaults As DataCollectorDefaultsConfig,
                                   field As DataCollectorFieldConfig,
                                   eval As DataCollectorFieldEvaluation,
                                   errors As List(Of DataCollectorMessageItem))
        Dim validation As DataCollectorFieldValidationConfig = If(field.validation, New DataCollectorFieldValidationConfig())
        Dim hasValue As Boolean = eval.outputToken IsNot Nothing AndAlso eval.outputToken.Type <> JTokenType.Null
        Dim originalValue As Object = TokenToPlainObject(eval.originalToken)
        Dim normalizedValue As Object = TokenToPlainObject(eval.outputToken)

        If field.required AndAlso Not hasValue Then
            errors.Add(CreateMessage(
                "REQUIRED_FIELD_MISSING",
                $"The field '{field.name}' is required.",
                field.name,
                originalValue,
                normalizedValue))
            Exit Sub
        End If

        If Not hasValue Then
            Exit Sub
        End If

        If TypeOf normalizedValue Is String Then
            Dim text As String = CStr(normalizedValue)

            If text.Length > defaults.maxFieldLength Then
                errors.Add(CreateMessage(
                    "FIELD_TOO_LONG",
                    $"The field '{field.name}' exceeds the configured maximum field length.",
                    field.name,
                    originalValue,
                    normalizedValue))
            End If

            If validation.minLength.HasValue AndAlso text.Length < validation.minLength.Value Then
                errors.Add(CreateMessage(
                    "MIN_LENGTH_VIOLATION",
                    $"The field '{field.name}' must have at least {validation.minLength.Value} characters.",
                    field.name,
                    originalValue,
                    normalizedValue))
            End If

            If validation.maxLength.HasValue AndAlso text.Length > validation.maxLength.Value Then
                errors.Add(CreateMessage(
                    "MAX_LENGTH_VIOLATION",
                    $"The field '{field.name}' must have at most {validation.maxLength.Value} characters.",
                    field.name,
                    originalValue,
                    normalizedValue))
            End If

            If Not String.IsNullOrWhiteSpace(validation.regex) AndAlso Not Regex.IsMatch(text, validation.regex) Then
                errors.Add(CreateMessage(
                    "REGEX_MISMATCH",
                    If(validation.customErrorMessage, $"The field '{field.name}' does not match the required pattern."),
                    field.name,
                    originalValue,
                    normalizedValue))
            End If
        End If

        Select Case If(field.type, "").Trim().ToLowerInvariant()
            Case "decimal", "integer"
                Dim numericValue As Decimal = System.Convert.ToDecimal(eval.normalizedValue, CultureInfo.InvariantCulture)

                If validation.min.HasValue AndAlso numericValue < validation.min.Value Then
                    errors.Add(CreateMessage("VALUE_BELOW_MIN", $"The field '{field.name}' must be at least {validation.min.Value}.", field.name, originalValue, normalizedValue))
                End If
                If validation.max.HasValue AndAlso numericValue > validation.max.Value Then
                    errors.Add(CreateMessage("VALUE_ABOVE_MAX", $"The field '{field.name}' must be at most {validation.max.Value}.", field.name, originalValue, normalizedValue))
                End If

            Case "date", "datetime"
                Dim dt As DateTime = DirectCast(eval.normalizedValue, DateTime)

                If validation.allowFutureDate.HasValue AndAlso Not validation.allowFutureDate.Value AndAlso dt.Date > DateTime.Now.Date Then
                    errors.Add(CreateMessage("DATE_IN_FUTURE_NOT_ALLOWED", $"The field '{field.name}' must not be in the future.", field.name, originalValue, normalizedValue))
                End If

                Dim minDate As DateTime
                If TryParseValidationDate(validation.minDate, minDate) AndAlso dt.Date < minDate.Date Then
                    errors.Add(CreateMessage("DATE_BELOW_MIN", $"The field '{field.name}' must be on or after {minDate:yyyy-MM-dd}.", field.name, originalValue, normalizedValue))
                End If

                Dim maxDate As DateTime
                If TryParseValidationDate(validation.maxDate, maxDate) AndAlso dt.Date > maxDate.Date Then
                    errors.Add(CreateMessage("DATE_ABOVE_MAX", $"The field '{field.name}' must be on or before {maxDate:yyyy-MM-dd}.", field.name, originalValue, normalizedValue))
                End If

            Case "email"
                Dim addr As MailAddress = Nothing
                Try
                    addr = New MailAddress(CStr(eval.normalizedValue))
                Catch
                End Try
                If addr Is Nothing Then
                    errors.Add(CreateMessage("EMAIL_INVALID", $"The field '{field.name}' must contain a valid email address.", field.name, originalValue, normalizedValue))
                End If

            Case "uri"
                Dim uri As Uri = Nothing
                If Not Uri.TryCreate(CStr(eval.normalizedValue), UriKind.Absolute, uri) Then
                    errors.Add(CreateMessage("URI_INVALID", $"The field '{field.name}' must contain a valid URI.", field.name, originalValue, normalizedValue))
                End If

            Case "enum"
                If validation.allowedValues IsNot Nothing AndAlso validation.allowedValues.Count > 0 Then
                    If Not validation.allowedValues.Any(Function(x) String.Equals(x, CStr(eval.normalizedValue), StringComparison.OrdinalIgnoreCase)) Then
                        errors.Add(CreateMessage("ENUM_NOT_ALLOWED", $"The field '{field.name}' must match one of the configured allowed values.", field.name, originalValue, normalizedValue))
                    End If
                End If
        End Select
    End Sub

    Private Function BuildOutputRecord(useCase As DataCollectorUseCaseConfig,
                                       request As DataCollectorCollectionRequest,
                                       evaluations As Dictionary(Of String, DataCollectorFieldEvaluation),
                                       warnings As List(Of DataCollectorMessageItem),
                                       errors As List(Of DataCollectorMessageItem)) As JObject
        Dim record As New JObject()

        For Each field In useCase.fields
            Dim eval As DataCollectorFieldEvaluation = Nothing
            If evaluations.TryGetValue(field.name, eval) Then
                record(field.name) = If(eval.outputToken IsNot Nothing, eval.outputToken.DeepClone(), JValue.CreateNull())
            Else
                record(field.name) = JValue.CreateNull()
            End If
        Next

        If useCase.extraction IsNot Nothing AndAlso useCase.extraction.includeOriginalRequest Then
            Dim fieldName As String = If(useCase.extraction.originalRequestFieldName, "originalRequest").Trim()

            If request.originalRequest IsNot Nothing AndAlso request.originalRequest.include Then
                If String.IsNullOrWhiteSpace(request.originalRequest.content) Then
                    warnings.Add(CreateMessage("ORIGINAL_REQUEST_EMPTY", $"The use case '{useCase.id}' allows originalRequest, but no original request content was supplied."))
                Else
                    record(fieldName) = request.originalRequest.content
                End If
            End If
        End If

        If SerializeJson(record).Length > 250000 Then
            errors.Add(CreateMessage("RECORD_TOO_LARGE", "The normalized record exceeds the configured maximum record size."))
        End If

        Return record
    End Function

    Private Function TryResolveTargetPath(config As DataCollectorConfiguration,
                                          useCase As DataCollectorUseCaseConfig,
                                          request As DataCollectorCollectionRequest,
                                          record As JObject,
                                          ByRef targetPath As String,
                                          errors As List(Of DataCollectorMessageItem)) As Boolean
        targetPath = Nothing

        Try
            Dim baseDirectory As String = NormalizeFullPath(ExpandEnvironmentVariables(config.allowedBaseDirectory))
            Dim relativeDirectory As String = If(useCase.target.directory, "").Trim()

            If Path.IsPathRooted(relativeDirectory) Then
                errors.Add(CreateMessage("TARGET_PATH_NOT_ALLOWED", "Absolute target directories are not allowed."))
                Return False
            End If

            Dim targetDirectory As String = NormalizeFullPath(Path.Combine(baseDirectory, relativeDirectory))
            If Not IsPathWithinBaseDirectory(targetDirectory, baseDirectory) Then
                errors.Add(CreateMessage("TARGET_PATH_NOT_ALLOWED", "The resolved target directory is outside allowedBaseDirectory."))
                Return False
            End If

            Dim fileName As String = ResolveFileNameTemplate(
                useCase,
                request,
                record,
                useCase.target.fileNameTemplate,
                useCase.target.fileNameMissingFieldPolicy,
                useCase.target.fileNameMissingFieldFallbackValue,
                errors)

            If errors.Count > 0 Then
                Return False
            End If

            Dim combined As String = NormalizeFullPath(Path.Combine(targetDirectory, fileName))
            If Not IsPathWithinBaseDirectory(combined, baseDirectory) Then
                errors.Add(CreateMessage("TARGET_PATH_NOT_ALLOWED", "The resolved target file path is outside allowedBaseDirectory."))
                Return False
            End If

            targetPath = combined
            Return True
        Catch ex As Exception
            errors.Add(CreateMessage("TARGET_PATH_NOT_ALLOWED", $"The target path could not be resolved safely: {ex.Message}"))
            Return False
        End Try
    End Function

    Private Function ResolveFileNameTemplate(useCase As DataCollectorUseCaseConfig,
                                             request As DataCollectorCollectionRequest,
                                             record As JObject,
                                             template As String,
                                             missingFieldPolicy As String,
                                             fallbackValue As String,
                                             errors As List(Of DataCollectorMessageItem)) As String
        Dim nowLocal As DateTime = DateTime.Now
        Dim missingFields As New List(Of String)()

        Dim resolved As String = Regex.Replace(template, "\{([^{}]+)\}",
            Function(m As Match) As String
                Dim token As String = m.Groups(1).Value.Trim()
                Dim replacement As String = ""

                Select Case token
                    Case "yyyy" : replacement = nowLocal.ToString("yyyy", CultureInfo.InvariantCulture)
                    Case "MM" : replacement = nowLocal.ToString("MM", CultureInfo.InvariantCulture)
                    Case "dd" : replacement = nowLocal.ToString("dd", CultureInfo.InvariantCulture)
                    Case "HH" : replacement = nowLocal.ToString("HH", CultureInfo.InvariantCulture)
                    Case "mm" : replacement = nowLocal.ToString("mm", CultureInfo.InvariantCulture)
                    Case "ss" : replacement = nowLocal.ToString("ss", CultureInfo.InvariantCulture)
                    Case "fff" : replacement = nowLocal.ToString("fff", CultureInfo.InvariantCulture)
                    Case "useCaseId" : replacement = useCase.id
                    Case "messageId" : replacement = request.source.messageId
                    Case "senderDomain" : replacement = ExtractSenderDomain(request.source.from)
                    Case Else
                        If token.StartsWith("field:", StringComparison.OrdinalIgnoreCase) Then
                            Dim fieldName As String = token.Substring(6).Trim()
                            Dim t As JToken = record(fieldName)

                            If t Is Nothing OrElse t.Type = JTokenType.Null OrElse String.IsNullOrWhiteSpace(t.ToString()) Then
                                missingFields.Add(fieldName)
                                replacement = ResolveMissingPlaceholderValue(missingFieldPolicy, fallbackValue)
                            Else
                                replacement = t.ToString()
                            End If
                        End If
                End Select

                Return SanitizeFileNameComponent(replacement)
            End Function)

        If missingFields.Count > 0 AndAlso String.Equals(If(missingFieldPolicy, "reject"), "reject", StringComparison.OrdinalIgnoreCase) Then
            errors.Add(CreateMessage("FILENAME_PLACEHOLDER_MISSING", $"The target file name requires missing field value(s): {String.Join(", ", missingFields)}."))
        End If

        If String.IsNullOrWhiteSpace(resolved) Then
            errors.Add(CreateMessage("FILENAME_TEMPLATE_INVALID", "The resolved target file name is empty."))
        End If

        Return resolved
    End Function

    Private Shared Function ResolveMissingPlaceholderValue(policy As String, fallbackValue As String) As String
        Select Case If(policy, "reject").Trim()
            Case "useEmpty"
                Return ""
            Case "useFallback"
                Return If(fallbackValue, "")
            Case Else
                Return ""
        End Select
    End Function

    Private Function CheckForDuplicate(config As DataCollectorConfiguration,
                                       useCase As DataCollectorUseCaseConfig,
                                       targetPath As String,
                                       record As JObject) As DataCollectorDuplicateCheckResult
        Dim result As New DataCollectorDuplicateCheckResult() With {
            .enabled = False,
            .duplicateFound = False
        }

        Dim duplicateControl As DataCollectorDuplicateControlConfig = If(useCase.duplicateControl, New DataCollectorDuplicateControlConfig())
        Dim mustCheck As Boolean =
            duplicateControl.enabled OrElse
            String.Equals(If(useCase.target?.mode, ""), "upsert", StringComparison.OrdinalIgnoreCase) OrElse
            String.Equals(GetDuplicatePolicy(useCase), "upsert", StringComparison.OrdinalIgnoreCase)

        If Not mustCheck Then
            Return result
        End If

        result.enabled = True

        If duplicateControl.recordKey Is Nothing OrElse duplicateControl.recordKey.Count = 0 Then
            result.blockingErrors.Add(CreateMessage("DUPLICATE_KEY_MISSING", "Duplicate checking is enabled but duplicateControl.recordKey is missing."))
            Return result
        End If

        For Each keyField In duplicateControl.recordKey
            Dim keyToken As JToken = record(keyField)
            result.recordKey(keyField) = If(keyToken IsNot Nothing, keyToken.DeepClone(), JValue.CreateNull())

            If (keyToken Is Nothing OrElse keyToken.Type = JTokenType.Null OrElse String.IsNullOrWhiteSpace(keyToken.ToString())) AndAlso
               Not duplicateControl.allowMissingDuplicateKeys Then
                result.blockingErrors.Add(CreateMessage("DUPLICATE_KEY_VALUE_MISSING", $"The duplicate key field '{keyField}' is missing or empty."))
            End If
        Next

        If result.blockingErrors.Count > 0 Then
            Return result
        End If

        Dim readResult As DataCollectorReadResult = ReadExistingRecords(targetPath, useCase)
        If Not readResult.success Then
            For Each errItem In readResult.errors
                result.blockingErrors.Add(errItem)
            Next
            Return result
        End If

        For Each warn In readResult.warnings
            result.warnings.Add(warn)
        Next

        Dim fieldMap As Dictionary(Of String, DataCollectorFieldConfig) =
            useCase.fields.ToDictionary(Function(f) f.name, Function(f) f, StringComparer.OrdinalIgnoreCase)

        For i As Integer = 0 To readResult.records.Count - 1
            Dim existing As JObject = readResult.records(i)
            Dim isMatch As Boolean = True

            For Each keyField In duplicateControl.recordKey
                Dim fieldDef As DataCollectorFieldConfig = Nothing
                fieldMap.TryGetValue(keyField, fieldDef)

                Dim leftValue As String = BuildComparableDuplicateValue(record(keyField), fieldDef, duplicateControl)
                Dim rightValue As String = BuildComparableDuplicateValue(existing(keyField), fieldDef, duplicateControl)

                If Not String.Equals(leftValue, rightValue, StringComparison.Ordinal) Then
                    isMatch = False
                    Exit For
                End If
            Next

            If isMatch Then
                result.duplicateFound = True
                result.existingRecord = existing
                result.existingIndex = i
                Exit For
            End If
        Next

        Return result
    End Function

    Private Function ReadExistingRecords(targetPath As String, useCase As DataCollectorUseCaseConfig) As DataCollectorReadResult
        Dim result As New DataCollectorReadResult()

        If Not File.Exists(targetPath) Then
            Return result
        End If

        Try
            Select Case If(useCase.target.format, "").Trim().ToLowerInvariant()
                Case "csv"
                    Return ReadCsvRecords(targetPath, useCase)
                Case "json"
                    Return ReadJsonArrayRecords(targetPath)
                Case "jsonl"
                    Return ReadJsonLinesRecords(targetPath, useCase)
                Case Else
                    result.success = False
                    result.errors.Add(CreateMessage("UNSUPPORTED_FORMAT", $"The target format '{useCase.target.format}' is not supported."))
                    Return result
            End Select
        Catch ex As IOException
            result.success = False
            result.errors.Add(CreateMessage("TARGET_FILE_LOCKED", $"The target file could not be read: {ex.Message}"))
        Catch ex As Exception
            result.success = False
            result.errors.Add(CreateMessage("TARGET_FILE_INVALID", $"The target file could not be parsed: {ex.Message}"))
        End Try

        Return result
    End Function

    Private Function ReadCsvRecords(targetPath As String, useCase As DataCollectorUseCaseConfig) As DataCollectorReadResult
        Dim result As New DataCollectorReadResult()
        Dim content As String = File.ReadAllText(targetPath, DetectEncoding(ResolveOutputEncodingName(useCase.target, Nothing)))
        Dim delimiter As Char = ResolveCsvDelimiter(useCase)
        Dim rows As List(Of List(Of String)) = ParseCsv(content, delimiter)

        If rows.Count = 0 Then
            Return result
        End If

        Dim columnNames As List(Of String)
        Dim firstDataRow As Integer = 0

        If useCase.csv IsNot Nothing AndAlso useCase.csv.includeHeader AndAlso rows.Count > 0 Then
            columnNames = rows(0)
            firstDataRow = 1
        Else
            columnNames = GetConfiguredOutputColumnNames(useCase)
        End If

        For i As Integer = firstDataRow To rows.Count - 1
            Dim row As List(Of String) = rows(i)
            Dim obj As New JObject()

            For c As Integer = 0 To columnNames.Count - 1
                obj(columnNames(c)) = If(c < row.Count, row(c), "")
            Next

            result.records.Add(obj)
        Next

        Return result
    End Function

    Private Function ReadJsonArrayRecords(targetPath As String) As DataCollectorReadResult
        Dim result As New DataCollectorReadResult()
        Dim content As String = File.ReadAllText(targetPath, DetectEncoding("utf-8"))
        Dim token As JToken = JToken.Parse(content)

        If TypeOf token IsNot JArray Then
            result.success = False
            result.errors.Add(CreateMessage("TARGET_FILE_INVALID", "The target JSON file must contain a root array."))
            Return result
        End If

        For Each item As JToken In DirectCast(token, JArray)
            If TypeOf item IsNot JObject Then
                result.success = False
                result.errors.Add(CreateMessage("TARGET_FILE_INVALID", "The target JSON array contains a non-object entry."))
                Return result
            End If

            result.records.Add(DirectCast(item, JObject))
        Next

        Return result
    End Function

    Private Function ReadJsonLinesRecords(targetPath As String, useCase As DataCollectorUseCaseConfig) As DataCollectorReadResult
        Dim result As New DataCollectorReadResult()
        Dim policy As String = NormalizeJsonlInvalidLinePolicy(If(useCase.jsonl?.invalidLinePolicy, "reject"))

        Using reader As New StreamReader(targetPath, DetectEncoding(ResolveOutputEncodingName(useCase.target, Nothing)), True)
            Dim lineNumber As Integer = 0

            Do While Not reader.EndOfStream
                Dim line As String = reader.ReadLine()
                lineNumber += 1

                If String.IsNullOrWhiteSpace(line) Then
                    Continue Do
                End If

                Try
                    Dim token As JToken = JToken.Parse(line)
                    If TypeOf token Is JObject Then
                        result.records.Add(DirectCast(token, JObject))
                    ElseIf policy = "ignore" Then
                        result.warnings.Add(CreateMessage("JSONL_INVALID_LINE_IGNORED", $"Ignored non-object JSONL line {lineNumber}."))
                    Else
                        result.success = False
                        result.errors.Add(CreateMessage("TARGET_FILE_INVALID", $"Invalid JSONL object at line {lineNumber}."))
                        Return result
                    End If
                Catch ex As Exception
                    Select Case policy
                        Case "ignore"
                            result.warnings.Add(CreateMessage("JSONL_INVALID_LINE_IGNORED", $"Ignored invalid JSONL line {lineNumber}: {ex.Message}"))
                        Case "allow_append_warn"
                            result.warnings.Add(CreateMessage("JSONL_INVALID_LINE_PRESENT", $"Invalid JSONL line {lineNumber} was present during duplicate scan: {ex.Message}"))
                        Case Else
                            result.success = False
                            result.errors.Add(CreateMessage("TARGET_FILE_INVALID", $"Invalid JSONL line {lineNumber}: {ex.Message}"))
                            Return result
                    End Select
                End Try
            Loop
        End Using

        Return result
    End Function

    Private Function CreateDataCollectorOutputWriter(format As String) As IDataCollectorOutputWriter
        Select Case If(format, "").Trim().ToLowerInvariant()
            Case "csv"
                Return New CsvOutputWriter(Me)
            Case "json"
                Return New JsonArrayOutputWriter(Me)
            Case "jsonl"
                Return New JsonLinesOutputWriter(Me)
            Case Else
                Return Nothing
        End Select
    End Function

    Private Function BuildCsvPreview(record As JObject, useCase As DataCollectorUseCaseConfig) As JObject
        Dim delimiter As Char = ResolveCsvDelimiter(useCase)
        Dim line As String = SerializeCsvRow(
            GetConfiguredOutputColumnNames(useCase).
                Select(Function(name) GetCsvSerializableValue(record(name), useCase)).
                ToList(),
            delimiter,
            ResolveCsvQuoteMode(useCase))

        Return New JObject(
            New JProperty("format", "csv"),
            New JProperty("csvLine", line),
            New JProperty("record", record.DeepClone()))
    End Function

    Private Function WriteCsvRecordCore(targetPath As String,
                                        record As JObject,
                                        useCase As DataCollectorUseCaseConfig,
                                        config As DataCollectorConfiguration,
                                        effectiveWriteMode As String,
                                        duplicateResult As DataCollectorDuplicateCheckResult) As DataCollectorWriteResult
        Dim result As New DataCollectorWriteResult()
        Dim delimiter As Char = ResolveCsvDelimiter(useCase)
        Dim includeHeader As Boolean = If(useCase.csv IsNot Nothing, useCase.csv.includeHeader, True)
        Dim newline As String = ResolveCsvNewLine(useCase)
        Dim quoteMode As String = ResolveCsvQuoteMode(useCase)
        Dim encoding As Encoding = DetectEncoding(ResolveOutputEncodingName(useCase.target, config.defaults))
        Dim columns As List(Of String) = GetConfiguredOutputColumnNames(useCase)

        Try
            EnsureDirectoryExists(Path.GetDirectoryName(targetPath), config.defaults)

            Select Case effectiveWriteMode
                Case "appendOrCreate", "append"
                    If effectiveWriteMode = "append" AndAlso Not File.Exists(targetPath) Then
                        result.success = False
                        result.errors.Add(CreateMessage("WRITE_MODE_APPEND_REQUIRES_FILE", "Write mode 'append' requires an existing target file."))
                        Return result
                    End If

                    Dim sb As New StringBuilder()

                    If Not File.Exists(targetPath) OrElse New FileInfo(targetPath).Length = 0 Then
                        If includeHeader Then
                            sb.Append(SerializeCsvRow(columns, delimiter, quoteMode))
                            sb.Append(newline)
                        End If
                    End If

                    sb.Append(SerializeCsvRow(
                        columns.Select(Function(name) GetCsvSerializableValue(record(name), useCase)).ToList(),
                        delimiter,
                        quoteMode))

                    AppendText(targetPath, sb.ToString(), encoding)
                    result.status = AP_DC_Status_Written

                Case "create"
                    If File.Exists(targetPath) Then
                        result.success = False
                        result.errors.Add(CreateMessage("WRITE_MODE_CREATE_FILE_EXISTS", "Write mode 'create' rejects existing target files."))
                        Return result
                    End If

                    WriteAllTextAtomic(
                        targetPath,
                        BuildCsvContent(New List(Of JObject) From {record}, columns, delimiter, includeHeader, quoteMode, newline, useCase),
                        encoding,
                        config.defaults.atomicWrites,
                        False)
                    result.status = AP_DC_Status_Written

                Case "overwrite"
                    WriteAllTextAtomic(
                        targetPath,
                        BuildCsvContent(New List(Of JObject) From {record}, columns, delimiter, includeHeader, quoteMode, newline, useCase),
                        encoding,
                        config.defaults.atomicWrites,
                        config.defaults.backupBeforeStructuredUpdate)
                    result.status = AP_DC_Status_Written

                Case "rejectIfExists"
                    If File.Exists(targetPath) Then
                        result.success = False
                        result.errors.Add(CreateMessage("WRITE_MODE_REJECT_IF_EXISTS", "Write mode 'rejectIfExists' rejects existing target files."))
                        Return result
                    End If

                    WriteAllTextAtomic(
                        targetPath,
                        BuildCsvContent(New List(Of JObject) From {record}, columns, delimiter, includeHeader, quoteMode, newline, useCase),
                        encoding,
                        config.defaults.atomicWrites,
                        False)
                    result.status = AP_DC_Status_Written

                Case "upsert"
                    Dim readResult As DataCollectorReadResult = ReadCsvRecords(targetPath, useCase)
                    If Not readResult.success Then
                        result.success = False
                        For Each errItem In readResult.errors
                            result.errors.Add(errItem)
                        Next
                        Return result
                    End If

                    Dim records As List(Of JObject) = readResult.records
                    If duplicateResult IsNot Nothing AndAlso duplicateResult.duplicateFound AndAlso duplicateResult.existingIndex >= 0 AndAlso duplicateResult.existingIndex < records.Count Then
                        records(duplicateResult.existingIndex) = DirectCast(record.DeepClone(), JObject)
                        result.status = AP_DC_Status_Updated
                    Else
                        records.Add(DirectCast(record.DeepClone(), JObject))
                        result.status = AP_DC_Status_Written
                    End If

                    WriteAllTextAtomic(
                        targetPath,
                        BuildCsvContent(records, columns, delimiter, includeHeader, quoteMode, newline, useCase),
                        encoding,
                        config.defaults.atomicWrites,
                        config.defaults.backupBeforeStructuredUpdate)

                Case Else
                    result.success = False
                    result.errors.Add(CreateMessage("UNSUPPORTED_WRITE_MODE", $"Unsupported write mode '{effectiveWriteMode}'."))
            End Select
        Catch ex As IOException
            result.success = False
            result.errors.Add(CreateMessage("TARGET_FILE_LOCKED", ex.Message))
        Catch ex As Exception
            result.success = False
            result.errors.Add(CreateMessage("WRITE_FAILED", ex.Message))
        End Try

        Return result
    End Function

    Private Function WriteJsonArrayRecordCore(targetPath As String,
                                              record As JObject,
                                              useCase As DataCollectorUseCaseConfig,
                                              config As DataCollectorConfiguration,
                                              effectiveWriteMode As String,
                                              duplicateResult As DataCollectorDuplicateCheckResult) As DataCollectorWriteResult
        Dim result As New DataCollectorWriteResult()
        Dim encoding As Encoding = DetectEncoding(ResolveOutputEncodingName(useCase.target, config.defaults))

        Try
            EnsureDirectoryExists(Path.GetDirectoryName(targetPath), config.defaults)

            Dim arr As JArray = Nothing
            Dim exists As Boolean = File.Exists(targetPath)

            Select Case effectiveWriteMode
                Case "appendOrCreate"
                    arr = If(exists, LoadJsonArrayForWrite(targetPath), New JArray())
                    arr.Add(record.DeepClone())
                    result.status = AP_DC_Status_Written

                Case "append"
                    If Not exists Then
                        result.success = False
                        result.errors.Add(CreateMessage("WRITE_MODE_APPEND_REQUIRES_FILE", "Write mode 'append' requires an existing target file."))
                        Return result
                    End If

                    arr = LoadJsonArrayForWrite(targetPath)
                    arr.Add(record.DeepClone())
                    result.status = AP_DC_Status_Written

                Case "create"
                    If exists Then
                        result.success = False
                        result.errors.Add(CreateMessage("WRITE_MODE_CREATE_FILE_EXISTS", "Write mode 'create' rejects existing target files."))
                        Return result
                    End If

                    arr = New JArray(record.DeepClone())
                    result.status = AP_DC_Status_Written

                Case "overwrite"
                    arr = New JArray(record.DeepClone())
                    result.status = AP_DC_Status_Written

                Case "rejectIfExists"
                    If exists Then
                        result.success = False
                        result.errors.Add(CreateMessage("WRITE_MODE_REJECT_IF_EXISTS", "Write mode 'rejectIfExists' rejects existing target files."))
                        Return result
                    End If

                    arr = New JArray(record.DeepClone())
                    result.status = AP_DC_Status_Written

                Case "upsert"
                    arr = If(exists, LoadJsonArrayForWrite(targetPath), New JArray())
                    If duplicateResult IsNot Nothing AndAlso duplicateResult.duplicateFound AndAlso duplicateResult.existingIndex >= 0 AndAlso duplicateResult.existingIndex < arr.Count Then
                        arr(duplicateResult.existingIndex) = record.DeepClone()
                        result.status = AP_DC_Status_Updated
                    Else
                        arr.Add(record.DeepClone())
                        result.status = AP_DC_Status_Written
                    End If

                Case Else
                    result.success = False
                    result.errors.Add(CreateMessage("UNSUPPORTED_WRITE_MODE", $"Unsupported write mode '{effectiveWriteMode}'."))
                    Return result
            End Select

            WriteAllTextAtomic(
                targetPath,
                arr.ToString(Newtonsoft.Json.Formatting.Indented),
                encoding,
                config.defaults.atomicWrites,
                config.defaults.backupBeforeStructuredUpdate)
        Catch ex As JsonReaderException
            result.success = False
            result.errors.Add(CreateMessage("TARGET_FILE_INVALID", ex.Message))
        Catch ex As IOException
            result.success = False
            result.errors.Add(CreateMessage("TARGET_FILE_LOCKED", ex.Message))
        Catch ex As Exception
            result.success = False
            result.errors.Add(CreateMessage("WRITE_FAILED", ex.Message))
        End Try

        Return result
    End Function

    Private Function WriteJsonLinesRecordCore(targetPath As String,
                                              record As JObject,
                                              useCase As DataCollectorUseCaseConfig,
                                              config As DataCollectorConfiguration,
                                              effectiveWriteMode As String,
                                              duplicateResult As DataCollectorDuplicateCheckResult) As DataCollectorWriteResult
        Dim result As New DataCollectorWriteResult()
        Dim encoding As Encoding = DetectEncoding(ResolveOutputEncodingName(useCase.target, config.defaults))
        Dim line As String = record.ToString(Newtonsoft.Json.Formatting.None) & vbCrLf

        Try
            EnsureDirectoryExists(Path.GetDirectoryName(targetPath), config.defaults)

            Select Case effectiveWriteMode
                Case "appendOrCreate"
                    AppendText(targetPath, line, encoding)
                    result.status = AP_DC_Status_Written

                Case "append"
                    If Not File.Exists(targetPath) Then
                        result.success = False
                        result.errors.Add(CreateMessage("WRITE_MODE_APPEND_REQUIRES_FILE", "Write mode 'append' requires an existing target file."))
                        Return result
                    End If

                    AppendText(targetPath, line, encoding)
                    result.status = AP_DC_Status_Written

                Case "create"
                    If File.Exists(targetPath) Then
                        result.success = False
                        result.errors.Add(CreateMessage("WRITE_MODE_CREATE_FILE_EXISTS", "Write mode 'create' rejects existing target files."))
                        Return result
                    End If

                    WriteAllTextAtomic(targetPath, line, encoding, config.defaults.atomicWrites, False)
                    result.status = AP_DC_Status_Written

                Case "overwrite"
                    WriteAllTextAtomic(targetPath, line, encoding, config.defaults.atomicWrites, config.defaults.backupBeforeStructuredUpdate)
                    result.status = AP_DC_Status_Written

                Case "rejectIfExists"
                    If File.Exists(targetPath) Then
                        result.success = False
                        result.errors.Add(CreateMessage("WRITE_MODE_REJECT_IF_EXISTS", "Write mode 'rejectIfExists' rejects existing target files."))
                        Return result
                    End If

                    WriteAllTextAtomic(targetPath, line, encoding, config.defaults.atomicWrites, False)
                    result.status = AP_DC_Status_Written

                Case "upsert"
                    Dim readResult As DataCollectorReadResult = ReadJsonLinesRecords(targetPath, useCase)
                    If Not readResult.success Then
                        result.success = False
                        For Each errItem In readResult.errors
                            result.errors.Add(errItem)
                        Next
                        Return result
                    End If

                    Dim records As List(Of JObject) = readResult.records
                    If duplicateResult IsNot Nothing AndAlso duplicateResult.duplicateFound AndAlso duplicateResult.existingIndex >= 0 AndAlso duplicateResult.existingIndex < records.Count Then
                        records(duplicateResult.existingIndex) = DirectCast(record.DeepClone(), JObject)
                        result.status = AP_DC_Status_Updated
                    Else
                        records.Add(DirectCast(record.DeepClone(), JObject))
                        result.status = AP_DC_Status_Written
                    End If

                    Dim content As String = String.Join(vbCrLf, records.Select(Function(x) x.ToString(Newtonsoft.Json.Formatting.None)))
                    If content <> "" Then
                        content &= vbCrLf
                    End If

                    WriteAllTextAtomic(targetPath, content, encoding, config.defaults.atomicWrites, config.defaults.backupBeforeStructuredUpdate)

                Case Else
                    result.success = False
                    result.errors.Add(CreateMessage("UNSUPPORTED_WRITE_MODE", $"Unsupported write mode '{effectiveWriteMode}'."))
            End Select
        Catch ex As IOException
            result.success = False
            result.errors.Add(CreateMessage("TARGET_FILE_LOCKED", ex.Message))
        Catch ex As Exception
            result.success = False
            result.errors.Add(CreateMessage("WRITE_FAILED", ex.Message))
        End Try

        Return result
    End Function

    Private Function LoadJsonArrayForWrite(targetPath As String) As JArray
        If Not File.Exists(targetPath) Then
            Return New JArray()
        End If

        Dim content As String = File.ReadAllText(targetPath, DetectEncoding("utf-8"))
        Dim token As JToken = JToken.Parse(content)

        If TypeOf token IsNot JArray Then
            Throw New JsonReaderException("The target JSON file must contain a root array.")
        End If

        Return DirectCast(token, JArray)
    End Function

    Private Sub TryWriteAuditLogSafe(config As DataCollectorConfiguration,
                                     useCase As DataCollectorUseCaseConfig,
                                     request As DataCollectorCollectionRequest,
                                     status As String,
                                     operation As String,
                                     targetPath As String,
                                     originalValues As JObject,
                                     normalizedValues As JObject,
                                     duplicateResult As DataCollectorDuplicateCheckResult,
                                     warnings As List(Of DataCollectorMessageItem),
                                     errors As List(Of DataCollectorMessageItem),
                                     dryRun As Boolean,
                                     shouldWriteAudit As Boolean,
                                     warningSink As List(Of DataCollectorMessageItem))
        If Not shouldWriteAudit Then
            Exit Sub
        End If

        Try
            WriteAuditLog(
                config,
                useCase,
                request,
                status,
                operation,
                targetPath,
                originalValues,
                normalizedValues,
                duplicateResult,
                warnings,
                errors,
                dryRun)
        Catch ex As Exception
            warningSink.Add(CreateMessage("LOG_WRITE_FAILED", $"The audit log could not be written: {ex.Message}"))
        End Try
    End Sub

    Private Sub WriteAuditLog(config As DataCollectorConfiguration,
                              useCase As DataCollectorUseCaseConfig,
                              request As DataCollectorCollectionRequest,
                              status As String,
                              operation As String,
                              targetPath As String,
                              originalValues As JObject,
                              normalizedValues As JObject,
                              duplicateResult As DataCollectorDuplicateCheckResult,
                              warnings As List(Of DataCollectorMessageItem),
                              errors As List(Of DataCollectorMessageItem),
                              dryRun As Boolean)
        If config.log Is Nothing OrElse Not config.log.enabled Then
            Exit Sub
        End If

        Dim baseDirectory As String = NormalizeFullPath(ExpandEnvironmentVariables(config.allowedBaseDirectory))
        Dim logDirectory As String = NormalizeFullPath(Path.Combine(baseDirectory, config.log.directory))

        If Not IsPathWithinBaseDirectory(logDirectory, baseDirectory) Then
            Throw New InvalidOperationException("The resolved log directory is outside allowedBaseDirectory.")
        End If

        EnsureDirectoryExists(logDirectory, config.defaults)

        Dim logFileName As String = ResolveLogFileName(config.log.fileNameTemplate, useCase, request)
        Dim logPath As String = NormalizeFullPath(Path.Combine(logDirectory, logFileName))

        If Not IsPathWithinBaseDirectory(logPath, baseDirectory) Then
            Throw New InvalidOperationException("The resolved log path is outside allowedBaseDirectory.")
        End If

        Dim entry As New JObject(
            New JProperty("timestamp", DateTimeOffset.Now.ToString("o", CultureInfo.InvariantCulture)),
            New JProperty("toolVersion", AP_DataCollectorToolVersion),
            New JProperty("schemaVersion", If(config.schemaVersion, "")),
            New JProperty("configVersion", If(config.configVersion, "")),
            New JProperty("useCaseId", If(useCase?.id, "")),
            New JProperty("status", status),
            New JProperty("operation", operation),
            New JProperty("dryRun", dryRun),
            New JProperty("source", BuildSourceObject(request.source)),
            New JProperty("target", New JObject(
                New JProperty("path", If(targetPath, "")),
                New JProperty("format", If(useCase?.target?.format, "")))),
            New JProperty("values", New JObject(
                New JProperty("original", If(originalValues, New JObject())),
                New JProperty("normalized", If(normalizedValues, New JObject())))),
            New JProperty("duplicateCheck", BuildDuplicateAuditObject(duplicateResult)),
            New JProperty("modelNotes", JArray.FromObject(If(request.modelNotes, New List(Of DataCollectorModelNote)()))),
            New JProperty("warnings", JArray.FromObject(If(warnings, New List(Of DataCollectorMessageItem)()))),
            New JProperty("errors", JArray.FromObject(If(errors, New List(Of DataCollectorMessageItem)()))))

        AppendText(logPath, entry.ToString(Newtonsoft.Json.Formatting.None) & vbCrLf, DetectEncoding("utf-8"))
    End Sub

    Private Shared Function BuildDuplicateAuditObject(duplicateResult As DataCollectorDuplicateCheckResult) As JObject
        If duplicateResult Is Nothing Then
            Return New JObject(
                New JProperty("enabled", False),
                New JProperty("duplicateFound", False),
                New JProperty("recordKey", New JObject()))
        End If

        Return New JObject(
            New JProperty("enabled", duplicateResult.enabled),
            New JProperty("duplicateFound", duplicateResult.duplicateFound),
            New JProperty("recordKey", If(duplicateResult.recordKey, New JObject())))
    End Function

    Private Shared Function BuildSourceObject(source As DataCollectorRequestSource) As JObject
        If source Is Nothing Then
            Return New JObject()
        End If

        Return New JObject(
            New JProperty("messageId", If(source.messageId, "")),
            New JProperty("threadId", If(source.threadId, "")),
            New JProperty("from", If(source.from, "")),
            New JProperty("to", JArray.FromObject(If(source.to, New List(Of String)()))),
            New JProperty("cc", JArray.FromObject(If(source.cc, New List(Of String)()))),
            New JProperty("subject", If(source.subject, "")),
            New JProperty("receivedAt", If(source.receivedAt, "")))
    End Function

    Private Shared Function ResolveLogFileName(template As String,
                                               useCase As DataCollectorUseCaseConfig,
                                               request As DataCollectorCollectionRequest) As String
        Dim nowLocal As DateTime = DateTime.Now
        Dim value As String = If(template, "collector_log_{yyyy}-{MM}.jsonl")

        value = value.Replace("{yyyy}", nowLocal.ToString("yyyy", CultureInfo.InvariantCulture))
        value = value.Replace("{MM}", nowLocal.ToString("MM", CultureInfo.InvariantCulture))
        value = value.Replace("{dd}", nowLocal.ToString("dd", CultureInfo.InvariantCulture))
        value = value.Replace("{HH}", nowLocal.ToString("HH", CultureInfo.InvariantCulture))
        value = value.Replace("{mm}", nowLocal.ToString("mm", CultureInfo.InvariantCulture))
        value = value.Replace("{ss}", nowLocal.ToString("ss", CultureInfo.InvariantCulture))
        value = value.Replace("{fff}", nowLocal.ToString("fff", CultureInfo.InvariantCulture))
        value = value.Replace("{useCaseId}", SanitizeFileNameComponent(If(useCase?.id, "")))
        value = value.Replace("{messageId}", SanitizeFileNameComponent(If(request?.source?.messageId, "")))
        value = value.Replace("{senderDomain}", SanitizeFileNameComponent(ExtractSenderDomain(If(request?.source?.from, ""))))

        Return value
    End Function

    Private Shared Function ResolveCulture(defaults As DataCollectorDefaultsConfig) As CultureInfo
        Dim cultureName As String = If(defaults?.culture, "invariant").Trim()

        If cultureName = "" OrElse String.Equals(cultureName, "invariant", StringComparison.OrdinalIgnoreCase) Then
            Return CultureInfo.InvariantCulture
        End If

        Try
            Return CultureInfo.GetCultureInfo(cultureName)
        Catch
            Return CultureInfo.InvariantCulture
        End Try
    End Function

    Private Shared Function NormalizeString(value As String,
                                            norm As DataCollectorFieldNormalizationConfig,
                                            ByRef applied As Boolean) As String
        Dim result As String = If(value, "")

        If norm IsNot Nothing AndAlso norm.trim Then
            Dim trimmed As String = result.Trim()
            applied = applied OrElse Not String.Equals(trimmed, result, StringComparison.Ordinal)
            result = trimmed
        End If

        If norm IsNot Nothing AndAlso norm.collapseWhitespace Then
            Dim collapsed As String = Regex.Replace(result, "\s+", " ").Trim()
            applied = applied OrElse Not String.Equals(collapsed, result, StringComparison.Ordinal)
            result = collapsed
        End If

        If norm IsNot Nothing AndAlso norm.uppercase Then
            Dim upper As String = result.ToUpperInvariant()
            applied = applied OrElse Not String.Equals(upper, result, StringComparison.Ordinal)
            result = upper
        End If

        If norm IsNot Nothing AndAlso norm.lowercase Then
            Dim lower As String = result.ToLowerInvariant()
            applied = applied OrElse Not String.Equals(lower, result, StringComparison.Ordinal)
            result = lower
        End If

        Return result
    End Function

    Private Shared Function TryParseBoolean(text As String, ByRef value As Boolean) As Boolean
        Select Case If(text, "").Trim().ToUpperInvariant()
            Case "TRUE", "YES", "Y", "1"
                value = True
                Return True
            Case "FALSE", "NO", "N", "0"
                value = False
                Return True
            Case Else
                Return Boolean.TryParse(text, value)
        End Select
    End Function

    Private Shared Function TryParseInteger(rawText As String,
                                            norm As DataCollectorFieldNormalizationConfig,
                                            ByRef value As Long) As Boolean
        Dim decValue As Decimal
        If Not TryParseDecimalValue(rawText, norm, decValue) Then
            Return False
        End If

        If decValue <> Decimal.Truncate(decValue) Then
            Return False
        End If

        Try
            value = System.Convert.ToInt64(decValue, CultureInfo.InvariantCulture)
            Return True
        Catch
            Return False
        End Try
    End Function

    Private Shared Function TryParseDecimalValue(rawText As String,
                                                 norm As DataCollectorFieldNormalizationConfig,
                                                 ByRef value As Decimal) As Boolean
        Dim text As String = If(rawText, "")
        Dim applied As Boolean = False
        text = NormalizeString(text, norm, applied)

        If norm IsNot Nothing AndAlso norm.removeCurrencySymbols Then
            text = Regex.Replace(text, "[^\d\-\+\.,' ]", "")
        End If

        text = text.Replace(vbTab, " ").Trim()
        If text = "" Then
            Return False
        End If

        Dim decimalIndex As Integer = -1
        For i As Integer = text.Length - 1 To 0 Step -1
            Dim ch As Char = text(i)
            If ch = "."c OrElse ch = ","c Then
                decimalIndex = i
                Exit For
            End If
        Next

        Dim sb As New StringBuilder()
        For i As Integer = 0 To text.Length - 1
            Dim ch As Char = text(i)

            If Char.IsDigit(ch) Then
                sb.Append(ch)
            ElseIf (ch = "-"c OrElse ch = "+"c) AndAlso sb.Length = 0 Then
                sb.Append(ch)
            ElseIf i = decimalIndex AndAlso (ch = "."c OrElse ch = ","c) Then
                sb.Append("."c)
            End If
        Next

        Return Decimal.TryParse(
            sb.ToString(),
            NumberStyles.AllowLeadingSign Or NumberStyles.AllowDecimalPoint,
            CultureInfo.InvariantCulture,
            value)
    End Function

    Private Shared Function TryParseDateValue(rawText As String,
                                              norm As DataCollectorFieldNormalizationConfig,
                                              culture As CultureInfo,
                                              ByRef value As DateTime,
                                              dateOnly As Boolean) As Boolean
        Dim text As String = If(rawText, "")
        Dim applied As Boolean = False
        text = NormalizeString(text, norm, applied)

        If norm IsNot Nothing AndAlso norm.inputFormats IsNot Nothing AndAlso norm.inputFormats.Count > 0 Then
            For Each fmt In norm.inputFormats
                If DateTime.TryParseExact(text, fmt, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, value) Then
                    If dateOnly Then value = value.Date
                    Return True
                End If

                If DateTime.TryParseExact(text, fmt, culture, DateTimeStyles.AllowWhiteSpaces, value) Then
                    If dateOnly Then value = value.Date
                    Return True
                End If
            Next
        End If

        If DateTime.TryParse(text, culture, DateTimeStyles.AllowWhiteSpaces, value) Then
            If dateOnly Then value = value.Date
            Return True
        End If

        If DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, value) Then
            If dateOnly Then value = value.Date
            Return True
        End If

        Return False
    End Function

    Private Shared Function TryParseValidationDate(text As String, ByRef value As DateTime) As Boolean
        If String.IsNullOrWhiteSpace(text) Then
            Return False
        End If

        Return DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, value)
    End Function

    Private Shared Function BuildComparableDuplicateValue(token As JToken,
                                                          field As DataCollectorFieldConfig,
                                                          duplicateControl As DataCollectorDuplicateControlConfig) As String
        Dim text As String = ""
        If token IsNot Nothing AndAlso token.Type <> JTokenType.Null Then
            text = token.ToString()
        End If

        If duplicateControl.trimWhitespace Then
            text = text.Trim()
        End If

        If duplicateControl.normalizeBeforeCompare AndAlso field IsNot Nothing AndAlso field.normalization IsNot Nothing Then
            Dim applied As Boolean = False
            text = NormalizeString(text, field.normalization, applied)
        End If

        If Not duplicateControl.caseSensitive Then
            text = text.ToUpperInvariant()
        End If

        Return text
    End Function

    Private Shared Function GetDuplicatePolicy(useCase As DataCollectorUseCaseConfig) As String
        Return If(useCase?.duplicateControl?.policy, "allow").Trim()
    End Function

    Private Shared Function BuildOriginalValuesObject(evaluations As Dictionary(Of String, DataCollectorFieldEvaluation)) As JObject
        Dim obj As New JObject()

        If evaluations Is Nothing Then
            Return obj
        End If

        For Each kvp In evaluations
            obj(kvp.Key) = If(kvp.Value.originalToken IsNot Nothing, kvp.Value.originalToken.DeepClone(), JValue.CreateNull())
        Next

        Return obj
    End Function

    Private Shared Function BuildNormalizedValuesObject(evaluations As Dictionary(Of String, DataCollectorFieldEvaluation),
                                                        useCase As DataCollectorUseCaseConfig,
                                                        request As DataCollectorCollectionRequest) As JObject
        Dim obj As New JObject()

        If evaluations IsNot Nothing Then
            For Each fieldEval In evaluations.Values
                obj(fieldEval.definition.name) = If(fieldEval.outputToken IsNot Nothing, fieldEval.outputToken.DeepClone(), JValue.CreateNull())
            Next
        End If

        If useCase IsNot Nothing AndAlso
           useCase.extraction IsNot Nothing AndAlso
           useCase.extraction.includeOriginalRequest AndAlso
           request IsNot Nothing AndAlso
           request.originalRequest IsNot Nothing AndAlso
           request.originalRequest.include AndAlso
           Not String.IsNullOrWhiteSpace(request.originalRequest.content) Then

            obj(If(useCase.extraction.originalRequestFieldName, "originalRequest")) = request.originalRequest.content
        End If

        Return obj
    End Function

    Private Shared Function ResolveOutputEncodingName(target As DataCollectorTargetConfig,
                                                      defaults As DataCollectorDefaultsConfig) As String
        If target IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(target.encoding) Then
            Return target.encoding
        End If

        If defaults IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(defaults.encoding) Then
            Return defaults.encoding
        End If

        Return "utf-8"
    End Function

    Private Shared Function DetectEncoding(name As String) As Encoding
        Select Case If(name, "utf-8").Trim().ToLowerInvariant()
            Case "utf-8-bom"
                Return New UTF8Encoding(True)
            Case "utf-8"
                Return New UTF8Encoding(False)
            Case Else
                Try
                    Return Encoding.GetEncoding(name)
                Catch
                    Return New UTF8Encoding(False)
                End Try
        End Select
    End Function

    Private Shared Function ResolveCsvDelimiter(useCase As DataCollectorUseCaseConfig) As Char
        Dim delimiter As String = If(useCase.csv?.delimiter, ";")
        If delimiter = "" Then
            Return ";"c
        End If
        Return delimiter(0)
    End Function

    Private Shared Function ResolveCsvQuoteMode(useCase As DataCollectorUseCaseConfig) As String
        Return If(useCase.csv?.quoteMode, "whenNeeded")
    End Function

    Private Shared Function ResolveCsvNewLine(useCase As DataCollectorUseCaseConfig) As String
        If String.Equals(If(useCase.csv?.newline, "CRLF"), "LF", StringComparison.OrdinalIgnoreCase) Then
            Return vbLf
        End If
        Return vbCrLf
    End Function

    Private Function BuildCsvContent(records As List(Of JObject),
                                     columns As List(Of String),
                                     delimiter As Char,
                                     includeHeader As Boolean,
                                     quoteMode As String,
                                     newline As String,
                                     useCase As DataCollectorUseCaseConfig) As String
        Dim sb As New StringBuilder()

        If includeHeader Then
            sb.Append(SerializeCsvRow(columns, delimiter, quoteMode))
            If records.Count > 0 Then
                sb.Append(newline)
            End If
        End If

        For i As Integer = 0 To records.Count - 1
            sb.Append(SerializeCsvRow(
                columns.Select(Function(name) GetCsvSerializableValue(records(i)(name), useCase)).ToList(),
                delimiter,
                quoteMode))

            If i < records.Count - 1 Then
                sb.Append(newline)
            End If
        Next

        Return sb.ToString()
    End Function

    Private Function GetCsvSerializableValue(token As JToken, useCase As DataCollectorUseCaseConfig) As String
        Dim value As String = If(token Is Nothing OrElse token.Type = JTokenType.Null, "", token.ToString())

        If useCase.csv IsNot Nothing AndAlso useCase.csv.protectAgainstFormulaInjection Then
            If value.StartsWith("=") OrElse value.StartsWith("+") OrElse value.StartsWith("-") OrElse value.StartsWith("@") Then
                value = "'" & value
            End If
        End If

        Return value
    End Function

    Private Shared Function SerializeCsvRow(values As IList(Of String),
                                            delimiter As Char,
                                            quoteMode As String) As String
        Dim parts As New List(Of String)()

        For Each value In values
            Dim text As String = If(value, "")
            Dim mustQuote As Boolean =
                String.Equals(quoteMode, "always", StringComparison.OrdinalIgnoreCase) OrElse
                text.Contains(delimiter) OrElse
                text.Contains(""""c) OrElse
                text.Contains(vbCr) OrElse
                text.Contains(vbLf)

            If mustQuote Then
                text = """" & text.Replace("""", """""") & """"
            End If

            parts.Add(text)
        Next

        Return String.Join(delimiter, parts)
    End Function

    Private Shared Function ParseCsv(content As String, delimiter As Char) As List(Of List(Of String))
        Dim rows As New List(Of List(Of String))()
        Dim row As New List(Of String)()
        Dim field As New StringBuilder()
        Dim inQuotes As Boolean = False
        Dim i As Integer = 0

        While i < content.Length
            Dim ch As Char = content(i)

            If ch = """"c Then
                If inQuotes AndAlso i + 1 < content.Length AndAlso content(i + 1) = """"c Then
                    field.Append(""""c)
                    i += 1
                Else
                    inQuotes = Not inQuotes
                End If
            ElseIf ch = delimiter AndAlso Not inQuotes Then
                row.Add(field.ToString())
                field.Clear()
            ElseIf (ch = vbCr(0) OrElse ch = vbLf(0)) AndAlso Not inQuotes Then
                row.Add(field.ToString())
                field.Clear()
                rows.Add(New List(Of String)(row))
                row.Clear()

                If ch = vbCr(0) AndAlso i + 1 < content.Length AndAlso content(i + 1) = vbLf(0) Then
                    i += 1
                End If
            Else
                field.Append(ch)
            End If

            i += 1
        End While

        If inQuotes Then
            row.Add(field.ToString())
            rows.Add(New List(Of String)(row))
            Return rows
        End If

        If field.Length > 0 OrElse row.Count > 0 Then
            row.Add(field.ToString())
            rows.Add(New List(Of String)(row))
        End If

        Return rows
    End Function

    Private Shared Function GetConfiguredOutputColumnNames(useCase As DataCollectorUseCaseConfig) As List(Of String)
        Dim names As New List(Of String)()

        If useCase.fields IsNot Nothing Then
            For Each field In useCase.fields
                names.Add(field.name)
            Next
        End If

        If useCase.extraction IsNot Nothing AndAlso useCase.extraction.includeOriginalRequest Then
            names.Add(If(useCase.extraction.originalRequestFieldName, "originalRequest"))
        End If

        Return names
    End Function

    Private Shared Function NormalizeJsonlInvalidLinePolicy(value As String) As String
        Select Case If(value, "reject").Trim().ToLowerInvariant()
            Case "ignore", "ignoreinvalidlines", "ignore_invalid_lines"
                Return "ignore"
            Case "failduplicatecheckbutallowappend", "allowappend", "allow_append_warn"
                Return "allow_append_warn"
            Case Else
                Return "reject"
        End Select
    End Function

    Private Shared Function NormalizeFullPath(path As String) As String
        If String.IsNullOrWhiteSpace(path) Then
            Return ""
        End If

        Return System.IO.Path.GetFullPath(path.Trim())
    End Function

    Private Shared Function IsPathWithinBaseDirectory(candidatePath As String, baseDirectory As String) As Boolean
        Dim candidate As String = NormalizeFullPath(candidatePath)
        Dim baseDir As String = NormalizeFullPath(baseDirectory)

        If candidate = "" OrElse baseDir = "" Then
            Return False
        End If

        If Not baseDir.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal) Then
            baseDir &= Path.DirectorySeparatorChar
        End If

        Return candidate.StartsWith(baseDir, StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(candidate.TrimEnd(Path.DirectorySeparatorChar), baseDir.TrimEnd(Path.DirectorySeparatorChar), StringComparison.OrdinalIgnoreCase)
    End Function

    Private Shared Function ExtractSenderDomain(address As String) As String
        Try
            Dim value As String = If(address, "").Trim()
            Dim atIndex As Integer = value.LastIndexOf("@"c)
            If atIndex >= 0 AndAlso atIndex < value.Length - 1 Then
                Return value.Substring(atIndex + 1)
            End If
        Catch
        End Try

        Return ""
    End Function

    Private Shared Function SanitizeFileNameComponent(value As String) As String
        Dim text As String = If(value, "")

        For Each ch In Path.GetInvalidFileNameChars()
            text = text.Replace(ch, "_"c)
        Next

        text = text.Replace("..", "_")
        text = text.Replace("/", "_").Replace("\", "_").Trim()

        If text = "" Then
            text = "value"
        End If

        Return text
    End Function

    Private Shared Function TokenToPlainObject(token As JToken) As Object
        If token Is Nothing Then
            Return Nothing
        End If

        Select Case token.Type
            Case JTokenType.Null
                Return Nothing
            Case JTokenType.Integer
                Return token.Value(Of Long)()
            Case JTokenType.Float
                Return token.Value(Of Decimal)()
            Case JTokenType.Boolean
                Return token.Value(Of Boolean)()
            Case Else
                Return token.ToObject(Of Object)()
        End Select
    End Function

    Private Shared Function CreateMessage(code As String,
                                          message As String,
                                          Optional field As String = Nothing,
                                          Optional originalValue As Object = Nothing,
                                          Optional normalizedValue As Object = Nothing,
                                          Optional details As Object = Nothing,
                                          Optional recordKey As Object = Nothing) As DataCollectorMessageItem
        Return New DataCollectorMessageItem With {
            .code = code,
            .message = message,
            .field = field,
            .originalValue = originalValue,
            .normalizedValue = normalizedValue,
            .details = details,
            .recordKey = recordKey
        }
    End Function

    Private Shared Function SerializeJson(token As JToken) As String
        Return token.ToString(Newtonsoft.Json.Formatting.None)
    End Function

    Private Function BuildOperationResultFailure(status As String,
                                                 errors As List(Of DataCollectorMessageItem),
                                                 Optional warnings As List(Of DataCollectorMessageItem) = Nothing,
                                                 Optional normalization As List(Of DataCollectorNormalizationItem) = Nothing,
                                                 Optional useCaseId As String = Nothing) As DataCollectorOperationResult
        Dim obj As New JObject(
            New JProperty("success", False),
            New JProperty("status", status),
            New JProperty("useCaseId", If(useCaseId, "")),
            New JProperty("errors", JArray.FromObject(If(errors, New List(Of DataCollectorMessageItem)()))),
            New JProperty("warnings", JArray.FromObject(If(warnings, New List(Of DataCollectorMessageItem)()))))

        If normalization IsNot Nothing Then
            obj("normalization") = JArray.FromObject(normalization)
        End If

        Return New DataCollectorOperationResult With {
            .success = False,
            .responseJson = SerializeJson(obj),
            .errorMessage = If(errors IsNot Nothing AndAlso errors.Count > 0, errors(0).message, status),
            .errorCode = status
        }
    End Function

    Private Shared Sub EnsureDirectoryExists(directoryPath As String, defaults As DataCollectorDefaultsConfig)
        If String.IsNullOrWhiteSpace(directoryPath) Then
            Exit Sub
        End If

        If defaults Is Nothing OrElse defaults.createDirectories Then
            Directory.CreateDirectory(directoryPath)
        ElseIf Not Directory.Exists(directoryPath) Then
            Throw New DirectoryNotFoundException($"Directory not found: {directoryPath}")
        End If
    End Sub

    Private Shared Sub AppendText(path As String, content As String, encoding As Encoding)
        Using fs As New FileStream(path, FileMode.Append, FileAccess.Write, FileShare.Read)
            Using writer As New StreamWriter(fs, encoding)
                writer.Write(content)
                writer.Flush()
            End Using
        End Using
    End Sub

    Private Shared Sub WriteAllTextAtomic(targetPath As String,
                                          content As String,
                                          encoding As Encoding,
                                          atomicWrites As Boolean,
                                          createBackup As Boolean)
        Dim directoryPath As String = Path.GetDirectoryName(targetPath)
        If Not Directory.Exists(directoryPath) Then
            Directory.CreateDirectory(directoryPath)
        End If

        If Not atomicWrites Then
            File.WriteAllText(targetPath, content, encoding)
            Exit Sub
        End If

        Dim tempPath As String = Path.Combine(directoryPath, Path.GetFileName(targetPath) & ".tmp." & Guid.NewGuid().ToString("N"))
        File.WriteAllText(tempPath, content, encoding)

        If File.Exists(targetPath) Then
            If createBackup Then
                Dim backupPath As String = targetPath & ".bak"
                File.Replace(tempPath, targetPath, backupPath, True)
            Else
                File.Replace(tempPath, targetPath, Nothing, True)
            End If
        Else
            File.Move(tempPath, targetPath)
        End If

        If File.Exists(tempPath) Then
            Try
                File.Delete(tempPath)
            Catch
            End Try
        End If
    End Sub


End Class
