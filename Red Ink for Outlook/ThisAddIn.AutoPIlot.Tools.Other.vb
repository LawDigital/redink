' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.Tools.Other.vb
' Purpose:
'   Defines and executes AutoPilot internal tools for miscellaneous operations
'   within Outlook AutoPilot Chat-Agent runs, including attachment inspection,
'   search, content generation, and utility operations.
'
' Tools Provided:
'   - read_attachment: Reads and returns text content from supported attachments
'     (DOCX, PDF, TXT, CSV, HTML, XML, JSON, XLSX, PPTX, .msg/.eml with unpacking)
'   - list_attachments: Lists all attachments and tool-generated output files
'     with metadata (filename, type, size, processing status)
'   - search_in_attachments: Searches for text terms across attachments with
'     context extraction and file-by-file result reporting
'   - generate_image: Generates images via AI model from text descriptions with
'     optional reference image for editing
'   - create_audio_file: Creates audio files via text-to-speech processing
'   - web_grounding: Delegates to the model's native web-search capability
'   - manage_scheduled_tasks: CRUD operations (create, list, get, update, delete,
'     pause, resume) for the AutoPilot Scheduler
'   - manage_user_memory: Manages persistent user/context memory
'   - manage_user_files: Manages workspace file operations and staging
'   - complete_word_tables: Auto-completes table cells in Word attachments
'   - report_inability: Reports when a requested operation is unavailable
'
' Tool Interface Architecture:
'   - Registration:
'       * Tools are exposed as `ModelConfig` entries (`Tool=True`, `ToolOnly=True`)
'         so they participate in the same tool-calling pipeline as external tools.
'       * Tool metadata (`ToolDefinition`, `ToolInstructionsPrompt`) is generated
'         inline and consumed by `ExecuteToolCall` / `ExecuteToolingLoop`.
'   - Dispatch:
'       * `TryExecuteAutoPilotTool` routes parsed tool calls to strongly scoped
'         executor methods (`ExecuteReadAttachmentTool`, `ExecuteListAttachmentsTool`,
'         `ExecuteSearchInAttachmentsTool`, `ExecuteGenerateImageTool`, etc.) and
'         returns `ToolResponse` payloads.
'   - Session scope:
'       * All tools use AutoPilot session state from `ThisAddIn.Autopilot.vb`:
'           - `_apCurrentAttachments`: attachment registry for input/output lookups
'           - `_apCurrentTempDir`: per-mail temp directory for file creation
'           - `_apCurrentMailInfo`: metadata about the current email session
'       * Supports tool chaining via output registration (`OutputFiles`) and
'         attachment lookup via `FindAttachment` (original + prior tool outputs).
'   - Content extraction:
'       * Multi-format support with cached text reuse (`CachedText`, `CachedDocxHint`).
'       * Automatic .msg/.eml unpacking with recursive attachment extraction.
'       * Embedded file references are resolved and made available as separate
'         attachments for subsequent tool operations.
'   - AI integration:
'       * `generate_image` and `create_audio_file` switch to alternate models
'         (ImageGeneration, AudioGeneration) when configured.
'       * `web_grounding` delegates to native model web-search capability.
'   - Scheduler integration:
'       * `manage_scheduled_tasks` provides CRUD access to AutoPilot Scheduler
'         via `SchedulerCreateTask` / `SchedulerListTasks` / `SchedulerFindTask`
'         from `ThisAddIn.AutoPilot.Scheduler.vb`.
'   - Error handling:
'       * Returns structured `ToolResponse` with success flag, message, and
'         error details. Missing/unavailable tools report gracefully.
'   - Logging and UX:
'       * Emits execution traces to tooling context (`context.Log`) and
'         AutoPilot dashboard (`ApDashboardLog`) with concise status summaries.
'
' Security & Safety:
'   - Path containment:
'       * All tool outputs are created in `_apCurrentTempDir` and re-used only
'         via resolved attachment/output references.
'   - File validation:
'       * Size checks prevent oversized attachments from processing.
'       * Format detection prevents mishandling of unknown file types.
'   - Model access control:
'       * Image and audio generation only proceed if alternate models are
'         configured; otherwise, graceful error messages are returned.
'
' =============================================================================



Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.IO.Compression
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Xml
Imports Microsoft.Office.Interop.Outlook
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn





    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: generate_image
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Generates an image by switching to the ImageGeneration special task model,
    ''' calling the LLM with the user's description (no predefined prompt), and
    ''' saving the result to the per-mail temp directory for attachment.
    ''' The LLM/HandleObject pipeline decodes the image from the JSON response
    ''' and saves it via <see cref="ImageDecoder.DecodeAndSaveImage"/>.
    ''' </summary>
    Private Async Function ExecuteGenerateImageTool(
            toolCall As ToolCall,
            context As ToolExecutionContext,
            ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim description = GetArgString(toolCall.Arguments, "description")
            If String.IsNullOrWhiteSpace(description) Then
                response.Success = False
                response.Response = "Missing required parameter: description"
                Return response
            End If

            Dim outputFileName = GetArgString(toolCall.Arguments, "output_filename")

            context.Log($"Generating image: {If(description.Length > 120, description.Substring(0, 120) & "...", description)}")
            ApDashboardLog("🎨 Generating image...", "step")

            ' ── Switch to ImageGeneration model ──
            If String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                response.Success = False
                response.Response = "No alternate model configuration file is configured. Image generation requires an ImageGeneration model."
                Return response
            End If

            Dim backupConfig As ModelConfig = GetCurrentConfig(_context)
            Dim previousUseSecondApi As Boolean = _apUseSecondApi
            Dim modelSwitched As Boolean = False

            Try
                modelSwitched = GetSpecialTaskModel(_context, INI_AlternateModelPath, "ImageGeneration")
            Catch
            End Try

            If Not modelSwitched Then
                response.Success = False
                response.Response = "No ImageGeneration model is configured in the alternate models file. Cannot generate images."
                Return response
            End If

            Try
                _apUseSecondApi = True

                ' ── Resolve optional reference image for editing ──
                Dim referenceImagePath As String = Nothing
                Dim refImageName = GetArgString(toolCall.Arguments, "image_attachment_name")
                If Not String.IsNullOrWhiteSpace(refImageName) Then
                    Dim refAtt = FindAttachment(refImageName)
                    If refAtt IsNot Nothing AndAlso refAtt.TempFilePath IsNot Nothing AndAlso
                       File.Exists(refAtt.TempFilePath) Then
                        referenceImagePath = refAtt.TempFilePath
                        context.Log($"Using reference image for editing: {refImageName}")
                        ApDashboardLog($"🖼 Reference image: {refImageName}", "step")
                    Else
                        context.Log($"Reference image '{refImageName}' not found — generating from scratch")
                        ApDashboardLog($"⚠ Reference image '{refImageName}' not found, generating from scratch", "warn")
                    End If
                End If

                ' ── Call LLM with just the description — no system prompt wrapping ──
                ' The binaryOutputDirectory directs ImageDecoder.DecodeAndSaveImage
                ' to save the image into the per-mail temp directory.
                Dim llmResult = Await LLM(
                    "", description,
                    UseSecondAPI:=True,
                    HideSplash:=True,
                    EnsureUI:=False,
                    cancellationToken:=ct,
                    binaryOutputDirectory:=_apCurrentTempDir,
                    FileObject:=If(referenceImagePath, ""))

                ' ── Locate the saved image file ──
                ' ImageDecoder.DecodeAndSaveImage returns the path in the LLM response
                ' as "Image saved to: <path>". We also scan the temp dir for new image files.
                Dim savedImagePath As String = Nothing

                ' Strategy 1: Parse the "Image saved to:" path from the LLM result
                If Not String.IsNullOrWhiteSpace(llmResult) Then
                    Dim imgMatch = System.Text.RegularExpressions.Regex.Match(
                        llmResult, "Image saved to:\s*(.+?)(?:\r?\n|$)")
                    If imgMatch.Success Then
                        Dim candidate = imgMatch.Groups(1).Value.Trim().Replace("\\", "\")
                        If File.Exists(candidate) Then
                            savedImagePath = candidate
                        End If
                    End If
                End If

                ' Strategy 2: Scan temp dir for the newest AI_Image_* file
                If savedImagePath Is Nothing AndAlso Directory.Exists(_apCurrentTempDir) Then
                    Dim imageFiles = Directory.GetFiles(_apCurrentTempDir, "AI_Image_*.*")
                    If imageFiles.Length > 0 Then
                        savedImagePath = imageFiles.OrderByDescending(Function(f) File.GetCreationTimeUtc(f)).First()
                    End If
                End If

                If savedImagePath Is Nothing OrElse Not File.Exists(savedImagePath) Then
                    response.Success = False
                    response.Response = "The image generation model did not return a valid image. " &
                        "The model may not support image output, or the response format is not recognized."
                    ApDashboardLog("⚠ Image generation: no image found in response", "warn")
                    Return response
                End If

                ' ── Optionally rename the file to the user's requested filename ──
                If Not String.IsNullOrWhiteSpace(outputFileName) Then
                    ' Sanitize
                    For Each c In Path.GetInvalidFileNameChars()
                        outputFileName = outputFileName.Replace(c, "_"c)
                    Next
                    ' Preserve the original extension from the generated file
                    Dim ext = Path.GetExtension(savedImagePath)
                    If Not outputFileName.EndsWith(ext, StringComparison.OrdinalIgnoreCase) Then
                        outputFileName &= ext
                    End If
                    Dim renamedPath = Path.Combine(_apCurrentTempDir, outputFileName)
                    ' Handle collision
                    Dim counter = 1
                    While File.Exists(renamedPath)
                        Dim baseName = Path.GetFileNameWithoutExtension(outputFileName)
                        renamedPath = Path.Combine(_apCurrentTempDir, baseName & $"_{counter}" & ext)
                        counter += 1
                    End While
                    Try
                        File.Move(savedImagePath, renamedPath)
                        savedImagePath = renamedPath
                    Catch
                        ' Keep original name on move failure
                    End Try
                End If

                ' ── Register as output file for attachment to reply ──
                If _apCurrentAttachments IsNot Nothing AndAlso _apCurrentAttachments.Count > 0 Then
                    _apCurrentAttachments(0).OutputFiles.Add(savedImagePath)
                End If

                Dim finalFileName = Path.GetFileName(savedImagePath)
                Dim sizeKb = New FileInfo(savedImagePath).Length / 1024

                response.Success = True
                response.Response = $"Image generated: {finalFileName} ({sizeKb:F0} KB). The file will be attached to the reply."
                ApDashboardLog($"✓ Image generated: {finalFileName} ({sizeKb:F0} KB)", "info")

            Finally
                ' ── Restore the original model configuration ──
                _apUseSecondApi = previousUseSecondApi
                If backupConfig IsNot Nothing Then
                    RestoreDefaults(_context, backupConfig)
                End If
            End Try

        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation was cancelled."
            response.Response = response.ErrorMessage
        Catch ex As System.Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error generating image: {ex.Message}"
            ApDashboardLog($"⚠ Image generation error: {ex.Message}", "warn")
        End Try

        Return response
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: web_grounding
    ' ═══════════════════════════════════════════════════════════════════════════


    Private Async Function ExecuteWebGroundingTool(
        toolCall As ToolCall,
        context As ToolExecutionContext,
        Optional cancellationToken As CancellationToken = Nothing) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
        .CallId = toolCall.CallId,
        .ToolName = toolCall.ToolName,
        .Timestamp = DateTime.UtcNow,
        .OriginalCallJson = toolCall.RawJson
    }

        Try
            response.Response =
            Await SharedLibrary.Agents.WebGroundingTool.ExecuteAsync(
                _context,
                toolCall.Arguments,
                cancellationToken,
                logStep:=Sub(message)
                             context.Log(message)
                             ApDashboardLog(message, "step")
                         End Sub,
                logInfo:=Sub(message)
                             context.Log(message)
                             ApDashboardLog("✓ " & message, "info")
                         End Sub,
                logWarn:=Sub(message)
                             context.Log(message, "warn")
                             ApDashboardLog("⚠ " & message, "warn")
                         End Sub)

            response.Success = Not String.IsNullOrWhiteSpace(response.Response)

            If Not response.Success Then
                response.ErrorMessage = "web_grounding returned no usable result."
                response.Response = response.ErrorMessage
            End If

        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation was cancelled."
            response.Response = response.ErrorMessage

        Catch ex As System.Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error during web grounding: {ex.Message}"
            context.Log(response.Response, "warn")
            ApDashboardLog($"⚠ Web grounding error: {ex.Message}", "warn")
        End Try

        Return response
    End Function




    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: describe_binary_attachment
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Async Function ExecuteDescribeBinaryTool(
            toolCall As ToolCall, context As ToolExecutionContext, ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim fileName = GetArgString(toolCall.Arguments, "attachment_name")
            Dim prompt = If(GetArgString(toolCall.Arguments, "prompt"), "Describe or transcribe this file.")

            If String.IsNullOrWhiteSpace(fileName) Then
                response.Success = False
                response.Response = "Missing required parameter: attachment_name"
                Return response
            End If

            Dim att = FindAttachment(fileName)
            If att Is Nothing Then response.Success = False : response.Response = $"Attachment '{fileName}' not found." : Return response
            If att.IsOverSizeLimit Then response.Success = False : response.Response = $"Attachment '{fileName}' exceeds the size limit." : Return response
            If att.TempFilePath Is Nothing OrElse Not File.Exists(att.TempFilePath) Then
                response.Success = False : response.Response = $"Attachment '{fileName}' could not be read." : Return response
            End If

            Dim ext As String = Path.GetExtension(att.TempFilePath).ToLowerInvariant()
            If Not IsBinaryMediaExtension(ext) Then
                response.Success = False
                response.Response = $"The file format '{ext}' is not supported for binary analysis. " &
                    "Supported formats include images (.png, .jpg, .gif, .webp, .tiff), " &
                    "audio (.mp3, .wav, .ogg, .m4a, .flac, .aac), and video (.mp4, .mov, .webm)."
                Return response
            End If

            context.Log($"Analyzing binary attachment: {fileName}")
            ApDashboardLog($"🖼 Sending binary attachment to AI: {fileName} ({ext})", "step")

            Dim useSecond As Boolean = (_apConfig IsNot Nothing AndAlso _apConfig.UseSecondApi)
            Dim llmResult As String = Await SharedMethods.LLM(
                _context, prompt, "", UseSecondAPI:=useSecond, Hidesplash:=True,
                FileObject:=att.TempFilePath, cancellationToken:=ct)

            If String.IsNullOrWhiteSpace(llmResult) Then
                response.Success = False
                response.Response = $"The AI model could not process the file '{fileName}'. The model may not support this file type."
                ApDashboardLog($"⚠ Binary analysis returned no result for: {fileName}", "warn")
                Return response
            End If

            response.Success = True
            response.Response = $"Analysis of '{fileName}':" & vbCrLf & llmResult
            ApDashboardLog($"✓ Binary analysis completed for: {fileName} ({llmResult.Length:N0} chars)", "info")

        Catch ex As System.Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error analyzing binary attachment: {ex.Message}"
        End Try

        Return response
    End Function


    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: read_attachment 
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Async Function ExecuteReadAttachmentTool(
            toolCall As ToolCall, context As ToolExecutionContext, ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            ' Support both single and batch mode
            Dim fileNames As New List(Of String)()
            Dim singleName = GetArgString(toolCall.Arguments, "attachment_name")
            Dim batchNames = GetArgStringArray(toolCall.Arguments, "attachment_names")

            If batchNames.Count > 0 Then
                fileNames.AddRange(batchNames)
            ElseIf Not String.IsNullOrWhiteSpace(singleName) Then
                fileNames.Add(singleName)
            End If

            If fileNames.Count = 0 Then
                response.Success = False
                response.Response = "Missing required parameter: attachment_name or attachment_names"
                Return response
            End If

            Dim sb As New StringBuilder()
            Dim anySuccess As Boolean = False

            For Each fileName In fileNames
                Dim att = FindAttachment(fileName)

                If att Is Nothing Then
                    sb.AppendLine($"[{fileName}]")
                    sb.AppendLine($"Attachment '{fileName}' not found.")
                    sb.AppendLine()
                    Continue For
                End If

                If att.IsOverSizeLimit Then
                    sb.AppendLine($"[{fileName}]")
                    sb.AppendLine($"Attachment '{fileName}' exceeds the size limit and cannot be processed.")
                    sb.AppendLine()
                    Continue For
                End If

                If att.TempFilePath Is Nothing OrElse Not File.Exists(att.TempFilePath) Then
                    sb.AppendLine($"[{fileName}]")
                    sb.AppendLine($"Attachment '{fileName}' could not be read.")
                    sb.AppendLine()
                    Continue For
                End If

                context.Log($"Reading attachment: {fileName}")
                Dim text = Await ReadSingleAttachmentText(att, context)

                If Not String.IsNullOrWhiteSpace(text) Then
                    If text.Length > 50000 Then
                        text = text.Substring(0, 50000) & vbCrLf & "[... content truncated at 50,000 characters ...]"
                    End If

                    If fileNames.Count > 1 Then sb.AppendLine($"[{fileName}]")
                    sb.AppendLine(text)

                    ' Append docx metadata hint if applicable
                    Dim hint = GetDocxMetadataHint(att)
                    If Not String.IsNullOrWhiteSpace(hint) Then sb.AppendLine(hint)

                    sb.AppendLine()
                    anySuccess = True
                Else
                    sb.AppendLine($"[{fileName}]")
                    sb.AppendLine($"Could not extract text from '{fileName}'. The file format may not be supported.")
                    sb.AppendLine()
                End If
            Next

            response.Success = anySuccess
            response.Response = sb.ToString().TrimEnd()
            If Not anySuccess Then
                response.Response = If(fileNames.Count = 1,
                    $"Could not extract text from '{fileNames(0)}'. The file format may not be supported.",
                    response.Response)
            End If

        Catch ex As System.Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error reading attachment: {ex.Message}"
        End Try

        Return response
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: list_attachments
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Function ExecuteListAttachmentsTool(toolCall As ToolCall, context As ToolExecutionContext) As ToolResponse
        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            If _apCurrentAttachments Is Nothing OrElse _apCurrentAttachments.Count = 0 Then
                response.Success = True
                response.Response = "No attachments in this email."
                Return response
            End If

            Dim sb As New StringBuilder()
            sb.AppendLine($"Attachments ({_apCurrentAttachments.Count}):")
            For i As Integer = 0 To _apCurrentAttachments.Count - 1
                Dim att = _apCurrentAttachments(i)
                Dim sizeStr = If(att.SizeBytes > 0, $"{att.SizeBytes / 1024:F0} KB", "unknown size")
                Dim statusStr = If(att.IsOverSizeLimit, " [OVER SIZE LIMIT]",
                               If(att.TempFilePath IsNot Nothing, " [available for processing]", " [not available]"))
                Dim pdfStr As String = ""
                If att.PageCount > 0 Then
                    pdfStr = $", {att.PageCount} page(s)"
                    If Not String.IsNullOrWhiteSpace(att.PageOrientation) Then
                        pdfStr &= $", {att.PageOrientation}"
                    End If
                    If Not String.IsNullOrWhiteSpace(att.PageSize) Then
                        pdfStr &= $", {att.PageSize}"
                    End If
                End If
                sb.AppendLine($"  {i + 1}. {att.OriginalFileName} ({att.Extension}, {sizeStr}{pdfStr}){statusStr}")
            Next

            ' List output files produced by earlier tool calls
            Dim outputFileCount = 0
            For Each att In _apCurrentAttachments
                If att.OutputFiles IsNot Nothing Then
                    For Each outputPath In att.OutputFiles
                        If Not String.IsNullOrEmpty(outputPath) AndAlso File.Exists(outputPath) Then
                            outputFileCount += 1
                        End If
                    Next
                End If
            Next

            If outputFileCount > 0 Then
                sb.AppendLine()
                sb.AppendLine($"Tool output files ({outputFileCount}):")
                Dim outIdx = 1
                For Each att In _apCurrentAttachments
                    If att.OutputFiles Is Nothing Then Continue For
                    For Each outputPath In att.OutputFiles
                        If Not String.IsNullOrEmpty(outputPath) AndAlso File.Exists(outputPath) Then
                            Dim outName = Path.GetFileName(outputPath)
                            Dim outExt = Path.GetExtension(outputPath).ToLowerInvariant()
                            Dim outSize = New FileInfo(outputPath).Length
                            sb.AppendLine($"  {outIdx}. {outName} ({outExt}, {outSize / 1024:F0} KB) [tool output — available for processing]")
                            outIdx += 1
                        End If
                    Next
                Next
            End If

            response.Success = True
            response.Response = sb.ToString().TrimEnd()

        Catch ex As System.Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error listing attachments: {ex.Message}"
        End Try

        Return response
    End Function





    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: search_in_attachments
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Async Function ExecuteSearchInAttachmentsTool(
            toolCall As ToolCall, context As ToolExecutionContext, ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim searchTerm = GetArgString(toolCall.Arguments, "search_term")
            If String.IsNullOrWhiteSpace(searchTerm) Then
                response.Success = False
                response.Response = "Missing required parameter: search_term"
                Return response
            End If

            Dim targetNames = GetArgStringArray(toolCall.Arguments, "attachment_names")

            Dim toSearch As List(Of AutoPilotAttachmentInfo)
            If targetNames.Count > 0 Then
                toSearch = _apCurrentAttachments?.Where(
                    Function(a) targetNames.Any(Function(n) a.OriginalFileName.Equals(n, StringComparison.OrdinalIgnoreCase)) AndAlso
                                Not a.IsOverSizeLimit AndAlso a.TempFilePath IsNot Nothing).ToList()
            Else
                toSearch = _apCurrentAttachments?.Where(
                    Function(a) Not a.IsOverSizeLimit AndAlso a.TempFilePath IsNot Nothing AndAlso
                                Not IsBinaryMediaExtension(a.Extension)).ToList()
            End If

            If toSearch Is Nothing OrElse toSearch.Count = 0 Then
                response.Success = False
                response.Response = "No searchable attachments found."
                Return response
            End If

            context.Log($"Searching for '{searchTerm}' in {toSearch.Count} attachment(s)")
            ApDashboardLog($"🔍 Searching for: {searchTerm}", "step")

            Dim sb As New StringBuilder()
            Dim totalMatches = 0

            For Each att In toSearch
                Dim text = Await ReadSingleAttachmentText(att, context)
                If String.IsNullOrWhiteSpace(text) Then Continue For

                Dim lines = text.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
                Dim matchLines As New List(Of String)()

                For i = 0 To lines.Length - 1
                    If lines(i).IndexOf(searchTerm, StringComparison.OrdinalIgnoreCase) >= 0 Then
                        Dim lineNum = i + 1
                        Dim excerpt = lines(i).Trim()
                        If excerpt.Length > 200 Then excerpt = excerpt.Substring(0, 200) & "..."
                        matchLines.Add($"  Line {lineNum}: {excerpt}")
                    End If
                Next

                If matchLines.Count > 0 Then
                    sb.AppendLine($"[{att.OriginalFileName}] — {matchLines.Count} match(es)")
                    For Each ml In matchLines.Take(20)
                        sb.AppendLine(ml)
                    Next
                    If matchLines.Count > 20 Then sb.AppendLine($"  ... and {matchLines.Count - 20} more match(es)")
                    sb.AppendLine()
                    totalMatches += matchLines.Count
                End If
            Next

            If totalMatches > 0 Then
                response.Success = True
                response.Response = $"Found {totalMatches} match(es) for '{searchTerm}':" & vbCrLf & sb.ToString().TrimEnd()
            Else
                response.Success = True
                response.Response = $"No matches found for '{searchTerm}' in any attachment."
            End If

        Catch ex As System.Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error searching attachments: {ex.Message}"
        End Try

        Return response
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: summarize_thread
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Function ExecuteSummarizeThreadTool(toolCall As ToolCall, context As ToolExecutionContext) As ToolResponse
        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            If _apCurrentMailInfo Is Nothing Then
                response.Success = False
                response.Response = "No current mail context available."
                Return response
            End If

            context.Log("Extracting email thread")
            ApDashboardLog("📧 Extracting email thread", "step")

            ' Get the monitored mailbox address to exclude autopilot's own messages
            Dim monitoredMailbox = If(_apConfig?.MonitoredMailbox, "").Trim().ToLowerInvariant()

            Dim body = _apCurrentMailInfo.Body
            If String.IsNullOrWhiteSpace(body) Then
                response.Success = True
                response.Response = "The email has no body text."
                Return response
            End If

            ' Parse the email thread by looking for common forwarding/reply separators
            Dim sb As New StringBuilder()
            Dim threadMessages As New List(Of (Sender As String, DateStr As String, Body As String))

            ' Add the current message first
            threadMessages.Add((_apCurrentMailInfo.SenderName & " <" & _apCurrentMailInfo.SenderEmail & ">",
                               _apCurrentMailInfo.ReceivedTime.ToString("yyyy-MM-dd HH:mm"),
                               ""))

            ' Split body on common thread separators
            Dim separatorPatterns = {
                "-----Original Message-----",
                "-----Ursprüngliche Nachricht-----",
                "-----Message d'origine-----",
                "-----Messaggio originale-----",
                "________________________________",
                "From: ", "Von: ", "De : ", "Da: "
            }

            ' Simple thread extraction: split on separator patterns
            Dim currentBody As New StringBuilder()
            Dim lines = body.Split({vbCrLf, vbLf}, StringSplitOptions.None)
            Dim msgIdx = 0

            For Each line In lines
                Dim isSeparator = False
                For Each sep In separatorPatterns
                    If line.TrimStart().StartsWith(sep, StringComparison.OrdinalIgnoreCase) Then
                        ' Save current body to current message
                        If msgIdx < threadMessages.Count Then
                            Dim msg = threadMessages(msgIdx)
                            msg.Body = currentBody.ToString().Trim()
                            threadMessages(msgIdx) = msg
                        End If
                        currentBody.Clear()

                        ' Try to extract sender from "From:" line
                        Dim senderLine = line.Trim()
                        If senderLine.StartsWith("From:", StringComparison.OrdinalIgnoreCase) OrElse
                           senderLine.StartsWith("Von:", StringComparison.OrdinalIgnoreCase) OrElse
                           senderLine.StartsWith("De :", StringComparison.OrdinalIgnoreCase) OrElse
                           senderLine.StartsWith("Da:", StringComparison.OrdinalIgnoreCase) Then
                            Dim senderPart = senderLine.Substring(senderLine.IndexOf(":"c) + 1).Trim()
                            threadMessages.Add((senderPart, "", ""))
                            msgIdx = threadMessages.Count - 1
                        Else
                            threadMessages.Add(("(previous message)", "", ""))
                            msgIdx = threadMessages.Count - 1
                        End If

                        isSeparator = True
                        Exit For
                    End If
                Next

                If Not isSeparator Then
                    ' Check for "Sent:" / "Date:" lines to capture date
                    Dim trimmed = line.TrimStart()
                    If trimmed.StartsWith("Sent:", StringComparison.OrdinalIgnoreCase) OrElse
                       trimmed.StartsWith("Gesendet:", StringComparison.OrdinalIgnoreCase) OrElse
                       trimmed.StartsWith("Date:", StringComparison.OrdinalIgnoreCase) OrElse
                       trimmed.StartsWith("Datum:", StringComparison.OrdinalIgnoreCase) Then
                        If msgIdx < threadMessages.Count Then
                            Dim msg = threadMessages(msgIdx)
                            msg.DateStr = trimmed.Substring(trimmed.IndexOf(":"c) + 1).Trim()
                            threadMessages(msgIdx) = msg
                        End If
                    Else
                        currentBody.AppendLine(line)
                    End If
                End If
            Next

            ' Save last body
            If msgIdx < threadMessages.Count Then
                Dim msg = threadMessages(msgIdx)
                msg.Body = currentBody.ToString().Trim()
                threadMessages(msgIdx) = msg
            End If

            ' Build output, excluding messages from/to the monitored mailbox
            sb.AppendLine($"Email Thread ({threadMessages.Count} message(s)):")
            sb.AppendLine()

            Dim displayIdx = 1
            For Each msg In threadMessages
                ' Exclude monitored mailbox messages
                If Not String.IsNullOrWhiteSpace(monitoredMailbox) AndAlso
                   msg.Sender.ToLowerInvariant().Contains(monitoredMailbox) Then
                    Continue For
                End If

                sb.AppendLine($"── Message {displayIdx} ──")
                sb.AppendLine($"From: {msg.Sender}")
                If Not String.IsNullOrWhiteSpace(msg.DateStr) Then sb.AppendLine($"Date: {msg.DateStr}")
                sb.AppendLine()
                If Not String.IsNullOrWhiteSpace(msg.Body) Then
                    Dim bodyText = msg.Body
                    If bodyText.Length > 5000 Then
                        bodyText = bodyText.Substring(0, 5000) & vbCrLf & "[... truncated ...]"
                    End If
                    sb.AppendLine(bodyText)
                Else
                    sb.AppendLine("(no body text)")
                End If
                sb.AppendLine()
                displayIdx += 1
            Next

            response.Success = True
            response.Response = sb.ToString().TrimEnd()
            ApDashboardLog($"✓ Thread extracted: {displayIdx - 1} message(s)", "info")

        Catch ex As System.Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error extracting email thread: {ex.Message}"
        End Try

        Return response
    End Function


    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: extract_data_from_attachments
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Extracts structured/tabular data from one or more attachments using the
    ''' <see cref="SharedLibrary.FactExtractionService.RunFactExtractionAsync"/> pipeline.
    ''' Returns schema + rows as JSON for downstream tool-chaining (create_excel, create_word, etc.).
    ''' </summary>
    Private Async Function ExecuteExtractDataFromAttachmentsTool(
            toolCall As ToolCall,
            context As ToolExecutionContext,
            ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim instruction = GetArgString(toolCall.Arguments, "instruction")
            If String.IsNullOrWhiteSpace(instruction) Then
                response.Success = False
                response.Response = "Missing required parameter: instruction"
                Return response
            End If

            Dim schemaSpec = GetArgString(toolCall.Arguments, "schema")
            Dim targetNames = GetArgStringArray(toolCall.Arguments, "attachment_names")
            Dim sortColumn = GetArgInt(toolCall.Arguments, "sort_column", 0)
            Dim sortDirection = If(GetArgString(toolCall.Arguments, "sort_direction"), "ASC").Trim().ToUpperInvariant()
            If sortDirection <> "ASC" AndAlso sortDirection <> "DESC" Then sortDirection = "ASC"
            Dim dateColumnsText = If(GetArgString(toolCall.Arguments, "date_columns"), "")

            ' ── Resolve attachments to process ──
            Dim toProcess As List(Of AutoPilotAttachmentInfo)
            If targetNames.Count > 0 Then
                toProcess = New List(Of AutoPilotAttachmentInfo)()
                For Each name In targetNames
                    Dim att = FindAttachment(name)
                    If att IsNot Nothing AndAlso Not att.IsOverSizeLimit AndAlso att.TempFilePath IsNot Nothing Then
                        toProcess.Add(att)
                    End If
                Next
            Else
                ' Process all readable non-binary attachments, plus binary media if the model supports them
                toProcess = _apCurrentAttachments?.Where(
                    Function(a) Not a.IsOverSizeLimit AndAlso a.TempFilePath IsNot Nothing AndAlso
                                (Not IsBinaryMediaExtension(a.Extension) OrElse
                                 SharedMethods.IsBinaryMediaSupported(_context, a.Extension,
                                     SharedMethods.TaskFlagForExtension(a.Extension)))).ToList()
            End If

            If toProcess Is Nothing OrElse toProcess.Count = 0 Then
                response.Success = False
                response.Response = "No processable attachments found for data extraction."
                Return response
            End If

            context.Log($"Extracting structured data from {toProcess.Count} file(s): {instruction}")
            ApDashboardLog($"📊 Fact extraction: {toProcess.Count} file(s)", "step")

            ' ── Parse optional fixed schema ──
            Dim fixedSchema As System.Collections.Generic.List(Of FactExtractionService.ExtractionSchemaColumn) = Nothing
            If Not String.IsNullOrWhiteSpace(schemaSpec) Then
                fixedSchema = FactExtractionService.ParseUserSchemaSpec(schemaSpec)
                ' Detect sort column from * marker if not explicitly set
                If (fixedSchema IsNot Nothing AndAlso fixedSchema.Count > 0) AndAlso sortColumn = 0 Then
                    sortColumn = FactExtractionService.DetectSortColumnFromSpec(schemaSpec)
                End If
            End If

            ' ── Parse date columns ──
            Dim dateCols As New System.Collections.Generic.List(Of Integer)
            For Each part In dateColumnsText.Split(New Char() {","c, ";"c}, StringSplitOptions.RemoveEmptyEntries)
                Dim n As Integer
                If Integer.TryParse(part.Trim(), n) AndAlso n > 0 Then dateCols.Add(n)
            Next

            ' ── Build file path list from resolved attachments ──
            Dim filePaths As New List(Of String)()
            For Each att In toProcess
                filePaths.Add(att.TempFilePath)
            Next

            ' ── Set up the extraction instruction for InterpolateAtRuntime ──
            Dim savedOtherPrompt = OtherPrompt
            Dim savedOutputLanguage = OutputLanguage
            Try
                OtherPrompt = instruction

                ' Let the LLM decide the output language; fall back to the user's primary language
                Dim requestedLanguage = GetArgString(toolCall.Arguments, "output_language")
                OutputLanguage = If(Not String.IsNullOrWhiteSpace(requestedLanguage), requestedLanguage.Trim(), INI_Language1)

                Dim useSecondApi As Boolean = (_apConfig IsNot Nothing AndAlso _apConfig.UseSecondApi)
                Dim cancelled As Boolean = False

                ' ── Adapt GetFileContent for the service ──
                ' RunFactExtractionAsync expects Func(Of String, Boolean, Boolean, Boolean, Task(Of String))
                ' Parameters: (path, silent, doOcr, askUser)
                ' We bridge to the Outlook text extraction pipeline (ReadSingleAttachmentText cache + fallbacks)
                Dim getFileContentFunc As Func(Of String, Boolean, Boolean, Boolean, Task(Of String)) =
                    Async Function(filePath As String, silent As Boolean, doOcr As Boolean, askUser As Boolean) As Task(Of String)
                        ' First, try to find the file in the current attachments by path and reuse cached text
                        Dim matchedAtt = toProcess.FirstOrDefault(
                            Function(a) a.TempFilePath IsNot Nothing AndAlso
                                        a.TempFilePath.Equals(filePath, StringComparison.OrdinalIgnoreCase))
                        If matchedAtt IsNot Nothing Then
                            Dim cachedText = Await ReadSingleAttachmentText(matchedAtt, context)
                            If Not String.IsNullOrWhiteSpace(cachedText) Then Return cachedText
                        End If

                        ' Fallback: try extraction methods directly on path
                        Dim text As String = Nothing
                        Dim label As String = Nothing
                        Try
                            If TryExtractOfficeText(filePath, text, label) AndAlso Not String.IsNullOrWhiteSpace(text) Then Return text
                        Catch : End Try
                        Try
                            Dim ext = Path.GetExtension(filePath).ToLowerInvariant()
                            If ext = ".xlsx" OrElse ext = ".xls" Then
                                text = ExtractExcelText(filePath)
                                If Not String.IsNullOrWhiteSpace(text) AndAlso Not text.StartsWith("Error") Then Return text
                            ElseIf ext = ".pptx" Then
                                text = ExtractPowerPointText(filePath)
                                If Not String.IsNullOrWhiteSpace(text) AndAlso Not text.StartsWith("Error") Then Return text
                            End If
                        Catch : End Try
                        Try
                            If TryExtractTextLike(filePath, text, label) AndAlso Not String.IsNullOrWhiteSpace(text) Then Return text
                        Catch : End Try
                        If Path.GetExtension(filePath).ToLowerInvariant() = ".pdf" Then
                            Try
                                text = Await SharedMethods.ReadPdfAsText(filePath, ReturnErrorInsteadOfEmpty:=True, DoOCR:=doOcr, AskUser:=False)
                                If Not String.IsNullOrWhiteSpace(text) Then Return text
                            Catch : End Try
                        End If
                        Return ""
                    End Function

                ' ── Adapt LLM function for the service ──
                ' RunFactExtractionAsync expects Func(Of String, String, String, String, Integer, Boolean, Boolean, Task(Of String))
                Dim llmFunc As Func(Of String, String, String, String, Integer, Boolean, Boolean, Task(Of String)) =
                    Async Function(sysPrompt As String, userText As String, model As String, temp As String,
                                   timeout As Integer, useSecond As Boolean, hideSplash As Boolean) As Task(Of String)
                        Return Await ThisAddIn.LLM(sysPrompt, userText, model, temp, timeout, useSecond, True, cancellationToken:=ct)
                    End Function

                ' ── Run the extraction ──
                Dim result = Await FactExtractionService.RunFactExtractionAsync(
                    filePaths,
                    instruction,
                    dateCols,
                    sortColumn,
                    sortDirection,
                    False,          ' doOcr — we already extracted text above
                    useSecondApi,
                    _apCurrentTempDir,
                    AddressOf InterpolateAtRuntime,
                    llmFunc,
                    getFileContentFunc,
                    _context,
                    fixedSchema,
                    Nothing,        ' clampFrom
                    Nothing,        ' clampTo
                    Sub(cur, total, label)
                        ApDashboardLog($"  📊 [{cur}/{total}] {label}", "step")
                    End Sub,
                    0,              ' mergeDateColumn
                    False,          ' mergeRowsViaLlm
                    Nothing,        ' mergeInstruction
                    Function() cancelled OrElse ct.IsCancellationRequested,
                    llmWithFileFunc:=Async Function(sys, usr, mdl, tmp, tmo, use2nd, hide, fileObj)
                                         Return Await ThisAddIn.LLM(sys, usr, mdl, tmp, tmo, use2nd, True, cancellationToken:=ct, FileObject:=fileObj)
                                     End Function)

                If result Is Nothing OrElse result.Rows.Count = 0 Then
                    Dim errMsg = "No data could be extracted from the provided file(s)."
                    If result IsNot Nothing AndAlso result.Errors.Count > 0 Then
                        errMsg &= " Errors: " & String.Join("; ", result.Errors.Take(5))
                    End If
                    If result IsNot Nothing AndAlso result.FailedFileNames.Count > 0 Then
                        errMsg &= " Failed files: " & String.Join(", ", result.FailedFileNames)
                    End If
                    response.Success = False
                    response.Response = errMsg
                    ApDashboardLog($"⚠ Fact extraction returned no data", "warn")
                    Return response
                End If

                ' ── Format result as JSON for the LLM ──
                Dim jResult As New JObject()

                ' Schema array
                Dim jSchema As New JArray()
                For Each col In result.Schema
                    jSchema.Add(New JObject From {
                        {"name", col.Name},
                        {"type", col.Type}
                    })
                Next
                jResult("schema") = jSchema

                ' Rows as array of arrays
                Dim jRows As New JArray()
                For Each row In result.Rows
                    Dim jRow As New JArray()
                    For Each cellVal In row.Values
                        jRow.Add(If(cellVal Is Nothing, "", cellVal.ToString()))
                    Next
                    jRows.Add(jRow)
                Next
                jResult("rows") = jRows

                ' Metadata
                jResult("total_rows") = result.Rows.Count
                jResult("total_columns") = result.Schema.Count
                jResult("files_processed") = result.ProcessedFiles
                jResult("files_failed") = result.FailedFiles
                If result.FailedFileNames.Count > 0 Then
                    jResult("failed_files") = New JArray(result.FailedFileNames.ToArray())
                End If
                If result.Errors.Count > 0 Then
                    jResult("errors") = New JArray(result.Errors.Take(10).ToArray())
                End If

                Dim jsonString = jResult.ToString(Newtonsoft.Json.Formatting.None)

                ' Truncate if extremely large to stay within LLM context limits
                If jsonString.Length > 200000 Then
                    ' Rebuild with fewer rows
                    Dim maxRows = Math.Max(1, CInt(result.Rows.Count * (200000.0 / jsonString.Length)))
                    Dim jRowsTruncated As New JArray()
                    For i = 0 To Math.Min(maxRows - 1, result.Rows.Count - 1)
                        Dim jRow As New JArray()
                        For Each cellVal In result.Rows(i).Values
                            jRow.Add(If(cellVal Is Nothing, "", cellVal.ToString()))
                        Next
                        jRowsTruncated.Add(jRow)
                    Next
                    jResult("rows") = jRowsTruncated
                    jResult("total_rows") = result.Rows.Count
                    jResult("rows_returned") = jRowsTruncated.Count
                    jResult("truncated") = True
                    jsonString = jResult.ToString(Newtonsoft.Json.Formatting.None)
                End If

                response.Success = True
                response.Response = jsonString

                Dim schemaNames = String.Join(", ", result.Schema.Select(Function(c) c.Name))
                ApDashboardLog($"✓ Fact extraction: {result.Rows.Count} rows, {result.Schema.Count} columns ({schemaNames}), {result.ProcessedFiles} file(s)", "info")

            Finally
                OtherPrompt = savedOtherPrompt
                OutputLanguage = savedOutputLanguage
            End Try

        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation was cancelled."
            response.Response = response.ErrorMessage
        Catch ex As System.Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error extracting data from attachments: {ex.Message}"
        End Try

        Return response
    End Function


    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: report_inability
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Handles the report_inability tool call. Generates helpful suggestions by:
    ''' (1) noting attachment size limits if relevant, (2) consulting the HelpMeInky
    ''' manual for Red Ink add-in features, and (3) optionally querying the
    ''' InternetResearch model for alternative online tools.
    ''' The suggestions are returned in the tool response for LLM incorporation,
    ''' with explicit instruction that suggestions may be rephrased but not omitted.
    ''' </summary>
    Private Async Function ExecuteReportInabilityTool(
            toolCall As ToolCall,
            context As ToolExecutionContext,
            ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = AP_Tool_ReportInability,
            .OriginalCallJson = toolCall.RawJson
        }

        Dim reason = GetArgString(toolCall.Arguments, "reason")
        If String.IsNullOrWhiteSpace(reason) Then reason = "unspecified"
        ApDashboardLog($"📋 Inability reported: {reason}", "warn")

        Dim sb As New StringBuilder()

        ' ── Attachment size limit detail ──
        If _apCurrentAttachments IsNot Nothing AndAlso _apCurrentAttachments.Any(Function(a) a.IsOverSizeLimit) Then
            Dim limitMB = _apConfig.MaxAttachmentBytes / 1024.0 / 1024.0
            Dim oversizedNames = _apCurrentAttachments.Where(Function(a) a.IsOverSizeLimit).
                Select(Function(a) $"{a.OriginalFileName} ({a.SizeBytes / 1024.0 / 1024.0:F1} MB)").ToList()
            sb.AppendLine($"The maximum permitted attachment size is {limitMB:F0} MB. " &
                          $"The following file(s) exceeded this limit: {String.Join(", ", oversizedNames)}. " &
                          $"Advise the sender to send smaller files or split large documents.")
            sb.AppendLine()
        End If

        ' ── Red Ink add-in suggestion (via HelpMeInky manual) ──
        Try
            Dim redInkSuggestion = Await GetRedInkSuggestionAsync(
                _apCurrentMailInfo, reason, ct)
            If Not String.IsNullOrWhiteSpace(redInkSuggestion) Then
                sb.AppendLine(redInkSuggestion.Trim())
                sb.AppendLine()
            End If
        Catch ex As System.Exception
            ApDashboardLog($"Red Ink suggestion failed: {ex.Message}", "warn")
        End Try

        ' ── Internet alternative suggestion ──
        Dim hasInternetSuggestion As Boolean = False
        Try
            Dim internetSuggestion = Await GetInternetAlternativeSuggestionAsync(
                _apCurrentMailInfo, reason, ct)
            If Not String.IsNullOrWhiteSpace(internetSuggestion) Then
                sb.AppendLine(internetSuggestion.Trim())
                sb.AppendLine()
                hasInternetSuggestion = True
            End If
        Catch ex As System.Exception
            ApDashboardLog($"Internet suggestion failed: {ex.Message}", "warn")
        End Try

        If sb.Length = 0 Then
            sb.AppendLine("No specific suggestions available. Advise the sender to try again, rephrase their request, or contact the operator for assistance.")
        End If

        ' ── Instruction to the LLM on how to use this tool output ──
        sb.AppendLine()
        sb.AppendLine("INSTRUCTIONS FOR YOUR RESPONSE:")
        sb.AppendLine("Include ALL of the above suggestions in your reply to the sender. " &
                      "You may rephrase the suggestions to fit naturally into your response, " &
                      "but do NOT omit any of them.")
        If hasInternetSuggestion Then
            sb.AppendLine("MANDATORY: Your reply MUST end with the following disclaimer paragraph, " &
                          "copied VERBATIM (do not rephrase, shorten or omit it):")
            sb.AppendLine("Please note: Third-party services and tools may only be used if permitted by your organization's policies. " &
                          "Before using any external service or tool, ensure it meets your corporate security, confidentiality, and data protection requirements.")
        End If

        response.Success = True
        response.Response = sb.ToString().TrimEnd()
        ApDashboardLog($"✓ Inability suggestions generated ({response.Response.Length} chars)", "step")
        Return response
    End Function


    Private Shared Function TryParseSchedulerUtcArgument(value As Object, ByRef utcValue As DateTime) As Boolean
        utcValue = DateTime.MinValue

        If value Is Nothing Then Return False

        Dim token As JToken = TryCast(value, JToken)

        If token Is Nothing Then
            Try
                token = JToken.FromObject(value)
            Catch
                token = Nothing
            End Try
        End If

        Return TryParseSchedulerUtcToken(token, utcValue)
    End Function

    Private Shared Function TryParseSchedulerUtcToken(token As JToken, ByRef utcValue As DateTime) As Boolean
        If token Is Nothing OrElse
           token.Type = JTokenType.Null OrElse
           token.Type = JTokenType.Undefined Then
            Return False
        End If

        Select Case token.Type
            Case JTokenType.String, JTokenType.Date
                Dim parsed As DateTime
                If DateTime.TryParse(
                    token.ToString(),
                    Nothing,
                    Globalization.DateTimeStyles.AdjustToUniversal Or Globalization.DateTimeStyles.AssumeUniversal,
                    parsed) Then

                    utcValue = parsed.ToUniversalTime()
                    Return True
                End If

            Case JTokenType.Object
                Dim obj = DirectCast(token, JObject)

                Dim preferredPropertyNames As String() = {
                    "next_due_utc",
                    "end_date_utc",
                    "utc",
                    "value",
                    "text",
                    "timestamp",
                    "iso",
                    "iso8601",
                    "result"
                }

                For Each propertyName In preferredPropertyNames
                    Dim child = obj(propertyName)
                    If child IsNot Nothing AndAlso TryParseSchedulerUtcToken(child, utcValue) Then
                        Return True
                    End If
                Next

                For Each prop In obj.Properties()
                    If TryParseSchedulerUtcToken(prop.Value, utcValue) Then
                        Return True
                    End If
                Next

            Case JTokenType.Array
                For Each child As JToken In DirectCast(token, JArray)
                    If TryParseSchedulerUtcToken(child, utcValue) Then
                        Return True
                    End If
                Next

            Case Else
                Dim parsed As DateTime
                If DateTime.TryParse(
                    token.ToString(Newtonsoft.Json.Formatting.None),
                    Nothing,
                    Globalization.DateTimeStyles.AdjustToUniversal Or Globalization.DateTimeStyles.AssumeUniversal,
                    parsed) Then

                    utcValue = parsed.ToUniversalTime()
                    Return True
                End If
        End Select

        Return False
    End Function

    Private Function GetLocalChatPrimaryMailboxSmtpAddress() As String
        Try
            Dim ns As Microsoft.Office.Interop.Outlook.NameSpace = Application.GetNamespace("MAPI")
            Dim targetStore As Microsoft.Office.Interop.Outlook.Store = Nothing

            Try
                Dim explorer As Microsoft.Office.Interop.Outlook.Explorer = Application.ActiveExplorer()
                If explorer IsNot Nothing Then
                    Dim currentFolder As Microsoft.Office.Interop.Outlook.MAPIFolder =
                        TryCast(explorer.CurrentFolder, Microsoft.Office.Interop.Outlook.MAPIFolder)
                    If currentFolder IsNot Nothing Then
                        targetStore = currentFolder.Store
                    End If
                End If
            Catch
            End Try

            If targetStore Is Nothing Then
                Try
                    targetStore = ns.DefaultStore
                Catch
                End Try
            End If

            If targetStore IsNot Nothing Then
                For i As Integer = 1 To ns.Accounts.Count
                    Try
                        Dim acct As Microsoft.Office.Interop.Outlook.Account = ns.Accounts(i)
                        If acct Is Nothing Then Continue For

                        Dim deliveryStore As Microsoft.Office.Interop.Outlook.Store = Nothing
                        Try : deliveryStore = acct.DeliveryStore : Catch : End Try

                        If deliveryStore IsNot Nothing AndAlso
                           deliveryStore.StoreID.Equals(targetStore.StoreID, StringComparison.OrdinalIgnoreCase) AndAlso
                           Not String.IsNullOrWhiteSpace(acct.SmtpAddress) Then
                            Return acct.SmtpAddress.Trim()
                        End If
                    Catch
                    End Try
                Next
            End If

            For i As Integer = 1 To ns.Accounts.Count
                Try
                    Dim acct As Microsoft.Office.Interop.Outlook.Account = ns.Accounts(i)
                    If acct IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(acct.SmtpAddress) Then
                        Return acct.SmtpAddress.Trim()
                    End If
                Catch
                End Try
            Next
        Catch
        End Try

        Return ""
    End Function
    Private Function GetSchedulerCallerOwnerAddress() As String
        If _apCurrentMailInfo IsNot Nothing AndAlso
           Not String.IsNullOrWhiteSpace(_apCurrentMailInfo.SenderEmail) Then
            Return _apCurrentMailInfo.SenderEmail.Trim()
        End If

        Return GetLocalChatPrimaryMailboxSmtpAddress()
    End Function

    Private Function TryNormalizeSchedulerDeliverToForCaller(rawValue As Object,
                                                             ownerAddress As String,
                                                             allowEmpty As Boolean,
                                                             ByRef normalizedDeliverTo As List(Of String),
                                                             ByRef errorMessage As String) As Boolean
        normalizedDeliverTo = New List(Of String)()
        errorMessage = ""

        Dim requested As New List(Of String)()

        If rawValue Is Nothing Then
            If Not allowEmpty AndAlso Not String.IsNullOrWhiteSpace(ownerAddress) Then
                normalizedDeliverTo.Add(ownerAddress)
            End If

            Return True
        End If

        If TypeOf rawValue Is JArray Then
            requested.AddRange(
                DirectCast(rawValue, JArray).
                    Select(Function(t) t.ToString().Trim()).
                    Where(Function(s) s.Length > 0))
        Else
            Dim rawText As String = rawValue.ToString().Trim()
            If rawText <> "" Then
                requested.AddRange(
                    rawText.Split({","c, ";"c}, StringSplitOptions.RemoveEmptyEntries).
                        Select(Function(s) s.Trim()).
                        Where(Function(s) s.Length > 0))
            End If
        End If

        requested = requested.
            Distinct(StringComparer.OrdinalIgnoreCase).
            ToList()

        If requested.Count = 0 Then
            If Not allowEmpty AndAlso Not String.IsNullOrWhiteSpace(ownerAddress) Then
                normalizedDeliverTo.Add(ownerAddress)
            End If

            Return True
        End If

        If String.IsNullOrWhiteSpace(ownerAddress) Then
            errorMessage = "Could not determine the permitted mailbox address for this scheduled task."
            Return False
        End If

        If requested.Count > 1 Then
            errorMessage = $"Only the mailbox address '{ownerAddress}' may be stored for this scheduled task."
            Return False
        End If

        If Not requested(0).Equals(ownerAddress, StringComparison.OrdinalIgnoreCase) Then
            errorMessage = $"Only the mailbox address '{ownerAddress}' may be stored for this scheduled task."
            Return False
        End If

        normalizedDeliverTo.Add(ownerAddress)
        Return True
    End Function

    Private Function IsSchedulerTaskOwnedByCaller(task As ScheduledTask, ownerAddress As String) As Boolean
        If task Is Nothing OrElse String.IsNullOrWhiteSpace(ownerAddress) Then
            Return False
        End If

        If Not String.IsNullOrWhiteSpace(task.CreatedBy) Then
            Return task.CreatedBy.Trim().Equals(ownerAddress.Trim(), StringComparison.OrdinalIgnoreCase)
        End If

        Return task.DeliverTo IsNot Nothing AndAlso
            task.DeliverTo.Any(Function(addr) Not String.IsNullOrWhiteSpace(addr) AndAlso
                                             addr.Trim().Equals(ownerAddress.Trim(), StringComparison.OrdinalIgnoreCase))
    End Function

    Private Function FindOwnedScheduledTask(idOrQuery As String, ownerAddress As String) As ScheduledTask
        If String.IsNullOrWhiteSpace(idOrQuery) Then
            Return Nothing
        End If

        Dim ownedTasks = SchedulerListTasks().
            Where(Function(t) IsSchedulerTaskOwnedByCaller(t, ownerAddress)).
            ToList()

        Dim byId = ownedTasks.FirstOrDefault(
            Function(t) t.Id.Equals(idOrQuery, StringComparison.OrdinalIgnoreCase) OrElse
                         t.Id.StartsWith(idOrQuery, StringComparison.OrdinalIgnoreCase))
        If byId IsNot Nothing Then
            Return byId
        End If

        Return ownedTasks.FirstOrDefault(
            Function(t) Not String.IsNullOrWhiteSpace(t.Instruction) AndAlso
                         t.Instruction.IndexOf(idOrQuery, StringComparison.OrdinalIgnoreCase) >= 0)
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  MANAGE_SCHEDULED_TASKS TOOL EXECUTOR
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>Executes the manage_scheduled_tasks tool call.</summary>
    Private Async Function ExecuteManageScheduledTasksTool(
            toolCall As ToolCall,
            context As ToolExecutionContext,
            ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName,
            .Timestamp = DateTime.Now
        }


        Try
            Dim args = toolCall.Arguments
            If args Is Nothing Then
                response.Success = False
                response.ErrorMessage = "No arguments provided"
                response.Response = "Error: No arguments provided."
                Return response
            End If

            If Not _apActive Then
                If INI_WebServerBlock = 4 Then
                    response.Success = False
                    response.Response = "Error: scheduling is disabled because WebServerBlock = 4."
                    Return response
                End If
            End If

            Dim action As String = ""
            If args.ContainsKey("action") Then action = args("action")?.ToString()?.Trim()?.ToLowerInvariant()

            Dim schedulerOwnerAddress As String = GetSchedulerCallerOwnerAddress()
            If String.IsNullOrWhiteSpace(schedulerOwnerAddress) Then
                response.Success = False
                response.Response = "Error: could not determine the mailbox identity for scheduler access."
                Return response
            End If

            Select Case action
                Case "create"
                    Dim task As New ScheduledTask()
                    task.Id = Guid.NewGuid().ToString("N")
                    task.CreatedUtc = DateTime.UtcNow

                    If args.ContainsKey("instruction") Then task.Instruction = args("instruction")?.ToString()
                    If args.ContainsKey("subject") Then task.Subject = args("subject")?.ToString()
                    If args.ContainsKey("schedule_description") Then task.ScheduleDescription = args("schedule_description")?.ToString()
                    If args.ContainsKey("rrule") Then task.Rrule = args("rrule")?.ToString()
                    If args.ContainsKey("time_of_day_local") Then task.TimeOfDayLocal = args("time_of_day_local")?.ToString()

                    ' Parse next_due_utc
                    If args.ContainsKey("next_due_utc") Then
                        Dim parsedUtc As DateTime
                        If TryParseSchedulerUtcArgument(args("next_due_utc"), parsedUtc) Then
                            task.NextDueUtc = parsedUtc
                        End If
                    End If

                    ' Fallback: if the LLM did not provide next_due_utc, compute it from
                    ' the RRULE and time_of_day_local so the task fires at the correct time
                    ' rather than immediately.
                    If task.NextDueUtc = DateTime.MaxValue Then
                        task.NextDueUtc = ComputeFirstDueUtc(task.Rrule, task.TimeOfDayLocal)
                    End If

                    ' Parse end_date_utc
                    If args.ContainsKey("end_date_utc") Then
                        Dim parsedUtc As DateTime
                        If TryParseSchedulerUtcArgument(args("end_date_utc"), parsedUtc) Then
                            task.EndDateUtc = parsedUtc
                        End If
                    End If

                    ' Parse remaining_occurrences
                    If args.ContainsKey("remaining_occurrences") Then
                        Dim occ As Integer
                        If Integer.TryParse(args("remaining_occurrences")?.ToString(), occ) Then task.RemainingOccurrences = occ
                    End If

                    Dim isLocalChatOrigin As Boolean =
                        (Not _apActive AndAlso INI_WebServerBlock <> 4)

                    Dim rawDeliverTo As Object = Nothing
                    If args.ContainsKey("deliver_to") Then
                        rawDeliverTo = args("deliver_to")
                    End If

                    If isLocalChatOrigin AndAlso Not INI_AutoPilotSchedulerLocalChat AndAlso rawDeliverTo IsNot Nothing Then
                        Dim requestedLocalDelivery As New List(Of String)()

                        If TypeOf rawDeliverTo Is JArray Then
                            requestedLocalDelivery.AddRange(
                                DirectCast(rawDeliverTo, JArray).
                                    Select(Function(t) t.ToString().Trim()).
                                    Where(Function(s) s.Length > 0))
                        Else
                            Dim rawText As String = rawDeliverTo.ToString().Trim()
                            If rawText <> "" Then
                                requestedLocalDelivery.AddRange(
                                    rawText.Split({","c, ";"c}, StringSplitOptions.RemoveEmptyEntries).
                                        Select(Function(s) s.Trim()).
                                        Where(Function(s) s.Length > 0))
                            End If
                        End If

                        If requestedLocalDelivery.Count > 0 Then
                            response.Success = False
                            response.ErrorCode = "scheduler_localchat_email_disabled"
                            response.ErrorMessage = "Local Chat scheduled e-mail delivery is disabled."
                            response.Response = "Error: Local Chat scheduled e-mail delivery is disabled."
                            Return response
                        End If

                        rawDeliverTo = Nothing
                    End If

                    Dim deliverToError As String = ""
                    If Not TryNormalizeSchedulerDeliverToForCaller(
                        rawDeliverTo,
                        schedulerOwnerAddress,
                                                allowEmpty:=isLocalChatOrigin AndAlso Not INI_AutoPilotSchedulerLocalChat,
                        normalizedDeliverTo:=task.DeliverTo,
                        errorMessage:=deliverToError) Then

                        response.Success = False
                        response.ErrorCode = "scheduler_delivery_not_allowed"
                        response.ErrorMessage = deliverToError
                        response.Response = "Error: " & deliverToError
                        Return response
                    End If

                    task.ExecutionMode = ResolveScheduledTaskExecutionMode(isLocalChatOrigin, task.DeliverTo)
                    task.CreatedBy = schedulerOwnerAddress


                    ' Store attachments from current mail if requested
                    Dim storeNames As List(Of String) = Nothing
                    If args.ContainsKey("store_attachment_names") Then
                        Dim storeObj = args("store_attachment_names")
                        If TypeOf storeObj Is JArray Then
                            storeNames = DirectCast(storeObj, JArray).Select(Function(t) t.ToString().Trim()).
                                Where(Function(s) s.Length > 0).ToList()
                        End If
                    End If

                    Dim taskId = SchedulerCreateTask(task)

                    ' Store requested attachments
                    If storeNames IsNot Nothing AndAlso storeNames.Count > 0 AndAlso _apCurrentAttachments IsNot Nothing Then
                        For Each name In storeNames
                            Dim att = _apCurrentAttachments.FirstOrDefault(
                                Function(a) a.OriginalFileName.Equals(name, StringComparison.OrdinalIgnoreCase))
                            If att IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(att.TempFilePath) AndAlso File.Exists(att.TempFilePath) Then
                                SchedulerStoreAttachment(taskId, att.TempFilePath)
                            End If
                        Next
                    End If

                    response.Success = True

                    Dim modeLabel As String =
                        If(task.ExecutionMode.Equals("browser_prompt", StringComparison.OrdinalIgnoreCase),
                           "Local Agent browser (interactive approval required)",
                           "E-mail delivery")

                    Dim apWarning = ""
                    If task.ExecutionMode.Equals(AP_TaskExecutionModeBrowserPrompt, StringComparison.OrdinalIgnoreCase) Then
                        apWarning = vbCrLf &
                            "⚠ This task will run only while the local webserver is running, and each due execution will first ask the user for approval."
                    ElseIf Not _apActive Then
                        apWarning = vbCrLf &
                            "⚠ This task will run while the local webserver is running and will send the result by e-mail to the mailbox recorded for this task."
                    End If

                    response.Response = $"Task created successfully." & vbCrLf &
                        $"Task ID: {taskId}" & vbCrLf &
                        "Instruction:" & vbCrLf &
                        If(task.Instruction, "") & vbCrLf &
                        $"Schedule: {If(task.ScheduleDescription, "one-time")}" & vbCrLf &
                        $"Execution mode: {modeLabel}" & vbCrLf &
                        $"Next execution: {task.NextDueUtc.ToLocalTime():yyyy-MM-dd HH:mm} (local)" &
                        If(task.DeliverTo IsNot Nothing AndAlso task.DeliverTo.Count > 0,
                           vbCrLf & $"Deliver to: {String.Join(", ", task.DeliverTo)}",
                           "") &
                        If(task.AttachmentFiles.Count > 0, vbCrLf & $"Stored attachments: {String.Join(", ", task.AttachmentFiles)}", "") &
                        apWarning

                Case "list"
                    Dim statusFilter As String = Nothing
                    If args.ContainsKey("status_filter") Then statusFilter = args("status_filter")?.ToString()

                    Dim tasks = SchedulerListTasks(statusFilter).
                        Where(Function(t) IsSchedulerTaskOwnedByCaller(t, schedulerOwnerAddress)).
                        ToList()

                    response.Success = True
                    response.Response = FormatTaskListForDisplay(tasks)

                Case "get"
                    Dim taskId As String = ""
                    If args.ContainsKey("task_id") Then taskId = args("task_id")?.ToString()
                    If String.IsNullOrWhiteSpace(taskId) Then
                        response.Success = False
                        response.Response = "Error: task_id is required for the 'get' action."
                        Return response
                    End If
                    Dim task = FindOwnedScheduledTask(taskId, schedulerOwnerAddress)
                    If task Is Nothing Then
                        response.Success = False
                        response.Response = $"No task found matching '{taskId}'."
                    Else
                        response.Success = True
                        response.Response = FormatTaskListForDisplay(New List(Of ScheduledTask) From {task})
                    End If

                Case "update"
                    Dim taskId As String = ""
                    If args.ContainsKey("task_id") Then taskId = args("task_id")?.ToString()
                    If String.IsNullOrWhiteSpace(taskId) Then
                        response.Success = False
                        response.Response = "Error: task_id is required for the 'update' action."
                        Return response
                    End If
                    Dim task = FindOwnedScheduledTask(taskId, schedulerOwnerAddress)
                    If task Is Nothing Then
                        response.Success = False
                        response.Response = $"No task found matching '{taskId}'."
                        Return response
                    End If

                    ' Apply updates
                    If args.ContainsKey("instruction") Then task.Instruction = args("instruction")?.ToString()
                    If args.ContainsKey("subject") Then task.Subject = args("subject")?.ToString()
                    If args.ContainsKey("schedule_description") Then task.ScheduleDescription = args("schedule_description")?.ToString()
                    If args.ContainsKey("rrule") Then task.Rrule = args("rrule")?.ToString()
                    If args.ContainsKey("time_of_day_local") Then task.TimeOfDayLocal = args("time_of_day_local")?.ToString()
                    If args.ContainsKey("status") Then task.Status = args("status")?.ToString()

                    If args.ContainsKey("next_due_utc") Then
                        Dim parsedUtc As DateTime
                        If TryParseSchedulerUtcArgument(args("next_due_utc"), parsedUtc) Then
                            task.NextDueUtc = parsedUtc
                        End If
                    End If
                    If args.ContainsKey("end_date_utc") Then
                        Dim parsedUtc As DateTime
                        If TryParseSchedulerUtcArgument(args("end_date_utc"), parsedUtc) Then
                            task.EndDateUtc = parsedUtc
                        End If
                    End If
                    If args.ContainsKey("remaining_occurrences") Then
                        Dim occ As Integer
                        If Integer.TryParse(args("remaining_occurrences")?.ToString(), occ) Then task.RemainingOccurrences = occ
                    End If


                    Dim isLocalChatOrigin As Boolean =
                        task.ExecutionMode.Equals(AP_TaskExecutionModeBrowserPrompt, StringComparison.OrdinalIgnoreCase) OrElse
                        (Not _apActive AndAlso INI_WebServerBlock <> 4)

                    Dim rawDeliverTo As Object = Nothing
                    If args.ContainsKey("deliver_to") Then
                        rawDeliverTo = args("deliver_to")
                    ElseIf task.DeliverTo IsNot Nothing AndAlso task.DeliverTo.Count > 0 AndAlso
                           (Not isLocalChatOrigin OrElse INI_AutoPilotSchedulerLocalChat) Then
                        rawDeliverTo = New JArray(task.DeliverTo)
                    End If

                    If isLocalChatOrigin AndAlso Not INI_AutoPilotSchedulerLocalChat AndAlso args.ContainsKey("deliver_to") Then
                        Dim requestedLocalDelivery As New List(Of String)()

                        If TypeOf rawDeliverTo Is JArray Then
                            requestedLocalDelivery.AddRange(
                                DirectCast(rawDeliverTo, JArray).
                                    Select(Function(t) t.ToString().Trim()).
                                    Where(Function(s) s.Length > 0))
                        ElseIf rawDeliverTo IsNot Nothing Then
                            Dim rawText As String = rawDeliverTo.ToString().Trim()
                            If rawText <> "" Then
                                requestedLocalDelivery.AddRange(
                                    rawText.Split({","c, ";"c}, StringSplitOptions.RemoveEmptyEntries).
                                        Select(Function(s) s.Trim()).
                                        Where(Function(s) s.Length > 0))
                            End If
                        End If

                        If requestedLocalDelivery.Count > 0 Then
                            response.Success = False
                            response.ErrorCode = "scheduler_localchat_email_disabled"
                            response.ErrorMessage = "Local Chat scheduled e-mail delivery is disabled."
                            response.Response = "Error: Local Chat scheduled e-mail delivery is disabled."
                            Return response
                        End If

                        rawDeliverTo = Nothing
                    End If

                    Dim deliverToError As String = ""
                    If Not TryNormalizeSchedulerDeliverToForCaller(
                        rawDeliverTo,
                        schedulerOwnerAddress,
                                                allowEmpty:=isLocalChatOrigin AndAlso Not INI_AutoPilotSchedulerLocalChat,
                        normalizedDeliverTo:=task.DeliverTo,
                        errorMessage:=deliverToError) Then

                        response.Success = False
                        response.ErrorCode = "scheduler_delivery_not_allowed"
                        response.ErrorMessage = deliverToError
                        response.Response = "Error: " & deliverToError
                        Return response
                    End If

                    task.ExecutionMode = ResolveScheduledTaskExecutionMode(isLocalChatOrigin, task.DeliverTo)
                    task.CreatedBy = schedulerOwnerAddress

                    ' Store new attachments from current mail if requested
                    Dim storeNames As List(Of String) = Nothing
                    If args.ContainsKey("store_attachment_names") Then
                        Dim storeObj = args("store_attachment_names")
                        If TypeOf storeObj Is JArray Then
                            storeNames = DirectCast(storeObj, JArray).Select(Function(t) t.ToString().Trim()).
                                Where(Function(s) s.Length > 0).ToList()
                        End If
                    End If

                    If SchedulerUpdateTask(task) Then
                        ' Store attachments after the update succeeds
                        Dim storedCount = 0
                        If storeNames IsNot Nothing AndAlso storeNames.Count > 0 AndAlso _apCurrentAttachments IsNot Nothing Then
                            For Each name In storeNames
                                Dim att = _apCurrentAttachments.FirstOrDefault(
                                    Function(a) a.OriginalFileName.Equals(name, StringComparison.OrdinalIgnoreCase))
                                If att IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(att.TempFilePath) AndAlso File.Exists(att.TempFilePath) Then
                                    If SchedulerStoreAttachment(task.Id, att.TempFilePath) Then storedCount += 1
                                End If
                            Next
                        End If

                        response.Success = True
                        response.Response = $"Task {taskId} updated successfully." &
                            If(storedCount > 0, vbCrLf & $"Stored {storedCount} new attachment(s).", "") & vbCrLf &
                            FormatTaskListForDisplay(New List(Of ScheduledTask) From {task})
                    Else
                        response.Success = False
                        response.Response = $"Failed to update task {taskId}."
                    End If

                Case "delete"
                    Dim taskId As String = ""
                    If args.ContainsKey("task_id") Then taskId = args("task_id")?.ToString()
                    If String.IsNullOrWhiteSpace(taskId) Then
                        response.Success = False
                        response.Response = "Error: task_id is required for the 'delete' action."
                        Return response
                    End If
                    Dim task = FindOwnedScheduledTask(taskId, schedulerOwnerAddress)
                    If task Is Nothing Then
                        response.Success = False
                        response.Response = $"No task found matching '{taskId}'."
                        Return response
                    End If

                    If SchedulerDeleteTask(task.Id) Then
                        response.Success = True
                        response.Response = $"Task {task.Id.Substring(0, Math.Min(8, task.Id.Length))}... deleted successfully."
                    Else
                        response.Success = False
                        response.Response = $"Failed to delete task."
                    End If

                Case Else
                    response.Success = False
                    response.Response = $"Unknown action: '{action}'. Supported: create, list, get, update, delete."
            End Select

        Catch ex As OperationCanceledException
            Throw
        Catch ex As System.Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Scheduler error: {ex.Message}"
            context.Log($"Scheduler tool error: {ex.Message}", "error")
        End Try

        Return response
    End Function

    ''' <summary>
    ''' Computes the first NextDueUtc for a newly created task based on its RRULE
    ''' and time-of-day. If the computed time today has already passed, advances to
    ''' the next occurrence. Falls back to UtcNow only if no schedule can be derived.
    ''' </summary>
    Private Shared Function ComputeFirstDueUtc(rrule As String, timeOfDayLocal As String) As DateTime
        Try
            ' If we have a time-of-day, try scheduling for that time today or tomorrow
            If Not String.IsNullOrWhiteSpace(timeOfDayLocal) Then
                Dim parsed As DateTime
                If DateTime.TryParse(timeOfDayLocal, parsed) Then
                    Dim todayAtTime = DateTime.Today.Add(parsed.TimeOfDay)

                    ' If that time is still in the future today, use it
                    If todayAtTime > DateTime.Now Then
                        Return todayAtTime.ToUniversalTime()
                    End If

                    ' Time already passed today — compute next occurrence from now
                    If Not String.IsNullOrWhiteSpace(rrule) Then
                        Dim nextUtc = ComputeNextOccurrence(rrule, todayAtTime.ToUniversalTime(), timeOfDayLocal)
                        If nextUtc IsNot Nothing Then Return nextUtc.Value
                    End If

                    ' No RRULE or couldn't compute — use tomorrow at that time
                    Return todayAtTime.AddDays(1).ToUniversalTime()
                End If
            End If

            ' No time-of-day but has RRULE — compute next from now
            If Not String.IsNullOrWhiteSpace(rrule) Then
                Dim nextUtc = ComputeNextOccurrence(rrule, DateTime.UtcNow, timeOfDayLocal)
                If nextUtc IsNot Nothing Then Return nextUtc.Value
            End If

            ' No schedule info at all — execute immediately
            Return DateTime.UtcNow

        Catch
            Return DateTime.UtcNow
        End Try
    End Function


    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: manage_user_memory
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Async Function ExecuteManageUserMemoryTool(
            toolCall As ToolCall,
            context As ToolExecutionContext,
            ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim action = GetArgString(toolCall.Arguments, "action")
            If String.IsNullOrWhiteSpace(action) Then
                response.Success = False
                response.Response = "Missing required parameter: action"
                Return response
            End If

            ' SECURITY: sender e-mail comes from the mail being processed, not from user input
            Dim senderEmail = _apCurrentMailInfo?.SenderEmail
            If String.IsNullOrWhiteSpace(senderEmail) Then
                response.Success = False
                response.Response = "Error: Could not determine sender e-mail address."
                Return response
            End If

            Select Case action.ToLowerInvariant()

                Case "enable"
                    EnableUserMemory(senderEmail)
                    response.Success = True
                    response.Response = "Memory has been enabled. I will now learn and remember your preferences across sessions. " &
                        "You can ask me to remember specific things, or I will automatically pick up on your preferences. " &
                        "Use 'disable' to opt out at any time (this will delete all stored preferences)."
                    ApDashboardLog($"🧠 User memory enabled for: {senderEmail}", "info")

                Case "disable"
                    DisableUserMemory(senderEmail)
                    response.Success = True
                    response.Response = "Memory has been disabled and all stored preferences have been deleted. " &
                        "I will no longer remember preferences across sessions."
                    ApDashboardLog($"🧠 User memory disabled for: {senderEmail}", "info")

                Case "list"
                    If Not IsUserMemoryEnabled(senderEmail) Then
                        response.Success = True
                        response.Response = "Memory is not enabled for this user. Use action='enable' to activate it."
                        Return response
                    End If
                    Dim memoryContent = ReadUserMemory(senderEmail, _context.INI_InkyMemoryCap)
                    If String.IsNullOrWhiteSpace(memoryContent) Then
                        response.Success = True
                        response.Response = "Memory is enabled but currently empty. No preferences stored yet."
                    Else
                        response.Success = True
                        response.Response = "Current memory items:" & vbCrLf & memoryContent
                    End If

                Case "add"
                    Dim value = GetArgString(toolCall.Arguments, "value")
                    If String.IsNullOrWhiteSpace(value) Then
                        response.Success = False
                        response.Response = "Missing required parameter: value"
                        Return response
                    End If
                    If Not IsUserMemoryEnabled(senderEmail) Then EnableUserMemory(senderEmail)
                    Dim ops As New List(Of SharedMethods.MemoryOperation)()
                    ops.Add(New SharedMethods.MemoryOperation With {
                        .Type = SharedMethods.MemoryOperation.OpType.Add, .Value = value
                    })
                    SharedMethods.ApplyMemoryOperationsToFile(GetUserMemoryFilePath(senderEmail), ops, _context.INI_InkyMemoryCap)
                    response.Success = True
                    response.Response = $"Memory item added: {value}"
                    ApDashboardLog($"🧠 Memory add for {senderEmail}: {value}", "step")

                Case "remove"
                    Dim value = GetArgString(toolCall.Arguments, "value")
                    If String.IsNullOrWhiteSpace(value) Then
                        response.Success = False
                        response.Response = "Missing required parameter: value"
                        Return response
                    End If
                    If Not IsUserMemoryEnabled(senderEmail) Then
                        response.Success = True
                        response.Response = "Memory is not enabled — nothing to remove."
                        Return response
                    End If
                    Dim ops As New List(Of SharedMethods.MemoryOperation)()
                    ops.Add(New SharedMethods.MemoryOperation With {
                        .Type = SharedMethods.MemoryOperation.OpType.Remove, .Value = value
                    })
                    SharedMethods.ApplyMemoryOperationsToFile(GetUserMemoryFilePath(senderEmail), ops, _context.INI_InkyMemoryCap)
                    response.Success = True
                    response.Response = $"Memory item removed (if matched): {value}"
                    ApDashboardLog($"🧠 Memory remove for {senderEmail}: {value}", "step")

                Case "amend"
                    Dim value = GetArgString(toolCall.Arguments, "value")
                    Dim newValue = GetArgString(toolCall.Arguments, "new_value")
                    If String.IsNullOrWhiteSpace(value) OrElse String.IsNullOrWhiteSpace(newValue) Then
                        response.Success = False
                        response.Response = "Missing required parameters: value and new_value"
                        Return response
                    End If
                    If Not IsUserMemoryEnabled(senderEmail) Then EnableUserMemory(senderEmail)
                    Dim ops As New List(Of SharedMethods.MemoryOperation)()
                    ops.Add(New SharedMethods.MemoryOperation With {
                        .Type = SharedMethods.MemoryOperation.OpType.Amend, .Value = value, .NewValue = newValue
                    })
                    SharedMethods.ApplyMemoryOperationsToFile(GetUserMemoryFilePath(senderEmail), ops, _context.INI_InkyMemoryCap)
                    response.Success = True
                    response.Response = $"Memory item amended: '{value}' → '{newValue}'"
                    ApDashboardLog($"🧠 Memory amend for {senderEmail}: {value} → {newValue}", "step")

                Case "clear"
                    If IsUserMemoryEnabled(senderEmail) Then
                        SharedMethods.WriteFileWithRetry(GetUserMemoryFilePath(senderEmail), SharedMethods.GetDefaultMemoryFileContent())
                        response.Success = True
                        response.Response = "All memory items have been cleared. Memory remains enabled."
                        ApDashboardLog($"🧠 Memory cleared for: {senderEmail}", "info")
                    Else
                        response.Success = True
                        response.Response = "Memory is not enabled — nothing to clear."
                    End If

                Case "toggle_auto_learn"
                    ' Auto-learn is controlled by the <INKY_MEMORY> block in the system prompt.
                    ' We track it via a special memory item.
                    Dim autoLearn = GetArgBool(toolCall.Arguments, "auto_learn", True)
                    If Not IsUserMemoryEnabled(senderEmail) Then EnableUserMemory(senderEmail)

                    Dim filePath = GetUserMemoryFilePath(senderEmail)
                    Dim markerOff = "AUTO_LEARN_DISABLED"

                    If autoLearn Then
                        ' Remove the disable marker
                        Dim ops As New List(Of SharedMethods.MemoryOperation)()
                        ops.Add(New SharedMethods.MemoryOperation With {
                            .Type = SharedMethods.MemoryOperation.OpType.Remove, .Value = markerOff
                        })
                        SharedMethods.ApplyMemoryOperationsToFile(filePath, ops, _context.INI_InkyMemoryCap)
                        response.Success = True
                        response.Response = "Automatic learning is now ON. I will automatically learn from your preferences in conversations."
                    Else
                        ' Add the disable marker
                        Dim ops As New List(Of SharedMethods.MemoryOperation)()
                        ops.Add(New SharedMethods.MemoryOperation With {
                            .Type = SharedMethods.MemoryOperation.OpType.Add, .Value = markerOff
                        })
                        SharedMethods.ApplyMemoryOperationsToFile(filePath, ops, _context.INI_InkyMemoryCap)
                        response.Success = True
                        response.Response = "Automatic learning is now OFF. I will only update memory when you explicitly ask me to remember or forget something."
                    End If
                    ApDashboardLog($"🧠 Auto-learn {If(autoLearn, "ON", "OFF")} for: {senderEmail}", "info")

                Case Else
                    response.Success = False
                    response.Response = $"Unknown action: {action}. Valid actions: enable, disable, list, add, remove, amend, clear, toggle_auto_learn."
            End Select

        Catch ex As OperationCanceledException
            response.Success = False
            response.Response = "Operation was cancelled."
        Catch ex As System.Exception
            response.Success = False
            response.Response = $"Error managing user memory: {ex.Message}"
        End Try

        Return response
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: manage_user_files
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Async Function ExecuteManageUserFilesTool(
            toolCall As ToolCall,
            context As ToolExecutionContext,
            ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim action = GetArgString(toolCall.Arguments, "action")
            If String.IsNullOrWhiteSpace(action) Then
                response.Success = False
                response.Response = "Missing required parameter: action"
                Return response
            End If

            ' SECURITY: sender e-mail comes from the mail being processed, not from user input
            Dim senderEmail = _apCurrentMailInfo?.SenderEmail
            If String.IsNullOrWhiteSpace(senderEmail) Then
                response.Success = False
                response.Response = "Error: Could not determine sender e-mail address."
                Return response
            End If

            Dim fileName = GetArgString(toolCall.Arguments, "file_name")
            Dim targetName = GetArgString(toolCall.Arguments, "target_name")

            Select Case action.ToLowerInvariant()

                Case "list"
                    Dim files = ListUserHomeFiles(senderEmail)
                    If files.Count = 0 Then
                        response.Success = True
                        response.Response = "No files stored in your home directory."
                    Else
                        Dim sb As New StringBuilder()
                        sb.AppendLine($"Files in home directory ({files.Count}):")
                        Dim totalSize As Long = 0
                        For Each f In files
                            sb.AppendLine($"  - {f.Name} ({f.SizeBytes / 1024:F0} KB)")
                            totalSize += f.SizeBytes
                        Next
                        sb.AppendLine($"Total: {totalSize / 1024 / 1024:F1} MB / {AP_UserHomeMaxBytes / 1024 / 1024:F0} MB limit")
                        response.Success = True
                        response.Response = sb.ToString()
                    End If

                Case "add", "replace"
                    If String.IsNullOrWhiteSpace(fileName) Then
                        response.Success = False
                        response.Response = "Missing required parameter: file_name (the attachment to store)"
                        Return response
                    End If

                    ' Find the attachment in the current session
                    Dim att = FindAttachment(fileName)
                    If att Is Nothing OrElse att.TempFilePath Is Nothing OrElse Not IO.File.Exists(att.TempFilePath) Then
                        response.Success = False
                        response.Response = $"Attachment '{fileName}' not found in the current e-mail. Available: {String.Join(", ", GetAllAvailableFileNames())}"
                        Return response
                    End If

                    Dim storeName = If(String.IsNullOrWhiteSpace(targetName), att.OriginalFileName, targetName.Trim())
                    Dim result = StoreFileToUserHome(senderEmail, att.TempFilePath, storeName)
                    response.Success = Not result.StartsWith("Error")
                    response.Response = result
                    If response.Success Then
                        ApDashboardLog($"📁 File stored for {senderEmail}: {storeName} ({att.SizeBytes / 1024:F0} KB)", "info")
                    End If

                Case "remove"
                    If String.IsNullOrWhiteSpace(fileName) Then
                        response.Success = False
                        response.Response = "Missing required parameter: file_name"
                        Return response
                    End If
                    Dim result = RemoveFileFromUserHome(senderEmail, fileName)
                    response.Success = Not result.StartsWith("Error")
                    response.Response = result
                    If response.Success Then
                        ApDashboardLog($"📁 File removed for {senderEmail}: {fileName}", "info")
                    End If

                Case "checkout"
                    If String.IsNullOrWhiteSpace(fileName) Then
                        response.Success = False
                        response.Response = "Missing required parameter: file_name"
                        Return response
                    End If
                    Dim loaded = LoadFileFromUserHome(senderEmail, fileName)
                    If loaded Is Nothing Then
                        response.Success = False
                        response.Response = $"File '{fileName}' not found in home directory, or session context unavailable."
                    Else
                        ' Register as output so it gets attached to the reply
                        loaded.OutputFiles.Add(loaded.TempFilePath)
                        response.Success = True
                        response.Response = $"File '{fileName}' retrieved from home directory and will be attached to the reply."
                        ApDashboardLog($"📁 File checkout for {senderEmail}: {fileName}", "info")
                    End If

                Case "use"
                    If String.IsNullOrWhiteSpace(fileName) Then
                        response.Success = False
                        response.Response = "Missing required parameter: file_name"
                        Return response
                    End If
                    Dim loaded = LoadFileFromUserHome(senderEmail, fileName)
                    If loaded Is Nothing Then
                        response.Success = False
                        response.Response = $"File '{fileName}' not found in home directory, or session context unavailable."
                    Else
                        response.Success = True
                        response.Response = $"File '{fileName}' loaded into the current session. " &
                            "It is now available as an attachment and can be referenced by other tools " &
                            $"(e.g., process_word_document, read_attachment) using the name '{loaded.OriginalFileName}'."
                        ApDashboardLog($"📁 File loaded into session for {senderEmail}: {fileName}", "step")
                    End If

                Case Else
                    response.Success = False
                    response.Response = $"Unknown action: {action}. Valid actions: list, add, remove, replace, checkout, use."
            End Select

        Catch ex As OperationCanceledException
            response.Success = False
            response.Response = "Operation was cancelled."
        Catch ex As System.Exception
            response.Success = False
            response.Response = $"Error managing user files: {ex.Message}"
        End Try

        Return response
    End Function

End Class
