' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.Tools.PDF.vb
' Purpose:
'   Defines and executes AutoPilot internal tools for PDF document operations
'   within Outlook AutoPilot Chat-Agent runs, including text extraction,
'   merging, splitting, watermarking, commenting, redacting, and overlaying.
'
' Tools Provided:
'   - extract_pdf_text: Extracts text content from PDF attachments
'   - merge_pdfs: Merges multiple PDF attachments into a single PDF
'   - split_pdf: Extracts a page range from a PDF into a new PDF
'   - add_pdf_watermark: Adds a watermark to all pages of a PDF
'   - comment_pdf_document: Adds annotations/comments to a PDF
'   - redact_pdf: Redacts sensitive text from a PDF (prepare/finalize modes)
'   - overlay_pdf: Overlays content, shapes, or text onto PDF pages
'
' Tool Interface Architecture:
'   - Registration:
'       * Tools are exposed as `ModelConfig` entries (`Tool=True`, `ToolOnly=True`)
'         so they participate in the same tool-calling pipeline as external tools.
'       * Tool metadata (`ToolDefinition`, `ToolInstructionsPrompt`) is generated
'         inline and consumed by `ExecuteToolCall` / `ExecuteToolingLoop`.
'   - Dispatch:
'       * `TryExecuteAutoPilotTool` routes parsed tool calls to strongly scoped
'         executor methods (`ExecuteExtractPdfTextTool`, `ExecuteMergePdfsTool`,
'         `ExecuteCommentPdfTool`, `ExecuteRedactPdfTool`, etc.) and returns
'         `ToolResponse` payloads.
'   - Session scope:
'       * All tools use AutoPilot session state from `ThisAddIn.Autopilot.vb`:
'           - `_apCurrentAttachments`: attachment registry for input/output lookups
'           - `_apCurrentTempDir`: per-mail temp directory for file creation
'           - `_apCurrentMailInfo`: metadata about the current email session
'       * Supports tool chaining via output registration (`OutputFiles`) and
'         attachment lookup via `FindAttachment` (original + prior tool outputs).
'   - PDF processing:
'       * Direct PDF operations via PdfSharp for watermarking and merging.
'       * PdfPig for text extraction with OCR-aware handling.
'       * Rasterization fallback for encrypted/restricted PDFs.
'       * Font resolution via `EnsureApPdfSharpFontResolver`.
'   - Error handling:
'       * Returns structured `ToolResponse` with success flag, message, and
'         error details. Encryption and corruption are handled gracefully with
'         fallback strategies.
'       * Temp files are cleaned up even on failure.
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
'       * PDF format validation and encryption detection.
'       * Filename collision prevention via counter-based renaming.
'   - Encryption handling:
'       * Encrypted PDFs are detected and handled via rasterization or user
'         notification as appropriate.
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
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn


    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: comment_pdf_document
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Async Function ExecuteCommentPdfTool(
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
                response.ErrorMessage = "Missing required parameter: instruction"
                response.Response = response.ErrorMessage
                Return response
            End If

            Dim author = GetArgString(toolCall.Arguments, "author")
            Dim targetNames = GetArgStringArray(toolCall.Arguments, "attachment_names")

            Dim toProcess As List(Of AutoPilotAttachmentInfo)
            If targetNames.Count > 0 Then
                toProcess = _apCurrentAttachments?.Where(
                    Function(a) targetNames.Any(
                        Function(n) a.OriginalFileName.Equals(n, StringComparison.OrdinalIgnoreCase)
                    ) AndAlso Not a.IsOverSizeLimit AndAlso a.TempFilePath IsNot Nothing
                ).ToList()
            Else
                toProcess = _apCurrentAttachments?.Where(
                    Function(a) a.Extension = ".pdf" AndAlso
                                Not a.IsOverSizeLimit AndAlso a.TempFilePath IsNot Nothing
                ).ToList()
            End If

            If toProcess Is Nothing OrElse toProcess.Count = 0 Then
                response.Success = False
                response.Response = "No processable PDF attachments found."
                Return response
            End If

            Dim effectiveAuthor = If(String.IsNullOrWhiteSpace(author), AN6, author.Trim())
            Dim authorNote = If(effectiveAuthor.Equals(AN6, StringComparison.OrdinalIgnoreCase), "", $" (author: {effectiveAuthor})")
            Dim resultMessages As New List(Of String)()

            For Each att In toProcess
                context.Log($"Adding PDF comments to: {att.OriginalFileName} with instruction: {instruction}{authorNote}")
                ApDashboardLog($"💬 Adding PDF comments to: {att.OriginalFileName}{authorNote}", "step")

                If Not att.TempFilePath.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) Then
                    resultMessages.Add($"✗ {att.OriginalFileName}: Only PDF files are supported for PDF comment insertion.")
                    Continue For
                End If

                Dim outputName = Path.GetFileNameWithoutExtension(att.OriginalFileName) & "_commented.pdf"
                Dim outputPath = Path.Combine(_apCurrentTempDir, outputName)

                ' Prevent filename collision
                Dim counter = 1
                While File.Exists(outputPath)
                    outputName = Path.GetFileNameWithoutExtension(att.OriginalFileName) & $"_commented_{counter}.pdf"
                    outputPath = Path.Combine(_apCurrentTempDir, outputName)
                    counter += 1
                End While

                Dim success = Await CommentPdfForAutoPilot(att.TempFilePath, outputPath, instruction, ct, author)

                If success Then
                    att.OutputFiles.Add(outputPath)
                    resultMessages.Add($"✓ {att.OriginalFileName}: PDF comments added successfully. Output: {outputName}")
                    ApDashboardLog($"✓ PDF comments added to: {att.OriginalFileName}", "info")
                Else
                    resultMessages.Add($"✗ {att.OriginalFileName}: Failed to add PDF comments (document may be empty, image-only, or unsupported).")
                    ApDashboardLog($"⚠ Failed to add PDF comments to: {att.OriginalFileName}", "warn")
                End If
            Next

            response.Success = resultMessages.Any(Function(m) m.StartsWith("✓"))
            response.Response = String.Join(vbCrLf, resultMessages)

        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation was cancelled."
            response.Response = response.ErrorMessage
        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error adding comments to PDF(s): {ex.Message}"
        End Try

        Return response
    End Function


    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: create_code_file
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Async Function ExecuteCreateCodeFileTool(
            toolCall As ToolCall, context As ToolExecutionContext, ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim fileName = GetArgString(toolCall.Arguments, "file_name")
            Dim content = GetArgString(toolCall.Arguments, "content")
            Dim description = GetArgString(toolCall.Arguments, "description")

            If String.IsNullOrWhiteSpace(fileName) Then
                response.Success = False
                response.Response = "Missing required parameter: file_name"
                Return response
            End If

            If String.IsNullOrWhiteSpace(content) Then
                response.Success = False
                response.Response = "Missing required parameter: content"
                Return response
            End If

            ' Sanitize filename — preserve the extension but clean invalid chars
            For Each c In Path.GetInvalidFileNameChars()
                fileName = fileName.Replace(c, "_"c)
            Next
            fileName = fileName.Trim()

            ' Ensure the file has an extension; default to .txt if none provided
            If String.IsNullOrWhiteSpace(Path.GetExtension(fileName)) Then
                fileName &= ".txt"
            End If

            ' Guard against binary/Office extensions that should use dedicated tools
            Dim ext = Path.GetExtension(fileName).ToLowerInvariant()
            Dim blockedExtensions = {".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt", ".pdf", ".exe", ".dll", ".zip", ".rar"}
            If blockedExtensions.Contains(ext) Then
                response.Success = False
                response.Response = $"Cannot create binary file with extension '{ext}' using this tool. " &
                    "Use the dedicated tools (create_word_document, create_excel_spreadsheet, create_powerpoint, create_pdf_from_text) instead."
                Return response
            End If

            Dim outputPath = Path.Combine(_apCurrentTempDir, fileName)

            ' Prevent filename collision
            Dim counter = 1
            While File.Exists(outputPath)
                Dim baseName = Path.GetFileNameWithoutExtension(fileName)
                Dim extension = Path.GetExtension(fileName)
                fileName = baseName & $"_{counter}{extension}"
                outputPath = Path.Combine(_apCurrentTempDir, fileName)
                counter += 1
            End While

            context.Log($"Creating code file: {fileName}")
            ApDashboardLog($"💻 Creating code file: {fileName}", "step")

            ' Write the file with UTF-8 encoding (with BOM for maximum compatibility)
            Await Task.Run(Sub() File.WriteAllText(outputPath, content, Encoding.UTF8), ct)

            If File.Exists(outputPath) Then
                ' Register as output on the first attachment if available
                If _apCurrentAttachments IsNot Nothing AndAlso _apCurrentAttachments.Count > 0 Then
                    _apCurrentAttachments(0).OutputFiles.Add(outputPath)
                End If

                Dim sizeKb = New FileInfo(outputPath).Length / 1024
                Dim lineCount = content.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None).Length

                Dim resultMsg As New StringBuilder()
                resultMsg.Append($"Code file created: {fileName} ({lineCount} lines, {sizeKb:F0} KB)")
                If Not String.IsNullOrWhiteSpace(description) Then
                    resultMsg.Append($". {description}")
                End If
                resultMsg.Append(". The file will be attached to the reply.")

                response.Success = True
                response.Response = resultMsg.ToString()
                ApDashboardLog($"✓ Code file created: {fileName} ({lineCount} lines)", "info")
            Else
                response.Success = False
                response.Response = "Failed to create code file."
            End If

        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation was cancelled."
            response.Response = response.ErrorMessage
        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error creating code file: {ex.Message}"
        End Try

        Return response
    End Function


    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: extract_pdf_text
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Async Function ExecuteExtractPdfTextTool(
            toolCall As ToolCall, context As ToolExecutionContext, ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim targetNames = GetArgStringArray(toolCall.Arguments, "attachment_names")

            Dim toProcess As List(Of AutoPilotAttachmentInfo)
            If targetNames.Count > 0 Then
                toProcess = _apCurrentAttachments?.Where(
                    Function(a) targetNames.Any(Function(n) a.OriginalFileName.Equals(n, StringComparison.OrdinalIgnoreCase)) AndAlso
                                Not a.IsOverSizeLimit AndAlso a.TempFilePath IsNot Nothing).ToList()
            Else
                toProcess = _apCurrentAttachments?.Where(
                    Function(a) a.Extension = ".pdf" AndAlso Not a.IsOverSizeLimit AndAlso a.TempFilePath IsNot Nothing).ToList()
            End If

            If toProcess Is Nothing OrElse toProcess.Count = 0 Then
                response.Success = False
                response.Response = "No PDF attachments found to extract."
                Return response
            End If

            Dim sb As New StringBuilder()
            For Each att In toProcess
                context.Log($"Extracting text from: {att.OriginalFileName}")
                ApDashboardLog($"📄 Extracting text from: {att.OriginalFileName}", "step")

                Dim pdfResult As PdfReadResult = Await SharedMethods.ReadPdfAsTextEx(
                    att.TempFilePath, ReturnErrorInsteadOfEmpty:=True, DoOCR:=False, AskUser:=False, context:=_context)

                Dim text As String = If(pdfResult IsNot Nothing, pdfResult.Content, "")
                Dim usedOcr As Boolean = False

                Dim ocrAvailable As Boolean = SharedMethods.IsOcrAvailable(_context)
                Dim needsOcr As Boolean = String.IsNullOrWhiteSpace(text) OrElse
                    (pdfResult IsNot Nothing AndAlso pdfResult.OcrWasSkippedDueToHeuristics)

                If needsOcr AndAlso ocrAvailable Then
                    ApDashboardLog($"🔍 Running OCR on: {att.OriginalFileName}", "step")
                    context.Log($"OCR: {att.OriginalFileName}")
                    Dim ocrResult As PdfReadResult = Await SharedMethods.ReadPdfAsTextEx(
                        att.TempFilePath, ReturnErrorInsteadOfEmpty:=True, DoOCR:=True, AskUser:=False, context:=_context)
                    Dim ocrText As String = If(ocrResult IsNot Nothing, ocrResult.Content, "")
                    If Not String.IsNullOrWhiteSpace(ocrText) Then
                        text = ocrText
                        usedOcr = True
                        ApDashboardLog($"✓ OCR completed for: {att.OriginalFileName} ({ocrText.Length:N0} chars)", "info")
                    Else
                        ApDashboardLog($"⚠ OCR returned no content for: {att.OriginalFileName}, using standard extraction", "warn")
                    End If
                End If

                sb.AppendLine($"[{att.OriginalFileName}]")
                If String.IsNullOrWhiteSpace(text) Then
                    sb.AppendLine(If(Not ocrAvailable,
                        "(no extractable text; OCR is not available in the current configuration)",
                        "(no extractable text)"))
                    ApDashboardLog($"⚠ No text extracted from: {att.OriginalFileName}", "warn")
                Else
                    If Not usedOcr AndAlso pdfResult IsNot Nothing AndAlso pdfResult.OcrWasSkippedDueToHeuristics AndAlso Not ocrAvailable Then
                        sb.AppendLine("(Note: This PDF may contain scanned images. Some content may be missing because OCR is not available.)")
                    End If
                    sb.AppendLine(text)
                End If
                sb.AppendLine()
            Next

            response.Success = True
            response.Response = sb.ToString().TrimEnd()

        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error extracting PDF text: {ex.Message}"
        End Try

        Return response
    End Function



    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: merge_pdfs
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Async Function ExecuteMergePdfsTool(
            toolCall As ToolCall, context As ToolExecutionContext, ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim targetNames = GetArgStringArray(toolCall.Arguments, "attachment_names")
            Dim outputName = If(GetArgString(toolCall.Arguments, "output_filename"), "merged.pdf")

            Dim toMerge As List(Of AutoPilotAttachmentInfo)
            If targetNames.Count > 0 Then
                toMerge = New List(Of AutoPilotAttachmentInfo)()
                For Each name In targetNames
                    Dim found = _apCurrentAttachments?.FirstOrDefault(
                        Function(a) a.OriginalFileName.Equals(name, StringComparison.OrdinalIgnoreCase) AndAlso
                                    Not a.IsOverSizeLimit AndAlso a.TempFilePath IsNot Nothing)
                    If found IsNot Nothing Then toMerge.Add(found)
                Next
            Else
                toMerge = _apCurrentAttachments?.Where(
                    Function(a) a.Extension = ".pdf" AndAlso Not a.IsOverSizeLimit AndAlso a.TempFilePath IsNot Nothing).ToList()
            End If

            If toMerge Is Nothing OrElse toMerge.Count < 2 Then
                response.Success = False
                response.Response = "Need at least 2 PDF attachments to merge."
                Return response
            End If

            context.Log($"Merging {toMerge.Count} PDFs into {outputName}")
            Dim outputPath = Path.Combine(_apCurrentTempDir, outputName)

            Using outputDoc As New PdfSharp.Pdf.PdfDocument()
                For Each att In toMerge
                    Using inputDoc = PdfSharp.Pdf.IO.PdfReader.Open(att.TempFilePath, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import)
                        For Each page In inputDoc.Pages
                            outputDoc.AddPage(page)
                        Next
                    End Using
                Next
                outputDoc.Save(outputPath)
            End Using

            toMerge(0).OutputFiles.Add(outputPath)
            response.Success = True
            response.Response = $"Successfully merged {toMerge.Count} PDFs into {outputName} ({New FileInfo(outputPath).Length / 1024:F0} KB)."

        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error merging PDFs: {ex.Message}"
        End Try

        Return response
    End Function




    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: create_pdf_from_text
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Function ExecuteCreatePdfFromTextTool(toolCall As ToolCall, context As ToolExecutionContext) As ToolResponse
        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim content = GetArgString(toolCall.Arguments, "content")
            If String.IsNullOrWhiteSpace(content) Then
                response.Success = False
                response.Response = "Missing required parameter: content"
                Return response
            End If

            Dim outputName = If(GetArgString(toolCall.Arguments, "output_filename"), "output.pdf")
            Dim title = GetArgString(toolCall.Arguments, "title")
            Dim outputPath = Path.Combine(_apCurrentTempDir, outputName)

            context.Log($"Creating PDF: {outputName}")
            ApDashboardLog($"📝 Creating PDF: {outputName}", "step")

            ' Ensure font resolver is configured before any XFont usage
            EnsureApPdfSharpFontResolver()

            Using doc As New PdfSharp.Pdf.PdfDocument()
                doc.Info.Title = If(title, "Generated Document")

                Dim font = New PdfSharp.Drawing.XFont("Arial", 11)
                Dim titleFont = New PdfSharp.Drawing.XFont("Arial", 16, PdfSharp.Drawing.XFontStyleEx.Bold)
                Dim margin = 50.0
                Dim pageWidth = 595.0 ' A4
                Dim pageHeight = 842.0
                Dim usableWidth = pageWidth - 2 * margin
                Dim lineHeight = 15.0
                Dim y = margin

                Dim page = doc.AddPage()
                page.Width = pageWidth
                page.Height = pageHeight
                Dim gfx = PdfSharp.Drawing.XGraphics.FromPdfPage(page)

                ' Title
                If Not String.IsNullOrWhiteSpace(title) Then
                    gfx.DrawString(title, titleFont, PdfSharp.Drawing.XBrushes.Black,
                                   New PdfSharp.Drawing.XRect(margin, y, usableWidth, 30),
                                   PdfSharp.Drawing.XStringFormats.TopLeft)
                    y += 35
                End If

                ' Content lines
                Dim lines = content.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
                For Each line In lines
                    If y + lineHeight > pageHeight - margin Then
                        page = doc.AddPage()
                        page.Width = pageWidth
                        page.Height = pageHeight
                        gfx = PdfSharp.Drawing.XGraphics.FromPdfPage(page)
                        y = margin
                    End If

                    If Not String.IsNullOrEmpty(line) Then
                        gfx.DrawString(line, font, PdfSharp.Drawing.XBrushes.Black,
                                       New PdfSharp.Drawing.XRect(margin, y, usableWidth, lineHeight),
                                       PdfSharp.Drawing.XStringFormats.TopLeft)
                    End If
                    y += lineHeight
                Next

                doc.Save(outputPath)
            End Using

            ' Register as output
            If _apCurrentAttachments IsNot Nothing AndAlso _apCurrentAttachments.Count > 0 Then
                _apCurrentAttachments(0).OutputFiles.Add(outputPath)
            End If

            response.Success = True
            response.Response = $"PDF created: {outputName} ({New FileInfo(outputPath).Length / 1024:F0} KB)"
            ApDashboardLog($"✓ PDF created: {outputName}", "info")

        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error creating PDF: {ex.Message}"
        End Try

        Return response
    End Function



    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: split_pdf
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Function ExecuteSplitPdfTool(toolCall As ToolCall, context As ToolExecutionContext) As ToolResponse
        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim fileName = GetArgString(toolCall.Arguments, "attachment_name")
            Dim startPage = GetArgInt(toolCall.Arguments, "start_page", 0)
            Dim endPage = GetArgInt(toolCall.Arguments, "end_page", 0)
            Dim outputName = If(GetArgString(toolCall.Arguments, "output_filename"), "split.pdf")

            If String.IsNullOrWhiteSpace(fileName) OrElse startPage < 1 OrElse endPage < 1 Then
                response.Success = False
                response.Response = "Missing required parameters: attachment_name, start_page, end_page"
                Return response
            End If

            Dim att = FindAttachment(fileName)
            If att Is Nothing Then response.Success = False : response.Response = $"Attachment '{fileName}' not found." : Return response

            Dim outputPath = Path.Combine(_apCurrentTempDir, outputName)

            context.Log($"Splitting PDF: {fileName} pages {startPage}-{endPage}")

            Using inputDoc = PdfSharp.Pdf.IO.PdfReader.Open(att.TempFilePath, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import)
                If endPage > inputDoc.PageCount Then endPage = inputDoc.PageCount
                If startPage > inputDoc.PageCount Then
                    response.Success = False
                    response.Response = $"Start page {startPage} exceeds document page count ({inputDoc.PageCount})."
                    Return response
                End If

                Using outputDoc As New PdfSharp.Pdf.PdfDocument()
                    For i As Integer = startPage - 1 To endPage - 1
                        outputDoc.AddPage(inputDoc.Pages(i))
                    Next
                    outputDoc.Save(outputPath)
                End Using
            End Using

            att.OutputFiles.Add(outputPath)
            response.Success = True
            response.Response = $"Split PDF: pages {startPage}-{endPage} extracted to {outputName} ({New FileInfo(outputPath).Length / 1024:F0} KB)."

        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error splitting PDF: {ex.Message}"
        End Try

        Return response
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: add_pdf_watermark
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Function ExecuteAddPdfWatermarkTool(toolCall As ToolCall, context As ToolExecutionContext) As ToolResponse
        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim fileName = GetArgString(toolCall.Arguments, "attachment_name")
            Dim watermarkText = GetArgString(toolCall.Arguments, "watermark_text")
            Dim outputName = If(GetArgString(toolCall.Arguments, "output_filename"), "watermarked.pdf")

            If String.IsNullOrWhiteSpace(fileName) OrElse String.IsNullOrWhiteSpace(watermarkText) Then
                response.Success = False
                response.Response = "Missing required parameters: attachment_name, watermark_text"
                Return response
            End If

            Dim att = FindAttachment(fileName)
            If att Is Nothing Then response.Success = False : response.Response = $"Attachment '{fileName}' not found." : Return response

            Dim outputPath = Path.Combine(_apCurrentTempDir, outputName)

            context.Log($"Adding watermark to: {fileName}")
            ApDashboardLog($"💧 Adding watermark to: {fileName}", "step")

            ' Ensure font resolver is configured before any XFont usage
            EnsureApPdfSharpFontResolver()

            Dim inputSize As Long = New FileInfo(att.TempFilePath).Length
            Dim rasterizeWarning As String = Nothing

            ' Check for encryption — encrypted PDFs may silently produce invisible overlays
            Dim isEncrypted = APOverlayIsPdfEncrypted(att.TempFilePath)

            Dim directSuccess As Boolean = False

            If Not isEncrypted Then
                ' Try direct PdfSharp overlay first
                Dim tempOutputPath = outputPath & ".tmp_" & Guid.NewGuid().ToString("N") & ".pdf"
                Try
                    File.Copy(att.TempFilePath, tempOutputPath, True)
                    Using doc = PdfSharp.Pdf.IO.PdfReader.Open(tempOutputPath, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Modify)
                        Dim wmFont = New PdfSharp.Drawing.XFont("Arial", 60, PdfSharp.Drawing.XFontStyleEx.Bold)
                        Dim wmBrush = New PdfSharp.Drawing.XSolidBrush(
                            PdfSharp.Drawing.XColor.FromArgb(80, 180, 180, 180))

                        For Each page In doc.Pages
                            Using gfx = PdfSharp.Drawing.XGraphics.FromPdfPage(page, PdfSharp.Drawing.XGraphicsPdfPageOptions.Append)
                                Dim state = gfx.Save()
                                gfx.TranslateTransform(page.Width.Point / 2, page.Height.Point / 2)
                                gfx.RotateTransform(-45)
                                Dim size = gfx.MeasureString(watermarkText, wmFont)
                                gfx.DrawString(watermarkText, wmFont, wmBrush,
                                               New PdfSharp.Drawing.XRect(-size.Width / 2, -size.Height / 2, size.Width, size.Height),
                                               PdfSharp.Drawing.XStringFormats.Center)
                                gfx.Restore(state)
                            End Using
                        Next
                        doc.Save(tempOutputPath)
                    End Using

                    ' Verify output grew (watermark adds data)
                    Dim outputSize = New FileInfo(tempOutputPath).Length
                    If outputSize > inputSize Then
                        If File.Exists(outputPath) Then File.Delete(outputPath)
                        File.Move(tempOutputPath, outputPath)
                        directSuccess = True
                    End If
                Catch
                    directSuccess = False
                Finally
                    Try : If File.Exists(tempOutputPath) Then File.Delete(tempOutputPath)
                    Catch : End Try
                End Try

                ' Verify page count matches if direct succeeded
                If directSuccess Then
                    Dim pageWarning = APOverlayVerifyPageCount(att.TempFilePath, outputPath)
                    If pageWarning IsNot Nothing Then
                        ' Page count mismatch — discard and fall back to rasterize
                        Try : If File.Exists(outputPath) Then File.Delete(outputPath)
                        Catch : End Try
                        directSuccess = False
                    End If
                End If
            End If

            ' Fallback: rasterize all pages, then apply watermark on clean rasterized pages
            If Not directSuccess Then
                context.Log($"Direct watermark failed or PDF is encrypted — falling back to rasterize for: {fileName}")
                ApDashboardLog($"⚠ Rasterize fallback for watermark: {fileName}", "warn")

                Try
                    APOverlayViaRasterize(att.TempFilePath, outputPath,
                        Sub(gfx As PdfSharp.Drawing.XGraphics, pageW As Double, pageH As Double, pageIdx As Integer)
                            Dim wmFont = New PdfSharp.Drawing.XFont("Arial", 60, PdfSharp.Drawing.XFontStyleEx.Bold)
                            Dim wmBrush = New PdfSharp.Drawing.XSolidBrush(
                                PdfSharp.Drawing.XColor.FromArgb(80, 180, 180, 180))
                            Dim state = gfx.Save()
                            gfx.TranslateTransform(pageW / 2, pageH / 2)
                            gfx.RotateTransform(-45)
                            Dim size = gfx.MeasureString(watermarkText, wmFont)
                            gfx.DrawString(watermarkText, wmFont, wmBrush,
                                           New PdfSharp.Drawing.XRect(-size.Width / 2, -size.Height / 2, size.Width, size.Height),
                                           PdfSharp.Drawing.XStringFormats.Center)
                            gfx.Restore(state)
                        End Sub)

                    rasterizeWarning = If(isEncrypted,
                        "PDF is encrypted or restricted — was rasterized to ensure watermark visibility (text no longer selectable).",
                        "PDF could not be watermarked directly — was rasterized instead (text no longer selectable).")
                Catch rasterEx As Exception
                    Dim msg = rasterEx.Message
                    If String.IsNullOrWhiteSpace(msg) OrElse
                       msg.Equals("No error", StringComparison.OrdinalIgnoreCase) OrElse
                       msg.Equals("No error.", StringComparison.OrdinalIgnoreCase) Then
                        msg = "PDF appears to be encrypted or corrupt and could not be processed"
                    End If
                    response.Success = False
                    response.ErrorMessage = msg
                    response.Response = $"Error adding watermark: {msg}"
                    Return response
                End Try
            End If

            ' Validate output
            APOverlayValidateOutput(att.TempFilePath, outputPath)

            att.OutputFiles.Add(outputPath)
            response.Success = True
            Dim resultMsg = $"Watermark '{watermarkText}' added to {outputName} ({New FileInfo(outputPath).Length / 1024:F0} KB)."
            If Not String.IsNullOrWhiteSpace(rasterizeWarning) Then
                resultMsg &= $" Note: {rasterizeWarning}"
            End If
            response.Response = resultMsg
            ApDashboardLog($"✓ Watermark added: {outputName}", "info")

        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error adding watermark: {ex.Message}"
        End Try

        Return response
    End Function




    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: redact_pdf
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Async Function ExecuteRedactPdfTool(
            toolCall As ToolCall,
            context As ToolExecutionContext,
            ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim fileName = GetArgString(toolCall.Arguments, "attachment_name")
            If String.IsNullOrWhiteSpace(fileName) Then
                response.Success = False
                response.Response = "Missing required parameter: attachment_name"
                Return response
            End If

            Dim att = FindAttachment(fileName)
            If att Is Nothing Then
                response.Success = False
                response.Response = $"Attachment '{fileName}' not found."
                Return response
            End If

            If Not att.TempFilePath.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) Then
                response.Success = False
                response.Response = $"'{fileName}' is not a PDF file."
                Return response
            End If

            Dim instruction = GetArgString(toolCall.Arguments, "instruction")
            Dim mode = If(GetArgString(toolCall.Arguments, "mode"), "prepare").Trim().ToLowerInvariant()
            Dim includeReasonCodes = GetArgBool(toolCall.Arguments, "include_reason_codes", False)
            Dim outputName = GetArgString(toolCall.Arguments, "output_filename")

            ' Validate mode
            If mode <> "prepare" AndAlso mode <> "finalize" AndAlso mode <> "prepare_and_finalize" Then
                mode = "prepare"
            End If

            ' Instruction required for prepare modes
            If (mode = "prepare" OrElse mode = "prepare_and_finalize") AndAlso String.IsNullOrWhiteSpace(instruction) Then
                response.Success = False
                response.Response = "Missing required parameter: 'instruction' is required for prepare and prepare_and_finalize modes."
                Return response
            End If

            ' Determine output filename
            Dim baseName = Path.GetFileNameWithoutExtension(att.OriginalFileName)
            Dim suffix = If(mode = "finalize", "_final",
                         If(mode = "prepare_and_finalize", "_redacted_final", "_redacted"))

            If String.IsNullOrWhiteSpace(outputName) Then
                outputName = baseName & suffix & ".pdf"
            End If
            If Not outputName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) Then
                outputName &= ".pdf"
            End If

            Dim outputPath = Path.Combine(_apCurrentTempDir, outputName)

            ' Prevent filename collision
            Dim counter = 1
            While File.Exists(outputPath)
                outputName = baseName & suffix & $"_{counter}.pdf"
                outputPath = Path.Combine(_apCurrentTempDir, outputName)
                counter += 1
            End While

            context.Log($"PDF redaction ({mode}): {fileName}" &
                        If(Not String.IsNullOrWhiteSpace(instruction), $" — {instruction}", ""))
            ApDashboardLog($"🔒 PDF redaction ({mode}): {fileName}", "step")

            EnsureApPdfSharpFontResolver()

            If mode = "finalize" Then
                ' Finalize-only: burn in existing annotations
                Await Task.Run(Sub() APRedactFinalizeOnly(att.TempFilePath, outputPath, includeReasonCodes, 300))

                att.OutputFiles.Add(outputPath)
                response.Success = True
                response.Response = $"PDF finalized (annotations burned in): {outputName} ({New FileInfo(outputPath).Length / 1024:F0} KB). " &
                    "All redaction boxes are now permanent black rectangles."
                ApDashboardLog($"✓ PDF finalized: {outputName}", "info")

            Else
                ' Prepare (with optional finalize)
                Dim finalize As Boolean = (mode = "prepare_and_finalize")
                Dim result = Await APRedactPdf(att.TempFilePath, outputPath, instruction,
                                                finalize, includeReasonCodes, ct)

                If result Is Nothing Then
                    response.Success = False
                    response.Response = $"Redaction failed for '{fileName}'. The PDF may contain no extractable text " &
                        "(run OCR first), the AI may have returned an empty/unparseable response, or no matching " &
                        "text was found for the identified redactions. You may want to retry."
                    ApDashboardLog($"⚠ Redaction failed for: {fileName}", "warn")
                    Return response
                End If

                If result = "no_redactions" Then
                    response.Success = True
                    response.Response = $"The AI found nothing to redact in '{fileName}' based on the instruction: '{instruction}'."
                    ApDashboardLog($"ℹ No redactions found in: {fileName}", "info")
                    Return response
                End If

                att.OutputFiles.Add(outputPath)
                response.Success = True
                Dim sizeKb = If(File.Exists(outputPath), $" ({New FileInfo(outputPath).Length / 1024:F0} KB)", "")
                response.Response = $"PDF redacted: {result} Output: {outputName}{sizeKb}."
                If Not finalize Then
                    response.Response &= " Note: redaction boxes are currently removable annotations. " &
                        "Call this tool again with mode='finalize' on the output file to make them permanent."
                End If
                ApDashboardLog($"✓ PDF redacted: {outputName}", "info")
            End If

        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation was cancelled."
            response.Response = response.ErrorMessage
        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error during PDF redaction: {ex.Message}"
        End Try

        Return response
    End Function


    ' ═══════════════════════════════════════════════════════════════════════════
    '  HELPER: Ensure PdfSharp font resolver is configured
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Shared _apFontResolverConfigured As Boolean = False

    ''' <summary>
    ''' Ensures a PdfSharp IFontResolver is registered so that font-by-name lookups work.
    ''' Must be called before creating any XFont. Safe to call multiple times.
    ''' </summary>
    Private Shared Sub EnsureApPdfSharpFontResolver()
        If _apFontResolverConfigured Then Return
        Try
            If PdfSharp.Fonts.GlobalFontSettings.FontResolver Is Nothing Then
                PdfSharp.Fonts.GlobalFontSettings.FontResolver = New ApFontResolver()
            End If
            _apFontResolverConfigured = True
        Catch
            ' Already set or locked — ignore
            _apFontResolverConfigured = True
        End Try
    End Sub

    ''' <summary>
    ''' Minimal font resolver for PdfSharp that loads system fonts by reading .ttf files from the Windows Fonts folder.
    ''' </summary>
    Private Class ApFontResolver
        Implements PdfSharp.Fonts.IFontResolver

        Private Shared ReadOnly _fontCache As New Dictionary(Of String, Byte())(StringComparer.OrdinalIgnoreCase)

        ''' <summary>
        ''' Resolves a requested font family/style combination to a concrete font face key.
        ''' </summary>
        Public Function ResolveTypeface(familyName As String, bold As Boolean, italic As Boolean) As PdfSharp.Fonts.FontResolverInfo _
                Implements PdfSharp.Fonts.IFontResolver.ResolveTypeface

            ' Map common family names to Windows font filenames
            Dim key As String = "arial.ttf"
            Select Case familyName.ToLowerInvariant()
                Case "arial", "helvetica"
                    If bold AndAlso italic Then
                        key = "arialbi.ttf"
                    ElseIf bold Then
                        key = "arialbd.ttf"
                    ElseIf italic Then
                        key = "ariali.ttf"
                    Else
                        key = "arial.ttf"
                    End If
                Case "times new roman", "times"
                    If bold AndAlso italic Then
                        key = "timesbi.ttf"
                    ElseIf bold Then
                        key = "timesbd.ttf"
                    ElseIf italic Then
                        key = "timesi.ttf"
                    Else
                        key = "times.ttf"
                    End If
                Case "courier new", "courier"
                    If bold AndAlso italic Then
                        key = "courbi.ttf"
                    ElseIf bold Then
                        key = "courbd.ttf"
                    ElseIf italic Then
                        key = "couri.ttf"
                    Else
                        key = "cour.ttf"
                    End If
                Case "segoe ui"
                    If bold AndAlso italic Then
                        key = "segoeuiz.ttf"
                    ElseIf bold Then
                        key = "segoeuib.ttf"
                    ElseIf italic Then
                        key = "segoeuii.ttf"
                    Else
                        key = "segoeui.ttf"
                    End If
                Case "calibri"
                    If bold AndAlso italic Then
                        key = "calibriz.ttf"
                    ElseIf bold Then
                        key = "calibrib.ttf"
                    ElseIf italic Then
                        key = "calibrii.ttf"
                    Else
                        key = "calibri.ttf"
                    End If
                Case "verdana"
                    If bold AndAlso italic Then
                        key = "verdanaz.ttf"
                    ElseIf bold Then
                        key = "verdanab.ttf"
                    ElseIf italic Then
                        key = "verdanai.ttf"
                    Else
                        key = "verdana.ttf"
                    End If
                Case "tahoma"
                    If bold Then
                        key = "tahomabd.ttf"
                    Else
                        key = "tahoma.ttf"
                    End If
                Case "georgia"
                    If bold AndAlso italic Then
                        key = "georgiaz.ttf"
                    ElseIf bold Then
                        key = "georgiab.ttf"
                    ElseIf italic Then
                        key = "georgiai.ttf"
                    Else
                        key = "georgia.ttf"
                    End If
                Case "trebuchet ms"
                    If bold AndAlso italic Then
                        key = "trebucbi.ttf"
                    ElseIf bold Then
                        key = "trebucbd.ttf"
                    ElseIf italic Then
                        key = "trebucit.ttf"
                    Else
                        key = "trebuc.ttf"
                    End If
                Case Else
                    ' Fallback to Arial
                    If bold Then
                        key = "arialbd.ttf"
                    Else
                        key = "arial.ttf"
                    End If
            End Select

            Return New PdfSharp.Fonts.FontResolverInfo(key)
        End Function

        ''' <summary>
        ''' Loads raw font bytes for a resolved face name from Windows font locations.
        ''' </summary>
        Public Function GetFont(faceName As String) As Byte() _
                Implements PdfSharp.Fonts.IFontResolver.GetFont

            SyncLock _fontCache
                If _fontCache.ContainsKey(faceName) Then Return _fontCache(faceName)
            End SyncLock

            Dim fontsDir = Environment.GetFolderPath(Environment.SpecialFolder.Fonts)
            Dim fontPath = Path.Combine(fontsDir, faceName)

            ' Also check Windows\Fonts directly (SpecialFolder.Fonts may return user fonts folder on some systems)
            If Not File.Exists(fontPath) Then
                fontPath = Path.Combine(Environment.GetEnvironmentVariable("SystemRoot"), "Fonts", faceName)
            End If

            If File.Exists(fontPath) Then
                Dim data = File.ReadAllBytes(fontPath)
                SyncLock _fontCache
                    _fontCache(faceName) = data
                End SyncLock
                Return data
            End If

            ' Last resort: return Nothing — PdfSharp will throw a descriptive error
            Return Nothing
        End Function
    End Class



    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: overlay_pdf
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Async Function ExecuteOverlayPdfTool(
            toolCall As ToolCall,
            context As ToolExecutionContext,
            ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim fileName = GetArgString(toolCall.Arguments, "attachment_name")
            If String.IsNullOrWhiteSpace(fileName) Then
                response.Success = False
                response.Response = "Missing required parameter: attachment_name"
                Return response
            End If

            Dim att = FindAttachment(fileName)
            If att Is Nothing Then
                response.Success = False
                response.Response = $"Attachment '{fileName}' not found."
                Return response
            End If

            If Not att.TempFilePath.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) Then
                response.Success = False
                response.Response = $"'{fileName}' is not a PDF file."
                Return response
            End If

            ' Parse elements array
            Dim elementsArray As JArray = Nothing
            If toolCall.Arguments IsNot Nothing AndAlso toolCall.Arguments.ContainsKey("elements") Then
                Dim elemObj = toolCall.Arguments("elements")
                If TypeOf elemObj Is JArray Then elementsArray = DirectCast(elemObj, JArray)
            End If

            If elementsArray Is Nothing OrElse elementsArray.Count = 0 Then
                response.Success = False
                response.Response = "Missing required parameter: elements (must be a non-empty array)"
                Return response
            End If

            ' Determine output filename
            Dim outputName = GetArgString(toolCall.Arguments, "output_filename")
            If String.IsNullOrWhiteSpace(outputName) Then
                outputName = Path.GetFileNameWithoutExtension(att.OriginalFileName) & "_overlay.pdf"
            End If
            If Not outputName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) Then
                outputName &= ".pdf"
            End If

            Dim outputPath = Path.Combine(_apCurrentTempDir, outputName)

            ' Prevent filename collision
            Dim counter = 1
            While File.Exists(outputPath)
                outputName = Path.GetFileNameWithoutExtension(att.OriginalFileName) & $"_overlay_{counter}.pdf"
                outputPath = Path.Combine(_apCurrentTempDir, outputName)
                counter += 1
            End While

            context.Log($"Overlaying {elementsArray.Count} element(s) on: {fileName}")
            ApDashboardLog($"🖌 Overlaying {elementsArray.Count} element(s) on: {fileName}", "step")

            EnsureApPdfSharpFontResolver()

            ' Pre-resolve all image attachments to avoid repeated lookups
            Dim imageCache As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            For Each elemObj As JObject In elementsArray
                Dim elemType = If(elemObj.Value(Of String)("type"), "text").ToLowerInvariant()
                If elemType = "image" Then
                    Dim imgName = elemObj.Value(Of String)("image_attachment_name")
                    If Not String.IsNullOrWhiteSpace(imgName) AndAlso Not imageCache.ContainsKey(imgName) Then
                        Dim imgAtt = FindAttachment(imgName)
                        If imgAtt IsNot Nothing AndAlso imgAtt.TempFilePath IsNot Nothing AndAlso File.Exists(imgAtt.TempFilePath) Then
                            imageCache(imgName) = imgAtt.TempFilePath
                        End If
                    End If
                End If
            Next

            Dim inputSize As Long = New FileInfo(att.TempFilePath).Length
            Dim rasterizeWarning As String = Nothing
            Dim textCount = 0
            Dim imageCount = 0

            ' Check for encryption — encrypted PDFs may silently produce invisible overlays
            Dim isEncrypted = APOverlayIsPdfEncrypted(att.TempFilePath)

            Dim directSuccess As Boolean = False

            If Not isEncrypted Then
                ' ── Try direct PdfSharp overlay first ──
                Dim tempWorkPath = outputPath & ".tmp_" & Guid.NewGuid().ToString("N") & ".pdf"
                Dim directTextCount = 0
                Dim directImageCount = 0

                Try
                    File.Copy(att.TempFilePath, tempWorkPath, True)

                    Using doc = PdfSharp.Pdf.IO.PdfReader.Open(tempWorkPath, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Modify)
                        Dim totalPages = doc.PageCount

                        APOverlayDrawElements(doc, totalPages, elementsArray, imageCache, context,
                                              directTextCount, directImageCount)

                        doc.Save(tempWorkPath)
                    End Using

                    ' Verify output grew (overlay adds image/text data)
                    Dim outputSize = New FileInfo(tempWorkPath).Length
                    If outputSize > inputSize Then
                        If File.Exists(outputPath) Then File.Delete(outputPath)
                        File.Move(tempWorkPath, outputPath)
                        directSuccess = True
                        textCount = directTextCount
                        imageCount = directImageCount
                    End If
                Catch
                    directSuccess = False
                Finally
                    Try : If File.Exists(tempWorkPath) Then File.Delete(tempWorkPath)
                    Catch : End Try
                End Try

                ' Verify page count matches if direct succeeded
                If directSuccess Then
                    Dim pageWarning = APOverlayVerifyPageCount(att.TempFilePath, outputPath)
                    If pageWarning IsNot Nothing Then
                        ' Page count mismatch — discard and fall back to rasterize
                        Try : If File.Exists(outputPath) Then File.Delete(outputPath)
                        Catch : End Try
                        directSuccess = False
                    End If
                End If
            End If

            ' ── Fallback: rasterize affected pages, then re-apply overlay ──
            If Not directSuccess Then
                context.Log($"Direct overlay failed or PDF is encrypted — falling back to rasterize for: {fileName}")
                ApDashboardLog($"⚠ Rasterize fallback for overlay: {fileName}", "warn")

                textCount = 0
                imageCount = 0

                Try
                    ' Build a lookup: which page indices need overlay elements?
                    ' We rasterize ALL pages (to handle encrypted PDFs where even non-overlaid
                    ' pages can't be copied via PdfSharp Import mode), then draw elements on
                    ' the appropriate pages.
                    APOverlayViaRasterizeWithElements(att.TempFilePath, outputPath, elementsArray,
                                                      imageCache, context, textCount, imageCount)

                    rasterizeWarning = If(isEncrypted,
                        "PDF is encrypted or restricted — was rasterized to ensure overlay visibility (text no longer selectable).",
                        "PDF could not be overlaid directly — was rasterized instead (text no longer selectable).")
                Catch rasterEx As Exception
                    Dim msg = rasterEx.Message
                    If String.IsNullOrWhiteSpace(msg) OrElse
                       msg.Equals("No error", StringComparison.OrdinalIgnoreCase) OrElse
                       msg.Equals("No error.", StringComparison.OrdinalIgnoreCase) Then
                        msg = "PDF appears to be encrypted or corrupt and could not be processed"
                    End If
                    response.Success = False
                    response.ErrorMessage = msg
                    response.Response = $"Error overlaying PDF: {msg}"
                    Return response
                End Try
            End If

            ' Validate output
            APOverlayValidateOutput(att.TempFilePath, outputPath)

            If File.Exists(outputPath) Then
                att.OutputFiles.Add(outputPath)
                response.Success = True
                Dim resultMsg = $"PDF overlay complete: {textCount} text element(s) and {imageCount} image element(s) placed. " &
                    $"Output: {outputName} ({New FileInfo(outputPath).Length / 1024:F0} KB). The file will be attached to the reply."
                If Not String.IsNullOrWhiteSpace(rasterizeWarning) Then
                    resultMsg &= $" Note: {rasterizeWarning}"
                End If
                response.Response = resultMsg
                ApDashboardLog($"✓ PDF overlay: {outputName} ({textCount} text, {imageCount} image)", "info")
            Else
                response.Success = False
                response.Response = "Failed to create overlaid PDF."
            End If

        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation was cancelled."
            response.Response = response.ErrorMessage
        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error overlaying PDF: {ex.Message}"
        End Try

        Return response
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  OVERLAY PDF HELPERS
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Parses a page specification string into a list of 0-based page indices.
    ''' Supports: "all", "1", "1,3,5", "2-5", "1,3-5,8".
    ''' </summary>
    Private Shared Function ResolvePageIndices(pagesSpec As String, totalPages As Integer) As List(Of Integer)
        Dim result As New List(Of Integer)()
        If String.IsNullOrWhiteSpace(pagesSpec) OrElse pagesSpec = "all" Then
            For i = 0 To totalPages - 1
                result.Add(i)
            Next
            Return result
        End If

        ' Split on commas, then handle each token
        For Each token In pagesSpec.Split(","c)
            Dim trimmed = token.Trim()
            If String.IsNullOrEmpty(trimmed) Then Continue For

            Dim dashIdx = trimmed.IndexOf("-"c)
            If dashIdx > 0 Then
                ' Range: "2-5"
                Dim startStr = trimmed.Substring(0, dashIdx).Trim()
                Dim endStr = trimmed.Substring(dashIdx + 1).Trim()
                Dim startPage As Integer
                Dim endPage As Integer
                If Integer.TryParse(startStr, startPage) AndAlso Integer.TryParse(endStr, endPage) Then
                    startPage = Math.Max(1, startPage)
                    endPage = Math.Min(totalPages, endPage)
                    For i = startPage To endPage
                        If Not result.Contains(i - 1) Then result.Add(i - 1)
                    Next
                End If
            Else
                ' Single page: "3"
                Dim pageNum As Integer
                If Integer.TryParse(trimmed, pageNum) AndAlso pageNum >= 1 AndAlso pageNum <= totalPages Then
                    If Not result.Contains(pageNum - 1) Then result.Add(pageNum - 1)
                End If
            End If
        Next

        Return result
    End Function

    ''' <summary>
    ''' Reads a Double value from a JObject token with a fallback default.
    ''' </summary>
    Private Shared Function GetJDouble(obj As JObject, key As String, defaultVal As Double) As Double
        Dim token = obj(key)
        If token Is Nothing Then Return defaultVal
        Dim result As Double
        If Double.TryParse(token.ToString(), Globalization.NumberStyles.Any,
                          Globalization.CultureInfo.InvariantCulture, result) Then
            Return result
        End If
        Return defaultVal
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  PDF OVERLAY/WATERMARK HARDENING HELPERS
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Checks whether a PDF file is encrypted by inspecting PdfSharp's SecurityHandler
    ''' and falling back to a chunked byte-level scan for the /Encrypt marker.
    ''' Mirrors ExhibitStampService.IsPdfEncrypted from the Word add-in.
    ''' </summary>
    Private Shared Function APOverlayIsPdfEncrypted(pdfPath As String) As Boolean
        ' Method 1: Try PdfSharp — check SecurityHandler
        Try
            Using doc As PdfSharp.Pdf.PdfDocument =
                    PdfSharp.Pdf.IO.PdfReader.Open(pdfPath, PdfSharp.Pdf.IO.PdfDocumentOpenMode.InformationOnly)
                If doc.SecurityHandler IsNot Nothing Then
                    Return True
                End If
            End Using
        Catch
            ' PdfSharp couldn't open it at all — likely encrypted with a user password.
            ' Fall through to byte-level scan.
        End Try

        ' Method 2: Byte-level scan for /Encrypt marker (chunked for large files)
        Try
            Const chunkSize As Integer = 65536
            Dim overlap As Integer = 7 ' Length of "/Encrypt" minus 1

            Using fs As New FileStream(pdfPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                Dim buffer(chunkSize + overlap - 1) As Byte
                Dim carryOver As Integer = 0

                While True
                    Dim bytesRead As Integer = fs.Read(buffer, carryOver, chunkSize)
                    If bytesRead = 0 Then Exit While

                    Dim totalInBuffer As Integer = carryOver + bytesRead
                    Dim text As String = System.Text.Encoding.ASCII.GetString(buffer, 0, totalInBuffer)
                    If text.IndexOf("/Encrypt", StringComparison.Ordinal) >= 0 Then
                        Return True
                    End If

                    ' Keep the last few bytes for overlap into next chunk
                    If totalInBuffer > overlap Then
                        Array.Copy(buffer, totalInBuffer - overlap, buffer, 0, overlap)
                        carryOver = overlap
                    Else
                        carryOver = totalInBuffer
                    End If
                End While
            End Using

            Return False
        Catch
            Return True ' Can't read → assume encrypted
        End Try
    End Function

    ''' <summary>
    ''' Counts pages in a PDF file using PdfPig (read-only, handles most PDF types).
    ''' Returns -1 if the file cannot be read.
    ''' </summary>
    Private Shared Function APOverlayGetPdfPageCount(pdfPath As String) As Integer
        Try
            Using doc As UglyToad.PdfPig.PdfDocument = UglyToad.PdfPig.PdfDocument.Open(pdfPath)
                Return doc.NumberOfPages
            End Using
        Catch
            Return -1
        End Try
    End Function

    ''' <summary>
    ''' Compares the page count of an original PDF with its output.
    ''' Returns a warning string if they differ, or Nothing if they match.
    ''' </summary>
    Private Shared Function APOverlayVerifyPageCount(originalPath As String, outputPath As String) As String
        Dim origPages = APOverlayGetPdfPageCount(originalPath)
        Dim outPages = APOverlayGetPdfPageCount(outputPath)

        If origPages = -1 OrElse outPages = -1 Then
            Return $"Could not verify page count (original={If(origPages = -1, "unreadable", origPages.ToString())}, output={If(outPages = -1, "unreadable", outPages.ToString())})."
        End If

        If origPages <> outPages Then
            Return $"PAGE COUNT MISMATCH: original has {origPages} page(s) but output has {outPages} page(s)."
        End If

        Return Nothing
    End Function

    ''' <summary>
    ''' Validates that an output PDF file exists, is non-empty, and readable.
    ''' Throws <see cref="InvalidOperationException"/> on failure.
    ''' </summary>
    Private Shared Sub APOverlayValidateOutput(inputPath As String, outputPath As String)
        If Not File.Exists(outputPath) Then
            Throw New InvalidOperationException(
                "Output file was not created — the source PDF may be encrypted or corrupt.")
        End If

        Dim outputSize As Long = New FileInfo(outputPath).Length
        If outputSize = 0 Then
            Try : File.Delete(outputPath) : Catch : End Try
            Throw New InvalidOperationException(
                "Output file is empty — the source PDF may be encrypted or corrupt.")
        End If

        Dim pageCount = APOverlayGetPdfPageCount(outputPath)
        If pageCount = 0 Then
            Try : File.Delete(outputPath) : Catch : End Try
            Throw New InvalidOperationException(
                "Output PDF contains no pages — the source PDF may be encrypted or corrupt.")
        End If
    End Sub

    ''' <summary>
    ''' Draws overlay elements (text and image) onto PdfSharp document pages.
    ''' Shared by both the direct overlay path and the rasterize fallback.
    ''' </summary>
    Private Sub APOverlayDrawElements(
            doc As PdfSharp.Pdf.PdfDocument,
            totalPages As Integer,
            elementsArray As JArray,
            imageCache As Dictionary(Of String, String),
            context As ToolExecutionContext,
            ByRef textCount As Integer,
            ByRef imageCount As Integer)

        For Each elemObj As JObject In elementsArray
            Dim elemType = If(elemObj.Value(Of String)("type"), "text").ToLowerInvariant()
            Dim pagesSpec = If(elemObj.Value(Of String)("pages"), "all").Trim().ToLowerInvariant()
            Dim x As Double = GetJDouble(elemObj, "x", 0)
            Dim y As Double = GetJDouble(elemObj, "y", 0)
            Dim rotation As Double = GetJDouble(elemObj, "rotation", 0)
            Dim opacity As Double = GetJDouble(elemObj, "opacity", 1.0)

            ' Resolve target page indices (0-based)
            Dim pageIndices = ResolvePageIndices(pagesSpec, totalPages)
            If pageIndices.Count = 0 Then Continue For

            For Each pageIdx In pageIndices
                If pageIdx < 0 OrElse pageIdx >= totalPages Then Continue For
                Dim page = doc.Pages(pageIdx)

                Using gfx = PdfSharp.Drawing.XGraphics.FromPdfPage(page, PdfSharp.Drawing.XGraphicsPdfPageOptions.Append)
                    If elemType = "text" Then
                        ' ── TEXT ELEMENT ──
                        Dim text = If(elemObj.Value(Of String)("text"), "")
                        If String.IsNullOrEmpty(text) Then Continue For

                        ' Handle \n escape sequences for multi-line text
                        text = text.Replace("\\n", vbLf).Replace("\n", vbLf)

                        Dim fontFamily = If(elemObj.Value(Of String)("font_family"), "Arial")
                        Dim fontSize As Double = GetJDouble(elemObj, "font_size", 12)
                        Dim isBold = GetJBool(elemObj, "bold")
                        Dim isItalic = GetJBool(elemObj, "italic")
                        Dim hAlign = If(elemObj.Value(Of String)("h_align"), "left").ToLowerInvariant()
                        Dim maxWidth As Double = GetJDouble(elemObj, "max_width", 0)
                        Dim fontColorHex = If(elemObj.Value(Of String)("font_color"), "#000000")

                        ' Build font style
                        Dim fontStyle As PdfSharp.Drawing.XFontStyleEx = PdfSharp.Drawing.XFontStyleEx.Regular
                        If isBold AndAlso isItalic Then
                            fontStyle = PdfSharp.Drawing.XFontStyleEx.BoldItalic
                        ElseIf isBold Then
                            fontStyle = PdfSharp.Drawing.XFontStyleEx.Bold
                        ElseIf isItalic Then
                            fontStyle = PdfSharp.Drawing.XFontStyleEx.Italic
                        End If

                        Dim font = New PdfSharp.Drawing.XFont(fontFamily, fontSize, fontStyle)

                        ' Parse color
                        Dim brush As PdfSharp.Drawing.XBrush = PdfSharp.Drawing.XBrushes.Black
                        Try
                            Dim colorHex = fontColorHex.TrimStart("#"c)
                            If colorHex.Length = 6 Then
                                Dim r = System.Convert.ToInt32(colorHex.Substring(0, 2), 16)
                                Dim g = System.Convert.ToInt32(colorHex.Substring(2, 2), 16)
                                Dim b = System.Convert.ToInt32(colorHex.Substring(4, 2), 16)
                                Dim alphaInt = CInt(Math.Round(Math.Max(0, Math.Min(1, opacity)) * 255))
                                brush = New PdfSharp.Drawing.XSolidBrush(
                                    PdfSharp.Drawing.XColor.FromArgb(alphaInt, r, g, b))
                            End If
                        Catch
                        End Try

                        ' Determine string format for alignment
                        Dim xFormat As New PdfSharp.Drawing.XStringFormat()
                        xFormat.LineAlignment = PdfSharp.Drawing.XLineAlignment.Near
                        Select Case hAlign
                            Case "center" : xFormat.Alignment = PdfSharp.Drawing.XStringAlignment.Center
                            Case "right" : xFormat.Alignment = PdfSharp.Drawing.XStringAlignment.Far
                            Case Else : xFormat.Alignment = PdfSharp.Drawing.XStringAlignment.Near
                        End Select

                        ' Apply rotation if specified
                        Dim state As PdfSharp.Drawing.XGraphicsState = Nothing
                        If rotation <> 0 Then
                            state = gfx.Save()
                            gfx.TranslateTransform(x, y)
                            gfx.RotateTransform(rotation)
                            gfx.TranslateTransform(-x, -y)
                        End If

                        ' Handle multi-line text
                        Dim lines = text.Split({vbLf}, StringSplitOptions.None)
                        Dim lineHeight = fontSize * 1.25
                        Dim currentY = y

                        For Each line In lines
                            If maxWidth > 0 Then
                                Dim rect As New PdfSharp.Drawing.XRect(x, currentY, maxWidth, lineHeight)
                                gfx.DrawString(line, font, brush, rect, xFormat)
                            Else
                                Dim drawPoint As New PdfSharp.Drawing.XPoint(x, currentY)
                                gfx.DrawString(line, font, brush, drawPoint, xFormat)
                            End If
                            currentY += lineHeight
                        Next

                        If state IsNot Nothing Then gfx.Restore(state)
                        textCount += 1

                    ElseIf elemType = "image" Then
                        ' ── IMAGE ELEMENT ──
                        Dim imgName = elemObj.Value(Of String)("image_attachment_name")
                        If String.IsNullOrWhiteSpace(imgName) Then Continue For

                        Dim imgPath As String = Nothing
                        If Not imageCache.TryGetValue(imgName, imgPath) OrElse
                           String.IsNullOrEmpty(imgPath) OrElse Not File.Exists(imgPath) Then
                            context.Log($"Image attachment not found: {imgName}")
                            Continue For
                        End If

                        Dim imgWidth As Double = GetJDouble(elemObj, "width", 0)
                        Dim imgHeight As Double = GetJDouble(elemObj, "height", 0)

                        ' Load image via stream to support all formats
                        Using imgStream As New FileStream(imgPath, FileMode.Open, FileAccess.Read, FileShare.Read)
                            Using xImg = PdfSharp.Drawing.XImage.FromStream(imgStream)
                                ' Default to native size if not specified
                                If imgWidth <= 0 AndAlso imgHeight <= 0 Then
                                    imgWidth = xImg.PointWidth
                                    imgHeight = xImg.PointHeight
                                ElseIf imgWidth > 0 AndAlso imgHeight <= 0 Then
                                    ' Scale proportionally
                                    imgHeight = xImg.PointHeight * (imgWidth / xImg.PointWidth)
                                ElseIf imgHeight > 0 AndAlso imgWidth <= 0 Then
                                    imgWidth = xImg.PointWidth * (imgHeight / xImg.PointHeight)
                                End If

                                ' Apply rotation if specified
                                Dim state As PdfSharp.Drawing.XGraphicsState = Nothing
                                If rotation <> 0 Then
                                    state = gfx.Save()
                                    Dim cx = x + imgWidth / 2
                                    Dim cy = y + imgHeight / 2
                                    gfx.TranslateTransform(cx, cy)
                                    gfx.RotateTransform(rotation)
                                    gfx.TranslateTransform(-cx, -cy)
                                End If

                                gfx.DrawImage(xImg, x, y, imgWidth, imgHeight)

                                If state IsNot Nothing Then gfx.Restore(state)
                                imageCount += 1
                            End Using
                        End Using
                    End If
                End Using
            Next
        Next
    End Sub

    ''' <summary>
    ''' Rasterize fallback for watermark: renders every page via PdfiumViewer, then
    ''' invokes a callback to draw the watermark on each rasterized page.
    ''' CRITICAL: PdfiumViewer types must NOT appear in the calling method's signature
    ''' to avoid JIT resolution before pdfium.dll is loaded.
    ''' </summary>
    Private Shared Sub APOverlayViaRasterize(
            inputPath As String,
            outputPath As String,
            drawCallback As Action(Of PdfSharp.Drawing.XGraphics, Double, Double, Integer))

        APRedactEnsurePdfiumLoaded()
        EnsureApPdfSharpFontResolver()
        APOverlayViaRasterizeCore(inputPath, outputPath, drawCallback)
    End Sub

    ''' <summary>
    ''' Core rasterize implementation for watermark, separated to ensure pdfium.dll
    ''' is loaded before PdfiumViewer types are JIT-resolved.
    ''' </summary>
    <Runtime.CompilerServices.MethodImpl(Runtime.CompilerServices.MethodImplOptions.NoInlining)>
    Private Shared Sub APOverlayViaRasterizeCore(
            inputPath As String,
            outputPath As String,
            drawCallback As Action(Of PdfSharp.Drawing.XGraphics, Double, Double, Integer))

        Const renderDpi As Integer = 200

        Using pdf As PdfiumViewer.PdfDocument = PdfiumViewer.PdfDocument.Load(inputPath)
            If pdf.PageCount = 0 Then
                Throw New InvalidOperationException("The PDF contains no pages.")
            End If

            Dim outDoc As New PdfSharp.Pdf.PdfDocument()

            Dim renderFlags As PdfiumViewer.PdfRenderFlags =
                PdfiumViewer.PdfRenderFlags.Annotations Or
                PdfiumViewer.PdfRenderFlags.LcdText Or
                PdfiumViewer.PdfRenderFlags.ForPrinting

            For pageIndex As Integer = 0 To pdf.PageCount - 1
                Dim sizePt As System.Drawing.SizeF = pdf.PageSizes(pageIndex)
                Dim widthPx As Integer = CInt(Math.Round(sizePt.Width / 72.0 * renderDpi))
                Dim heightPx As Integer = CInt(Math.Round(sizePt.Height / 72.0 * renderDpi))

                ' Declare outPage OUTSIDE the Using rendered block so it stays in scope
                Dim outPage As PdfSharp.Pdf.PdfPage = outDoc.AddPage()
                outPage.Width = PdfSharp.Drawing.XUnit.FromPoint(sizePt.Width)
                outPage.Height = PdfSharp.Drawing.XUnit.FromPoint(sizePt.Height)

                Using rendered As System.Drawing.Image =
                    pdf.Render(pageIndex, widthPx, heightPx, renderDpi, renderDpi, renderFlags)

                    ' Draw rasterized page image
                    Using ms As New MemoryStream()
                        Dim jpegEncoder As System.Drawing.Imaging.ImageCodecInfo = Nothing
                        For Each codec In System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders()
                            If codec.MimeType = "image/jpeg" Then jpegEncoder = codec : Exit For
                        Next
                        If jpegEncoder IsNot Nothing Then
                            Dim ep As New System.Drawing.Imaging.EncoderParameters(1)
                            ep.Param(0) = New System.Drawing.Imaging.EncoderParameter(
                                System.Drawing.Imaging.Encoder.Quality, 85L)
                            rendered.Save(ms, jpegEncoder, ep)
                        Else
                            rendered.Save(ms, System.Drawing.Imaging.ImageFormat.Png)
                        End If

                        ms.Position = 0
                        Using xgfx As PdfSharp.Drawing.XGraphics = PdfSharp.Drawing.XGraphics.FromPdfPage(outPage)
                            Using ximg As PdfSharp.Drawing.XImage = PdfSharp.Drawing.XImage.FromStream(ms)
                                xgfx.DrawImage(ximg, 0, 0, outPage.Width.Point, outPage.Height.Point)
                            End Using
                        End Using
                    End Using
                End Using

                ' Draw overlay content on top of the rasterized page
                Using gfx As PdfSharp.Drawing.XGraphics =
                    PdfSharp.Drawing.XGraphics.FromPdfPage(outPage, PdfSharp.Drawing.XGraphicsPdfPageOptions.Append)
                    drawCallback(gfx, outPage.Width.Point, outPage.Height.Point, pageIndex)
                End Using
            Next

            outDoc.Save(outputPath)
            outDoc.Close()
        End Using
    End Sub
    Private Sub APOverlayViaRasterizeWithElements(
            inputPath As String,
            outputPath As String,
            elementsArray As JArray,
            imageCache As Dictionary(Of String, String),
            context As ToolExecutionContext,
            ByRef textCount As Integer,
            ByRef imageCount As Integer)

        EnsureApPdfSharpFontResolver()

        Dim usedWindowsPdf As Boolean = False

#If HAS_WINRT Then
        If Environment.OSVersion.Version.Major >= 10 Then
            Try
                APOverlayViaRasterizeWithElementsWindowsPdf(inputPath, outputPath, elementsArray,
                                                             imageCache, context, textCount, imageCount)
                usedWindowsPdf = True
            Catch ex As Exception
                Debug.WriteLine($"APOverlayViaRasterizeWithElements: Windows.Data.Pdf failed: {ex.Message} — falling back to PdfiumViewer")
                usedWindowsPdf = False
            End Try
        End If
#End If

        If Not usedWindowsPdf Then
            APRedactEnsurePdfiumLoaded()
            APOverlayViaRasterizeWithElementsCore(inputPath, outputPath, elementsArray,
                                                   imageCache, context, textCount, imageCount)
        End If
    End Sub

    Private Sub APOverlayViaRasterizeWithElementsWindowsPdf(
            inputPath As String,
            outputPath As String,
            elementsArray As JArray,
            imageCache As Dictionary(Of String, String),
            context As ToolExecutionContext,
            ByRef textCount As Integer,
            ByRef imageCount As Integer)
#If Not HAS_WINRT Then
        Throw New PlatformNotSupportedException("Windows.Data.Pdf is not available.")
#Else
        Dim tc As Integer = 0
        Dim ic As Integer = 0
        Dim renderException As Exception = Nothing

        Dim staThread As New System.Threading.Thread(
            Sub()
                Try
                    APOverlayViaRasterizeWithElementsWindowsPdfCore(
                        inputPath, outputPath, elementsArray, imageCache, context, tc, ic)
                Catch ex As Exception
                    renderException = ex
                End Try
            End Sub)

        staThread.SetApartmentState(System.Threading.ApartmentState.STA)
        staThread.IsBackground = True
        staThread.Start()

        If Not staThread.Join(TimeSpan.FromMinutes(5)) Then
            Try : staThread.Abort() : Catch : End Try
            Throw New TimeoutException("Windows.Data.Pdf rendering timed out after 5 minutes.")
        End If

        textCount = tc
        imageCount = ic

        If renderException IsNot Nothing Then
            Throw renderException
        End If
#End If
    End Sub

    Private Sub APOverlayViaRasterizeWithElementsWindowsPdfCore(
            inputPath As String,
            outputPath As String,
            elementsArray As JArray,
            imageCache As Dictionary(Of String, String),
            context As ToolExecutionContext,
            ByRef textCount As Integer,
            ByRef imageCount As Integer)
#If Not HAS_WINRT Then
        Throw New PlatformNotSupportedException("Windows.Data.Pdf is not available.")
#Else
        Const renderDpi As Integer = 200

        Dim storageFile As Windows.Storage.StorageFile =
            Windows.Storage.StorageFile.GetFileFromPathAsync(inputPath).GetAwaiter().GetResult()
        Dim winPdf As Windows.Data.Pdf.PdfDocument =
            Windows.Data.Pdf.PdfDocument.LoadFromFileAsync(storageFile).GetAwaiter().GetResult()

        If winPdf.PageCount = 0 Then
            Throw New InvalidOperationException("The PDF contains no pages.")
        End If

        Dim outDoc As New PdfSharp.Pdf.PdfDocument()
        Dim totalPages As Integer = CInt(winPdf.PageCount)
        Dim scaleFactor As Double = renderDpi / 96.0

        ' Rasterize all pages
        For pageIndex As UInteger = 0 To CUInt(totalPages - 1)
            Dim page As Windows.Data.Pdf.PdfPage = winPdf.GetPage(pageIndex)

            Dim pageWidthPt As Double = page.Size.Width * 72.0 / 96.0
            Dim pageHeightPt As Double = page.Size.Height * 72.0 / 96.0
            Dim renderWidthPx As UInteger = CUInt(Math.Round(page.Size.Width * scaleFactor))
            Dim renderHeightPx As UInteger = CUInt(Math.Round(page.Size.Height * scaleFactor))

            Dim outPage As PdfSharp.Pdf.PdfPage = outDoc.AddPage()
            outPage.Width = PdfSharp.Drawing.XUnit.FromPoint(pageWidthPt)
            outPage.Height = PdfSharp.Drawing.XUnit.FromPoint(pageHeightPt)

            Using renderStream As New Windows.Storage.Streams.InMemoryRandomAccessStream()
                Dim renderOptions As New Windows.Data.Pdf.PdfPageRenderOptions()
                renderOptions.DestinationWidth = renderWidthPx
                renderOptions.DestinationHeight = renderHeightPx
                renderOptions.BitmapEncoderId = Windows.Graphics.Imaging.BitmapEncoder.PngEncoderId

                page.RenderToStreamAsync(renderStream, renderOptions).GetAwaiter().GetResult()
                renderStream.Seek(0)

                Dim netStream As System.IO.Stream =
                    System.IO.WindowsRuntimeStreamExtensions.AsStreamForRead(renderStream)

                Using pngMs As New MemoryStream()
                    netStream.CopyTo(pngMs)
                    pngMs.Position = 0

                    Using bmp As New System.Drawing.Bitmap(pngMs)
                        Using jpegMs As New MemoryStream()
                            Dim jpegEncoder As System.Drawing.Imaging.ImageCodecInfo = Nothing
                            For Each codec In System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders()
                                If codec.MimeType = "image/jpeg" Then jpegEncoder = codec : Exit For
                            Next
                            If jpegEncoder IsNot Nothing Then
                                Dim ep As New System.Drawing.Imaging.EncoderParameters(1)
                                ep.Param(0) = New System.Drawing.Imaging.EncoderParameter(
                                    System.Drawing.Imaging.Encoder.Quality, 85L)
                                bmp.Save(jpegMs, jpegEncoder, ep)
                            Else
                                bmp.Save(jpegMs, System.Drawing.Imaging.ImageFormat.Png)
                            End If
                            jpegMs.Position = 0

                            Using xgfx As PdfSharp.Drawing.XGraphics =
                                PdfSharp.Drawing.XGraphics.FromPdfPage(outPage)
                                Using ximg As PdfSharp.Drawing.XImage =
                                    PdfSharp.Drawing.XImage.FromStream(jpegMs)
                                    xgfx.DrawImage(ximg, 0, 0, outPage.Width.Point, outPage.Height.Point)
                                End Using
                            End Using
                        End Using
                    End Using
                End Using
            End Using

            page.Dispose()
        Next

        ' Draw overlay elements on the rasterized pages
        APOverlayDrawElements(outDoc, totalPages, elementsArray, imageCache, context,
                              textCount, imageCount)

        outDoc.Save(outputPath)
        outDoc.Close()
#End If
    End Sub

    ''' <summary>
    ''' Core rasterize + overlay implementation, separated to ensure pdfium.dll
    ''' is loaded before PdfiumViewer types are JIT-resolved.
    ''' </summary>
    <Runtime.CompilerServices.MethodImpl(Runtime.CompilerServices.MethodImplOptions.NoInlining)>
    Private Sub APOverlayViaRasterizeWithElementsCore(
            inputPath As String,
            outputPath As String,
            elementsArray As JArray,
            imageCache As Dictionary(Of String, String),
            context As ToolExecutionContext,
            ByRef textCount As Integer,
            ByRef imageCount As Integer)

        Const renderDpi As Integer = 200

        Using pdf As PdfiumViewer.PdfDocument = PdfiumViewer.PdfDocument.Load(inputPath)
            If pdf.PageCount = 0 Then
                Throw New InvalidOperationException("The PDF contains no pages.")
            End If

            Dim outDoc As New PdfSharp.Pdf.PdfDocument()
            Dim totalPages = pdf.PageCount

            Dim renderFlags As PdfiumViewer.PdfRenderFlags =
                PdfiumViewer.PdfRenderFlags.Annotations Or
                PdfiumViewer.PdfRenderFlags.LcdText Or
                PdfiumViewer.PdfRenderFlags.ForPrinting

            ' Rasterize all pages first
            For pageIndex As Integer = 0 To totalPages - 1
                Dim sizePt As System.Drawing.SizeF = pdf.PageSizes(pageIndex)
                Dim widthPx As Integer = CInt(Math.Round(sizePt.Width / 72.0 * renderDpi))
                Dim heightPx As Integer = CInt(Math.Round(sizePt.Height / 72.0 * renderDpi))

                ' Declare outPage OUTSIDE the Using rendered block so it stays in scope
                Dim outPage As PdfSharp.Pdf.PdfPage = outDoc.AddPage()
                outPage.Width = PdfSharp.Drawing.XUnit.FromPoint(sizePt.Width)
                outPage.Height = PdfSharp.Drawing.XUnit.FromPoint(sizePt.Height)

                Using rendered As System.Drawing.Image =
                    pdf.Render(pageIndex, widthPx, heightPx, renderDpi, renderDpi, renderFlags)

                    Using ms As New MemoryStream()
                        Dim jpegEncoder As System.Drawing.Imaging.ImageCodecInfo = Nothing
                        For Each codec In System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders()
                            If codec.MimeType = "image/jpeg" Then jpegEncoder = codec : Exit For
                        Next
                        If jpegEncoder IsNot Nothing Then
                            Dim ep As New System.Drawing.Imaging.EncoderParameters(1)
                            ep.Param(0) = New System.Drawing.Imaging.EncoderParameter(
                                System.Drawing.Imaging.Encoder.Quality, 85L)
                            rendered.Save(ms, jpegEncoder, ep)
                        Else
                            rendered.Save(ms, System.Drawing.Imaging.ImageFormat.Png)
                        End If

                        ms.Position = 0
                        Using xgfx As PdfSharp.Drawing.XGraphics = PdfSharp.Drawing.XGraphics.FromPdfPage(outPage)
                            Using ximg As PdfSharp.Drawing.XImage = PdfSharp.Drawing.XImage.FromStream(ms)
                                xgfx.DrawImage(ximg, 0, 0, outPage.Width.Point, outPage.Height.Point)
                            End Using
                        End Using
                    End Using
                End Using
            Next

            ' Now draw overlay elements on the rasterized pages
            APOverlayDrawElements(outDoc, totalPages, elementsArray, imageCache, context,
                                  textCount, imageCount)

            outDoc.Save(outputPath)
            outDoc.Close()
        End Using
    End Sub


End Class
