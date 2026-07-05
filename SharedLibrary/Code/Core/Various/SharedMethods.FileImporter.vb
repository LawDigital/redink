' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SharedMethods.FileImporter.vb
' Purpose: Provides helper functions to read text from common document formats
'          (plain text, RTF, Word documents, and PDF), returning either extracted
'          text or an error string (depending on caller preference).
'
' Architecture:
'  - Text files: Normalizes the input path, validates existence, then reads UTF-8
'    (with BOM detection) via `StreamReader`.
'  - RTF: Loads the file contents and uses a hidden `RichTextBox` to convert RTF
'    markup to plain text.
'  - Word: Uses Office interop; attempts to attach to an existing Word instance,
'    otherwise creates an invisible instance; opens the document read-only and
'    returns `doc.Content.Text`.
'  - PDF: Uses UglyToad.PdfPig to iterate pages and extract text with multiple
'    fallback strategies; optionally runs OCR via an LLM call when heuristics
'    indicate that the PDF likely contains scanned images / poor text layer.
'  - Binary/media files (images, audio, video, etc.): Sent as binary objects
'    directly to the LLM when the configured model supports the file's MIME type.
'
' External Dependencies:
'  - Microsoft.Office.Interop.Word (Word automation / COM interop)
'  - System.Windows.Forms.RichTextBox (RTF-to-text conversion)
'  - UglyToad.PdfPig (PDF parsing and text extraction)
'  - SharedLibrary.SharedContext.ISharedContext (OCR model/config access)
'  - Internal helpers used here: `ShowCustomYesNoBox`, `ShowCustomMessageBox`,
'    `LLM`, `GetSpecialTaskModel`, `RestoreDefaults`, and related configuration
'    fields (`originalConfigLoaded`, `originalConfig`).
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Word
Imports PdfSharp
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary
    Partial Public Class SharedMethods

        ''' <summary>
        ''' Result from reading a file, including content and metadata about potential incompleteness.
        ''' </summary>
        Public Class FileReadResult
            ''' <summary>The extracted text content from the file.</summary>
            Public Property Content As String = ""

            ''' <summary>True if the file was a PDF and heuristics suggested it may contain images but OCR was not performed.</summary>
            Public Property PdfMayBeIncomplete As Boolean = False

            ''' <summary>True if the user canceled an interactive file-reading choice (for example worksheet selection).</summary>
            Public Property UserCancelled As Boolean = False

            Public Sub New()
            End Sub

            Public Sub New(content As String, pdfMayBeIncomplete As Boolean, Optional userCancelled As Boolean = False)
                Me.Content = content
                Me.PdfMayBeIncomplete = pdfMayBeIncomplete
                Me.UserCancelled = userCancelled
            End Sub
        End Class

        ''' <summary>
        ''' Result from reading a PDF file, including content and metadata about OCR status.
        ''' </summary>
        Public Class PdfReadResult
            ''' <summary>The extracted text content from the PDF.</summary>
            Public Property Content As String = ""

            ''' <summary>True if heuristics suggested OCR but it was not performed (OCR unavailable or user declined).</summary>
            Public Property OcrWasSkippedDueToHeuristics As Boolean = False

            Public Sub New()
            End Sub

            Public Sub New(content As String, ocrSkipped As Boolean)
                Me.Content = content
                Me.OcrWasSkippedDueToHeuristics = ocrSkipped
            End Sub
        End Class

        ' ── Binary / media file extensions that require LLM-based extraction ──

        ''' <summary>
        ''' Image file extensions that can be processed by a vision-capable LLM.
        ''' </summary>
        Private Shared ReadOnly ImageExtensions As String() = {
            ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff", ".tif", ".webp", ".svg"
        }

        ''' <summary>
        ''' Audio file extensions that can be processed by an audio-capable LLM.
        ''' </summary>
        Private Shared ReadOnly AudioExtensions As String() = {
            ".mp3", ".wav", ".ogg", ".flac", ".m4a", ".aac", ".wma", ".opus", ".webm"
        }

        ''' <summary>
        ''' Video file extensions that can be processed by a video-capable LLM.
        ''' </summary>
        Private Shared ReadOnly VideoExtensions As String() = {
            ".mp4", ".avi", ".mkv", ".mov", ".wmv"
        }

        ''' <summary>
        ''' Returns True if the extension identifies a binary/media file that cannot be read as text
        ''' and must instead be sent to the LLM as a binary object.
        ''' </summary>
        ''' <param name="extension">File extension including the leading dot (e.g. ".png").</param>
        ''' <returns>True when the file is a binary/media type.</returns>
        Public Shared Function IsBinaryMediaExtension(extension As String) As Boolean
            If String.IsNullOrWhiteSpace(extension) Then Return False
            Dim ext = extension.ToLowerInvariant()
            Return ImageExtensions.Contains(ext) OrElse
                   AudioExtensions.Contains(ext) OrElse
                   VideoExtensions.Contains(ext)
        End Function

        ''' <summary>
        ''' Checks whether the APICall_Object configuration supports a specific MIME type prefix
        ''' (e.g. "image/", "audio/", "video/") or a wildcard ("*/*").
        ''' </summary>
        ''' <param name="apiCallObject">The INI_APICall_Object or INI_APICall_Object_2 string.</param>
        ''' <param name="mimePrefix">MIME prefix to look for, e.g. "image/", "audio/", "video/".</param>
        ''' <returns>True if the configuration accepts at least one matching MIME type.</returns>
        Public Shared Function IsApiCallObjectMimeCapable(apiCallObject As String, mimePrefix As String) As Boolean
            If String.IsNullOrWhiteSpace(apiCallObject) Then Return False

            Dim segments As String() = apiCallObject.Split(New Char() {"¦"c}, StringSplitOptions.RemoveEmptyEntries)
            Dim hasUnfilteredSegment As Boolean = False
            Dim hasMimeFilter As Boolean = False
            Dim allSegmentsHaveFilters As Boolean = True

            For Each segment As String In segments
                Dim trimmedSegment As String = segment.Trim()

                If trimmedSegment.StartsWith("[") Then
                    Dim closeBracketIdx As Integer = trimmedSegment.IndexOf("]"c)
                    If closeBracketIdx > 1 Then
                        Dim filterContent As String = trimmedSegment.Substring(1, closeBracketIdx - 1)
                        If filterContent.IndexOf(mimePrefix, StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                           filterContent.IndexOf("*/*", StringComparison.OrdinalIgnoreCase) >= 0 Then
                            hasMimeFilter = True
                        End If
                    End If
                Else
                    hasUnfilteredSegment = True
                    allSegmentsHaveFilters = False
                End If
            Next

            If hasUnfilteredSegment Then Return True
            If hasMimeFilter Then Return True
            If allSegmentsHaveFilters Then Return False
            Return True
        End Function

        ''' <summary>
        ''' Determines whether the configured model can accept a binary file of the given extension
        ''' by checking MIME type support in the APICall_Object configuration and alternate model paths.
        ''' </summary>
        ''' <param name="context">Shared context containing model and API configuration.</param>
        ''' <param name="extension">File extension including the leading dot (e.g. ".png").</param>
        ''' <param name="taskFlag">
        ''' Optional alternate-model task flag (e.g. "ImageExtraction", "AudioTranscription").
        ''' When supplied, the alternate-model INI is checked first.
        ''' </param>
        ''' <returns>True if a model capable of handling this file type is available.</returns>
        Public Shared Function IsBinaryMediaSupported(context As ISharedContext,
                                                      extension As String,
                                                      Optional taskFlag As String = Nothing) As Boolean
            If context Is Nothing Then Return False

            Dim mimePrefix As String = MimePrefixForExtension(extension)
            If String.IsNullOrWhiteSpace(mimePrefix) Then Return False

            If Not String.IsNullOrWhiteSpace(taskFlag) AndAlso
               Not String.IsNullOrWhiteSpace(context.INI_AlternateModelPath) Then

                Dim scope = CaptureModelConfigScope(context)

                Try
                    If GetSpecialTaskModel(context, context.INI_AlternateModelPath, taskFlag) Then
                        Return True
                    End If
                Catch
                Finally
                    RestoreModelConfigScope(context, scope)
                End Try
            End If

            Return IsApiCallObjectMimeCapable(context.INI_APICall_Object, mimePrefix)
        End Function

        ''' <summary>
        ''' Sends a binary/media file to the LLM as a file object and returns the LLM's textual response.
        ''' </summary>
        ''' <param name="filePath">Path to the binary file.</param>
        ''' <param name="context">Shared context containing model and API configuration.</param>
        ''' <param name="systemPrompt">
        ''' System prompt to use. When empty, falls back to <c>context.SP_InsertClipboard</c>.
        ''' </param>
        ''' <param name="askUser">If False, suppresses all UI dialogs.</param>
        ''' <param name="taskFlag">
        ''' Optional alternate-model task flag (e.g. "ImageExtraction", "AudioTranscription").
        ''' </param>
        ''' <returns>Text returned by the LLM, or an empty string on failure.</returns>
        Public Shared Async Function ReadBinaryFileViaLLM(filePath As String,
                                                          context As ISharedContext,
                                                          Optional systemPrompt As String = "",
                                                          Optional askUser As Boolean = True,
                                                          Optional taskFlag As String = Nothing) As Task(Of String)
            Dim scope = CaptureModelConfigScope(context)

            Try
                If String.IsNullOrWhiteSpace(filePath) OrElse Not IO.File.Exists(filePath) Then
                    Return ""
                End If

                Dim ext As String = IO.Path.GetExtension(filePath).ToLowerInvariant()
                If Not IsBinaryMediaSupported(context, ext, taskFlag) Then
                    If askUser Then
                        ShowCustomMessageBox($"The file type '{ext}' is not supported by your current model configuration.")
                    End If
                    Return ""
                End If

                Dim useSecondAPI As Boolean = False
                Dim timeOut = context.INI_Timeout

                If Not String.IsNullOrWhiteSpace(taskFlag) AndAlso
                   Not String.IsNullOrWhiteSpace(context.INI_AlternateModelPath) AndAlso
                   GetSpecialTaskModel(context, context.INI_AlternateModelPath, taskFlag) Then

                    useSecondAPI = True
                    timeOut = context.INI_Timeout_2
                End If

                Dim sysPrompt As String =
                    If(String.IsNullOrWhiteSpace(systemPrompt), context.SP_InsertClipboard, systemPrompt)

                Dim result As String =
                    Await LLM(context, sysPrompt, "", "", "", timeOut * 2, useSecondAPI, Not askUser, "", filePath)

                Return If(result, "")

            Catch ex As System.Exception
                Debug.WriteLine("ReadBinaryFileViaLLM failed: " & ex.Message)
                Return ""
            Finally
                RestoreModelConfigScope(context, scope)
            End Try
        End Function

        ''' <summary>
        ''' Returns the MIME prefix for a file extension (e.g. ".png" → "image/", ".mp3" → "audio/").
        ''' Returns an empty string for unknown extensions.
        ''' </summary>
        Private Shared Function MimePrefixForExtension(extension As String) As String
            If String.IsNullOrWhiteSpace(extension) Then Return ""
            Dim ext = extension.ToLowerInvariant()
            If ImageExtensions.Contains(ext) Then Return "image/"
            If AudioExtensions.Contains(ext) Then Return "audio/"
            If VideoExtensions.Contains(ext) Then Return "video/"
            Return ""
        End Function

        ''' <summary>
        ''' Returns the alternate-model task flag appropriate for a file extension.
        ''' </summary>
        Public Shared Function TaskFlagForExtension(extension As String) As String
            If String.IsNullOrWhiteSpace(extension) Then Return Nothing
            Dim ext = extension.ToLowerInvariant()
            If ImageExtensions.Contains(ext) Then Return "ImageExtraction"
            If AudioExtensions.Contains(ext) Then Return "AudioTranscription"
            If VideoExtensions.Contains(ext) Then Return "VideoExtraction"
            Return Nothing
        End Function

        ''' <summary>
        ''' Determines whether a file with the given extension can be processed for content extraction.
        ''' Text-based formats (.pdf, .docx, .txt, etc.) are always supported.
        ''' Binary/media formats (images, audio, video) require model capability and return <c>False</c>
        ''' when the configured model cannot handle the corresponding MIME type.
        ''' </summary>
        ''' <param name="context">Shared context containing model and API configuration.</param>
        ''' <param name="extension">File extension including the leading dot (e.g. ".png").</param>
        ''' <returns><c>True</c> if the file can be processed; <c>False</c> if it requires unsupported model capabilities.</returns>
        Public Shared Function IsModelCapableForExtension(context As ISharedContext, extension As String) As Boolean
            If String.IsNullOrWhiteSpace(extension) Then Return False
            Dim ext = extension.ToLowerInvariant()

            ' Text-based formats are always supported (no model capability needed)
            If Not IsBinaryMediaExtension(ext) Then Return True

            ' Binary/media formats require model capability
            Dim taskFlag = TaskFlagForExtension(ext)
            Return IsBinaryMediaSupported(context, ext, taskFlag)
        End Function

        ''' <summary>
        ''' Reads a text file as UTF-8 (with BOM detection) and returns its contents.
        ''' </summary>
        ''' <param name="filePath">Path to the file to read.</param>
        ''' <param name="ReturnErrorInsteadOfEmpty">
        ''' If <c>True</c>, returns an error message string on failure; otherwise returns an empty string.
        ''' </param>
        ''' <returns>The file contents, or an error string / empty string depending on <paramref name="ReturnErrorInsteadOfEmpty"/>.</returns>
        Public Shared Function ReadTextFile(filePath As String, Optional ReturnErrorInsteadOfEmpty As Boolean = True) As String
            Try
                ' Normalize and check the path
                filePath = Path.GetFullPath(filePath)
                If Not File.Exists(filePath) Then
                    Return If(ReturnErrorInsteadOfEmpty, "Error: File not found.", "")
                End If

                ' Use StreamReader for reading
                Using reader As New StreamReader(filePath, System.Text.Encoding.UTF8, True)
                    Dim content As String = reader.ReadToEnd()
                    Return content
                End Using
            Catch ex As System.Exception
                Return If(ReturnErrorInsteadOfEmpty, $"Error reading file: {ex.Message}", "")
            End Try
        End Function

        ''' <summary>
        ''' Reads an RTF file and returns its plain-text representation.
        ''' </summary>
        ''' <param name="rtfPath">Path to the RTF file to read.</param>
        ''' <param name="ReturnErrorInsteadOfEmpty">
        ''' If <c>True</c>, returns an error message string on failure; otherwise returns an empty string.
        ''' </param>
        ''' <returns>The extracted plain text, or an error string / empty string depending on <paramref name="ReturnErrorInsteadOfEmpty"/>.</returns>
        Public Shared Function ReadRtfAsText(ByVal rtfPath As String, Optional ReturnErrorInsteadOfEmpty As Boolean = True) As String
            Try
                Dim rtfContent As String = File.ReadAllText(rtfPath)
                Using rtb As New RichTextBox()
                    rtb.Visible = False
                    rtb.Rtf = rtfContent
                    Return rtb.Text
                End Using
            Catch ex As System.Exception
                Return If(ReturnErrorInsteadOfEmpty, $"Error reading RTF: {ex.Message}", "")
            End Try
        End Function

        ''' <summary>
        ''' Reads a Word document via Office interop and returns the document's text content.
        ''' </summary>
        ''' <param name="docPath">Path to the Word document to open.</param>
        ''' <param name="ReturnErrorInsteadOfEmpty">
        ''' If <c>True</c>, returns an error message string on failure; otherwise returns an empty string.
        ''' </param>
        ''' <returns>The extracted text, or an error string / empty string depending on <paramref name="ReturnErrorInsteadOfEmpty"/>.</returns>
        Public Shared Function ReadWordDocument(ByVal docPath As String, Optional ReturnErrorInsteadOfEmpty As Boolean = True) As String
            Dim app As Microsoft.Office.Interop.Word.Application = Nothing
            Dim doc As Document = Nothing
            Dim createdNewInstance As Boolean = False

            Try
                Try
                    ' Try to attach to an existing Word instance.
                    app = CType(Marshal.GetActiveObject("Word.Application"), Microsoft.Office.Interop.Word.Application)
                Catch ex As System.Exception
                    ' If Word is not running, create a new Word application.
                    app = New Microsoft.Office.Interop.Word.Application With {.Visible = False}
                    createdNewInstance = True
                End Try

                ' Open the Word document in read-only mode                
                Dim fileName As Object = docPath
                doc = app.Documents.Open(fileName, [ReadOnly]:=True, Visible:=False)

                ' Extract the content text
                Dim text As String = doc.Content.Text

                ' Close the document without saving changes
                doc.Close(SaveChanges:=False)

                ' Return the extracted text
                Return text

            Catch ex As System.Exception
                ' Ensure the document is closed in case of an error
                If doc IsNot Nothing Then
                    doc.Close(SaveChanges:=False)
                End If

                ' Return the error message (or empty string if ReturnErrorInsteadOfEmpty=False)
                Return If(ReturnErrorInsteadOfEmpty, $"Error reading Word document: {ex.Message}", "")

            Finally
                ' Only quit the application if it was newly created by this method
                If app IsNot Nothing AndAlso createdNewInstance Then
                    app.Quit()
                End If
            End Try
        End Function

        ''' <summary>
        ''' Reads a PDF using PdfPig and returns extracted text; optionally performs OCR via an LLM call
        ''' when heuristics indicate the PDF contains little or low-quality extractable text.
        ''' </summary>
        ''' <param name="pdfPath">Path to the PDF file to read.</param>
        ''' <param name="ReturnErrorInsteadOfEmpty">
        ''' If <c>True</c>, returns an error message string on failure; otherwise returns an empty string.
        ''' </param>
        ''' <param name="DoOCR">If <c>True</c>, enables OCR heuristics and (if confirmed) OCR execution.</param>
        ''' <param name="AskUser">If <c>True</c>, prompts the user before performing OCR.</param>
        ''' <param name="context">Shared context used for OCR-capable model configuration and LLM invocation.</param>
        ''' <param name="ocrAdditionalInstruction">Additional instructions for OCR processing when reading PDF files.</param>
        ''' <returns>A PdfReadResult containing the extracted text and whether OCR was skipped despite being suggested.</returns>
        Public Shared Async Function ReadPdfAsTextEx(ByVal pdfPath As String,
                                            Optional ByVal ReturnErrorInsteadOfEmpty As Boolean = True,
                                            Optional ByVal DoOCR As Boolean = False,
                                            Optional ByVal AskUser As Boolean = True,
                                            Optional ByVal context As ISharedContext = Nothing,
                                            Optional ByVal ocrAdditionalInstruction As String = Nothing,
                                            Optional ByVal ShowOcrProgressWindow As Boolean = False) As Task(Of PdfReadResult)

            Dim result As New PdfReadResult()

            Try
                If String.IsNullOrWhiteSpace(pdfPath) OrElse Not IO.File.Exists(pdfPath) Then
                    result.Content = If(ReturnErrorInsteadOfEmpty, "Error: File not found or path is empty.", "")
                    Return result
                End If

                Dim sb As New System.Text.StringBuilder()
                Dim pageCount As Integer = 0
                Dim totalChars As Integer = 0
                Dim hasLowQualityText As Boolean = False
                Dim reasons As New List(Of String)()
                Dim sparsePageCount As Integer = 0
                Dim perPageChars As New List(Of Integer)()
                Dim pagesWithImagesButNoText As Integer = 0
                Dim pagesWithGarbledText As Integer = 0

                Using document As UglyToad.PdfPig.PdfDocument = UglyToad.PdfPig.PdfDocument.Open(pdfPath)
                    pageCount = document.NumberOfPages

                    For Each page As UglyToad.PdfPig.Content.Page In document.GetPages()
                        Dim pageText As String = page.Text
                        sb.AppendLine(pageText)
                        Dim pageCharCount As Integer = If(pageText IsNot Nothing, pageText.Length, 0)
                        totalChars += pageCharCount
                        perPageChars.Add(pageCharCount)

                        ' Track pages with very little text (likely scanned/image pages)
                        If pageCharCount < 50 Then
                            sparsePageCount += 1
                        End If

                        ' Check for pages that have images but little/no text (scanned documents)
                        Try
                            Dim images = page.GetImages()
                            If images IsNot Nothing AndAlso images.Count > 0 AndAlso pageCharCount < 100 Then
                                pagesWithImagesButNoText += 1
                            End If
                        Catch
                            ' Some PDFs may fail image enumeration; ignore
                        End Try

                        ' Check for low-quality text indicators
                        If pageText IsNot Nothing Then
                            Dim words = pageText.Split({" "c, vbCr(0), vbLf(0)}, StringSplitOptions.RemoveEmptyEntries)
                            Dim avgWordLen = If(words.Length > 0, words.Average(Function(w) w.Length), 0)
                            If avgWordLen < 2 AndAlso words.Length > 10 Then
                                hasLowQualityText = True
                            End If

                            ' Check for garbled/non-printable characters (broken font encoding)
                            If pageCharCount > 20 Then
                                Dim nonPrintableCount As Integer = pageText.Count(Function(c) Char.IsControl(c) AndAlso c <> vbLf(0) AndAlso c <> vbCr(0) AndAlso c <> vbTab(0))
                                Dim replacementCount As Integer = pageText.Count(Function(c) c = ChrW(&HFFFD) OrElse c = "?"c)
                                Dim suspiciousRatio As Double = (nonPrintableCount + replacementCount) / pageCharCount
                                If suspiciousRatio > 0.15 Then
                                    pagesWithGarbledText += 1
                                End If
                            End If
                        End If
                    Next
                End Using

                Dim extractedText As String = sb.ToString().Trim()

                ' Heuristics to determine if OCR might be needed
                Dim shouldSuggestOcr As Boolean = False
                Dim avgCharsPerPage As Double = If(pageCount > 0, totalChars / pageCount, 0)

                If pageCount > 0 AndAlso avgCharsPerPage < 100 Then
                    shouldSuggestOcr = True
                    reasons.Add($"Very little text extracted ({avgCharsPerPage:F0} chars/page average)")
                End If

                If hasLowQualityText Then
                    shouldSuggestOcr = True
                    reasons.Add("Text appears to be low quality (possibly garbled OCR or image-based)")
                End If

                If String.IsNullOrWhiteSpace(extractedText) AndAlso pageCount > 0 Then
                    shouldSuggestOcr = True
                    reasons.Add("No text could be extracted from any page")
                End If

                ' Check if a significant portion of pages are sparse (mixed document scenario)
                If pageCount >= 2 AndAlso sparsePageCount > 0 Then
                    Dim sparseRatio As Double = sparsePageCount / pageCount
                    If sparseRatio >= 0.1 Then
                        shouldSuggestOcr = True
                        reasons.Add($"{sparsePageCount} of {pageCount} pages contain very little or no text (likely scanned images)")
                    End If
                End If

                ' Check for pages with images but no meaningful text (scanned pages)
                If pagesWithImagesButNoText > 0 Then
                    shouldSuggestOcr = True
                    reasons.Add($"{pagesWithImagesButNoText} of {pageCount} pages contain images but little or no extractable text")
                End If

                ' Check for garbled text (broken font encoding / CID mapping issues)
                If pagesWithGarbledText > 0 Then
                    shouldSuggestOcr = True
                    reasons.Add($"{pagesWithGarbledText} of {pageCount} pages contain garbled or non-printable characters (likely encoding issues)")
                End If

                ' Check for extreme variance between pages (some rich, some empty)
                If pageCount >= 3 AndAlso perPageChars.Count >= 3 Then
                    Dim maxChars As Integer = perPageChars.Max()
                    Dim minChars As Integer = perPageChars.Min()
                    If maxChars > 500 AndAlso minChars < 50 Then
                        Dim pagesAbove500 As Integer = perPageChars.Where(Function(c) c > 500).Count()
                        Dim pagesBelow50 As Integer = perPageChars.Where(Function(c) c < 50).Count()
                        If pagesAbove500 >= 1 AndAlso pagesBelow50 >= 1 AndAlso Not shouldSuggestOcr Then
                            shouldSuggestOcr = True
                            reasons.Add($"Large variation in text content across pages ({pagesBelow50} pages nearly empty, {pagesAbove500} pages with substantial text)")
                        End If
                    End If
                End If

                ' Disable OCR if no OCR-capable call is configured or context missing
                Dim ocrUnavailable As Boolean = False
                If DoOCR AndAlso (context Is Nothing OrElse Not IsOcrAvailable(context)) Then
                    DoOCR = False
                    ocrUnavailable = True
                End If

                ' If DoOCR is disabled → just return whatever text we found (or empty string)
                If Not DoOCR Then
                    ' If we would have suggested OCR but it's not available, flag and warn the user
                    If shouldSuggestOcr Then
                        result.OcrWasSkippedDueToHeuristics = True

                        If AskUser AndAlso ocrUnavailable Then
                            Dim formattedReasons As String = String.Join(Environment.NewLine, reasons.ConvertAll(Function(r) "- " & r))
                            ShowCustomMessageBox(
                                "The PDF appears to contain pages that may need OCR:" & Environment.NewLine & Environment.NewLine &
                                formattedReasons & Environment.NewLine & Environment.NewLine &
                                "OCR is not available with your current model configuration." & Environment.NewLine &
                                "The extracted text may be incomplete.")
                        End If
                    End If
                    result.Content = extractedText
                    Return result
                End If

                If shouldSuggestOcr Then
                    ' Check if OCR is actually available
                    If Not IsOcrAvailable(context) Then
                        ' OCR would be suggested but is not available - warn user if allowed
                        Debug.WriteLine("OCR suggested by heuristics but not available - skipping OCR prompt.")
                        result.OcrWasSkippedDueToHeuristics = True

                        If AskUser Then
                            Dim formattedReasons As String = String.Join(Environment.NewLine, reasons.ConvertAll(Function(r) "- " & r))
                            ShowCustomMessageBox(
                                "The PDF appears to contain pages that may need OCR:" & Environment.NewLine & Environment.NewLine &
                                formattedReasons & Environment.NewLine & Environment.NewLine &
                                "OCR is not available with your current model configuration." & Environment.NewLine &
                                "The extracted text may be incomplete.")
                        End If

                        result.Content = extractedText
                        Return result
                    End If

                    If AskUser Then
                        Dim formattedReasons As String = String.Join(Environment.NewLine, reasons.ConvertAll(Function(r) "- " & r))
                        Dim msg As String = $"The PDF appears to contain little or no extractable text:" & Environment.NewLine & Environment.NewLine &
                                            formattedReasons & Environment.NewLine & Environment.NewLine &
                                            "It's likely that the document consists mainly of scanned images." & Environment.NewLine & Environment.NewLine &
                                            "Would you like AI to perform OCR to extract text (if supported by your configured model)?"
                        Dim userChoice As Integer = ShowCustomYesNoBox(msg, "Yes, try OCR", "No, use what you have")
                        If userChoice <> 1 Then
                            result.OcrWasSkippedDueToHeuristics = True
                            result.Content = extractedText
                            Return result
                        End If
                    End If

                    Dim ocrText As String =
                        Await PerformOCR(pdfPath, context, AskUser, ocrAdditionalInstruction, ShowOcrProgressWindow)
                    If Not String.IsNullOrWhiteSpace(ocrText) Then
                        result.Content = ocrText
                        Return result
                    Else
                        ' OCR was attempted but returned empty - content may be incomplete
                        result.OcrWasSkippedDueToHeuristics = True
                    End If
                End If

                result.Content = extractedText
                Return result

            Catch ex As System.Exception
                result.Content = If(ReturnErrorInsteadOfEmpty, $"Error reading PDF: {ex.Message}", "")
                Return result
            End Try
        End Function

        ''' <summary>
        ''' Reads a PDF using PdfPig and returns extracted text (backward compatible wrapper).
        ''' </summary>
        Public Shared Async Function ReadPdfAsText(ByVal pdfPath As String,
                                            Optional ByVal ReturnErrorInsteadOfEmpty As Boolean = True,
                                            Optional ByVal DoOCR As Boolean = False,
                                            Optional ByVal AskUser As Boolean = True,
                                            Optional ByVal context As ISharedContext = Nothing,
                                            Optional ByVal ocrAdditionalInstruction As String = Nothing,
                                            Optional ByVal ShowOcrProgressWindow As Boolean = False) As Task(Of String)
            Dim result = Await ReadPdfAsTextEx(pdfPath,
                                               ReturnErrorInsteadOfEmpty,
                                               DoOCR,
                                               AskUser,
                                               context,
                                               ocrAdditionalInstruction,
                                               ShowOcrProgressWindow)
            Return result.Content
        End Function

        ''' <summary>
        ''' Extracts plain text content from a single PDF page using multiple strategies:
        ''' content-order extraction, word/line reconstruction, and finally a letter-gap heuristic.
        ''' </summary>
        ''' <param name="page">The PDF page to extract text from.</param>
        ''' <returns>Extracted page text (may be empty).</returns>
        Private Shared Function ExtractPageTextFromPdf(page As UglyToad.PdfPig.Content.Page) As String
            ' 1) Try PdfPig's content-order extractor (good spacing/reading order on many PDFs)
            Try
                Dim t As String = UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor.ContentOrderTextExtractor.GetText(page)
                If Not String.IsNullOrWhiteSpace(t) AndAlso (t.Contains(" ") OrElse t.Contains(vbTab) OrElse t.Contains(vbCr) OrElse t.Contains(vbLf)) Then
                    Return t
                End If
            Catch
                ' Older PdfPig versions or certain pages may not support this path; ignore and fallback.
            End Try

            ' 2) Word-based reconstruction using Nearest-Neighbour extractor (higher recall on tricky PDFs)
            Try
                Dim words As System.Collections.Generic.IEnumerable(Of UglyToad.PdfPig.Content.Word) =
            page.GetWords(UglyToad.PdfPig.DocumentLayoutAnalysis.WordExtractor.NearestNeighbourWordExtractor.Instance)

                If words IsNot Nothing AndAlso words.Count > 0 Then
                    ' Group words into lines by baseline with a tolerant threshold
                    Dim baselineTol As Double = Math.Max(0.5, page.Height * 0.002) ' ~0.2% of page height
                    Dim lines As New System.Collections.Generic.List(Of System.Collections.Generic.List(Of UglyToad.PdfPig.Content.Word))()

                    For Each w In words.OrderByDescending(Function(x) x.BoundingBox.Bottom).ThenBy(Function(x) x.BoundingBox.Left)
                        Dim placed As Boolean = False
                        For Each ln In lines
                            Dim ref = ln(0)
                            If Math.Abs(w.BoundingBox.Bottom - ref.BoundingBox.Bottom) <= baselineTol Then
                                ln.Add(w)
                                placed = True
                                Exit For
                            End If
                        Next
                        If Not placed Then
                            lines.Add(New System.Collections.Generic.List(Of UglyToad.PdfPig.Content.Word) From {w})
                        End If
                    Next

                    Dim sbLine As New System.Text.StringBuilder()
                    Dim first As Boolean = True
                    For Each ln In lines.OrderByDescending(Function(l) l.Average(Function(w) w.BoundingBox.Bottom))
                        If Not first Then sbLine.AppendLine()
                        first = False
                        Dim lineText = String.Join(" ", ln.OrderBy(Function(w) w.BoundingBox.Left).Select(Function(w) w.Text))
                        sbLine.Append(lineText)
                    Next

                    Dim s = sbLine.ToString()
                    If Not String.IsNullOrWhiteSpace(s) Then
                        Return s
                    End If
                End If
            Catch
                ' Ignore and fallback
            End Try

            ' 3) Letter-gap heuristic: insert spaces based on horizontal gaps; break lines on baseline changes
            Dim letters = page.Letters
            If letters Is Nothing OrElse letters.Count = 0 Then Return String.Empty

            Dim ordered = letters.OrderByDescending(Function(l) l.GlyphRectangle.Bottom).ThenBy(Function(l) l.GlyphRectangle.Left)
            Dim sb As New System.Text.StringBuilder()
            Dim prev As UglyToad.PdfPig.Content.Letter = Nothing

            For Each l In ordered
                If prev IsNot Nothing Then
                    Dim sameLine = Math.Abs(l.GlyphRectangle.Bottom - prev.GlyphRectangle.Bottom) <= Math.Max(0.5, prev.GlyphRectangle.Height * 0.6)
                    If Not sameLine Then
                        sb.AppendLine()
                    Else
                        Dim gap = l.GlyphRectangle.Left - prev.GlyphRectangle.Right
                        Dim spaceThreshold = Math.Max(prev.GlyphRectangle.Width * 0.6, 0.5) ' tune if needed
                        If gap > spaceThreshold Then sb.Append(" ")
                    End If
                End If
                sb.Append(l.Value)
                prev = l
            Next

            Return sb.ToString()
        End Function

        ''' <summary>
        ''' Performs OCR on a PDF by invoking an LLM call with the PDF path as a binary object input.
        ''' </summary>
        ''' <param name="pdfPath">Path to the PDF file to OCR.</param>
        ''' <param name="context">Shared context containing model and API configuration.</param>
        ''' <param name="askUser">If False, suppresses all UI dialogs (for non-interactive callers like AutoPilot).</param>
        ''' <param name="additionalInstruction">Additional instructions to include in the system prompt for OCR processing.</param>
        ''' <returns>OCR result text, or an empty string if OCR is not available or fails.</returns>
        Private Shared Async Function oldPerformOCR(ByVal pdfPath As String,
                                                 context As ISharedContext,
                                                 Optional askUser As Boolean = True,
                                                 Optional additionalInstruction As String = Nothing) As Task(Of String)
            If Not IsOcrAvailable(context) Then
                If askUser Then
                    ShowCustomMessageBox("OCR is not available with your current model configuration.")
                End If
                Return ""
            End If

            Dim scope = CaptureModelConfigScope(context)

            Try
                Dim useSecondAPI As Boolean = False
                Dim timeOut = context.INI_Timeout

                If Not String.IsNullOrWhiteSpace(context.INI_AlternateModelPath) AndAlso
                   GetSpecialTaskModel(context, context.INI_AlternateModelPath, "OCR") Then

                    useSecondAPI = True
                    timeOut = context.INI_Timeout_2
                End If

                Dim systemPrompt As String = context.SP_InsertClipboard
                If Not String.IsNullOrWhiteSpace(additionalInstruction) Then
                    systemPrompt &= Environment.NewLine & Environment.NewLine & additionalInstruction.Trim()
                End If

                Dim result As String =
                    Await LLM(context, systemPrompt, "", "", "", timeOut * 2, useSecondAPI, Not askUser, "", pdfPath)

                Return If(result, "")

            Finally
                RestoreModelConfigScope(context, scope)
            End Try
        End Function


        Private Shared Async Function PerformOCR(ByVal pdfPath As String,
                                                 context As ISharedContext,
                                                 Optional askUser As Boolean = True,
                                                 Optional additionalInstruction As String = Nothing,
                                                 Optional showProgressWindow As Boolean = False) As Task(Of String)

            If Not IsOcrAvailable(context) Then
                If askUser Then
                    ShowCustomMessageBox("OCR is not available with your current model configuration.")
                End If
                Return ""
            End If

            Const ChunkOcrMaxRounds As Integer = 3

            If askUser Then showProgressWindow = True

            Dim scope = CaptureModelConfigScope(context)

            Try
                Dim useSecondAPI As Boolean = False
                Dim timeOut = context.INI_Timeout

                If Not String.IsNullOrWhiteSpace(context.INI_AlternateModelPath) AndAlso
                   GetSpecialTaskModel(context, context.INI_AlternateModelPath, "OCR") Then

                    useSecondAPI = True
                    timeOut = context.INI_Timeout_2
                End If

                Dim systemPrompt As String = context.SP_InsertClipboard
                If Not String.IsNullOrWhiteSpace(additionalInstruction) Then
                    systemPrompt &= Environment.NewLine & Environment.NewLine & additionalInstruction.Trim()
                End If

                Dim pageCount As Integer = 0
                Try
                    Using document As UglyToad.PdfPig.PdfDocument = UglyToad.PdfPig.PdfDocument.Open(pdfPath)
                        pageCount = document.NumberOfPages
                    End Using
                Catch
                    pageCount = 0
                End Try

                Dim showStatusWindow As Boolean = askUser OrElse showProgressWindow

                If pageCount <= 0 OrElse context.INI_ChunkOCR <= 0 OrElse pageCount <= context.INI_ChunkOCR Then
                    Dim result As String =
                        Await PerformSinglePdfOcrRequest(pdfPath, context, systemPrompt, timeOut, useSecondAPI, Not askUser)

                    Return If(result, "")
                End If

                Dim statusDialog As OcrChunkStatusDialog = Nothing

                Try
                    If showStatusWindow Then
                        statusDialog = New OcrChunkStatusDialog()
                        statusDialog.Show(
                            "OCR is running." & Environment.NewLine & Environment.NewLine &
                            $"Pages done: 0 / {pageCount:N0}" & Environment.NewLine &
                            "Chunks done: 0")
                    End If

                    Dim chunkedResult As String =
                        Await PerformAdaptiveChunkedOcr(pdfPath,
                                                        context,
                                                        systemPrompt,
                                                        timeOut,
                                                        useSecondAPI,
                                                        pageCount,
                                                        context.INI_ChunkOCR,
                                                        ChunkOcrMaxRounds,
                                                        statusDialog)

                    If String.IsNullOrWhiteSpace(chunkedResult) Then
                        Return ""
                    End If

                    Return chunkedResult

                Finally
                    If statusDialog IsNot Nothing Then
                        statusDialog.Dispose()
                    End If
                End Try

            Catch ex As OperationCanceledException
                Return ""
            Catch ex As System.Exception
                If askUser Then
                    ShowCustomMessageBox($"OCR failed: {ex.Message}")
                End If
                Return ""
            Finally
                RestoreModelConfigScope(context, scope)
            End Try
        End Function

        Private Shared Async Function PerformSinglePdfOcrRequest(ByVal pdfPath As String,
                                                                 context As ISharedContext,
                                                                 systemPrompt As String,
                                                                 timeOut As Long,
                                                                 useSecondAPI As Boolean,
                                                                 suppressUi As Boolean) As Task(Of String)
            Dim result As String =
                Await LLM(context, systemPrompt, "", "", "", timeOut * 2, useSecondAPI, suppressUi, "", pdfPath)

            Return If(result, "")
        End Function

        Private Shared Async Function PerformAdaptiveChunkedOcr(ByVal pdfPath As String,
                                                                context As ISharedContext,
                                                                systemPrompt As String,
                                                                timeOut As Long,
                                                                useSecondAPI As Boolean,
                                                                totalPageCount As Integer,
                                                                initialChunkSize As Integer,
                                                                maxRetries As Integer,
                                                                statusDialog As OcrChunkStatusDialog) As Task(Of String)

            Dim progressState As New OcrChunkProgressState(totalPageCount)
            Dim sb As New System.Text.StringBuilder()
            Dim currentStartPage As Integer = 1

            While currentStartPage <= totalPageCount
                ThrowIfOcrCancelled(statusDialog)

                Dim currentEndPage As Integer = System.Math.Min(totalPageCount, currentStartPage + initialChunkSize - 1)

                Dim chunkText As String =
                    Await ProcessOcrRangeWithRetries(pdfPath,
                                                     currentStartPage,
                                                     currentEndPage,
                                                     context,
                                                     systemPrompt,
                                                     timeOut,
                                                     useSecondAPI,
                                                     currentAttempt:=1,
                                                     currentChunkSize:=initialChunkSize,
                                                     maxRetries:=maxRetries,
                                                     progressState:=progressState,
                                                     statusDialog:=statusDialog)

                If chunkText Is Nothing Then
                    Return ""
                End If

                If sb.Length > 0 AndAlso chunkText.Length > 0 Then
                    sb.AppendLine()
                    sb.AppendLine()
                End If

                sb.Append(chunkText)
                currentStartPage = currentEndPage + 1
            End While

            If statusDialog IsNot Nothing Then
                statusDialog.UpdateStatus(
                    "OCR finished." & Environment.NewLine & Environment.NewLine &
                    $"Pages done: {progressState.TotalPages:N0} / {progressState.TotalPages:N0}" & Environment.NewLine &
                    $"Chunks done: {progressState.GetCompletedChunks():N0}")
                Await System.Threading.Tasks.Task.Delay(150)
            End If

            Return sb.ToString()
        End Function

        Private Shared Async Function ProcessOcrRangeWithRetries(ByVal pdfPath As String,
                                                                 ByVal startPage As Integer,
                                                                 ByVal endPage As Integer,
                                                                 context As ISharedContext,
                                                                 systemPrompt As String,
                                                                 timeOut As Long,
                                                                 useSecondAPI As Boolean,
                                                                 currentAttempt As Integer,
                                                                 currentChunkSize As Integer,
                                                                 maxRetries As Integer,
                                                                 progressState As OcrChunkProgressState,
                                                                 statusDialog As OcrChunkStatusDialog) As Task(Of String)

            ThrowIfOcrCancelled(statusDialog)
            UpdateOcrChunkStatus(statusDialog, progressState, startPage, endPage, currentAttempt, maxRetries)

            Dim tempChunkPath As String = Nothing
            Dim chunkText As String = ""

            Try
                tempChunkPath = System.IO.Path.Combine(
                    System.IO.Path.GetTempPath(),
                    $"RedInk_OCR_{System.Guid.NewGuid():N}_{startPage}_{endPage}.pdf")

                CreatePdfChunkForOcr(pdfPath, tempChunkPath, startPage, endPage)

                chunkText = Await PerformSinglePdfOcrRequest(tempChunkPath,
                                                             context,
                                                             systemPrompt,
                                                             timeOut,
                                                             useSecondAPI,
                                                             suppressUi:=True)
            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine(
                    $"OCR chunk attempt failed for pages {startPage}-{endPage} (attempt {currentAttempt} of {maxRetries}): {ex.Message}")
            Finally
                If tempChunkPath IsNot Nothing Then
                    Try
                        If System.IO.File.Exists(tempChunkPath) Then
                            System.IO.File.Delete(tempChunkPath)
                        End If
                    Catch
                    End Try
                End If
            End Try

            If Not String.IsNullOrWhiteSpace(chunkText) Then
                progressState.MarkCompleted(startPage, endPage)
                UpdateOcrChunkStatus(statusDialog, progressState, startPage, endPage, currentAttempt, maxRetries)
                Return chunkText
            End If

            If currentAttempt >= maxRetries Then
                Return Nothing
            End If

            Dim pageCountInRange As Integer = endPage - startPage + 1

            If pageCountInRange <= 1 Then
                Return Await ProcessOcrRangeWithRetries(pdfPath,
                                                        startPage,
                                                        endPage,
                                                        context,
                                                        systemPrompt,
                                                        timeOut,
                                                        useSecondAPI,
                                                        currentAttempt + 1,
                                                        1,
                                                        maxRetries,
                                                        progressState,
                                                        statusDialog)
            End If

            Dim reducedChunkSize As Integer = GetReducedChunkSize(pageCountInRange, currentChunkSize)
            Dim sb As New System.Text.StringBuilder()
            Dim currentSubStart As Integer = startPage

            While currentSubStart <= endPage
                ThrowIfOcrCancelled(statusDialog)

                Dim currentSubEnd As Integer = System.Math.Min(endPage, currentSubStart + reducedChunkSize - 1)

                Dim subChunkText As String =
                    Await ProcessOcrRangeWithRetries(pdfPath,
                                                     currentSubStart,
                                                     currentSubEnd,
                                                     context,
                                                     systemPrompt,
                                                     timeOut,
                                                     useSecondAPI,
                                                     currentAttempt + 1,
                                                     reducedChunkSize,
                                                     maxRetries,
                                                     progressState,
                                                     statusDialog)

                If subChunkText Is Nothing Then
                    Return Nothing
                End If

                If sb.Length > 0 AndAlso subChunkText.Length > 0 Then
                    sb.AppendLine()
                    sb.AppendLine()
                End If

                sb.Append(subChunkText)
                currentSubStart = currentSubEnd + 1
            End While

            Return sb.ToString()
        End Function

        Private Shared Sub CreatePdfChunkForOcr(ByVal sourcePdfPath As String,
                                                ByVal outputPdfPath As String,
                                                ByVal startPage As Integer,
                                                ByVal endPage As Integer)
            Using inputDocument As PdfSharp.Pdf.PdfDocument =
                PdfSharp.Pdf.IO.PdfReader.Open(sourcePdfPath, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import)

                Using outputDocument As New PdfSharp.Pdf.PdfDocument()
                    For pageIndex As Integer = startPage To endPage
                        outputDocument.AddPage(inputDocument.Pages(pageIndex - 1))
                    Next

                    outputDocument.Save(outputPdfPath)
                End Using
            End Using
        End Sub

        Private Shared Function GetReducedChunkSize(pageCountInRange As Integer, currentChunkSize As Integer) As Integer
            Dim reducedChunkSize As Integer =
                System.Math.Max(1, CInt(System.Math.Ceiling(currentChunkSize / 2.0R)))

            If reducedChunkSize >= pageCountInRange AndAlso pageCountInRange > 1 Then
                reducedChunkSize = System.Math.Max(1, CInt(System.Math.Ceiling(pageCountInRange / 2.0R)))
            End If

            If reducedChunkSize >= pageCountInRange AndAlso pageCountInRange > 1 Then
                reducedChunkSize = pageCountInRange - 1
            End If

            Return System.Math.Max(1, reducedChunkSize)
        End Function

        Private Shared Sub ThrowIfOcrCancelled(statusDialog As OcrChunkStatusDialog)
            If statusDialog IsNot Nothing AndAlso statusDialog.IsCancelled Then
                Throw New OperationCanceledException("OCR was cancelled.")
            End If
        End Sub

        Private Shared Sub UpdateOcrChunkStatus(statusDialog As OcrChunkStatusDialog,
                                                progressState As OcrChunkProgressState,
                                                startPage As Integer,
                                                endPage As Integer,
                                                currentAttempt As Integer,
                                                maxRetries As Integer)
            If statusDialog Is Nothing OrElse progressState Is Nothing Then
                Return
            End If

            Dim completedPages As Integer = progressState.GetCompletedPages()
            Dim completedChunks As Integer = progressState.GetCompletedChunks()

            statusDialog.UpdateStatus(
                "OCR is running." & Environment.NewLine & Environment.NewLine &
                $"Pages done: {completedPages:N0} / {progressState.TotalPages:N0}" & Environment.NewLine &
                $"Chunks done: {completedChunks:N0}" & Environment.NewLine &
                $"Now: {startPage:N0}-{endPage:N0}" & Environment.NewLine &
                $"Round: {currentAttempt:N0} / {maxRetries:N0}")
        End Sub

        Private NotInheritable Class OcrChunkProgressState
            Private ReadOnly _syncRoot As New Object()
            Private _completedPages As Integer
            Private _completedChunks As Integer

            Public Sub New(totalPages As Integer)
                Me.TotalPages = totalPages
            End Sub

            Public ReadOnly Property TotalPages As Integer

            Public Sub MarkCompleted(startPage As Integer, endPage As Integer)
                Dim pagesCompleted As Integer = System.Math.Max(0, endPage - startPage + 1)

                SyncLock _syncRoot
                    _completedPages += pagesCompleted
                    _completedChunks += 1
                End SyncLock
            End Sub

            Public Function GetCompletedPages() As Integer
                SyncLock _syncRoot
                    Return _completedPages
                End SyncLock
            End Function

            Public Function GetCompletedChunks() As Integer
                SyncLock _syncRoot
                    Return _completedChunks
                End SyncLock
            End Function
        End Class

        Private NotInheritable Class OcrChunkStatusDialog
            Implements System.IDisposable

            Private ReadOnly _syncRoot As New Object()
            Private ReadOnly _readyEvent As New System.Threading.ManualResetEventSlim(False)
            Private _uiThread As System.Threading.Thread = Nothing
            Private _statusText As String = "Starting OCR..."
            Private _cancelled As Boolean = False
            Private _closeRequested As Boolean = False
            Private _form As System.Windows.Forms.Form = Nothing

            Public Sub Show(initialText As String)
                SyncLock _syncRoot
                    _statusText = initialText
                    _cancelled = False
                    _closeRequested = False
                End SyncLock

                _uiThread = New System.Threading.Thread(AddressOf UiThreadMain) With {
                    .IsBackground = True
                }
                _uiThread.SetApartmentState(System.Threading.ApartmentState.STA)
                _uiThread.Start()

                _readyEvent.Wait()
            End Sub

            Public Sub UpdateStatus(text As String)
                SyncLock _syncRoot
                    _statusText = text
                End SyncLock
            End Sub

            Public ReadOnly Property IsCancelled As Boolean
                Get
                    SyncLock _syncRoot
                        Return _cancelled
                    End SyncLock
                End Get
            End Property

            Private Sub RequestCancel()
                SyncLock _syncRoot
                    _cancelled = True
                    _statusText = "Cancelling OCR..."
                End SyncLock
            End Sub

            Private Shared Function GetLowerMiddleLocation(formSize As System.Drawing.Size) As System.Drawing.Point
                Dim wa As System.Drawing.Rectangle = System.Windows.Forms.Screen.PrimaryScreen.WorkingArea
                Dim x As Integer = wa.Left + ((wa.Width - formSize.Width) \ 2)
                Dim y As Integer = wa.Top + CInt((wa.Height * 0.75R) - (formSize.Height / 2.0R))

                If x < wa.Left Then x = wa.Left
                If y < wa.Top Then y = wa.Top
                If x + formSize.Width > wa.Right Then x = wa.Right - formSize.Width
                If y + formSize.Height > wa.Bottom Then y = wa.Bottom - formSize.Height

                Return New System.Drawing.Point(x, y)
            End Function

            Private Sub UiThreadMain()
                Try
                    Dim localForm As New System.Windows.Forms.Form() With {
                        .Opacity = 0,
                        .Text = SharedMethods.AN & " OCR",
                        .FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog,
                        .StartPosition = System.Windows.Forms.FormStartPosition.Manual,
                        .MaximizeBox = False,
                        .MinimizeBox = False,
                        .ShowInTaskbar = False,
                        .TopMost = True,
                        .KeyPreview = True,
                        .AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font,
                        .AutoSize = False
                    }

                    Dim bmpIcon As New System.Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
                    localForm.Icon = System.Drawing.Icon.FromHandle(bmpIcon.GetHicon())

                    Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
                    localForm.Font = standardFont

                    Dim wa As System.Drawing.Rectangle = System.Windows.Forms.Screen.PrimaryScreen.WorkingArea
                    Dim paddingAll As Integer = 20
                    Dim gapAboveButtons As Integer = 10
                    Dim spacerExtra As Integer = 20
                    Dim minContentWidth As Integer = 220
                    Dim maxWindowWidth As Integer = CInt(System.Math.Floor(wa.Width * 0.32))
                    Dim maxWindowHeight As Integer = CInt(System.Math.Floor(wa.Height * 0.35))

                    Dim cancelButton As New System.Windows.Forms.Button() With {
                        .Text = "Cancel",
                        .AutoSize = True,
                        .Font = standardFont,
                        .Margin = New System.Windows.Forms.Padding(0)
                    }

                    Dim bottomFlow As New System.Windows.Forms.FlowLayoutPanel() With {
                        .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
                        .AutoSize = True,
                        .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                        .Margin = New System.Windows.Forms.Padding(0)
                    }
                    bottomFlow.Controls.Add(cancelButton)
                    bottomFlow.PerformLayout()

                    Dim reservedBottomHeight As Integer = bottomFlow.PreferredSize.Height + gapAboveButtons

                    Dim statusLabel As New System.Windows.Forms.Label() With {
                        .Text = If(_statusText, String.Empty),
                        .Font = standardFont,
                        .AutoSize = True,
                        .Margin = New System.Windows.Forms.Padding(0)
                    }

                    Dim getLabelPreferred As System.Func(Of Integer, System.Drawing.Size) =
                        Function(w As Integer) As System.Drawing.Size
                            statusLabel.MaximumSize = New System.Drawing.Size(System.Math.Max(1, w), 0)
                            Return statusLabel.GetPreferredSize(New System.Drawing.Size(System.Math.Max(1, w), 0))
                        End Function

                    Dim maxContentWidth As Integer = System.Math.Max(minContentWidth, maxWindowWidth - 2 * paddingAll)
                    Dim pref As System.Drawing.Size = getLabelPreferred(maxContentWidth)
                    Dim contentWidth As Integer = System.Math.Max(minContentWidth, System.Math.Min(maxContentWidth, pref.Width))
                    pref = getLabelPreferred(contentWidth)

                    Dim maxBodyHeightNoScroll As Integer = System.Math.Max(100, maxWindowHeight - reservedBottomHeight - spacerExtra - 2 * paddingAll)

                    While (pref.Height > maxBodyHeightNoScroll) AndAlso ((contentWidth + 2 * paddingAll) < maxWindowWidth)
                        Dim stepW As Integer = System.Math.Max(20, (maxWindowWidth - 2 * paddingAll - contentWidth) \ 3)
                        contentWidth = System.Math.Min(maxWindowWidth - 2 * paddingAll, contentWidth + stepW)
                        pref = getLabelPreferred(contentWidth)
                    End While

                    Dim bodyPanelHeight As Integer = System.Math.Max(90, System.Math.Min(pref.Height, maxBodyHeightNoScroll))

                    Dim bodyPanel As New System.Windows.Forms.Panel() With {
                        .AutoSize = False,
                        .Size = New System.Drawing.Size(contentWidth, bodyPanelHeight),
                        .Margin = New System.Windows.Forms.Padding(0),
                        .Padding = New System.Windows.Forms.Padding(0)
                    }

                    statusLabel.MaximumSize = New System.Drawing.Size(contentWidth, 0)
                    bodyPanel.Controls.Add(statusLabel)
                    statusLabel.Location = New System.Drawing.Point(0, 0)

                    Dim table As New System.Windows.Forms.TableLayoutPanel() With {
                        .Dock = System.Windows.Forms.DockStyle.Fill,
                        .ColumnCount = 1,
                        .RowCount = 3,
                        .Padding = New System.Windows.Forms.Padding(paddingAll),
                        .AutoSize = False,
                        .Margin = New System.Windows.Forms.Padding(0)
                    }
                    table.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
                    table.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, bodyPanelHeight))
                    table.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, spacerExtra))
                    table.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))

                    table.Controls.Add(bodyPanel, 0, 0)

                    Dim spacer As New System.Windows.Forms.Panel() With {
                        .Height = spacerExtra,
                        .Width = 1,
                        .Margin = New System.Windows.Forms.Padding(0)
                    }
                    table.Controls.Add(spacer, 0, 1)

                    Dim bottomHost As New System.Windows.Forms.Panel() With {
                        .AutoSize = True,
                        .Margin = New System.Windows.Forms.Padding(0)
                    }
                    bottomHost.Padding = New System.Windows.Forms.Padding(0, gapAboveButtons, 0, 0)
                    bottomHost.Controls.Add(bottomFlow)
                    table.Controls.Add(bottomHost, 0, 2)

                    localForm.Controls.Clear()
                    localForm.Controls.Add(table)

                    Dim clientW As Integer = contentWidth + 2 * paddingAll
                    Dim clientH As Integer = bodyPanelHeight + spacerExtra + reservedBottomHeight + 2 * paddingAll
                    clientW = System.Math.Min(clientW, maxWindowWidth)
                    clientH = System.Math.Min(clientH, maxWindowHeight)
                    localForm.ClientSize = New System.Drawing.Size(clientW, clientH)
                    localForm.Location = GetLowerMiddleLocation(localForm.Size)

                    AddHandler cancelButton.Click,
                        Sub(sender As Object, e As System.EventArgs)
                            RequestCancel()
                        End Sub

                    AddHandler localForm.KeyDown,
                        Sub(sender As Object, e As System.Windows.Forms.KeyEventArgs)
                            If e.KeyCode = System.Windows.Forms.Keys.Escape Then
                                RequestCancel()
                                e.SuppressKeyPress = True
                            End If
                        End Sub

                    AddHandler localForm.FormClosing,
                        Sub(sender As Object, e As System.Windows.Forms.FormClosingEventArgs)
                            Dim closeRequested As Boolean

                            SyncLock _syncRoot
                                closeRequested = _closeRequested
                            End SyncLock

                            If Not closeRequested Then
                                RequestCancel()
                            End If
                        End Sub

                    AddHandler localForm.Shown,
                        Sub(sender As Object, e As System.EventArgs)
                            localForm.TopMost = False
                            localForm.TopMost = True
                            localForm.Activate()
                            localForm.BringToFront()
                        End Sub

                    Dim refreshTimer As New System.Windows.Forms.Timer() With {
                        .Interval = 100
                    }

                    AddHandler refreshTimer.Tick,
                        Sub(sender As Object, e As System.EventArgs)
                            Dim latestText As String = ""
                            Dim closeRequested As Boolean = False

                            SyncLock _syncRoot
                                latestText = _statusText
                                closeRequested = _closeRequested
                            End SyncLock

                            statusLabel.Text = latestText

                            If closeRequested Then
                                refreshTimer.Stop()
                                localForm.Close()
                            End If
                        End Sub

                    SyncLock _syncRoot
                        _form = localForm
                    End SyncLock

                    _readyEvent.Set()
                    refreshTimer.Start()
                    localForm.Opacity = 1

                    Dim owner As System.Windows.Forms.IWin32Window = SharedMethods.ResolveDialogOwner()
                    Dim ownerScope As System.IDisposable = Nothing

                    Try
                        ownerScope = SharedMethods.PushDialogOwner(localForm)

                        If owner IsNot Nothing Then
                            localForm.ShowDialog(owner)
                        Else
                            localForm.ShowDialog()
                        End If
                    Finally
                        If ownerScope IsNot Nothing Then
                            Try
                                ownerScope.Dispose()
                            Catch
                            End Try
                        End If
                    End Try

                    refreshTimer.Dispose()

                Catch
                    _readyEvent.Set()
                Finally
                    SyncLock _syncRoot
                        _form = Nothing
                    End SyncLock
                End Try
            End Sub

            Public Sub Close()
                SyncLock _syncRoot
                    _closeRequested = True
                End SyncLock

                Dim localForm As System.Windows.Forms.Form = Nothing

                SyncLock _syncRoot
                    localForm = _form
                End SyncLock

                If localForm IsNot Nothing AndAlso localForm.IsHandleCreated Then
                    Try
                        localForm.BeginInvoke(New System.Action(Sub() localForm.Close()))
                    Catch
                    End Try
                End If

                If _uiThread IsNot Nothing AndAlso _uiThread.IsAlive Then
                    Try
                        _uiThread.Join(2000)
                    Catch
                    End Try
                End If
            End Sub

            Public Sub Dispose() Implements System.IDisposable.Dispose
                Close()
                _readyEvent.Dispose()
            End Sub
        End Class


        ''' <summary>
        ''' Determines whether OCR is available based on the configured model capabilities.
        ''' </summary>
        ''' <param name="context">Shared context containing model and API configuration.</param>
        ''' <returns>True if OCR is available, False otherwise.</returns>
        Public Shared Function IsOcrAvailable(context As ISharedContext) As Boolean
            If context Is Nothing Then Return False

            If Not String.IsNullOrWhiteSpace(context.INI_AlternateModelPath) Then
                Dim scope = CaptureModelConfigScope(context)

                Try
                    If GetSpecialTaskModel(context, context.INI_AlternateModelPath, "OCR") Then
                        Return True
                    End If
                Catch
                Finally
                    RestoreModelConfigScope(context, scope)
                End Try
            End If

            Return IsApiCallObjectOcrCapable(context.INI_APICall_Object)
        End Function


        ''' <summary>
        ''' Checks if the given APICall_Object configuration string supports PDF/OCR.
        ''' </summary>
        ''' <param name="apiCallObject">The INI_APICall_Object or INI_APICall_Object_2 string.</param>
        ''' <returns>True if OCR/PDF is supported, False otherwise.</returns>
        Private Shared Function IsApiCallObjectOcrCapable(apiCallObject As String) As Boolean
            ' If null or empty, OCR is not available
            If String.IsNullOrWhiteSpace(apiCallObject) Then
                Return False
            End If

            ' Check if the string contains segment separators (¦)
            Dim segments As String() = apiCallObject.Split(New Char() {"¦"c}, StringSplitOptions.RemoveEmptyEntries)

            ' Track if we found any segment without a filter (means all types supported)
            ' or any segment with a filter that includes PDF
            Dim hasUnfilteredSegment As Boolean = False
            Dim hasPdfFilter As Boolean = False
            Dim allSegmentsHaveFilters As Boolean = True

            For Each segment As String In segments
                Dim trimmedSegment As String = segment.Trim()

                ' Check if this segment has a filter (starts with [...])
                If trimmedSegment.StartsWith("[") Then
                    ' Extract the filter content between [ and ]
                    Dim closeBracketIdx As Integer = trimmedSegment.IndexOf("]"c)
                    If closeBracketIdx > 1 Then
                        Dim filterContent As String = trimmedSegment.Substring(1, closeBracketIdx - 1)

                        ' Check if the filter contains application/pdf or pdf
                        If filterContent.IndexOf("application/pdf", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                           filterContent.IndexOf("pdf", StringComparison.OrdinalIgnoreCase) >= 0 Then
                            hasPdfFilter = True
                        End If
                    End If
                Else
                    ' No filter on this segment - means it accepts all types
                    hasUnfilteredSegment = True
                    allSegmentsHaveFilters = False
                End If
            Next

            ' OCR is available if:
            ' 1. There's at least one segment without a filter (accepts all), OR
            ' 2. There's a segment with a filter that includes PDF
            If hasUnfilteredSegment Then
                Return True
            End If

            If hasPdfFilter Then
                Return True
            End If

            ' If all segments have filters and none include PDF, OCR is not available
            If allSegmentsHaveFilters Then
                Return False
            End If

            ' Default: if we have content but couldn't parse filters, assume capable
            Return True
        End Function


        ''' <summary>
        ''' Determines whether audio transcription is available based on the configured model capabilities.
        ''' Checks for audio/* MIME type support in the APICall_Object configuration,
        ''' mirroring the logic of <see cref="IsOcrAvailable"/> for PDF.
        ''' </summary>
        ''' <param name="context">Shared context containing model and API configuration.</param>
        ''' <returns>True if audio transcription via binary input is available, False otherwise.</returns>
        Public Shared Function IsAudioTranscriptionAvailable(context As ISharedContext) As Boolean
            If context Is Nothing Then Return False

            ' First check: alternate model with AudioTranscription task flag
            If Not String.IsNullOrWhiteSpace(context.INI_AlternateModelPath) Then
                Dim savedConfig As ModelConfig = GetCurrentConfig(context)
                Dim savedConfigLoaded As Boolean = originalConfigLoaded

                Try
                    If GetSpecialTaskModel(context, context.INI_AlternateModelPath, "AudioTranscription") Then
                        RestoreDefaults(context, savedConfig)
                        originalConfigLoaded = savedConfigLoaded
                        Return True
                    End If
                Catch
                Finally
                    RestoreDefaults(context, savedConfig)
                    originalConfigLoaded = savedConfigLoaded
                End Try
            End If

            ' Second check: primary model's APICall_Object supports audio MIME types
            Return IsApiCallObjectAudioCapable(context.INI_APICall_Object)
        End Function

        ''' <summary>
        ''' Checks if the given APICall_Object configuration string supports audio input.
        ''' Mirrors <see cref="IsApiCallObjectOcrCapable"/> but checks for audio/* MIME types.
        ''' </summary>
        ''' <param name="apiCallObject">The INI_APICall_Object or INI_APICall_Object_2 string.</param>
        ''' <returns>True if audio input is supported, False otherwise.</returns>
        Public Shared Function IsApiCallObjectAudioCapable(apiCallObject As String) As Boolean
            If String.IsNullOrWhiteSpace(apiCallObject) Then Return False

            Dim segments As String() = apiCallObject.Split(New Char() {"¦"c}, StringSplitOptions.RemoveEmptyEntries)
            Dim hasUnfilteredSegment As Boolean = False
            Dim hasAudioFilter As Boolean = False
            Dim allSegmentsHaveFilters As Boolean = True

            For Each segment As String In segments
                Dim trimmedSegment As String = segment.Trim()

                If trimmedSegment.StartsWith("[") Then
                    Dim closeBracketIdx As Integer = trimmedSegment.IndexOf("]"c)
                    If closeBracketIdx > 1 Then
                        Dim filterContent As String = trimmedSegment.Substring(1, closeBracketIdx - 1)
                        If filterContent.IndexOf("audio/", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                           filterContent.IndexOf("audio/*", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                           filterContent.IndexOf("*/*", StringComparison.OrdinalIgnoreCase) >= 0 Then
                            hasAudioFilter = True
                        End If
                    End If
                Else
                    hasUnfilteredSegment = True
                    allSegmentsHaveFilters = False
                End If
            Next

            If hasUnfilteredSegment Then Return True
            If hasAudioFilter Then Return True
            If allSegmentsHaveFilters Then Return False
            Return True
        End Function


    End Class

End Namespace
