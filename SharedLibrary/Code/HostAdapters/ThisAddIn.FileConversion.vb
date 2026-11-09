' Part of "Red Ink for Word / Excel".
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' Shared host adapter source: both VSTO projects compile this linked partial class.
' Retains the complete file conversion, PDF rasterization, preparation and indexing workflow.
Option Explicit On
Option Strict On
Option Infer On

Imports System.IO
Imports System.Linq
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports Slib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn
    Private Const Add_OcrMarkdownInstruction As String =
            "Provide the output as Markdown. " &
            "Use Markdown formatting where it helps preserve the original structure, including headings, lists, tables and other tabular content. " &
            "Return only the Markdown content and no surrounding code fences."

    ''' <summary>
    ''' Converts a PDF to an image-only PDF by rasterizing each page via the Windows 10+
    ''' built-in PDF renderer (Windows.Data.Pdf). This API is maintained by Microsoft via
    ''' Windows Update, handles all modern PDF features (transparency groups, layered image
    ''' XObjects, form XObjects, compositing modes), and requires no third-party native DLLs.
    ''' Falls back to PdfiumViewer if the Windows API is unavailable.
    ''' </summary>
    ''' <param name="inputPath">Input PDF file path</param>
    ''' <param name="outputPath">Output PDF file path</param>
    ''' <param name="dpi">Rasterization DPI (default 200 for balance of size/quality)</param>
    Private Shared Sub FlattenPdfToImageOnly(inputPath As String, outputPath As String, Optional dpi As Integer = 200)

        Dim dbg As System.Action(Of String) = Sub(msg) System.Diagnostics.Debug.WriteLine($"[FlattenPDF] {msg}")
        dbg($"=== FlattenPdfToImageOnly START === dpi={dpi}")
        dbg($"Input:  {inputPath}")
        dbg($"Output: {outputPath}")

        Dim usedWindowsPdf As Boolean = False

#If HAS_WINRT Then
        If System.Environment.OSVersion.Version.Major >= 10 Then
            Try
                FlattenPdfViaWindowsApi(inputPath, outputPath, dpi)
                usedWindowsPdf = True
                dbg("Rendered via Windows.Data.Pdf — all visual elements preserved")
            Catch ex As System.Exception
                dbg($"Windows.Data.Pdf failed: {ex.Message} — falling back to PdfiumViewer")
                usedWindowsPdf = False
            End Try
        Else
            dbg($"OS version {System.Environment.OSVersion.Version} — Windows.Data.Pdf not available, using PdfiumViewer")
        End If
#End If

        If Not usedWindowsPdf Then
            dbg("WARNING: PdfiumViewer fallback may not preserve layered image objects with transparency")
            FlattenPdfToImageOnlyViaPdfium(inputPath, outputPath, dpi)
        End If

        dbg("=== FlattenPdfToImageOnly END ===")
    End Sub

    Private Shared Sub FlattenPdfViaWindowsApi(inputPath As String, outputPath As String, dpi As Integer)
#If Not HAS_WINRT Then
        Throw New System.PlatformNotSupportedException("Windows.Data.Pdf is not available (SDK not installed at build time).")
#Else
        Dim dbg As System.Action(Of String) = Sub(msg) System.Diagnostics.Debug.WriteLine($"[FlattenPDF-WinPdf] {msg}")

        Dim renderException As System.Exception = Nothing

        Dim staThread As New System.Threading.Thread(
            Sub()
                Try
                    FlattenPdfViaWindowsApiCore(inputPath, outputPath, dpi, dbg)
                Catch ex As System.Exception
                    renderException = ex
                End Try
            End Sub)

        staThread.SetApartmentState(System.Threading.ApartmentState.STA)
        staThread.IsBackground = True
        staThread.Start()

        If Not staThread.Join(System.TimeSpan.FromMinutes(5)) Then
            Try : staThread.Abort() : Catch : End Try
            Throw New System.TimeoutException("Windows.Data.Pdf rendering timed out after 5 minutes.")
        End If

        If renderException IsNot Nothing Then
            Throw renderException
        End If
#End If
    End Sub

    Private Shared Sub FlattenPdfViaWindowsApiCore(inputPath As String, outputPath As String,
                                                    dpi As Integer,
                                                    dbg As System.Action(Of String))
#If Not HAS_WINRT Then
        Throw New System.PlatformNotSupportedException("Windows.Data.Pdf is not available (SDK not installed at build time).")
#Else
        ' Open the PDF via Windows.Data.Pdf
        Dim storageFile As Windows.Storage.StorageFile =
            Windows.Storage.StorageFile.GetFileFromPathAsync(inputPath).GetAwaiter().GetResult()

        Dim winPdf As Windows.Data.Pdf.PdfDocument =
            Windows.Data.Pdf.PdfDocument.LoadFromFileAsync(storageFile).GetAwaiter().GetResult()

        dbg($"Opened via Windows.Data.Pdf — {winPdf.PageCount} page(s)")

        Dim outDoc As New PdfSharp.Pdf.PdfDocument()

        ' Copy metadata from source
        Try
            Using srcDoc As PdfSharp.Pdf.PdfDocument = PdfSharp.Pdf.IO.PdfReader.Open(inputPath, PdfSharp.Pdf.IO.PdfDocumentOpenMode.InformationOnly)
                CopyPdfMetadata(srcDoc, outDoc)
            End Using
        Catch
        End Try

        ' Scale factor: Windows.Data.Pdf renders at 96 DPI by default.
        Dim scaleFactor As Double = dpi / 96.0

        For pageIndex As UInteger = 0 To CUInt(winPdf.PageCount - 1)
            Dim page As Windows.Data.Pdf.PdfPage = winPdf.GetPage(pageIndex)

            Dim pageWidthPt As Double = page.Size.Width * 72.0 / 96.0
            Dim pageHeightPt As Double = page.Size.Height * 72.0 / 96.0

            Dim renderWidthPx As UInteger = CUInt(System.Math.Round(page.Size.Width * scaleFactor))
            Dim renderHeightPx As UInteger = CUInt(System.Math.Round(page.Size.Height * scaleFactor))

            dbg($"Page {pageIndex + 1}: {pageWidthPt:F0}x{pageHeightPt:F0} pts, render {renderWidthPx}x{renderHeightPx} px")

            Using renderStream As New Windows.Storage.Streams.InMemoryRandomAccessStream()
                Dim renderOptions As New Windows.Data.Pdf.PdfPageRenderOptions()
                renderOptions.DestinationWidth = renderWidthPx
                renderOptions.DestinationHeight = renderHeightPx
                renderOptions.BitmapEncoderId = Windows.Graphics.Imaging.BitmapEncoder.PngEncoderId

                page.RenderToStreamAsync(renderStream, renderOptions).GetAwaiter().GetResult()

                renderStream.Seek(0)
                Dim netStream As System.IO.Stream = System.IO.WindowsRuntimeStreamExtensions.AsStreamForRead(renderStream)

                Using ms As New System.IO.MemoryStream()
                    netStream.CopyTo(ms)
                    ms.Position = 0

                    Using bmp As New System.Drawing.Bitmap(ms)
                        Using jpegMs As New System.IO.MemoryStream()
                            Dim jpegEncoder As System.Drawing.Imaging.ImageCodecInfo = GetJpegEncoder()
                            If jpegEncoder IsNot Nothing Then
                                Dim encoderParams As New System.Drawing.Imaging.EncoderParameters(1)
                                encoderParams.Param(0) = New System.Drawing.Imaging.EncoderParameter(
                                    System.Drawing.Imaging.Encoder.Quality, 85L)
                                bmp.Save(jpegMs, jpegEncoder, encoderParams)
                            Else
                                bmp.Save(jpegMs, System.Drawing.Imaging.ImageFormat.Png)
                            End If

                            jpegMs.Position = 0

                            Dim outPage As PdfSharp.Pdf.PdfPage = outDoc.AddPage()
                            outPage.Width = PdfSharp.Drawing.XUnit.FromPoint(pageWidthPt)
                            outPage.Height = PdfSharp.Drawing.XUnit.FromPoint(pageHeightPt)

                            Using xgfx As PdfSharp.Drawing.XGraphics = PdfSharp.Drawing.XGraphics.FromPdfPage(outPage)
                                Using ximg As PdfSharp.Drawing.XImage = PdfSharp.Drawing.XImage.FromStream(jpegMs)
                                    xgfx.DrawImage(ximg, 0, 0, outPage.Width.Point, outPage.Height.Point)
                                End Using
                            End Using
                        End Using
                    End Using
                End Using
            End Using

            page.Dispose()
        Next

        outDoc.Save(outputPath)
        outDoc.Close()

        dbg($"Output saved: {New System.IO.FileInfo(outputPath).Length:N0} bytes")
#End If
    End Sub


    ''' <summary>
    ''' Rasterizes a PDF to an image-only PDF via PdfiumViewer.
    ''' Used as fallback if Windows.Data.Pdf is unavailable (pre-Windows 10).
    ''' </summary>
    Private Shared Sub FlattenPdfToImageOnlyViaPdfium(inputPath As String, outputPath As String,
                                                      Optional dpi As Integer = 200,
                                                      Optional metadataSourcePath As String = Nothing)

        EnsureConversionPdfiumLoaded()

        Using pdf As PdfiumViewer.PdfDocument = PdfiumViewer.PdfDocument.Load(inputPath)
            Dim outDoc As New PdfSharp.Pdf.PdfDocument()

            Dim metaPath As String = If(String.IsNullOrEmpty(metadataSourcePath), inputPath, metadataSourcePath)
            Try
                Using srcDoc As PdfSharp.Pdf.PdfDocument = PdfSharp.Pdf.IO.PdfReader.Open(metaPath, PdfSharp.Pdf.IO.PdfDocumentOpenMode.InformationOnly)
                    CopyPdfMetadata(srcDoc, outDoc)
                End Using
            Catch
            End Try

            For pageIndex As Integer = 0 To pdf.PageCount - 1
                Dim sizePt As System.Drawing.SizeF = pdf.PageSizes(pageIndex)
                Dim widthPx As Integer = CInt(System.Math.Round(sizePt.Width / 72.0 * dpi))
                Dim heightPx As Integer = CInt(System.Math.Round(sizePt.Height / 72.0 * dpi))

                Dim renderFlags As PdfiumViewer.PdfRenderFlags =
                    PdfiumViewer.PdfRenderFlags.Annotations Or
                    PdfiumViewer.PdfRenderFlags.LcdText Or
                    PdfiumViewer.PdfRenderFlags.ForPrinting

                Using rendered As System.Drawing.Image = pdf.Render(pageIndex, widthPx, heightPx, dpi, dpi, renderFlags)
                    Dim outPage As PdfSharp.Pdf.PdfPage = outDoc.AddPage()
                    outPage.Width = PdfSharp.Drawing.XUnit.FromPoint(sizePt.Width)
                    outPage.Height = PdfSharp.Drawing.XUnit.FromPoint(sizePt.Height)

                    Using ms As New System.IO.MemoryStream()
                        Dim jpegEncoder As System.Drawing.Imaging.ImageCodecInfo = GetJpegEncoder()
                        If jpegEncoder IsNot Nothing Then
                            Dim encoderParams As New System.Drawing.Imaging.EncoderParameters(1)
                            encoderParams.Param(0) = New System.Drawing.Imaging.EncoderParameter(
                                System.Drawing.Imaging.Encoder.Quality, 85L)
                            rendered.Save(ms, jpegEncoder, encoderParams)
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

            outDoc.Save(outputPath)
            outDoc.Close()
        End Using
    End Sub



    ''' <summary>
    ''' Gets the JPEG image encoder for file size optimization.
    ''' </summary>
    Private Shared Function GetJpegEncoder() As System.Drawing.Imaging.ImageCodecInfo
        Dim codecs As System.Drawing.Imaging.ImageCodecInfo() = System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders()
        For Each codec In codecs
            If codec.MimeType = "image/jpeg" Then
                Return codec
            End If
        Next
        Return Nothing
    End Function

    ''' <summary>
    ''' Copies PDF metadata from source to destination document.
    ''' </summary>
    ''' <summary>
    ''' Copies PDF metadata from source to destination document.
    ''' Handles DateTimeKind conversion to avoid DateTimeInvalidLocalFormat exceptions.
    ''' </summary>
    Private Shared Sub CopyPdfMetadata(src As PdfSharp.Pdf.PdfDocument, dest As PdfSharp.Pdf.PdfDocument)
        If src Is Nothing OrElse dest Is Nothing Then Return

        Try
            dest.Info.Title = src.Info.Title
            dest.Info.Author = src.Info.Author
            dest.Info.Subject = src.Info.Subject
            dest.Info.Keywords = src.Info.Keywords
            dest.Info.Creator = src.Info.Creator

            ' Handle CreationDate - convert UTC to Local to avoid DateTimeInvalidLocalFormat MDA
            If src.Info.CreationDate <> Date.MinValue Then
                Dim creationDate As System.DateTime = src.Info.CreationDate
                If creationDate.Kind = System.DateTimeKind.Utc Then
                    creationDate = System.DateTime.SpecifyKind(creationDate.ToLocalTime(), System.DateTimeKind.Local)
                ElseIf creationDate.Kind = System.DateTimeKind.Unspecified Then
                    creationDate = System.DateTime.SpecifyKind(creationDate, System.DateTimeKind.Local)
                End If
                dest.Info.CreationDate = creationDate
            End If

            ' Handle ModificationDate - convert UTC to Local to avoid DateTimeInvalidLocalFormat MDA
            If src.Info.ModificationDate <> Date.MinValue Then
                Dim modificationDate As System.DateTime = src.Info.ModificationDate
                If modificationDate.Kind = System.DateTimeKind.Utc Then
                    modificationDate = System.DateTime.SpecifyKind(modificationDate.ToLocalTime(), System.DateTimeKind.Local)
                ElseIf modificationDate.Kind = System.DateTimeKind.Unspecified Then
                    modificationDate = System.DateTime.SpecifyKind(modificationDate, System.DateTimeKind.Local)
                End If
                dest.Info.ModificationDate = modificationDate
            End If
        Catch
            ' Ignore metadata copy errors
        End Try
    End Sub



    ''' <summary>
    ''' Prompts the user for a file or directory, reads content using GetFileContent(),
    ''' and saves the extracted text as a .txt file at the same location (or optionally in a subdirectory).
    ''' Shows a progress bar and allows cancellation. If 2+ PDF files are present and OCR is available,
    ''' prompts user once for OCR preference. Optionally flattens PDFs to image-only before OCR
    ''' to ensure all visual elements (layered images, transparency groups) are captured.
    ''' </summary>
    Public Async Sub ExportFileContentToText()
        Dim selectedPath As String = ""

        ' Show DragDropForm for file or directory selection
        Globals.ThisAddIn.DragDropFormLabel = "Select a file or folder to convert to text"

        Try
            Using frm As New DragDropForm(DragDropMode.FileOrDirectory)
                Dim __safeDialogOwner1936 As System.Windows.Forms.IWin32Window = SharedLibrary.SharedLibrary.SharedMethods.ResolveSameThreadDialogOwner()
                If If(__safeDialogOwner1936 IsNot Nothing, frm.ShowDialog(__safeDialogOwner1936), frm.ShowDialog()) = System.Windows.Forms.DialogResult.OK Then
                    selectedPath = frm.SelectedFilePath
                End If
            End Using
        Finally
            Globals.ThisAddIn.DragDropFormLabel = ""
        End Try

        If String.IsNullOrWhiteSpace(selectedPath) Then
            Return
        End If

        ' Determine if it's a file or directory
        Dim isDirectory As Boolean = IO.Directory.Exists(selectedPath)
        Dim isFile As Boolean = IO.File.Exists(selectedPath)

        If Not isFile AndAlso Not isDirectory Then
            ShowCustomMessageBox("The selected path does not exist.")
            Return
        End If

        ' Build supported extensions list aligned with GetFileContentEx
        Dim supportedExtensionsList As New System.Collections.Generic.List(Of String)({
            ".txt", ".rtf", ".docx", ".pdf", ".xlsx", ".pptx", ".ini", ".csv", ".log",
            ".json", ".xml", ".html", ".htm", ".md", ".yaml", ".yml",
            ".vb", ".cs", ".js", ".ts", ".py", ".java", ".cpp", ".c", ".h", ".sql",
            ".eml", ".msg",
            ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff", ".tif", ".webp", ".svg",
            ".mp3", ".wav", ".ogg", ".flac", ".m4a", ".aac", ".wma", ".opus", ".webm",
            ".mp4", ".avi", ".mkv", ".mov", ".wmv"
        })

        ' Conditionally include .doc (legacy Word format — may execute macros via COM)
        If INI_AllowLegacyDocFiles Then
            supportedExtensionsList.Add(".doc")
        End If

        ' Filter out binary/media types the current model can't handle
        Dim supportedExtensions As String() = supportedExtensionsList.
            Where(Function(ext) Not IsBinaryMediaExtension(ext) OrElse IsModelCapableForExtension(_context, ext)).
            ToArray()

        ' Collect files to process
        Dim filesToProcess As New System.Collections.Generic.List(Of String)()
        Dim unsupportedFiles As New System.Collections.Generic.List(Of String)()
        Dim legacyDocFilesSkipped As New System.Collections.Generic.List(Of String)()

        If isFile Then
            Dim ext As String = IO.Path.GetExtension(selectedPath).ToLowerInvariant()
            If supportedExtensions.Contains(ext) Then
                filesToProcess.Add(selectedPath)
            ElseIf ext = ".doc" AndAlso Not INI_AllowLegacyDocFiles Then
                ShowCustomMessageBox($"Legacy .doc files are disabled for security reasons (AllowLegacyDocFiles is off). Save as .docx first.")
                Return
            Else
                ShowCustomMessageBox($"File type '{ext}' is not supported for text extraction.")
                Return
            End If
        Else
            ' Directory - ask user about recursion
            Dim recurseChoice As Integer = ShowCustomYesNoBox(
                "Do you want to include files from subdirectories?",
                "Yes, include subdirectories",
                "No, top directory only")

            If recurseChoice = 0 Then
                Return ' User aborted
            End If

            Dim searchOption As IO.SearchOption = If(recurseChoice = 1, IO.SearchOption.AllDirectories, IO.SearchOption.TopDirectoryOnly)

            ' Collect all files
            Dim allFiles As String() = IO.Directory.GetFiles(selectedPath, "*.*", searchOption)

            For Each f In allFiles
                Dim ext As String = IO.Path.GetExtension(f).ToLowerInvariant()
                If supportedExtensions.Contains(ext) Then
                    filesToProcess.Add(f)
                ElseIf ext = ".doc" AndAlso Not INI_AllowLegacyDocFiles Then
                    legacyDocFilesSkipped.Add(f)
                Else
                    unsupportedFiles.Add(f)
                End If
            Next

            If filesToProcess.Count = 0 Then
                ShowCustomMessageBox($"No supported files found in the selected directory." &
                    If(unsupportedFiles.Count > 0, $" ({unsupportedFiles.Count} unsupported file(s) were ignored.)", "") &
                    If(legacyDocFilesSkipped.Count > 0, $" ({legacyDocFilesSkipped.Count} .doc file(s) skipped — AllowLegacyDocFiles is off.)", ""))
                Return
            End If

            ' Confirm if many files
            If filesToProcess.Count > 10 Then
                Dim confirmAnswer As Integer = ShowCustomYesNoBox(
                    $"The directory contains {filesToProcess.Count} supported files to process." &
                    If(unsupportedFiles.Count > 0, $" ({unsupportedFiles.Count} unsupported files will be ignored.)", "") &
                    If(legacyDocFilesSkipped.Count > 0, $" ({legacyDocFilesSkipped.Count} .doc files skipped — disabled for security.)", "") &
                    " Continue?",
                    "Yes, continue", "No, abort")
                If confirmAnswer <> 1 Then
                    Return
                End If
            End If
        End If

        ' Ask user about output location
        Dim useSubdirectory As Boolean = False
        Dim outputSubdirName As String = "extracted_text"

        If filesToProcess.Count > 1 Then
            Dim outputChoice As Integer = ShowCustomYesNoBox(
                "Where should the extracted text files be saved?",
                $"In a subdirectory '{outputSubdirName}'",
                "Same location as original files")

            If outputChoice = 0 Then
                Return ' User aborted
            End If
            useSubdirectory = (outputChoice = 1)
        End If

        ' Count PDF files to determine OCR behavior
        Dim pdfFiles As System.Collections.Generic.List(Of String) = filesToProcess.Where(
            Function(f) IO.Path.GetExtension(f).Equals(".pdf", System.StringComparison.OrdinalIgnoreCase)
        ).ToList()
        Dim pdfCount As Integer = pdfFiles.Count

        ' Detect files that support direct Markdown output
        Dim markdownEligibleFiles As System.Collections.Generic.List(Of String) = filesToProcess.Where(
            Function(f)
                Dim ext As String = IO.Path.GetExtension(f).ToLowerInvariant()
                Return ext = ".docx" OrElse ext = ".pdf"
            End Function
        ).ToList()

        ' Determine OCR settings
        Dim doOcr As Boolean = False
        Dim flattenBeforeOcr As Boolean = False
        Dim useMarkdownOutputWhenSupported As Boolean = False
        Dim ocrMarkdownInstruction As String = String.Empty
        Dim ocrCleanupMode As OcrMarkdownCleanupMode = OcrMarkdownCleanupMode.None
        Dim markdownPreparation As PdfMarkdownPreparationOptions = Nothing
        Dim markdownPreparationSkippedCount As System.Int32 = 0
        Dim conversionDiagnostics As New System.Text.StringBuilder()
        Dim indexMetadataWarnings As New System.Collections.Generic.List(Of System.String)()

        If pdfCount >= 2 AndAlso SharedMethods.IsOcrAvailable(_context) Then
            Dim ocrChoice As Integer = ShowCustomYesNoBox(
                $"There are {pdfCount} PDF files to process." & vbCrLf & vbCrLf &
                "Some PDFs may require OCR (optical character recognition) to extract text from scanned documents or images." & vbCrLf & vbCrLf &
                "How would you like to handle OCR for these PDF files?",
                "Enable OCR for all PDFs",
                "Skip OCR for all PDFs")

            If ocrChoice = 0 Then
                Return ' User aborted
            ElseIf ocrChoice = 1 Then
                doOcr = True
            Else
                doOcr = False
            End If
        ElseIf pdfCount = 1 AndAlso SharedMethods.IsOcrAvailable(_context) Then
            Dim ocrChoice As Integer = ShowCustomYesNoBox(
                "The PDF may require OCR to extract text from scanned content." & vbCrLf & vbCrLf &
                "Do you want to enable OCR?",
                "Yes, enable OCR",
                "No, skip OCR")
            If ocrChoice = 0 Then
                Return ' User aborted
            End If
            doOcr = (ocrChoice = 1)
        End If

        ' Per-conversion override: never mutate the user's persistent model configuration.
        Dim conversionOcrBatchPages As System.Int32 = -1
        If doOcr AndAlso pdfCount > 0 Then
            Dim chunkChoice As System.Int32 = ShowCustomYesNoBox(
                "OCR normally uses ChunkOCR=" & _context.INI_ChunkOCR.ToString(System.Globalization.CultureInfo.InvariantCulture) & "." & vbCrLf & vbCrLf &
                "Choose a temporary page batch size for all PDFs in this conversion? 0 means no chunking; failed batches are automatically reduced.",
                "Choose for this conversion", "Use configured ChunkOCR")
            If chunkChoice = 0 Then Return
            If chunkChoice = 1 Then
                Dim validChunkSize As System.Boolean = False
                Do
                    Dim input As System.String = ShowCustomInputBox("Pages per OCR request (0 = no chunking). Applies to all PDFs in this conversion only:", AN, True,
                        System.Math.Max(0, _context.INI_ChunkOCR).ToString(System.Globalization.CultureInfo.InvariantCulture))
                    If System.String.IsNullOrWhiteSpace(input) Then Return
                    validChunkSize = System.Int32.TryParse(input, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, conversionOcrBatchPages) AndAlso conversionOcrBatchPages >= 0
                    If Not validChunkSize Then ShowCustomMessageBox("Enter a whole number greater than or equal to zero.")
                Loop Until validChunkSize
            End If
        End If

        ' Offer PDF flattening before OCR when OCR is enabled
        If doOcr AndAlso pdfCount >= 1 Then
            Dim flattenChoice As Integer = ShowCustomYesNoBox(
                "Some scanned PDFs contain layered images or transparency groups that OCR may miss." & vbCrLf & vbCrLf &
                "Would you like to flatten PDFs to image-only format before OCR? " &
                "This rasterizes each page (like printing to image), ensuring all visual elements are captured." & vbCrLf & vbCrLf &
                "⚠ This adds processing time (a few seconds per page) but can significantly improve OCR results for complex PDFs.",
                "Yes, flatten first (recommended for scans)",
                "No, OCR directly")
            If flattenChoice = 0 Then
                Return ' User aborted
            End If
            flattenBeforeOcr = (flattenChoice = 1)
        End If

        If markdownEligibleFiles.Count > 0 Then
            Dim markdownChoice As Integer = ShowCustomYesNoBox(
                "You can have your Word documents (.docx) and PDF files converted to Markdown, preserving headings, lists and tables where possible. Readable PDFs can be converted directly; scanned PDFs use the OCR choice above." & vbCrLf & vbCrLf &
                "If the quality is not what you need, you may want to try to have your files converted to PDF, flattened to an image and OCR'd by your model, if it supports so. You can use this conversion helper to process the PDF." & vbCrLf & vbCrLf &
                "Do you want supported files to be saved as Markdown (.md) instead of plain text (.txt)?",
                "Yes, use Markdown where supported",
                "No, keep plain text")
            If markdownChoice = 0 Then
                Return ' User aborted
            End If

            useMarkdownOutputWhenSupported = (markdownChoice = 1)

            If useMarkdownOutputWhenSupported Then
                ocrMarkdownInstruction = Add_OcrMarkdownInstruction
            End If
        End If

        If useMarkdownOutputWhenSupported AndAlso pdfCount > 0 Then
            markdownPreparation = SelectPdfMarkdownPreparationOptions(ocrCleanupMode)
            If markdownPreparation Is Nothing Then Return
        End If

        ' Create output subdirectory if needed
        Dim outputBaseDir As String = ""
        If useSubdirectory Then
            If isDirectory Then
                outputBaseDir = IO.Path.Combine(selectedPath, outputSubdirName)
            Else
                outputBaseDir = IO.Path.Combine(IO.Path.GetDirectoryName(selectedPath), outputSubdirName)
            End If

            Try
                If Not IO.Directory.Exists(outputBaseDir) Then
                    IO.Directory.CreateDirectory(outputBaseDir)
                End If
            Catch ex As System.Exception
                ShowCustomMessageBox($"Failed to create output directory: {ex.Message}")
                Return
            End Try
        End If

        ' Show progress bar
        ProgressBarModule.GlobalProgressValue = 0
        ProgressBarModule.GlobalProgressMax = filesToProcess.Count
        ProgressBarModule.GlobalProgressLabel = "Initializing..."
        ProgressBarModule.CancelOperation = False
        ProgressBarModule.ShowProgressBarInSeparateThread(AN & " Convert to Text", "Starting text extraction...")

        ' Process files
        Dim successCount As Integer = 0
        Dim failedFiles As New System.Collections.Generic.List(Of String)()
        Dim emptyContentFiles As New System.Collections.Generic.List(Of String)()
        Dim flattenedPdfCount As Integer = 0
        Dim markdownOutputCount As Integer = 0
        Dim markdownPdfCount As Integer = 0
        Dim markdownDocxCount As Integer = 0
        ' Collected successfully-extracted content for the optional combine/index step.
        Dim extractedDocuments As New System.Collections.Generic.List(Of (Name As String, Content As String))()
        Dim singleExtractedOutputPath As String = Nothing

        Try
            For i As Integer = 0 To filesToProcess.Count - 1
                ' Check for cancellation
                If ProgressBarModule.CancelOperation Then
                    Exit For
                End If

                Dim filePath As String = filesToProcess(i)
                Dim fileName As String = IO.Path.GetFileName(filePath)

                ' Update progress
                ProgressBarModule.GlobalProgressValue = i
                ProgressBarModule.GlobalProgressLabel = $"Processing file {i + 1} of {filesToProcess.Count}: {fileName}"

                Try
                    ' Determine OCR and Markdown settings for this file
                    Dim fileExtension As String = IO.Path.GetExtension(filePath).ToLowerInvariant()
                    Dim isPdf As Boolean = fileExtension = ".pdf"
                    Dim isDocx As Boolean = fileExtension = ".docx"
                    Dim useOcrForThisFile As Boolean = isPdf AndAlso doOcr
                    Dim useMarkdownForThisFile As Boolean =
                        useMarkdownOutputWhenSupported AndAlso (isPdf OrElse isDocx)
                    Dim outputExtension As String = If(useMarkdownForThisFile, ".md", ".txt")

                    ' If flattening is requested for PDFs, flatten to temp file first
                    Dim effectiveFilePath As String = filePath
                    Dim tempFlattenedPath As String = Nothing

                    If isPdf AndAlso flattenBeforeOcr AndAlso useOcrForThisFile Then
                        Try
                            ProgressBarModule.GlobalProgressLabel = $"Flattening PDF {i + 1} of {filesToProcess.Count}: {fileName}..."
                            tempFlattenedPath = IO.Path.Combine(IO.Path.GetTempPath(), $"{AN2}_flatten_{Guid.NewGuid():N}.pdf")
                            Await System.Threading.Tasks.Task.Run(Sub() FlattenPdfToImageOnly(filePath, tempFlattenedPath, 200))
                            effectiveFilePath = tempFlattenedPath
                            flattenedPdfCount += 1
                            ProgressBarModule.GlobalProgressLabel = $"OCR processing file {i + 1} of {filesToProcess.Count}: {fileName}..."
                        Catch ex As System.Exception
                            conversionDiagnostics.AppendLine("=== " & fileName & " / PDF flattening ===")
                            conversionDiagnostics.AppendLine("Using original PDF after flattening failure: " & ex.ToString())
                            conversionDiagnostics.AppendLine()
                            ' Flattening failed — fall back to original PDF
                            effectiveFilePath = filePath
                            If tempFlattenedPath IsNot Nothing Then
                                Try : If IO.File.Exists(tempFlattenedPath) Then IO.File.Delete(tempFlattenedPath)
                                Catch : End Try
                            End If
                            ' Retain the path so the final cleanup can retry if deletion failed.
                        End Try
                    End If

                    Try
                        If isPdf AndAlso useMarkdownForThisFile AndAlso markdownPreparation IsNot Nothing Then
                            markdownPreparation.ShowProgressWindow = True
                            markdownPreparation.SourceName = "File " & (i + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) & " / " & filesToProcess.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) & ": " & fileName
                        End If
                        ' Read file content — askUser=False ensures Hidesplash=True in PerformOCR/ReadBinaryFileViaLLM,
                        ' preventing the countdown splash from interrupting batch processing.
                        ' The user has already made all choices upfront (OCR on/off, flatten on/off, Markdown on/off).
                        Dim fileRead As FileReadResult = Await GetFileContentEx(
                            effectiveFilePath,
                            Silent:=True,
                            DoOCR:=useOcrForThisFile,
                            AskUser:=False,
                            OcrAdditionalInstruction:=If(useMarkdownForThisFile AndAlso useOcrForThisFile, ocrMarkdownInstruction, ""),
                            ShowOCRProgress:=True,
                            ReturnMarkdown:=useMarkdownForThisFile,
                            OcrCleanupMode:=If(useMarkdownForThisFile AndAlso useOcrForThisFile, ocrCleanupMode, OcrMarkdownCleanupMode.None),
                            OcrBatchPages:=conversionOcrBatchPages,
                            MarkdownPreparation:=If(isPdf AndAlso useMarkdownForThisFile, markdownPreparation, Nothing))
                        ' Capture diagnostics before empty/error Continue For paths. No report side files.
                        If fileRead.ExtractionWarnings.Count > 0 OrElse Not System.String.IsNullOrWhiteSpace(fileRead.OcrCleanupReport) OrElse fileRead.MarkdownPreparationSkipped Then
                            conversionDiagnostics.AppendLine("=== " & fileName & " / OCR and Markdown ===")
                            For Each warning As System.String In fileRead.ExtractionWarnings
                                conversionDiagnostics.AppendLine(warning)
                            Next
                            If Not System.String.IsNullOrWhiteSpace(fileRead.OcrCleanupReport) Then conversionDiagnostics.AppendLine(fileRead.OcrCleanupReport)
                            If fileRead.MarkdownPreparationSkipped Then conversionDiagnostics.AppendLine("Markdown preparation skipped; original extraction retained.")
                            conversionDiagnostics.AppendLine()
                        End If
                        If fileRead.MarkdownPreparationSkipped Then markdownPreparationSkippedCount += 1
                        Dim content As System.String = fileRead.Content

                        If String.IsNullOrWhiteSpace(content) Then
                            emptyContentFiles.Add($"{fileName} ({IO.Path.GetExtension(filePath).ToLowerInvariant()})")
                            Continue For
                        End If

                        If content.StartsWith("Error", System.StringComparison.OrdinalIgnoreCase) AndAlso content.Length < 200 Then
                            failedFiles.Add($"{fileName}: {content}")
                            Continue For
                        End If

                        ' Determine output path
                        Dim outputPath As String
                        If useSubdirectory Then
                            ' Preserve relative directory structure if processing subdirectories
                            If isDirectory Then
                                Dim relativePath As String = filePath.Substring(selectedPath.Length).TrimStart(IO.Path.DirectorySeparatorChar)
                                Dim relativeDir As String = IO.Path.GetDirectoryName(relativePath)
                                If Not String.IsNullOrEmpty(relativeDir) Then
                                    Dim targetDir As String = IO.Path.Combine(outputBaseDir, relativeDir)
                                    If Not IO.Directory.Exists(targetDir) Then
                                        IO.Directory.CreateDirectory(targetDir)
                                    End If
                                End If
                                outputPath = IO.Path.Combine(outputBaseDir, relativePath & outputExtension)
                            Else
                                outputPath = IO.Path.Combine(outputBaseDir, fileName & outputExtension)
                            End If
                        Else
                            outputPath = filePath & outputExtension
                        End If

                        IO.File.WriteAllText(outputPath, content, System.Text.Encoding.UTF8)
                        If Not System.String.IsNullOrEmpty(fileRead.OcrCleanupReport) Then
                            ' The report was captured above for the conversion clipboard log.
                            If markdownPreparation IsNot Nothing AndAlso markdownPreparation.SaveRawCopy AndAlso
                               Not System.String.Equals(fileRead.OcrCleanupRawContent, content, System.StringComparison.Ordinal) Then
                                Dim rawPath As System.String = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(outputPath),
                                    System.IO.Path.GetFileNameWithoutExtension(outputPath) & ".unbereinigt.md")
                                Dim rawBase As System.String = System.IO.Path.GetFileNameWithoutExtension(rawPath)
                                Dim suffix As System.Int32 = 2
                                While System.IO.File.Exists(rawPath)
                                    rawPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(outputPath), rawBase & " (" & suffix.ToString(System.Globalization.CultureInfo.InvariantCulture) & ").md")
                                    suffix += 1
                                End While
                                System.IO.File.WriteAllText(rawPath, fileRead.OcrCleanupRawContent, New System.Text.UTF8Encoding(False))
                                conversionDiagnostics.AppendLine("Uncleaned Markdown copy saved: " & rawPath)
                            End If
                        End If
                        successCount += 1

                        If useMarkdownForThisFile Then
                            markdownOutputCount += 1
                            If isPdf Then
                                markdownPdfCount += 1
                            ElseIf isDocx Then
                                markdownDocxCount += 1
                            End If
                        End If

                        extractedDocuments.Add((fileName, content))
                        If successCount = 1 Then
                            singleExtractedOutputPath = outputPath
                        End If

                    Finally
                        ' Clean up on success, failure and Continue For paths.
                        If tempFlattenedPath IsNot Nothing Then
                            Try
                                If System.IO.File.Exists(tempFlattenedPath) Then System.IO.File.Delete(tempFlattenedPath)
                            Catch cleanupFailure As System.Exception
                                conversionDiagnostics.AppendLine("=== " & fileName & " / temporary PDF cleanup ===")
                                conversionDiagnostics.AppendLine(cleanupFailure.ToString())
                                failedFiles.Add(fileName & ": could not remove temporary working PDF: " & tempFlattenedPath)
                            End Try
                        End If
                    End Try

                Catch ex As System.OperationCanceledException
                    conversionDiagnostics.AppendLine("=== " & fileName & " / conversion cancelled ===")
                    conversionDiagnostics.AppendLine("Cancelled by the user; no partially prepared Markdown was saved.")
                    Exit For
                Catch ex As System.Exception
                    failedFiles.Add($"{fileName}: {ex.Message}")
                    conversionDiagnostics.AppendLine("=== " & fileName & " / conversion exception ===")
                    conversionDiagnostics.AppendLine(ex.ToString())
                    conversionDiagnostics.AppendLine()
                End Try
            Next

        Finally
            ' Close progress bar
            ProgressBarModule.CancelOperation = True
        End Try

        ' Optional: combine multiple extracted documents into a single file and/or create a
        ' semantic-search index from them.
        Dim combinedOutputPath As String = Nothing
        Dim indexOutputPath As String = Nothing

        If extractedDocuments.Count > 1 Then
            Dim combineChoice As Integer = ShowCustomYesNoBox(
                $"{extractedDocuments.Count:N0} documents were extracted." & vbCrLf & vbCrLf &
                "Do you want to combine them into a single file using the standard document separators? " &
                $"This is useful for feeding all documents to the AI at once or for creating a searchable index for {AN} (you can choose to create one as the next step).",
                "Yes, combine into one file",
                "No, keep separate files")

            If combineChoice = 1 Then
                Try
                    ' Combine using the canonical wrapper separator so the semantic-search generator
                    ' can align segments to document boundaries.
                    Dim combinedBuilder As New System.Text.StringBuilder()
                    For docIndex As Integer = 0 To extractedDocuments.Count - 1
                        Dim docNumber As Integer = docIndex + 1
                        Dim safeName As String = If(extractedDocuments(docIndex).Name, "").Replace("""", "'")
                        combinedBuilder.Append($"<document{docNumber} name=""{safeName}"">").AppendLine()
                        combinedBuilder.Append(If(extractedDocuments(docIndex).Content, "")).AppendLine()
                        combinedBuilder.Append($"</document{docNumber}>").AppendLine()
                        combinedBuilder.AppendLine()
                    Next

                    Dim combineBaseDir As String =
                        If(useSubdirectory,
                           outputBaseDir,
                           If(isDirectory, selectedPath, IO.Path.GetDirectoryName(selectedPath)))

                    Dim combineBaseName As String =
                        If(isDirectory, IO.Path.GetFileName(selectedPath.TrimEnd(IO.Path.DirectorySeparatorChar)), IO.Path.GetFileNameWithoutExtension(selectedPath))
                    If String.IsNullOrWhiteSpace(combineBaseName) Then
                        combineBaseName = "combined"
                    End If

                    combinedOutputPath = IO.Path.Combine(combineBaseDir, combineBaseName & ".combined.txt")
                    IO.File.WriteAllText(combinedOutputPath, combinedBuilder.ToString().TrimEnd(), System.Text.Encoding.UTF8)

                    ' Offer to also create a semantic-search index from the combined file.
                    Dim indexChoice As Integer = ShowCustomYesNoBox(
                        "The documents were combined into a single file." & vbCrLf & vbCrLf &
                        "Do you also want to create a semantic-search index file from it? " &
                        $"This allows {AN} to find and retrieve content from the file even if it is too large to be sent to the AI at once. 'Discuss this, Inky' and the chatbot in Word supports these index files, for example.",
                        "Yes, create an index file",
                        "No, just keep the combined file")

                    If indexChoice = 1 Then
                        indexOutputPath = Await CreateSemanticSearchIndexWithUserOptionsAsync(
                            combinedOutputPath,
                            IO.Path.Combine(combineBaseDir, combineBaseName & ".indexed.txt"),
                            "Preparing combined content ...",
                            indexMetadataWarnings).ConfigureAwait(True)
                    End If

                Catch ex As System.Exception
                    ShowCustomMessageBox($"Failed to combine or index the extracted documents: {ex.Message}")
                    combinedOutputPath = Nothing
                    indexOutputPath = Nothing
                End Try
            End If
        End If

        If String.IsNullOrWhiteSpace(indexOutputPath) AndAlso
           successCount = 1 AndAlso
           Not String.IsNullOrWhiteSpace(singleExtractedOutputPath) Then

            Dim singleFileIndexChoice As Integer = ShowCustomYesNoBox(
                "One file was successfully extracted." & vbCrLf & vbCrLf &
                "Do you also want to create a semantic-search index file from it? " &
                $"This allows {AN} to find and retrieve content from the file even if it is too large to be sent to the AI at once. 'Discuss this, Inky' and the chatbot in Word supports these index files, for example.",
                "Yes, create an index file",
                "No, keep the extracted file only")

            If singleFileIndexChoice = 1 Then
                Try
                    indexOutputPath = Await CreateSemanticSearchIndexWithUserOptionsAsync(
                        singleExtractedOutputPath,
                        IO.Path.Combine(
                            IO.Path.GetDirectoryName(singleExtractedOutputPath),
                            IO.Path.GetFileNameWithoutExtension(singleExtractedOutputPath) & ".indexed.txt"),
                        "Preparing extracted content ...",
                        indexMetadataWarnings).ConfigureAwait(True)
                Catch ex As System.Exception
                    ShowCustomMessageBox($"Failed to create the semantic-search index: {ex.Message}")
                    indexOutputPath = Nothing
                End Try
            End If
        End If

        ' A source-derived fallback is deliberately visible: the source is fully
        ' indexed, but affected segments have less precise semantic metadata.
        If indexMetadataWarnings.Count > 0 Then
            conversionDiagnostics.AppendLine("=== SEMANTIC INDEX METADATA FALLBACKS ===")
            conversionDiagnostics.AppendLine("The source text and byte offsets were preserved. The following segments used source-derived title/summary after metadata-generation failures:")
            For Each warningText As System.String In indexMetadataWarnings
                conversionDiagnostics.AppendLine(warningText)
            Next
            conversionDiagnostics.AppendLine()
        End If

        ' Build summary
        Dim wasCancelled As Boolean = (successCount + failedFiles.Count + emptyContentFiles.Count) < filesToProcess.Count

        Dim summary As New System.Text.StringBuilder()

        If wasCancelled Then
            summary.AppendLine("Operation was cancelled by user.")
            summary.AppendLine()
        End If

        summary.AppendLine($"Successfully converted: {successCount} file(s)")

        If useSubdirectory Then
            summary.AppendLine($"Output directory: {outputBaseDir}")
        End If

        If Not String.IsNullOrWhiteSpace(combinedOutputPath) Then
            summary.AppendLine()
            summary.AppendLine($"Combined file: {combinedOutputPath}")
            If Not String.IsNullOrWhiteSpace(indexOutputPath) Then
                summary.AppendLine($"Index file: {indexOutputPath}")
            End If
        ElseIf Not String.IsNullOrWhiteSpace(indexOutputPath) Then
            summary.AppendLine()
            summary.AppendLine($"Index file: {indexOutputPath}")
        End If

        If indexMetadataWarnings.Count > 0 Then
            summary.AppendLine()
            summary.AppendLine("Warning: " & indexMetadataWarnings.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                " index segment(s) use source-derived metadata after LLM failures; the complete original text remains indexed.")
            summary.AppendLine("The affected segment IDs and failure categories are in the clipboard log and the index file.")
        End If

        If doOcr Then
            summary.AppendLine("OCR was enabled for PDF files.")
            If flattenBeforeOcr Then
                summary.AppendLine($"PDFs were flattened to images before OCR ({flattenedPdfCount} file(s)).")
            End If
        End If

        If useMarkdownOutputWhenSupported Then
            summary.AppendLine($"Markdown output was used for {markdownOutputCount} supported file(s).")
            If markdownDocxCount > 0 Then
                summary.AppendLine($"  • DOCX → Markdown: {markdownDocxCount}")
            End If
            If markdownPdfCount > 0 Then
                summary.AppendLine($"  • PDF → Markdown: {markdownPdfCount}")
            End If
        End If

        If emptyContentFiles.Count > 0 Then
            summary.AppendLine()
            summary.AppendLine($"⚠ Files with no extractable content ({emptyContentFiles.Count} file(s)):")
            For Each f In emptyContentFiles.Take(10)
                summary.AppendLine($"  • {f}")
            Next
            If emptyContentFiles.Count > 10 Then
                summary.AppendLine($"  ... and {emptyContentFiles.Count - 10} more")
            End If
            summary.AppendLine("  (These files were read but returned no text content)")
        End If

        If unsupportedFiles.Count > 0 Then
            summary.AppendLine()
            summary.AppendLine($"Skipped due to unsupported file type ({unsupportedFiles.Count} file(s)):")
            ' Group by extension for readability
            Dim byExtension = unsupportedFiles.
                GroupBy(Function(f) IO.Path.GetExtension(f).ToLowerInvariant()).
                OrderByDescending(Function(g) g.Count())
            For Each extGroup In byExtension
                Dim extLabel As String = If(String.IsNullOrEmpty(extGroup.Key), "(no extension)", extGroup.Key)
                If extGroup.Count() <= 3 Then
                    For Each f In extGroup
                        summary.AppendLine($"  • {IO.Path.GetFileName(f)}")
                    Next
                Else
                    summary.AppendLine($"  • {extGroup.Count()} {extLabel} files (e.g., {IO.Path.GetFileName(extGroup.First())})")
                End If
            Next
        End If

        If legacyDocFilesSkipped.Count > 0 Then
            summary.AppendLine()
            summary.AppendLine($"Skipped .doc files ({legacyDocFilesSkipped.Count} — disabled for security):")
            For Each f In legacyDocFilesSkipped.Take(5)
                summary.AppendLine($"  • {IO.Path.GetFileName(f)}")
            Next
            If legacyDocFilesSkipped.Count > 5 Then
                summary.AppendLine($"  ... and {legacyDocFilesSkipped.Count - 5} more")
            End If
        End If

        ' Build detailed failure/skip log for clipboard
        Dim clipboardLog As New System.Text.StringBuilder()

        If failedFiles.Count > 0 Then
            summary.AppendLine()
            summary.AppendLine($"Failed: {failedFiles.Count} file(s)")
            For Each f In failedFiles.Take(10)
                summary.AppendLine($"  • {f}")
            Next
            If failedFiles.Count > 10 Then
                summary.AppendLine($"  ... and {failedFiles.Count - 10} more")
            End If

            clipboardLog.AppendLine("=== FAILED FILES ===")
            For Each f In failedFiles
                clipboardLog.AppendLine(f)
            Next
            clipboardLog.AppendLine()
        End If

        If markdownPreparationSkippedCount > 0 Then
            summary.AppendLine("Markdown preparation was skipped for " & markdownPreparationSkippedCount.ToString(System.Globalization.CultureInfo.InvariantCulture) & " file(s); original extraction retained.")
            summary.AppendLine("See the OCR/Markdown details in the log copied to the clipboard.")
        End If

        If emptyContentFiles.Count > 0 Then
            clipboardLog.AppendLine("=== EMPTY CONTENT FILES ===")
            For Each f In emptyContentFiles
                clipboardLog.AppendLine(f)
            Next
            clipboardLog.AppendLine()
        End If

        If unsupportedFiles.Count > 0 Then
            clipboardLog.AppendLine("=== UNSUPPORTED FILES (Ignored) ===")
            For Each f In unsupportedFiles
                clipboardLog.AppendLine(IO.Path.GetFileName(f))
            Next
            clipboardLog.AppendLine()
        End If

        If legacyDocFilesSkipped.Count > 0 Then
            clipboardLog.AppendLine("=== LEGACY .DOC FILES (Disabled for security) ===")
            For Each f In legacyDocFilesSkipped
                clipboardLog.AppendLine(IO.Path.GetFileName(f))
            Next
            clipboardLog.AppendLine()
        End If

        If conversionDiagnostics.Length > 0 Then
            clipboardLog.AppendLine("=== OCR / MARKDOWN DIAGNOSTICS ===")
            clipboardLog.Append(conversionDiagnostics.ToString())
        End If

        ' Conversion diagnostics belong to this clipboard report, independently of usage logging.
        If failedFiles.Count > 0 OrElse emptyContentFiles.Count > 0 OrElse unsupportedFiles.Count > 0 OrElse legacyDocFilesSkipped.Count > 0 OrElse conversionDiagnostics.Length > 0 Then
            Dim logText As String = clipboardLog.ToString().TrimEnd()
            SharedMethods.PutInClipboard(logText)
            summary.AppendLine()
            summary.AppendLine("(Detailed conversion log copied to clipboard; paste it into a text editor with Ctrl+V.)")
        End If

        ShowCustomMessageBox(summary.ToString().TrimEnd(), AN & " Convert to Text")
    End Sub

    Private Async Function CreateSemanticSearchIndexWithUserOptionsAsync(
        sourcePath As String,
        defaultOutputPath As String,
        preparationMessage As String,
        metadataWarnings As System.Collections.Generic.List(Of System.String)) As System.Threading.Tasks.Task(Of String)

        If String.IsNullOrWhiteSpace(sourcePath) OrElse
           String.IsNullOrWhiteSpace(defaultOutputPath) Then
            Return Nothing
        End If

        Dim selectedProfile As SharedMethods.SemanticSearchMetadataProfile =
            SharedMethods.SemanticSearchMetadataProfile.Generic
        Dim targetBytes As Integer = SharedMethods.SemanticSearchDefaultTargetBytes
        Dim minimumBytes As Integer = SharedMethods.SemanticSearchDefaultMinimumBytes
        Dim maximumBytes As Integer = SharedMethods.SemanticSearchDefaultMaximumBytes
        Dim overwriteOutput As Boolean = False
        Dim outputPath As String = defaultOutputPath

        Dim availableProfiles As System.Collections.Generic.List(Of SharedMethods.SemanticSearchMetadataProfile) =
            SharedMethods.GetSemanticSearchMetadataProfiles()
        Dim profileDisplayMap As New System.Collections.Generic.Dictionary(Of String, SharedMethods.SemanticSearchMetadataProfile)(
            System.StringComparer.Ordinal)
        Dim profileDisplayOptions As New System.Collections.Generic.List(Of String)()

        For Each profile As SharedMethods.SemanticSearchMetadataProfile In availableProfiles
            Dim displayName As String = SharedMethods.GetSemanticSearchMetadataProfileDisplayName(profile)
            If Not profileDisplayMap.ContainsKey(displayName) Then
                profileDisplayMap.Add(displayName, profile)
                profileDisplayOptions.Add(displayName)
            End If
        Next

        Dim selectedProfileDisplay As String =
            SharedMethods.GetSemanticSearchMetadataProfileDisplayName(selectedProfile)

        Dim generationParameters() As Slib.InputParameter = {
            New Slib.InputParameter("Profile", selectedProfileDisplay, profileDisplayOptions),
            New Slib.InputParameter("Target bytes", targetBytes),
            New Slib.InputParameter("Minimum bytes", minimumBytes),
            New Slib.InputParameter("Maximum bytes", maximumBytes)
        }

        Dim generationPrompt As String =
            "Please configure the semantic-search index generation settings." & vbCrLf & vbCrLf &
            "All size values are UTF-8 byte counts, not character counts." & vbCrLf & vbCrLf &
            "Profile: chooses what kind of metadata the indexer should emphasize for each segment." & vbCrLf &
            "Target bytes: the preferred segment size. The generator tries to end each segment near this size." & vbCrLf &
            "Minimum bytes: the lower bound before the generator starts looking for a natural break point." & vbCrLf &
            "Maximum bytes: the hard segment-size ceiling. A segment will not intentionally exceed this size."

        If Not ShowCustomVariableInputForm(
            generationPrompt,
            $"{AN} Semantic-search index",
            generationParameters) Then
            Return Nothing
        End If

        selectedProfileDisplay = CStr(generationParameters(0).Value)
        targetBytes = CInt(generationParameters(1).Value)
        minimumBytes = CInt(generationParameters(2).Value)
        maximumBytes = CInt(generationParameters(3).Value)

        If Not profileDisplayMap.TryGetValue(selectedProfileDisplay, selectedProfile) Then
            Throw New System.InvalidOperationException("The selected semantic-search profile could not be resolved.")
        End If

        If targetBytes <= 0 OrElse minimumBytes <= 0 OrElse maximumBytes <= 0 Then
            Throw New System.ArgumentException("Target bytes, minimum bytes, and maximum bytes must all be greater than zero.")
        End If

        If minimumBytes > targetBytes Then
            Throw New System.ArgumentException("Minimum bytes must be less than or equal to target bytes.")
        End If

        If targetBytes > maximumBytes Then
            Throw New System.ArgumentException("Target bytes must be less than or equal to maximum bytes.")
        End If

        If IO.File.Exists(outputPath) Then
            Dim overwriteAnswer As Integer = ShowCustomYesNoBox(
                "The target output file already exists:" & vbCrLf & vbCrLf &
                outputPath & vbCrLf & vbCrLf &
                "Do you want to overwrite it?",
                "Yes, overwrite",
                "No, cancel")

            If overwriteAnswer <> 1 Then
                Return Nothing
            End If

            overwriteOutput = True
        End If

        Using indexProgress As New ProgressScope(
            "Generating semantic-search index",
            preparationMessage,
            1)

            Dim generationProgress As New System.Progress(Of SharedMethods.SemanticSearchIndexGenerationProgress)(
                Sub(update As SharedMethods.SemanticSearchIndexGenerationProgress)
                    Dim segCount As Integer = System.Math.Max(1, update.SegmentCount)
                    Dim segNumber As Integer = System.Math.Max(0, System.Math.Min(update.SegmentNumber, segCount))
                    Dim statusMessage As String = If(update.Message, "").Trim()

                    If String.IsNullOrWhiteSpace(statusMessage) Then
                        statusMessage = "Generating semantic metadata"
                    End If

                    ProgressScope.Report(
                        segNumber,
                        segCount,
                        statusMessage & " (" &
                        segNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) & "/" &
                        segCount.ToString(System.Globalization.CultureInfo.InvariantCulture) & ")")
                End Sub)

            Dim generationResult As SharedMethods.SemanticSearchIndexGenerationResult =
                Await SharedMethods.CreateSemanticSearchIndexedTextFileAsync(
                sourcePath,
                outputPath,
                _context,
                New SharedMethods.SemanticSearchIndexGeneratorOptions() With {
                    .TargetBytes = targetBytes,
                    .MinimumBytes = minimumBytes,
                    .MaximumBytes = maximumBytes,
                    .SpecialTaskName = "Indexer",
                    .MetadataProfile = selectedProfile,
                    .OverwriteOutput = overwriteOutput
                },
                generationProgress,
                indexProgress.Token).ConfigureAwait(True)

            If metadataWarnings IsNot Nothing Then
                ' Project boundary: keep the structured warning type in SharedLibrary and
                ' project it into a host diagnostic line only for UI logging.
                For Each warning In generationResult.MetadataGenerationWarnings
                    metadataWarnings.Add(warning.SegmentId & " [" & warning.ErrorType & "]: " & warning.Message)
                Next
            End If
        End Using

        Return outputPath
    End Function


    Private Function SelectOcrMarkdownCleanupMode(includeDisabled As System.Boolean) As SharedLibrary.SharedLibrary.SharedMethods.OcrMarkdownCleanupMode
        Dim items As New System.Collections.Generic.List(Of SharedLibrary.SharedLibrary.SharedMethods.SelectionItem)()
        If includeDisabled Then items.Add(New SharedLibrary.SharedLibrary.SharedMethods.SelectionItem("Keep OCR output unchanged (default)", 1))
        items.Add(New SharedLibrary.SharedLibrary.SharedMethods.SelectionItem("Join wrapped prose lines; keep all text", 2))
        items.Add(New SharedLibrary.SharedLibrary.SharedMethods.SelectionItem("Also remove exact repeated page headers; log the cleanup report", 3))
        Dim selected As System.Int32 = SharedLibrary.SharedLibrary.SharedMethods.SelectValue(items, If(includeDisabled, 1, 2),
            "Optional OCR cleanup. Numbers, footers, signatures, code, footnote identifiers and ambiguous hyphens are retained. Header removal needs actual page boundaries (form feeds in a text file).",
            AN & " OCR cleanup")
        Select Case selected
            Case 2
                Return SharedLibrary.SharedLibrary.SharedMethods.OcrMarkdownCleanupMode.JoinWrappedLines
            Case 3
                Return SharedLibrary.SharedLibrary.SharedMethods.OcrMarkdownCleanupMode.JoinWrappedLinesAndRepeatedHeaders
            Case Else
                Return SharedLibrary.SharedLibrary.SharedMethods.OcrMarkdownCleanupMode.None
        End Select
    End Function

    Private Function SelectPdfMarkdownPreparationOptions(ByRef legacyMode As SharedLibrary.SharedLibrary.SharedMethods.OcrMarkdownCleanupMode) As SharedLibrary.SharedLibrary.SharedMethods.PdfMarkdownPreparationOptions
        legacyMode = SharedLibrary.SharedLibrary.SharedMethods.OcrMarkdownCleanupMode.None
        Dim items As New System.Collections.Generic.List(Of SharedLibrary.SharedLibrary.SharedMethods.SelectionItem) From {
            New SharedLibrary.SharedLibrary.SharedMethods.SelectionItem("Keep extracted Markdown unchanged (default)", 1),
            New SharedLibrary.SharedLibrary.SharedMethods.SelectionItem("Reading copy: prose flow and repeated page artifacts", 2),
            New SharedLibrary.SharedLibrary.SharedMethods.SelectionItem("Document structure: footnotes, margins, headings and page continuations", 3),
            New SharedLibrary.SharedLibrary.SharedMethods.SelectionItem("Existing conservative OCR cleanup", 4)}
        Dim choice As System.Int32 = SharedLibrary.SharedLibrary.SharedMethods.SelectValue(items, 1,
            "Choose Markdown preparation for the PDFs in this conversion. Structure preparation uses additional model requests; the original wording is retained. This choice does not change your configuration.", AN & " PDF to Markdown")
        If choice = 0 Then Return Nothing
        Dim options As New SharedLibrary.SharedLibrary.SharedMethods.PdfMarkdownPreparationOptions()
        If choice = 1 Then Return options
        If choice = 4 Then
            legacyMode = SelectOcrMarkdownCleanupMode(True)
            Dim rawParameters As SharedLibrary.SharedLibrary.SharedMethods.InputParameter() = {
                New SharedLibrary.SharedLibrary.SharedMethods.InputParameter("Also save an uncleaned .md copy", False)}
            If Not SharedLibrary.SharedLibrary.SharedMethods.ShowCustomVariableInputForm("The cleanup report is copied to the clipboard at the end. Only the final Markdown is saved unless selected below.", AN & " OCR cleanup", rawParameters) Then Return Nothing
            options.SaveRawCopy = System.Convert.ToBoolean(rawParameters(0).Value, System.Globalization.CultureInfo.InvariantCulture)
            Return options
        End If
        Dim structured As System.Boolean = choice = 3
        Dim parameters As SharedLibrary.SharedLibrary.SharedMethods.InputParameter() = {
            New SharedLibrary.SharedLibrary.SharedMethods.InputParameter("Join wrapped prose lines (preserve code, lists, tables and explicit breaks)", True),
            New SharedLibrary.SharedLibrary.SharedMethods.InputParameter("Remove repeated page-edge headers/footers and verified page-number sequences", True),
            New SharedLibrary.SharedLibrary.SharedMethods.InputParameter("Collect all identified footnotes at the document end (Markdown notes)", structured),
            New SharedLibrary.SharedLibrary.SharedMethods.InputParameter("Represent marginal numbers/text consistently", structured),
            New SharedLibrary.SharedLibrary.SharedMethods.InputParameter("Reconcile heading levels across the entire document", structured),
            New SharedLibrary.SharedLibrary.SharedMethods.InputParameter("Join confirmed paragraph continuations across pages", structured),
            New SharedLibrary.SharedLibrary.SharedMethods.InputParameter("Use model structure analysis (additional requests; needed for structural inference)", structured),
            New SharedLibrary.SharedLibrary.SharedMethods.InputParameter("Also save an uncleaned .md copy", False)}
        If Not SharedLibrary.SharedLibrary.SharedMethods.ShowCustomVariableInputForm(
            "These options apply to all PDFs in this conversion. Uncertain or invalid transformations retain the source and are reported in the detailed log copied to the clipboard at the end (Ctrl+V to paste). All identified footnotes are collected when selected; unmatched references are retained and logged. Printed OCR footnotes need model structure analysis; existing Markdown notes do not. Without model analysis, only existing footnotes, prose lines and repeated page-edge patterns are processed. For PDFs read without OCR, the existing layout reader uses source geometry; the later Markdown pass preserves its tables and notes.",
            AN & " Markdown preparation", parameters, checkboxRightClearance:=24) Then Return Nothing
        options.JoinProseLines = System.Convert.ToBoolean(parameters(0).Value, System.Globalization.CultureInfo.InvariantCulture)
        options.RemovePageArtifacts = System.Convert.ToBoolean(parameters(1).Value, System.Globalization.CultureInfo.InvariantCulture)
        options.CollectFootnotes = System.Convert.ToBoolean(parameters(2).Value, System.Globalization.CultureInfo.InvariantCulture)
        options.NormalizeMarginNotes = System.Convert.ToBoolean(parameters(3).Value, System.Globalization.CultureInfo.InvariantCulture)
        options.NormalizeHeadings = System.Convert.ToBoolean(parameters(4).Value, System.Globalization.CultureInfo.InvariantCulture)
        options.JoinPageParagraphs = System.Convert.ToBoolean(parameters(5).Value, System.Globalization.CultureInfo.InvariantCulture)
        options.UseModelStructure = System.Convert.ToBoolean(parameters(6).Value, System.Globalization.CultureInfo.InvariantCulture)
        options.SaveRawCopy = System.Convert.ToBoolean(parameters(7).Value, System.Globalization.CultureInfo.InvariantCulture)
        Return options
    End Function

    ''' <summary>
    ''' Lightweight IDisposable scope to show a modeless progress window with cancellation support
    ''' for long-running operations. Integrates with ProgressBarModule and DPIProgressForm/ProgressForm.
    ''' The progress form runs on a dedicated STA UI thread with its own message loop.
    ''' </summary>
    Public NotInheritable Class ProgressScope
        Implements System.IDisposable

        Private ReadOnly _cts As System.Threading.CancellationTokenSource = New System.Threading.CancellationTokenSource()
        Private _uiThread As System.Threading.Thread
        Private _form As System.Windows.Forms.Form
        Private ReadOnly _useDpiForm As Boolean
        Private _closed As Integer = 0

        ''' <summary>
        ''' Initializes the progress scope and displays a modeless progress window.
        ''' The form updates its display by polling ProgressBarModule.* every 250ms.
        ''' </summary>
        ''' <param name="headerText">The header text displayed on the progress form.</param>
        ''' <param name="initialLabel">The initial status label text.</param>
        ''' <param name="max">Maximum progress value (default 100).</param>
        ''' <param name="useDpiForm">If True, uses DPIProgressForm; otherwise uses ProgressForm.</param>
        Public Sub New(headerText As String,
                   initialLabel As String,
                   Optional max As Integer = 100,
                   Optional useDpiForm As Boolean = True)
            _useDpiForm = useDpiForm

            ' Initialize global progress state
            ProgressBarModule.CancelOperation = False
            ProgressBarModule.GlobalProgressMax = System.Math.Max(1, max)
            ProgressBarModule.GlobalProgressValue = 0
            ProgressBarModule.GlobalProgressLabel = If(initialLabel, "")

            ' The independent progress STA owns its complete form/context lifetime.
            ' Disposal never waits for that STA from the Office message loop.
            Dim hostExitToken As System.Threading.CancellationToken = Globals.ThisAddIn.HostWindowLifetimeToken
            _uiThread = New System.Threading.Thread(
                Sub()
                    Try
                        If System.Threading.Volatile.Read(_closed) <> 0 OrElse hostExitToken.IsCancellationRequested Then Return
                        System.Threading.SynchronizationContext.SetSynchronizationContext(Nothing)
                        System.Windows.Forms.Application.EnableVisualStyles()
                        Using progressForm As System.Windows.Forms.Form =
                            If(_useDpiForm,
                               CType(New DPIProgressForm(headerText, initialLabel), System.Windows.Forms.Form),
                               CType(New ProgressForm(headerText, initialLabel), System.Windows.Forms.Form))
                            System.Threading.Volatile.Write(_form, progressForm)
                            If System.Threading.Volatile.Read(_closed) <> 0 OrElse hostExitToken.IsCancellationRequested Then Return
                            Using closeTimer As New System.Windows.Forms.Timer() With {.Interval = 100}
                                AddHandler closeTimer.Tick,
                                    Sub()
                                        If System.Threading.Volatile.Read(_closed) <> 0 OrElse hostExitToken.IsCancellationRequested Then
                                            closeTimer.Stop()
                                            progressForm.Close()
                                        End If
                                    End Sub
                                closeTimer.Start()
                                System.Windows.Forms.Application.Run(progressForm)
                            End Using
                        End Using
                    Catch ex As System.Exception
                        System.Diagnostics.Debug.WriteLine("Progress window: " & ex.Message)
                    Finally
                        System.Threading.Volatile.Write(_form, Nothing)
                        Try
                            System.Windows.Forms.Application.ExitThread()
                        Catch ex As System.Exception
                            System.Diagnostics.Debug.WriteLine("Progress message loop cleanup: " & ex.Message)
                        End Try
                        Try
                            Dim workerContext As System.Windows.Forms.WindowsFormsSynchronizationContext =
                                TryCast(System.Threading.SynchronizationContext.Current, System.Windows.Forms.WindowsFormsSynchronizationContext)
                            If workerContext IsNot Nothing Then workerContext.Dispose()
                        Catch ex As System.Exception
                            System.Diagnostics.Debug.WriteLine("Progress context cleanup: " & ex.Message)
                        Finally
                            System.Threading.SynchronizationContext.SetSynchronizationContext(Nothing)
                        End Try
                    End Try
                End Sub)
            _uiThread.IsBackground = True
            _uiThread.SetApartmentState(System.Threading.ApartmentState.STA)
            _uiThread.Start()
        End Sub

        ''' <summary>
        ''' Reports progress in a thread-safe manner via the global ProgressBarModule.
        ''' </summary>
        ''' <param name="current">Current progress value.</param>
        ''' <param name="max">Optional new maximum value (if >= 1).</param>
        ''' <param name="label">Optional new status label text.</param>
        Public Shared Sub Report(current As Integer,
                             Optional max As Integer = -1,
                             Optional label As String = Nothing)
            If max >= 1 Then ProgressBarModule.GlobalProgressMax = max
            If label IsNot Nothing Then ProgressBarModule.GlobalProgressLabel = label
            ProgressBarModule.GlobalProgressValue = System.Math.Max(0, System.Math.Min(current, ProgressBarModule.GlobalProgressMax))
        End Sub

        ''' <summary>
        ''' Requests cancellation of the operation.
        ''' Also triggered by the Cancel button in the progress UI.
        ''' </summary>
        Public Sub RequestCancel()
            _cts.Cancel()
            ProgressBarModule.CancelOperation = True
        End Sub

        ''' <summary>
        ''' Gets a value indicating whether cancellation has been requested.
        ''' Check this frequently at safe points (between steps/chunks) to bail out early.
        ''' </summary>
        Public ReadOnly Property CancelRequested As Boolean
            Get
                Return ProgressBarModule.CancelOperation OrElse _cts.IsCancellationRequested
            End Get
        End Property

        ''' <summary>
        ''' Gets the CancellationToken associated with this scope.
        ''' </summary>
        Public ReadOnly Property Token As System.Threading.CancellationToken
            Get
                Return _cts.Token
            End Get
        End Property

        ''' <summary>
        ''' Disposes the progress scope, signaling cancellation and asynchronously
        ''' closing the progress form on its own STA. A delayed worker startup checks
        ''' the same closed flag before and after constructing its form.
        ''' </summary>
        Public Sub Dispose() Implements System.IDisposable.Dispose
            If System.Threading.Interlocked.Exchange(_closed, 1) <> 0 Then Return
            ProgressBarModule.CancelOperation = True
            Dim f As System.Windows.Forms.Form = System.Threading.Volatile.Read(_form)
            If f Is Nothing Then Return
            Try
                If f.IsHandleCreated AndAlso Not f.IsDisposed Then
                    f.BeginInvoke(New System.Action(
                        Sub()
                            If Not f.IsDisposed Then f.Close()
                        End Sub))
                End If
            Catch ex As System.InvalidOperationException
                ' The worker's local timer/Finally handles the concurrent shutdown.
            End Try
        End Sub
    End Class

End Class
