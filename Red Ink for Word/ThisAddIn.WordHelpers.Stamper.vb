' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.WordHelpers.Stamper.vb
' Purpose: Exhibit stamp overlay for PDF files. Places a user-provided image
'          (exhibit stamp graphic) on the top-right area of every page and
'          renders an exhibit number text on top of it. The exhibit number is
'          extracted from each filename using either regex or LLM-based parsing.
'
' Architecture:
'  - UI Parameter Collection: Uses SharedMethods.InputParameter and
'    ShowCustomVariableInputForm to collect stamp parameters (image path,
'    offsets, text size, prefix/suffix, extraction mode, output paths).
'    All parameters are persisted in My.Settings for reuse.
'  - Exhibit Number Extraction:
'      - Regex mode: Extracts any number-like pattern (including dotted,
'        hyphenated, alphanumeric identifiers) from the filename stem.
'      - LLM mode: Sends filenames in batches to the LLM for intelligent
'        exhibit number extraction; returns JSON array of results.
'  - Stamp Overlay (PdfSharp): Opens each PDF, loads the stamp image,
'    computes placement based on user-defined offsets and desired width
'    (preserving aspect ratio), draws the image and overlays the exhibit
'    text (right-aligned) on every page.
'  - Single-file & Multi-file Modes: Uses DragDropForm(FileOrDirectory)
'    to auto-detect whether a file or folder was selected.
'  - Unit Conversion: Offset/size values entered with "cm" or "in" suffix
'    are automatically converted to points (1 cm ≈ 28.35 pts, 1 in = 72 pts).
'  - Font Handling: Own EnsureFontResolver mirrors PdfRedactionService's
'    ArialFontResolver setup for PdfSharp text rendering; falls back to
'    Helvetica.
'  - External Dependencies: PdfSharp for PDF manipulation; SharedLibrary
'    for UI, LLM, and progress utilities.
' =============================================================================

Option Explicit On
Option Strict On

Imports System.Diagnostics
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Entry point for exhibit stamping. Collects parameters, selects file(s),
    ''' extracts exhibit numbers, and applies stamp overlay to every PDF page.
    ''' </summary>
    Public Async Sub StampExhibitPDF()
        If INILoadFail() Then Return

        Try
            ' ── 1) Load persisted defaults from My.Settings (strongly-typed) ──
            Dim defaultImagePath As System.String = If(My.Settings.Stamp_ImagePath, "")
            Dim defaultTopOffset As System.Double = My.Settings.Stamp_TopOffset
            Dim defaultRightOffset As System.Double = My.Settings.Stamp_RightOffset
            Dim defaultStampWidth As System.Double = My.Settings.Stamp_Width
            Dim defaultTextTopOffset As System.Double = My.Settings.Stamp_TextTopOffset
            Dim defaultTextRightOffset As System.Double = My.Settings.Stamp_TextRightOffset
            Dim defaultTextSize As System.Double = My.Settings.Stamp_TextSize
            Dim defaultPrefix As System.String = If(My.Settings.Stamp_Prefix, "")
            Dim defaultSuffix As System.String = If(My.Settings.Stamp_Suffix, "")
            Dim defaultUseLlm As System.Boolean = My.Settings.Stamp_UseLLM
            Dim defaultSubdir As System.String = If(My.Settings.Stamp_OutputSubdir, "")
            Dim defaultExtension As System.String = If(My.Settings.Stamp_OutputExtension, "")
            Dim defaultRegex As System.String = If(My.Settings.Stamp_Regex, "")

            ' Apply sensible defaults if settings are at zero (first run)
            If defaultTopOffset <= 0 Then defaultTopOffset = 20.0
            If defaultRightOffset <= 0 Then defaultRightOffset = 20.0
            If defaultStampWidth <= 0 Then defaultStampWidth = 100.0
            If defaultTextTopOffset <= 0 Then defaultTextTopOffset = 30.0
            If defaultTextRightOffset <= 0 Then defaultTextRightOffset = 25.0
            If defaultTextSize <= 0 Then defaultTextSize = 12.0

            ' ── 2) Collect parameters via single form ──
            ' Conversion note: 1 cm ≈ 28.35 pts, 1 inch = 72 pts. Append "cm" or "in" to auto-convert.
            Dim p00 As SLib.InputParameter = New SLib.InputParameter("Stamp image file path (or leave empty to pick via drag && drop)", defaultImagePath)
            Dim p01 As SLib.InputParameter = New SLib.InputParameter("Image distance from top (pts; append 'cm'/'in' to convert)", CInt(defaultTopOffset))
            Dim p02 As SLib.InputParameter = New SLib.InputParameter("Image distance from right (pts; append 'cm'/'in' to convert)", CInt(defaultRightOffset))
            Dim p03 As SLib.InputParameter = New SLib.InputParameter("Image width (pts; append 'cm'/'in' to convert, keeps proportions)", CInt(defaultStampWidth))
            Dim p04 As SLib.InputParameter = New SLib.InputParameter("Text distance from top (pts; append 'cm'/'in' to convert)", CInt(defaultTextTopOffset))
            Dim p05 As SLib.InputParameter = New SLib.InputParameter("Text distance from right (pts; append 'cm'/'in' to convert)", CInt(defaultTextRightOffset))
            Dim p06 As SLib.InputParameter = New SLib.InputParameter("Text size (pts)", CInt(defaultTextSize))
            Dim p07 As SLib.InputParameter = New SLib.InputParameter("Exhibit number prefix", defaultPrefix)
            Dim p08 As SLib.InputParameter = New SLib.InputParameter("Exhibit number suffix", defaultSuffix)
            Dim p09 As SLib.InputParameter = New SLib.InputParameter("Use LLM to extract exhibit numbers", defaultUseLlm)
            Dim p10 As SLib.InputParameter = New SLib.InputParameter("Custom regex for number extraction", defaultRegex)
            Dim p11 As SLib.InputParameter = New SLib.InputParameter("Output subdirectory", defaultSubdir)
            Dim p12 As SLib.InputParameter = New SLib.InputParameter("Output filename extension", defaultExtension)
            Dim p13 As SLib.InputParameter = New SLib.InputParameter("Force rasterize before stamping (use if stamp is invisible on some PDFs)", False)

            Dim params() As SLib.InputParameter = {p00, p01, p02, p03, p04, p05, p06, p07, p08, p09, p10, p11, p12, p13}

            If ShowCustomVariableInputForm("Please set the exhibit stamp parameters (1 cm ≈ 28.35 pts, 1 in = 72 pts):",
                                           AN & " Exhibit Stamp", params) = False Then
                Return
            End If

            ' ── 3) Read back values (with cm/in → pts auto-conversion) ──
            Dim stampImagePath As System.String = System.Convert.ToString(params(0).Value)
            Dim topOffset As System.Double = ExhibitStampService.ParseWithUnitConversion(params(1).Value, 20.0)
            Dim rightOffset As System.Double = ExhibitStampService.ParseWithUnitConversion(params(2).Value, 20.0)
            Dim stampWidth As System.Double = ExhibitStampService.ParseWithUnitConversion(params(3).Value, 100.0)
            Dim textTopOffset As System.Double = ExhibitStampService.ParseWithUnitConversion(params(4).Value, 30.0)
            Dim textRightOffset As System.Double = ExhibitStampService.ParseWithUnitConversion(params(5).Value, 25.0)
            Dim textSize As System.Double = ExhibitStampService.ParseWithUnitConversion(params(6).Value, 12.0)

            Dim prefix As System.String = System.Convert.ToString(params(7).Value)
            Dim suffix As System.String = System.Convert.ToString(params(8).Value)
            Dim useLlm As System.Boolean = False
            Try : useLlm = System.Convert.ToBoolean(params(9).Value) : Catch : End Try
            Dim customRegex As System.String = System.Convert.ToString(params(10).Value)
            Dim outSubdir As System.String = System.Convert.ToString(params(11).Value)
            Dim outExtension As System.String = System.Convert.ToString(params(12).Value)
            Dim forceRasterize As System.Boolean = False
            Try : forceRasterize = System.Convert.ToBoolean(params(13).Value) : Catch : End Try

            ' Clamp minimums
            If stampWidth < 10 Then stampWidth = 10
            If textSize < 4 Then textSize = 4

            ' ── 4) Validate / acquire stamp image via drag & drop if not provided ──
            If System.String.IsNullOrWhiteSpace(stampImagePath) OrElse Not System.IO.File.Exists(stampImagePath) Then
                ' Offer drag & drop for image file
                DragDropFormLabel = "Stamp image file (PNG, JPG, BMP, etc.)"
                DragDropFormFilter = "Image Files|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff|All Files|*.*"

                Try
                    Using frm As New DragDropForm(DragDropMode.FileOnly)
                        Dim __safeDialogOwner134 As System.Windows.Forms.IWin32Window = SharedLibrary.SharedLibrary.SharedMethods.ResolveSameThreadDialogOwner()
                        If If(__safeDialogOwner134 IsNot Nothing, frm.ShowDialog(__safeDialogOwner134), frm.ShowDialog()) = DialogResult.OK AndAlso Not System.String.IsNullOrWhiteSpace(frm.SelectedFilePath) Then
                            stampImagePath = frm.SelectedFilePath
                        End If
                    End Using
                Finally
                    DragDropFormLabel = ""
                    DragDropFormFilter = ""
                End Try

                If System.String.IsNullOrWhiteSpace(stampImagePath) OrElse Not System.IO.File.Exists(stampImagePath) Then
                    ShowCustomMessageBox("No stamp image file has been provided - will abort.", AN & " Exhibit Stamp")
                    Return
                End If
            End If

            ' ── 5) Persist settings (strongly-typed) ──
            ' Note: ForceRasterize is intentionally NOT persisted — it's a temporary per-run option
            My.Settings.Stamp_ImagePath = If(stampImagePath, "")
            My.Settings.Stamp_TopOffset = topOffset
            My.Settings.Stamp_RightOffset = rightOffset
            My.Settings.Stamp_Width = stampWidth
            My.Settings.Stamp_TextTopOffset = textTopOffset
            My.Settings.Stamp_TextRightOffset = textRightOffset
            My.Settings.Stamp_TextSize = textSize
            My.Settings.Stamp_Prefix = If(prefix, "")
            My.Settings.Stamp_Suffix = If(suffix, "")
            My.Settings.Stamp_UseLLM = useLlm
            My.Settings.Stamp_OutputSubdir = If(outSubdir, "")
            My.Settings.Stamp_OutputExtension = If(outExtension, "")
            My.Settings.Stamp_Regex = If(customRegex, "")
            My.Settings.Save()

            ' ── 6) Build stamp config ──
            Dim config As New ExhibitStampService.StampConfig()
            config.StampImagePath = stampImagePath
            config.TopOffsetPts = topOffset
            config.RightOffsetPts = rightOffset
            config.StampWidthPts = stampWidth
            config.TextTopOffsetPts = textTopOffset
            config.TextRightOffsetPts = textRightOffset
            config.TextSizePts = textSize
            config.Prefix = If(prefix, "")
            config.Suffix = If(suffix, "")
            config.UseLlm = useLlm
            config.CustomRegex = If(customRegex, "")
            config.OutputSubdir = If(outSubdir, "")
            config.OutputExtension = If(outExtension, "")
            config.ForceRasterize = forceRasterize

            ' ── 7) Select PDF file or folder via unified drag & drop (FileOrDirectory) ──
            Dim selectedPath As System.String = Nothing
            Dim isDirectory As System.Boolean = False

            DragDropFormLabel = "Drop a PDF file or a folder of PDFs"
            DragDropFormFilter = "PDF Files|*.pdf|All Files|*.*"
            Try
                Using frm As New DragDropForm(DragDropMode.FileOrDirectory)
                    Dim __safeDialogOwner191 As System.Windows.Forms.IWin32Window = SharedLibrary.SharedLibrary.SharedMethods.ResolveSameThreadDialogOwner()
                    If If(__safeDialogOwner191 IsNot Nothing, frm.ShowDialog(__safeDialogOwner191), frm.ShowDialog()) = DialogResult.OK AndAlso Not System.String.IsNullOrWhiteSpace(frm.SelectedFilePath) Then
                        selectedPath = frm.SelectedFilePath
                        isDirectory = frm.IsDirectory
                    End If
                End Using
            Finally
                DragDropFormLabel = ""
                DragDropFormFilter = ""
            End Try

            If System.String.IsNullOrWhiteSpace(selectedPath) Then
                ShowCustomMessageBox("No file or folder has been selected - will abort.", AN & " Exhibit Stamp")
                Return
            End If

            If Not isDirectory Then
                ' === SINGLE-FILE MODE ===
                Dim filePath As System.String = selectedPath

                If Not System.IO.File.Exists(filePath) Then
                    ShowCustomMessageBox("The selected file does not exist - will abort.", AN & " Exhibit Stamp")
                    Return
                End If
                If Not System.String.Equals(System.IO.Path.GetExtension(filePath), ".pdf", System.StringComparison.OrdinalIgnoreCase) Then
                    ShowCustomMessageBox("The selected file '" & System.IO.Path.GetFileName(filePath) & "' is not a PDF - will abort.", AN & " Exhibit Stamp")
                    Return
                End If

                ' Extract exhibit number
                Dim exhibitNumber As System.String = Nothing
                If config.UseLlm Then
                    Dim llmResults As System.Collections.Generic.Dictionary(Of System.String, System.String) =
                        Await ExhibitStampService.ExtractExhibitNumbersViaLlmAsync(New System.String() {filePath})
                    If llmResults IsNot Nothing AndAlso llmResults.ContainsKey(filePath) Then
                        exhibitNumber = llmResults(filePath)
                    End If
                Else
                    exhibitNumber = ExhibitStampService.ExtractExhibitNumberRegex(filePath, config.CustomRegex)
                End If

                If System.String.IsNullOrWhiteSpace(exhibitNumber) Then
                    ShowCustomMessageBox("Could not extract an exhibit number from the filename '" &
                                         System.IO.Path.GetFileName(filePath) & "' - will abort.", AN & " Exhibit Stamp")
                    Return
                End If

                Dim splash As New SLib.SplashScreen("Stamping PDF ...")
                splash.Show()
                splash.Refresh()

                Try
                    Dim stampResult As ExhibitStampService.StampResult = Await System.Threading.Tasks.Task.Run(
                        Function() ExhibitStampService.ApplyStamp(filePath, exhibitNumber, config))
                    splash.Close()

                    ' Safety check: verify page count matches
                    Dim pageWarning As System.String = ExhibitStampService.VerifyPageCount(filePath, stampResult.OutputPath)

                    Dim resultMsg As System.String = "Exhibit stamp applied successfully." & vbCrLf & "Output: " & stampResult.OutputPath
                    If Not System.String.IsNullOrWhiteSpace(stampResult.RasterizeWarning) Then
                        resultMsg &= vbCrLf & vbCrLf & "⚠ " & stampResult.RasterizeWarning
                    End If
                    If Not System.String.IsNullOrWhiteSpace(pageWarning) Then
                        resultMsg &= vbCrLf & vbCrLf & "⚠ " & pageWarning
                    End If
                    ShowCustomMessageBox(resultMsg, AN & " Exhibit Stamp")

                Catch ex As System.Exception
                    splash.Close()
                    Dim errMsg As System.String = ex.Message
                    If System.String.IsNullOrWhiteSpace(errMsg) OrElse
                       errMsg.Equals("No error", System.StringComparison.OrdinalIgnoreCase) OrElse
                       errMsg.Equals("No error.", System.StringComparison.OrdinalIgnoreCase) Then
                        errMsg = "PDF could not be processed (may be encrypted, restricted, or corrupt)"
                    End If
                    ShowCustomMessageBox("Stamping failed: " & errMsg, AN & " Exhibit Stamp")
                End Try

            Else
                ' === MULTI-FILE (FOLDER) MODE ===
                Dim selectedFolder As System.String = selectedPath

                Dim pdfFiles As System.String()
                Try
                    pdfFiles = System.IO.Directory.GetFiles(selectedFolder, "*.pdf", SearchOption.TopDirectoryOnly)
                Catch ex As System.Exception
                    ShowCustomMessageBox("Failed to enumerate PDFs: " & ex.Message, AN & " Exhibit Stamp")
                    Return
                End Try

                If pdfFiles Is Nothing OrElse pdfFiles.Length = 0 Then
                    ShowCustomMessageBox("Selected folder contains no PDF files - nothing to do.", AN & " Exhibit Stamp")
                    Return
                End If

                ' Extract exhibit numbers (batch for LLM, individual for regex)
                Dim exhibitMap As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.OrdinalIgnoreCase)

                If config.UseLlm Then
                    Dim splash As New SLib.SplashScreen("Extracting exhibit numbers via AI ...")
                    splash.Show()
                    splash.Refresh()
                    Try
                        exhibitMap = Await ExhibitStampService.ExtractExhibitNumbersViaLlmAsync(pdfFiles)
                    Catch ex As System.Exception
                        splash.Close()
                        ShowCustomMessageBox("LLM exhibit number extraction failed: " & ex.Message, AN & " Exhibit Stamp")
                        Return
                    End Try
                    splash.Close()
                Else
                    For Each f As System.String In pdfFiles
                        Dim num As System.String = ExhibitStampService.ExtractExhibitNumberRegex(f, config.CustomRegex)
                        If Not System.String.IsNullOrWhiteSpace(num) Then
                            exhibitMap(f) = num
                        End If
                    Next
                End If

                If exhibitMap.Count = 0 Then
                    ShowCustomMessageBox("Could not extract exhibit numbers from any filename - nothing to do.", AN & " Exhibit Stamp")
                    Return
                End If

                ' Compute output directory
                Dim outputBaseDir As System.String = selectedFolder
                If Not System.String.IsNullOrWhiteSpace(config.OutputSubdir) Then
                    outputBaseDir = System.IO.Path.Combine(selectedFolder, config.OutputSubdir)
                    Try
                        System.IO.Directory.CreateDirectory(outputBaseDir)
                    Catch ex As System.Exception
                        ShowCustomMessageBox("Cannot create output folder '" & outputBaseDir & "': " & ex.Message, AN & " Exhibit Stamp")
                        Return
                    End Try
                End If

                ' If both subdir and extension are empty, set a default extension
                If System.String.IsNullOrWhiteSpace(config.OutputSubdir) AndAlso System.String.IsNullOrWhiteSpace(config.OutputExtension) Then
                    config.OutputExtension = "_stamped"
                End If

                ' Progress bar
                ShowProgressBarInSeparateThread(AN & " Exhibit Stamp", "Stamping PDFs...")
                ProgressBarModule.CancelOperation = False

                GlobalProgressMax = pdfFiles.Length
                GlobalProgressValue = 0
                GlobalProgressLabel = "Starting..."

                Dim processed As System.Int32 = 0
                Dim skipped As System.Int32 = 0
                Dim errors As New System.Collections.Generic.List(Of System.String)()
                Dim warnings As New System.Collections.Generic.List(Of System.String)()
                Dim pageWarnings As New System.Collections.Generic.List(Of System.String)()
                Dim cancelled As System.Boolean = False

                For i As System.Int32 = 0 To pdfFiles.Length - 1
                    If ProgressBarModule.CancelOperation Then
                        cancelled = True
                        Exit For
                    End If

                    Dim inPath As System.String = pdfFiles(i)
                    Dim fileNameOnly As System.String = System.IO.Path.GetFileName(inPath)
                    GlobalProgressValue = i
                    GlobalProgressLabel = "Processing " & fileNameOnly & " (" & (i + 1).ToString() & " of " & pdfFiles.Length.ToString() & ")"

                    ' Check if we have an exhibit number for this file
                    Dim exhibitNum As System.String = Nothing
                    If Not exhibitMap.TryGetValue(inPath, exhibitNum) OrElse System.String.IsNullOrWhiteSpace(exhibitNum) Then
                        skipped += 1
                        errors.Add(fileNameOnly & ": no exhibit number extracted (skipped)")
                        Continue For
                    End If

                    ' Compute output path
                    Dim nameOnly As System.String = System.IO.Path.GetFileNameWithoutExtension(inPath)
                    Dim extOnly As System.String = System.IO.Path.GetExtension(inPath)
                    Dim outName As System.String = nameOnly & config.OutputExtension & extOnly
                    Dim outPath As System.String = System.IO.Path.Combine(outputBaseDir, outName)

                    Try
                        Dim stampResult As ExhibitStampService.StampResult = Await System.Threading.Tasks.Task.Run(
                            Function() ExhibitStampService.ApplyStamp(inPath, exhibitNum, config, outPath))
                        processed += 1

                        ' Report rasterize fallback as a warning (not an error — file was processed)
                        If Not System.String.IsNullOrWhiteSpace(stampResult.RasterizeWarning) Then
                            warnings.Add(fileNameOnly & ": " & stampResult.RasterizeWarning)
                        End If

                        ' Safety check: verify page count matches
                        Dim pageWarning As System.String = ExhibitStampService.VerifyPageCount(inPath, stampResult.OutputPath)
                        If Not System.String.IsNullOrWhiteSpace(pageWarning) Then
                            pageWarnings.Add(fileNameOnly & ": " & pageWarning)
                        End If

                    Catch ex As System.Exception
                        Dim errMsg As System.String = ex.Message
                        If System.String.IsNullOrWhiteSpace(errMsg) OrElse
                           errMsg.Equals("No error", System.StringComparison.OrdinalIgnoreCase) OrElse
                           errMsg.Equals("No error.", System.StringComparison.OrdinalIgnoreCase) Then
                            errMsg = "PDF could not be processed (may be encrypted, restricted, or corrupt)"
                        End If
                        errors.Add(fileNameOnly & ": " & errMsg)
                    End Try
                Next

                ' Close progress bar
                ProgressBarModule.CancelOperation = True

                ' ── Build summary message (shown in message box) ──
                Dim msg As System.String
                If cancelled Then
                    msg = "Operation aborted by user." & vbCrLf
                Else
                    msg = "Exhibit stamping completed." & vbCrLf
                End If
                msg &= "Processed: " & processed.ToString() & " of " & pdfFiles.Length.ToString() & " file(s)."
                If skipped > 0 Then
                    msg &= vbCrLf & "Skipped (no exhibit number): " & skipped.ToString()
                End If
                If Not System.String.IsNullOrWhiteSpace(config.OutputSubdir) Then
                    msg &= vbCrLf & "Output folder: " & outputBaseDir
                End If
                If pageWarnings.Count > 0 Then
                    msg &= vbCrLf & vbCrLf & "⚠ PAGE COUNT MISMATCHES (" & pageWarnings.Count.ToString() & "):"
                    For Each pw As System.String In pageWarnings
                        msg &= vbCrLf & " • " & pw
                    Next
                End If
                ' Only actual errors (files that could not be processed) in the message box
                If errors.Count > 0 Then
                    msg &= vbCrLf & vbCrLf & "Errors (" & errors.Count.ToString() & "):"
                    For Each err As System.String In errors
                        msg &= vbCrLf & " • " & err
                    Next
                End If
                If warnings.Count > 0 Then
                    msg &= vbCrLf & vbCrLf & warnings.Count.ToString() & " file(s) were rasterized (details copied to clipboard)."
                End If

                ' ── Build full clipboard message (includes warnings detail) ──
                If pageWarnings.Count > 0 OrElse errors.Count > 0 OrElse warnings.Count > 0 Then
                    Dim clipMsg As System.String = msg
                    If warnings.Count > 0 Then
                        clipMsg &= vbCrLf & vbCrLf & "Rasterize warnings (" & warnings.Count.ToString() & "):"
                        For Each w As System.String In warnings
                            clipMsg &= vbCrLf & " • " & w
                        Next
                    End If
                    SLib.PutInClipboard(clipMsg)
                End If
                ShowCustomMessageBox(msg, AN & " Exhibit Stamp")
            End If

        Catch ex As System.Exception
            ShowCustomMessageBox("Exhibit Stamp operation failed: " & ex.Message, AN & " Exhibit Stamp")
        End Try
    End Sub

    ''' <summary>
    ''' Service class providing exhibit stamp operations: exhibit number extraction
    ''' (regex and LLM), image overlay, and text rendering on PDF pages.
    ''' </summary>
    Public Class ExhibitStampService

        ''' <summary>
        ''' Configuration parameters for exhibit stamping.
        ''' </summary>
        Public Class StampConfig
            ''' <summary>Full path to the stamp image file (PNG, JPG, BMP, etc.).</summary>
            Public Property StampImagePath As System.String
            ''' <summary>Distance from the top edge of the page to the top of the image, in points.</summary>
            Public Property TopOffsetPts As System.Double
            ''' <summary>Distance from the right edge of the page to the right side of the image, in points.</summary>
            Public Property RightOffsetPts As System.Double
            ''' <summary>Desired width of the stamp image in points (height is computed to preserve aspect ratio).</summary>
            Public Property StampWidthPts As System.Double
            ''' <summary>Distance from the top edge of the page to the top of the exhibit text, in points.</summary>
            Public Property TextTopOffsetPts As System.Double
            ''' <summary>Distance from the right edge of the page to the right edge of the exhibit text, in points.</summary>
            Public Property TextRightOffsetPts As System.Double
            ''' <summary>Font size for the exhibit number text, in points.</summary>
            Public Property TextSizePts As System.Double
            ''' <summary>Prefix prepended to the exhibit number (e.g., "Exhibit ").</summary>
            Public Property Prefix As System.String
            ''' <summary>Suffix appended to the exhibit number (e.g., "").</summary>
            Public Property Suffix As System.String
            ''' <summary>If True, use LLM for exhibit number extraction; otherwise use regex.</summary>
            Public Property UseLlm As System.Boolean
            ''' <summary>Optional custom regex pattern for exhibit number extraction (empty = default).</summary>
            Public Property CustomRegex As System.String
            ''' <summary>Output subdirectory relative to source folder (optional).</summary>
            Public Property OutputSubdir As System.String
            ''' <summary>Suffix appended to the filename before .pdf extension (optional).</summary>
            Public Property OutputExtension As System.String
            ''' <summary>
            ''' If True, always rasterize the first page before stamping (guarantees visibility
            ''' at the cost of losing text selectability). Not persisted — temporary per-run option.
            ''' </summary>
            Public Property ForceRasterize As System.Boolean
        End Class

        ' ── Conversion constants ──
        Private Const PtsPerCm As System.Double = 28.3464567
        Private Const PtsPerInch As System.Double = 72.0

        ''' <summary>
        ''' Parses a numeric value from user input, with optional "cm" or "in" suffix
        ''' for automatic conversion to points. Returns <paramref name="fallback"/> on failure.
        ''' </summary>
        ''' <param name="rawValue">Raw value from InputParameter (may be Integer, Double, or String with unit suffix).</param>
        ''' <param name="fallback">Default value if parsing fails.</param>
        ''' <returns>Value in points.</returns>
        Public Shared Function ParseWithUnitConversion(rawValue As System.Object, fallback As System.Double) As System.Double
            If rawValue Is Nothing Then Return fallback

            Dim text As System.String = System.Convert.ToString(rawValue).Trim()
            If System.String.IsNullOrWhiteSpace(text) Then Return fallback

            ' Check for "cm" suffix (case-insensitive)
            If text.EndsWith("cm", System.StringComparison.OrdinalIgnoreCase) Then
                Dim numPart As System.String = text.Substring(0, text.Length - 2).Trim()
                Dim cmVal As System.Double
                If System.Double.TryParse(numPart, System.Globalization.NumberStyles.Any,
                                          System.Globalization.CultureInfo.InvariantCulture, cmVal) Then
                    Return cmVal * PtsPerCm
                End If
                ' Also try current culture
                If System.Double.TryParse(numPart, cmVal) Then
                    Return cmVal * PtsPerCm
                End If
                Return fallback
            End If

            ' Check for "in" or "inch" suffix (case-insensitive)
            Dim inSuffix As System.String = ""
            If text.EndsWith("inch", System.StringComparison.OrdinalIgnoreCase) Then
                inSuffix = "inch"
            ElseIf text.EndsWith("in", System.StringComparison.OrdinalIgnoreCase) Then
                inSuffix = "in"
            End If
            If inSuffix.Length > 0 Then
                Dim numPart As System.String = text.Substring(0, text.Length - inSuffix.Length).Trim()
                Dim inVal As System.Double
                If System.Double.TryParse(numPart, System.Globalization.NumberStyles.Any,
                                          System.Globalization.CultureInfo.InvariantCulture, inVal) Then
                    Return inVal * PtsPerInch
                End If
                If System.Double.TryParse(numPart, inVal) Then
                    Return inVal * PtsPerInch
                End If
                Return fallback
            End If

            ' Plain numeric value (already in points)
            Dim pts As System.Double
            If System.Double.TryParse(text, System.Globalization.NumberStyles.Any,
                                      System.Globalization.CultureInfo.InvariantCulture, pts) Then
                Return pts
            End If
            If System.Double.TryParse(text, pts) Then
                Return pts
            End If

            ' Last resort: try Convert (handles Integer input etc.)
            Try
                Return System.Convert.ToDouble(rawValue)
            Catch
                Return fallback
            End Try
        End Function

        ' ── Default regex: captures numeric identifier after a multilingual exhibit label ──
        ' Supports: English (Exhibit, Ex., Exh.), German (Beilage, Anlage, Beweisstück,
        ' Beweismittel, Urkunde, Urk.), French (Pièce, Annexe), Italian (Allegato, Doc.),
        ' Spanish/Portuguese (Anexo, Exhibición, Documento).
        ' Two-pass approach: first try to find an explicit exhibit label pattern, then fall back
        ' to capturing the longest standalone number-like token.
        Private Const ExhibitLabeledRegex As System.String =
            "(?:" &
            "Exhibit|Exh?\.?\s*" &                                          ' English
            "|Beilage|Anlage|Beweisstück|Beweisst(?:ue|ü)ck" &             ' German (with ü and ue variant)
            "|Beweismittel|Urkunde|Urk\.?\s*" &                            ' German
            "|Pièce|Pi(?:e|è)ce|Annexe" &                                  ' French (with è and plain e variant)
            "|Allegato|Doc\.?\s*" &                                         ' Italian
            "|Anexo|Exhibición|Exhibici(?:o|ó)n|Documento" &               ' Spanish / Portuguese
            ")\s*[-:.]?\s*" &                                               ' Optional separator after label
            "([A-Za-z]?[\-.]?\d+(?:[.\-]\d+)*(?:[.\-][A-Za-z]\d*)*)"
        Private Const ExhibitFallbackRegex As System.String =
            "(\d+(?:[.\-]\d+)*(?:[.\-][A-Za-z]\d*)*)"

        ''' <summary>
        ''' Extracts an exhibit number from a PDF filename using regex.
        ''' First strips the extension and any leading/trailing "..." from the stem.
        ''' Tries a labeled pattern (Exhibit/Ex prefix) first, then falls back to
        ''' the longest standalone numeric token.
        ''' </summary>
        ''' <param name="filePath">Full path to the PDF file.</param>
        ''' <param name="customRegex">Optional user-supplied regex pattern (empty = default).</param>
        ''' <returns>Extracted exhibit number string, or Nothing if no match.</returns>
        Public Shared Function ExtractExhibitNumberRegex(filePath As System.String, customRegex As System.String) As System.String
            If System.String.IsNullOrWhiteSpace(filePath) Then Return Nothing

            Dim stem As System.String = System.IO.Path.GetFileNameWithoutExtension(filePath)
            If System.String.IsNullOrWhiteSpace(stem) Then Return Nothing

            ' Strip leading/trailing ellipsis characters that may appear in filenames
            stem = stem.Trim("."c, " "c, CChar(ChrW(&H2026)))

            ' If user supplied a custom regex, use it directly
            If Not System.String.IsNullOrWhiteSpace(customRegex) Then
                Try
                    Dim matches As MatchCollection = Regex.Matches(stem, customRegex, RegexOptions.IgnoreCase)
                    If matches.Count = 0 Then Return Nothing
                    Dim best As System.String = ""
                    For Each m As Match In matches
                        Dim val As System.String = m.Value.Trim()
                        If val.Length > best.Length Then best = val
                    Next
                    Return If(best.Length > 0, best, Nothing)
                Catch
                    Return Nothing
                End Try
            End If

            ' Pass 1: Try labeled pattern (e.g., "Exhibit A-1", "Ex. 32.3.1")
            Try
                Dim labelMatch As Match = Regex.Match(stem, ExhibitLabeledRegex, RegexOptions.IgnoreCase)
                If labelMatch.Success AndAlso labelMatch.Groups.Count > 1 Then
                    Dim captured As System.String = labelMatch.Groups(1).Value.Trim()
                    If captured.Length > 0 Then Return captured
                End If
            Catch
                ' Fall through to pass 2
            End Try

            ' Pass 2: Longest standalone number token (e.g., "32.3.1" from "Document 32.3.1")
            Try
                Dim matches As MatchCollection = Regex.Matches(stem, ExhibitFallbackRegex, RegexOptions.IgnoreCase)
                If matches.Count = 0 Then Return Nothing

                Dim best As System.String = ""
                For Each m As Match In matches
                    Dim val As System.String = m.Groups(1).Value.Trim()
                    If val.Length > best.Length Then best = val
                Next

                Return If(best.Length > 0, best, Nothing)
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' DTO for LLM exhibit number extraction response.
        ''' </summary>
        Private Class ExhibitNumberResponseDto
            <Newtonsoft.Json.JsonProperty("exhibits")>
            Public Property Exhibits As System.Collections.Generic.List(Of ExhibitFileDto)
        End Class

        Private Class ExhibitFileDto
            <Newtonsoft.Json.JsonProperty("filename")>
            Public Property Filename As System.String

            <Newtonsoft.Json.JsonProperty("number")>
            Public Property Number As System.String
        End Class

        ' LLM batch size: number of filenames sent per LLM call
        Private Const LlmBatchSize As System.Int32 = 50

        ''' <summary>
        ''' Extracts exhibit numbers from an array of PDF file paths using the LLM.
        ''' Sends filenames in batches for efficiency. Returns a dictionary mapping
        ''' full file path to extracted exhibit number.
        ''' </summary>
        ''' <param name="filePaths">Array of full PDF file paths.</param>
        ''' <returns>Dictionary(Of filePath, exhibitNumber). Entries with no result are omitted.</returns>
        Public Shared Async Function ExtractExhibitNumbersViaLlmAsync(
                filePaths As System.String()) As System.Threading.Tasks.Task(Of System.Collections.Generic.Dictionary(Of System.String, System.String))

            Dim result As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.OrdinalIgnoreCase)
            If filePaths Is Nothing OrElse filePaths.Length = 0 Then Return result

            ' Build a map from filename-only to full path (for reverse lookup)
            Dim nameToPath As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.OrdinalIgnoreCase)
            For Each fp As System.String In filePaths
                Dim fn As System.String = System.IO.Path.GetFileName(fp)
                If Not nameToPath.ContainsKey(fn) Then
                    nameToPath(fn) = fp
                End If
            Next

            Dim systemPrompt As System.String = SP_ExhibitNumber

            ' Process in batches
            Dim totalFiles As System.Int32 = filePaths.Length
            Dim batchStart As System.Int32 = 0

            While batchStart < totalFiles
                Dim batchEnd As System.Int32 = System.Math.Min(batchStart + LlmBatchSize, totalFiles) - 1
                Dim sb As New System.Text.StringBuilder()
                sb.AppendLine("Extract the exhibit numbers from these filenames:")
                For idx As System.Int32 = batchStart To batchEnd
                    sb.AppendLine(System.IO.Path.GetFileName(filePaths(idx)))
                Next

                Dim llmResponse As System.String = Await LLM(systemPrompt, sb.ToString(), "", "", 0, False, True)
                llmResponse = WebAgentInterpreter.SanitizeLlmResult(llmResponse)

                If Not System.String.IsNullOrWhiteSpace(llmResponse) Then
                    Try
                        Dim parsed As ExhibitNumberResponseDto =
                            Newtonsoft.Json.JsonConvert.DeserializeObject(Of ExhibitNumberResponseDto)(llmResponse)

                        If parsed IsNot Nothing AndAlso parsed.Exhibits IsNot Nothing Then
                            For Each item As ExhibitFileDto In parsed.Exhibits
                                If item Is Nothing OrElse System.String.IsNullOrWhiteSpace(item.Filename) Then Continue For
                                If System.String.IsNullOrWhiteSpace(item.Number) Then Continue For

                                Dim fullPath As System.String = Nothing
                                If nameToPath.TryGetValue(item.Filename, fullPath) Then
                                    result(fullPath) = item.Number.Trim()
                                End If
                            Next
                        End If
                    Catch
                        ' Ignore parse errors for this batch; continue with next
                    End Try
                End If

                batchStart = batchEnd + 1
            End While

            Return result
        End Function

        ''' <summary>
        ''' Configures PdfSharp global font resolver to use Arial from Windows Fonts folder.
        ''' Falls back to Base-14 fonts (Helvetica) if Arial is unavailable.
        ''' Mirrors PdfRedactionService.EnsurePdfSharpFontResolver which is Private.
        ''' </summary>
        ''' <returns>True if Arial configured; False if using Base-14 fallback.</returns>
        Private Shared Function EnsureFontResolver() As System.Boolean
            ' If already configured, keep it.
            If PdfSharp.Fonts.GlobalFontSettings.FontResolver IsNot Nothing Then
                Return True
            End If

            ' Try to locate Arial on Windows
            Dim fontsDir As System.String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Fonts)
            Dim arialRegular As System.String = System.IO.Path.Combine(fontsDir, "arial.ttf")
            If System.IO.File.Exists(arialRegular) Then
                PdfSharp.Fonts.GlobalFontSettings.FontResolver = New PdfRedactionService.ArialFontResolver()
                Return True
            End If

            ' No Arial installed → use Base14 later (Helvetica) without resolver
            Return False
        End Function


        ''' <summary>
        ''' Counts pages in a PDF file using PdfPig (read-only, handles most PDF types).
        ''' Returns -1 if the file cannot be read.
        ''' </summary>
        ''' <param name="pdfPath">Full path to the PDF file.</param>
        ''' <returns>Number of pages, or -1 on failure.</returns>
        Public Shared Function GetPdfPageCount(pdfPath As System.String) As System.Int32
            Try
                Using doc As UglyToad.PdfPig.PdfDocument = UglyToad.PdfPig.PdfDocument.Open(pdfPath)
                    Return doc.NumberOfPages
                End Using
            Catch
                Return -1
            End Try
        End Function

        ''' <summary>
        ''' Compares the page count of an original PDF with its stamped output.
        ''' Returns a warning string if they differ, or Nothing if they match.
        ''' </summary>
        ''' <param name="originalPath">Path to the original input PDF.</param>
        ''' <param name="stampedPath">Path to the stamped output PDF.</param>
        ''' <returns>Warning message if page counts differ; Nothing if OK.</returns>
        Public Shared Function VerifyPageCount(originalPath As System.String,
                                                stampedPath As System.String) As System.String
            Dim origPages As System.Int32 = GetPdfPageCount(originalPath)
            Dim stampedPages As System.Int32 = GetPdfPageCount(stampedPath)

            If origPages = -1 OrElse stampedPages = -1 Then
                Return "Could not verify page count (original=" &
                       If(origPages = -1, "unreadable", origPages.ToString()) &
                       ", stamped=" &
                       If(stampedPages = -1, "unreadable", stampedPages.ToString()) & ")."
            End If

            If origPages <> stampedPages Then
                Return "PAGE COUNT MISMATCH: original has " & origPages.ToString() &
                       " page(s) but stamped output has " & stampedPages.ToString() &
                       " page(s). Please verify the output file."
            End If

            Return Nothing
        End Function


        ''' <summary>
        ''' Result of a stamp operation, including the output path and any warnings.
        ''' </summary>
        Public Class StampResult
            ''' <summary>Full path to the stamped output PDF.</summary>
            Public Property OutputPath As System.String
            ''' <summary>Non-empty if the PDF had to be rasterized (e.g., encrypted or problematic source).</summary>
            Public Property RasterizeWarning As System.String
        End Class

        ''' <summary>
        ''' Applies the exhibit stamp image and exhibit number text to the first page of a PDF.
        ''' When ForceRasterize is set, always uses the rasterize path for guaranteed visibility.
        ''' Otherwise tries direct PdfSharp stamping first, verifies output grew, and falls back
        ''' to rasterize-then-stamp on any failure.
        ''' </summary>
        Public Shared Function ApplyStamp(inputPath As System.String,
                                          exhibitNumber As System.String,
                                          config As StampConfig,
                                          Optional outputPathOverride As System.String = Nothing) As StampResult

            ' Compute output path if not provided
            Dim outputPath As System.String = outputPathOverride
            If System.String.IsNullOrWhiteSpace(outputPath) Then
                Dim baseDir As System.String = System.IO.Path.GetDirectoryName(inputPath)
                If Not System.String.IsNullOrWhiteSpace(config.OutputSubdir) Then
                    baseDir = System.IO.Path.Combine(baseDir, config.OutputSubdir)
                    If Not System.IO.Directory.Exists(baseDir) Then
                        System.IO.Directory.CreateDirectory(baseDir)
                    End If
                End If
                Dim nameNoExt As System.String = System.IO.Path.GetFileNameWithoutExtension(inputPath)
                Dim sfx As System.String = If(Not System.String.IsNullOrWhiteSpace(config.OutputExtension),
                                              config.OutputExtension, "")
                If System.String.IsNullOrWhiteSpace(config.OutputSubdir) AndAlso System.String.IsNullOrWhiteSpace(sfx) Then
                    sfx = "_stamped"
                End If
                outputPath = System.IO.Path.Combine(baseDir, nameNoExt & sfx & ".pdf")
            End If

            ' Build the full exhibit label text
            Dim labelText As System.String = config.Prefix & exhibitNumber & config.Suffix

            ' Ensure font resolver is set up
            Dim canUseArial As System.Boolean = EnsureFontResolver()

            ' Load stamp image into memory once
            Dim imgBytes As System.Byte() = System.IO.File.ReadAllBytes(config.StampImagePath)

            Dim result As New StampResult()
            result.OutputPath = outputPath

            ' Detect encrypted PDFs early — PdfSharp may "succeed" on owner-password PDFs
            ' but produce an invisible stamp overlay, so force rasterize for these
            Dim forceRasterizeDueToEncryption As System.Boolean = False
            If Not config.ForceRasterize Then
                forceRasterizeDueToEncryption = IsPdfEncrypted(inputPath)
            End If

            If config.ForceRasterize OrElse forceRasterizeDueToEncryption Then
                ' Rasterize mode — either user-requested or due to encryption detection
                ApplyStampViaRasterize(inputPath, outputPath, labelText, canUseArial, imgBytes, config)
                If forceRasterizeDueToEncryption Then
                    result.RasterizeWarning = "PDF is encrypted or restricted — was rasterized to ensure stamp visibility (text no longer selectable)."
                End If
                ValidateOutput(inputPath, outputPath)
                Return result
            End If

            ' Record original file size for verification
            Dim inputSize As System.Int64 = New System.IO.FileInfo(inputPath).Length

            ' Try direct PdfSharp stamping first; on ANY failure, fall back to rasterize-then-stamp
            Dim directSuccess As System.Boolean = False
            Try
                directSuccess = ApplyStampDirect(inputPath, outputPath, labelText, canUseArial, imgBytes, config)
            Catch
                directSuccess = False
            End Try

            ' Verify: output must exist and be larger than input (stamp adds image + text data)
            If directSuccess Then
                Try
                    Dim outputSize As System.Int64 = New System.IO.FileInfo(outputPath).Length
                    If outputSize <= inputSize Then
                        directSuccess = False
                    End If
                Catch
                    directSuccess = False
                End Try
            End If

            If Not directSuccess Then
                ' Clean up failed direct output before retrying
                If System.IO.File.Exists(outputPath) Then
                    Try : System.IO.File.Delete(outputPath) : Catch : End Try
                End If

                ' Fallback: rasterize page 1 via PdfiumViewer, then stamp the clean rasterized page
                Try
                    ApplyStampViaRasterize(inputPath, outputPath, labelText, canUseArial, imgBytes, config)
                    result.RasterizeWarning = "PDF could not be stamped directly (may be encrypted or restricted) — was rasterized instead (text no longer selectable)."
                Catch rasterEx As System.Exception
                    ' Both direct and rasterize failed — throw a clear error
                    Dim msg As System.String = rasterEx.Message
                    If System.String.IsNullOrWhiteSpace(msg) OrElse
                       msg.Equals("No error", System.StringComparison.OrdinalIgnoreCase) OrElse
                       msg.Equals("No error.", System.StringComparison.OrdinalIgnoreCase) Then
                        msg = "PDF appears to be encrypted or corrupt and could not be processed"
                    End If
                    Throw New System.InvalidOperationException(msg, rasterEx)
                End Try
            End If

            ValidateOutput(inputPath, outputPath)
            Return result

        End Function


        ''' <summary>
        ''' Checks whether a PDF file is encrypted by attempting to open it with PdfSharp
        ''' and inspecting its security settings, then falling back to a chunked byte-level
        ''' scan of the entire file for the /Encrypt marker.
        ''' </summary>
        ''' <param name="pdfPath">Full path to the PDF file.</param>
        ''' <returns>True if the PDF appears to contain encryption; False otherwise.</returns>
        Private Shared Function IsPdfEncrypted(pdfPath As System.String) As System.Boolean
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
                Const chunkSize As System.Int32 = 65536
                Dim overlap As System.Int32 = 7 ' Length of "/Encrypt" minus 1

                Using fs As New System.IO.FileStream(pdfPath, System.IO.FileMode.Open,
                                                      System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite)
                    Dim buffer(chunkSize + overlap - 1) As System.Byte
                    Dim carryOver As System.Int32 = 0

                    While True
                        Dim bytesRead As System.Int32 = fs.Read(buffer, carryOver, chunkSize)
                        If bytesRead = 0 Then Exit While

                        Dim totalInBuffer As System.Int32 = carryOver + bytesRead
                        Dim text As System.String = System.Text.Encoding.ASCII.GetString(buffer, 0, totalInBuffer)
                        If text.IndexOf("/Encrypt", System.StringComparison.Ordinal) >= 0 Then
                            Return True
                        End If

                        ' Keep the last few bytes for overlap into next chunk
                        If totalInBuffer > overlap Then
                            System.Array.Copy(buffer, totalInBuffer - overlap, buffer, 0, overlap)
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
        ''' Validates that the stamped output file exists and is a plausible PDF.
        ''' Throws an <see cref="System.InvalidOperationException"/> if the output
        ''' is missing, empty, or the page count cannot be verified — which typically
        ''' indicates the source PDF is encrypted or corrupt.
        ''' </summary>
        ''' <param name="inputPath">Path to the original input PDF (for error messages).</param>
        ''' <param name="outputPath">Path to the stamped output PDF to validate.</param>
        Private Shared Sub ValidateOutput(inputPath As System.String, outputPath As System.String)
            If Not System.IO.File.Exists(outputPath) Then
                Throw New System.InvalidOperationException(
                    "Stamped output file was not created — the source PDF may be encrypted or corrupt.")
            End If

            Dim outputSize As System.Int64 = New System.IO.FileInfo(outputPath).Length
            If outputSize = 0 Then
                Try : System.IO.File.Delete(outputPath) : Catch : End Try
                Throw New System.InvalidOperationException(
                    "Stamped output file is empty — the source PDF may be encrypted or corrupt.")
            End If

            ' Verify the output is a readable PDF with at least one page
            Dim pageCount As System.Int32 = GetPdfPageCount(outputPath)
            If pageCount = 0 Then
                Try : System.IO.File.Delete(outputPath) : Catch : End Try
                Throw New System.InvalidOperationException(
                    "Stamped output contains no pages — the source PDF may be encrypted or corrupt.")
            End If
        End Sub

        ''' <summary>
        ''' Direct stamping: opens PDF in Modify mode, draws stamp on existing page.
        ''' Returns True on success, False if stamping should fall back to rasterize.
        ''' </summary>
        Private Shared Function ApplyStampDirect(inputPath As System.String,
                                                  outputPath As System.String,
                                                  labelText As System.String,
                                                  canUseArial As System.Boolean,
                                                  imgBytes As System.Byte(),
                                                  config As StampConfig) As System.Boolean

            Dim document As PdfSharp.Pdf.PdfDocument = Nothing
            Try
                document = PdfSharp.Pdf.IO.PdfReader.Open(inputPath, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Modify)
            Catch
                Return False ' Cannot open → fall back
            End Try

            Try
                If document.Pages.Count = 0 Then Return False

                Dim page As PdfSharp.Pdf.PdfPage = document.Pages(0)
                Dim rotation As System.Int32 = page.Rotate

                Dim pageW As System.Double
                Dim pageH As System.Double
                If rotation = 90 OrElse rotation = 270 Then
                    pageW = page.Height.Point
                    pageH = page.Width.Point
                Else
                    pageW = page.Width.Point
                    pageH = page.Height.Point
                End If

                ' Sanity: if dimensions look wrong, bail to rasterize fallback
                If pageW <= 0 OrElse pageH <= 0 OrElse pageW > 20000 OrElse pageH > 20000 Then
                    Return False
                End If

                Using gfx As PdfSharp.Drawing.XGraphics =
                        PdfSharp.Drawing.XGraphics.FromPdfPage(page, PdfSharp.Drawing.XGraphicsPdfPageOptions.Append)

                    Dim gfxState As PdfSharp.Drawing.XGraphicsState = gfx.Save()

                    DrawStampContent(gfx, pageW, labelText, canUseArial, imgBytes, config)

                    gfx.Restore(gfxState)
                End Using

                document.Save(outputPath)
                Return True

            Finally
                If document IsNot Nothing Then
                    document.Close()
                    document.Dispose()
                End If
            End Try
        End Function

        ''' <summary>
        ''' Rasterize fallback: renders page 1 via Windows.Data.Pdf (preferred) or PdfiumViewer,
        ''' creates a clean page, stamps it, then copies all remaining pages as-is from the
        ''' original PDF. Handles password-protected, corrupt, and problematic PDFs.
        ''' </summary>
        Private Shared Sub ApplyStampViaRasterize(inputPath As System.String,
                                                    outputPath As System.String,
                                                    labelText As System.String,
                                                    canUseArial As System.Boolean,
                                                    imgBytes As System.Byte(),
                                                    config As StampConfig)

            Dim tempPath As System.String = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                System.Guid.NewGuid().ToString("N") & ".pdf")

            Try
                Dim totalPageCount As System.Int32 = 0
                Dim outDoc As New PdfSharp.Pdf.PdfDocument()
                Const renderDpi As System.Int32 = 200

                ' ── Try Windows.Data.Pdf first for page 0 ──
                Dim usedWindowsPdf As System.Boolean = False
                Dim page0Result As System.Tuple(Of System.IO.MemoryStream, System.Double, System.Double) = Nothing

                Try
                    page0Result = RenderPageViaWindowsPdf(inputPath, 0, renderDpi)
                Catch
                End Try

                If page0Result IsNot Nothing Then
                    usedWindowsPdf = True
                    Dim jpegMs As System.IO.MemoryStream = page0Result.Item1
                    Dim pageWidthPt As System.Double = page0Result.Item2
                    Dim pageHeightPt As System.Double = page0Result.Item3

                    Dim outPage As PdfSharp.Pdf.PdfPage = outDoc.AddPage()
                    outPage.Width = PdfSharp.Drawing.XUnit.FromPoint(pageWidthPt)
                    outPage.Height = PdfSharp.Drawing.XUnit.FromPoint(pageHeightPt)

                    Using xgfx As PdfSharp.Drawing.XGraphics = PdfSharp.Drawing.XGraphics.FromPdfPage(outPage)
                        Using ximg As PdfSharp.Drawing.XImage = PdfSharp.Drawing.XImage.FromStream(jpegMs)
                            xgfx.DrawImage(ximg, 0, 0, outPage.Width.Point, outPage.Height.Point)
                        End Using
                    End Using
                    jpegMs.Dispose()

                    totalPageCount = ExhibitStampService.GetPdfPageCount(inputPath)
                    If totalPageCount < 1 Then totalPageCount = 1
                End If

                ' ── Fallback to PdfiumViewer if Windows.Data.Pdf failed ──
                If Not usedWindowsPdf Then
                    PdfRedactionService.EnsurePdfiumLoadedPublic()

                    Using pdf As PdfiumViewer.PdfDocument = PdfiumViewer.PdfDocument.Load(inputPath)
                        totalPageCount = pdf.PageCount
                        If totalPageCount = 0 Then
                            Throw New System.InvalidOperationException("The PDF contains no pages.")
                        End If

                        Dim sizePt As System.Drawing.SizeF = pdf.PageSizes(0)
                        Dim widthPx As System.Int32 = CInt(System.Math.Round(sizePt.Width / 72.0 * renderDpi))
                        Dim heightPx As System.Int32 = CInt(System.Math.Round(sizePt.Height / 72.0 * renderDpi))

                        Dim renderFlags As PdfiumViewer.PdfRenderFlags =
                            PdfiumViewer.PdfRenderFlags.Annotations Or
                            PdfiumViewer.PdfRenderFlags.LcdText Or
                            PdfiumViewer.PdfRenderFlags.ForPrinting

                        Using rendered As System.Drawing.Image = pdf.Render(0, widthPx, heightPx, renderDpi, renderDpi, renderFlags)
                            Dim outPage As PdfSharp.Pdf.PdfPage = outDoc.AddPage()
                            outPage.Width = PdfSharp.Drawing.XUnit.FromPoint(sizePt.Width)
                            outPage.Height = PdfSharp.Drawing.XUnit.FromPoint(sizePt.Height)

                            Using ms As New System.IO.MemoryStream()
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
                    End Using
                End If

                ' Draw stamp on the rasterized page 0
                Using gfx As PdfSharp.Drawing.XGraphics =
                        PdfSharp.Drawing.XGraphics.FromPdfPage(outDoc.Pages(0), PdfSharp.Drawing.XGraphicsPdfPageOptions.Append)
                    DrawStampContent(gfx, outDoc.Pages(0).Width.Point, labelText, canUseArial, imgBytes, config)
                End Using

                ' Copy remaining pages (1..N) as-is from the original PDF
                If totalPageCount > 1 Then
                    Dim srcDoc As PdfSharp.Pdf.PdfDocument = Nothing
                    Try
                        srcDoc = PdfSharp.Pdf.IO.PdfReader.Open(inputPath, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import)
                        For pageIdx As System.Int32 = 1 To srcDoc.PageCount - 1
                            outDoc.AddPage(srcDoc.Pages(pageIdx))
                        Next
                    Catch
                        ' If Import mode fails, rasterize all remaining pages
                        For pageIdx As System.Int32 = 1 To totalPageCount - 1
                            Dim pageResult As System.Tuple(Of System.IO.MemoryStream, System.Double, System.Double) = Nothing
                            Try
                                pageResult = RenderPageViaWindowsPdf(inputPath, pageIdx, renderDpi)
                            Catch
                            End Try

                            If pageResult IsNot Nothing Then
                                Dim pg As PdfSharp.Pdf.PdfPage = outDoc.AddPage()
                                pg.Width = PdfSharp.Drawing.XUnit.FromPoint(pageResult.Item2)
                                pg.Height = PdfSharp.Drawing.XUnit.FromPoint(pageResult.Item3)

                                Using pgGfx As PdfSharp.Drawing.XGraphics = PdfSharp.Drawing.XGraphics.FromPdfPage(pg)
                                    Using pgImg As PdfSharp.Drawing.XImage = PdfSharp.Drawing.XImage.FromStream(pageResult.Item1)
                                        pgGfx.DrawImage(pgImg, 0, 0, pg.Width.Point, pg.Height.Point)
                                    End Using
                                End Using
                                pageResult.Item1.Dispose()
                            Else
                                ' Final fallback: PdfiumViewer
                                PdfRedactionService.EnsurePdfiumLoadedPublic()
                                Using pdf2 As PdfiumViewer.PdfDocument = PdfiumViewer.PdfDocument.Load(inputPath)
                                    Dim pgSize As System.Drawing.SizeF = pdf2.PageSizes(pageIdx)
                                    Dim pgWPx As System.Int32 = CInt(System.Math.Round(pgSize.Width / 72.0 * renderDpi))
                                    Dim pgHPx As System.Int32 = CInt(System.Math.Round(pgSize.Height / 72.0 * renderDpi))

                                    Dim pgRenderFlags As PdfiumViewer.PdfRenderFlags =
                                        PdfiumViewer.PdfRenderFlags.Annotations Or
                                        PdfiumViewer.PdfRenderFlags.LcdText Or
                                        PdfiumViewer.PdfRenderFlags.ForPrinting

                                    Using pgRendered As System.Drawing.Image = pdf2.Render(pageIdx, pgWPx, pgHPx, renderDpi, renderDpi, pgRenderFlags)
                                        Dim pg As PdfSharp.Pdf.PdfPage = outDoc.AddPage()
                                        pg.Width = PdfSharp.Drawing.XUnit.FromPoint(pgSize.Width)
                                        pg.Height = PdfSharp.Drawing.XUnit.FromPoint(pgSize.Height)

                                        Using pgMs As New System.IO.MemoryStream()
                                            Dim jpgEnc As System.Drawing.Imaging.ImageCodecInfo = Nothing
                                            For Each codec In System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders()
                                                If codec.MimeType = "image/jpeg" Then jpgEnc = codec : Exit For
                                            Next
                                            If jpgEnc IsNot Nothing Then
                                                Dim ep2 As New System.Drawing.Imaging.EncoderParameters(1)
                                                ep2.Param(0) = New System.Drawing.Imaging.EncoderParameter(
                                                    System.Drawing.Imaging.Encoder.Quality, 85L)
                                                pgRendered.Save(pgMs, jpgEnc, ep2)
                                            Else
                                                pgRendered.Save(pgMs, System.Drawing.Imaging.ImageFormat.Png)
                                            End If

                                            pgMs.Position = 0
                                            Using pgGfx As PdfSharp.Drawing.XGraphics = PdfSharp.Drawing.XGraphics.FromPdfPage(pg)
                                                Using pgImg As PdfSharp.Drawing.XImage = PdfSharp.Drawing.XImage.FromStream(pgMs)
                                                    pgGfx.DrawImage(pgImg, 0, 0, pg.Width.Point, pg.Height.Point)
                                                End Using
                                            End Using
                                        End Using
                                    End Using
                                End Using
                            End If
                        Next
                    End Try
                End If

                outDoc.Save(outputPath)
                outDoc.Close()

            Finally
                If System.IO.File.Exists(tempPath) Then
                    Try : System.IO.File.Delete(tempPath) : Catch : End Try
                End If
            End Try
        End Sub

        ''' <summary>
        ''' Draws the stamp image and exhibit number text onto an XGraphics surface.
        ''' Shared by both direct and rasterize paths.
        ''' </summary>
        Private Shared Sub DrawStampContent(gfx As PdfSharp.Drawing.XGraphics,
                                             pageW As System.Double,
                                             labelText As System.String,
                                             canUseArial As System.Boolean,
                                             imgBytes As System.Byte(),
                                             config As StampConfig)

            ' ── Draw stamp image ──
            Using ms As New System.IO.MemoryStream(imgBytes)
                ms.Position = 0
                Using ximg As PdfSharp.Drawing.XImage = PdfSharp.Drawing.XImage.FromStream(ms)
                    Dim imgPixelW As System.Double = System.Math.Max(1, ximg.PixelWidth)
                    Dim imgPixelH As System.Double = System.Math.Max(1, ximg.PixelHeight)
                    Dim aspectRatio As System.Double = imgPixelH / imgPixelW
                    Dim stampWidth As System.Double = System.Math.Max(1, config.StampWidthPts)
                    Dim stampHeight As System.Double = System.Math.Max(1, stampWidth * aspectRatio)

                    Dim imgX As System.Double = System.Math.Max(0, pageW - config.RightOffsetPts - stampWidth)
                    Dim imgY As System.Double = System.Math.Max(0, config.TopOffsetPts)

                    gfx.DrawImage(ximg, imgX, imgY, stampWidth, stampHeight)
                End Using
            End Using

            ' ── Draw exhibit number text (right-aligned) ──
            If Not System.String.IsNullOrWhiteSpace(labelText) Then
                Dim family As System.String = If(canUseArial, "Arial", "Helvetica")

                Dim style As PdfSharp.Drawing.XFontStyleEx
                Try
                    style = CType([Enum].Parse(GetType(PdfSharp.Drawing.XFontStyleEx), "Bold", True), PdfSharp.Drawing.XFontStyleEx)
                Catch
                    style = CType([Enum].ToObject(GetType(PdfSharp.Drawing.XFontStyleEx), 1), PdfSharp.Drawing.XFontStyleEx)
                End Try

                Dim font As New PdfSharp.Drawing.XFont(family, config.TextSizePts, style)
                Dim textMeasure As PdfSharp.Drawing.XSize = gfx.MeasureString(labelText, font)

                Dim textX As System.Double = System.Math.Max(0, pageW - config.TextRightOffsetPts - textMeasure.Width)
                Dim textY As System.Double = System.Math.Max(0, config.TextTopOffsetPts)

                gfx.DrawString(labelText, font,
                               PdfSharp.Drawing.XBrushes.Black,
                               New PdfSharp.Drawing.XPoint(textX, textY + textMeasure.Height),
                               PdfSharp.Drawing.XStringFormats.Default)
            End If
        End Sub


    End Class

End Class
