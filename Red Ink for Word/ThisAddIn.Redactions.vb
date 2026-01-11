' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.Redactions.vb
' Purpose: LLM-powered PDF redaction preparation and finalization. Prepares
'          removable redaction annotations using AI-detected text ranges; 
'          finalizes by flattening annotations into rasterized images.
'
' Architecture:
'  - Instruction Library: Loads local/global *.txt instruction sets; each line "Title|Instruction"
'    (lines starting with ";" ignored). Manual instruction overrides prepared instruction.
'  - UI Parameter Collection: Uses SharedMethods.InputParameter and ShowCustomVariableInputForm
'    to collect redaction parameters (prepared/manual instruction, transparency, reason codes,
'    metadata removal, output paths, optional secondary model).
'  - Model Switching: Optional alternate model or secondary API selection based on configuration flags.
'  - Text Extraction (PDF): Uses UglyToad.PdfPig to extract text and character-level positions
'    from PDF pages, maintaining spatial layout for accurate box placement.
'  - LLM Invocation: Calls LLM(systemPrompt, userPrompt, ...) expecting JSON with redaction ranges
'    (start/end indices or exact_text). Sanitizes and parses response.
'  - Text Matching: Supports exact and normalized (whitespace-tolerant) text matching to map
'    LLM-provided text snippets to character positions.
'  - Redaction Box Generation: Converts character ranges to PDF rectangles with padding,
'    grouped by page and line; coordinates transformed from PdfPig (bottom-left) to PdfSharp (top-left).
'  - Annotation Creation: Adds /Square PDF annotations (removable boxes) in red (transparent) or black;
'    optionally includes reason codes in /Contents field; optionally strips metadata.
'  - Finalization (Flattening): Burns annotations into rasterized images using PdfiumViewer,
'    preserving metadata; sticky notes and popup annotations excluded from burning.
'  - Progress & Cancellation: ProgressBarModule tracks state in multi-file mode; user can abort mid-iteration.
'  - Font Handling: ArialFontResolver implements PdfSharp.Fonts.IFontResolver for rendering reason
'    code labels when flattening; falls back to Base-14 Helvetica if Arial unavailable.
'  - External Dependencies: SharedLibrary.SharedMethods supplies UI, LLM, progress, model selection;
'    UglyToad.PdfPig for text extraction; PdfSharp for PDF manipulation; PdfiumViewer for rendering.
' =============================================================================

Option Explicit On
Option Strict On

Imports System.Data
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Microsoft.Office.Interop.Word
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods
Imports PDFP = UglyToad.PdfPig

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Prepares PDF redactions by analyzing text with LLM and adding removable annotation boxes.
    ''' Supports single-file and multi-file modes with configurable parameters.
    ''' </summary>
    Public Async Sub PrepareRedactedPDF()
        If INILoadFail() Then Return

        Dim do2ndModel As Boolean = False
        Dim useSecondApi As Boolean = False

        Try
            ' 1) Load prepared instruction sets from local and global files/folders
            Dim displayToInstruction As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.OrdinalIgnoreCase)
            Dim displayOptions As New System.Collections.Generic.List(Of System.String)()

            Dim localPath As System.String = ExpandEnvironmentVariables(INI_RedactionInstructionsPathLocal)
            Dim globalPath As System.String = ExpandEnvironmentVariables(INI_RedactionInstructionsPath)

            ' Helper: enumerate files from a path that may be a folder or a single file
            Dim EnumerateInstructionFiles As System.Func(Of System.String, System.Collections.Generic.IEnumerable(Of System.String)) =
                Function(p As System.String) As System.Collections.Generic.IEnumerable(Of System.String)
                    Dim result As New System.Collections.Generic.List(Of System.String)()
                    If System.String.IsNullOrWhiteSpace(p) Then Return result
                    Try
                        If System.IO.Directory.Exists(p) Then
                            result.AddRange(System.IO.Directory.GetFiles(p, "*.txt", SearchOption.TopDirectoryOnly))
                        ElseIf System.IO.File.Exists(p) Then
                            result.Add(p)
                        End If
                    Catch
                        ' ignore
                    End Try
                    Return result
                End Function

            ' Helper: parse a single file and add entries
            Dim LoadFromFile As System.Action(Of System.String, System.Boolean) =
                Sub(file As System.String, isLocal As System.Boolean)
                    Try
                        For Each rawLine As System.String In System.IO.File.ReadAllLines(file)
                            Dim line As System.String = If(rawLine, "").Trim()
                            If line.Length = 0 OrElse line.StartsWith(";", System.StringComparison.Ordinal) Then Continue For
                            Dim barIdx As Integer = line.IndexOf("|"c)
                            If barIdx <= 0 Then Continue For
                            Dim title As System.String = line.Substring(0, barIdx).Trim()
                            Dim instr As System.String = line.Substring(barIdx + 1).Trim()
                            If title.Length = 0 OrElse instr.Length = 0 Then Continue For

                            Dim display As System.String = title
                            If isLocal Then display &= " (local)"
                            ' Ensure uniqueness of display text
                            Dim uniqueDisplay As System.String = MakeUniqueDisplay(display, displayToInstruction.Keys)
                            displayToInstruction(uniqueDisplay) = instr
                            displayOptions.Add(uniqueDisplay)
                        Next
                    Catch
                        ' ignore bad files
                    End Try
                End Sub

            ' Local first (as requested), then global
            For Each f As System.String In EnumerateInstructionFiles(localPath)
                LoadFromFile(f, True)
            Next
            For Each f As System.String In EnumerateInstructionFiles(globalPath)
                LoadFromFile(f, False)
            Next

            ' 2) Defaults from My.Settings (best-effort; tolerate missing properties)
            Dim defaultManual As System.String = ""
            Dim defaultSubdir As System.String = ""
            Dim defaultExtension As System.String = ""
            Try
                Dim v = My.Settings("Redaction_ManualInstruction")
                If v IsNot Nothing Then defaultManual = System.Convert.ToString(v)
            Catch
            End Try
            Try
                Dim v = My.Settings("Redaction_OutputSubdir")
                If v IsNot Nothing AndAlso System.Convert.ToString(v).Trim().Length > 0 Then defaultSubdir = System.Convert.ToString(v)
            Catch
            End Try
            Try
                Dim v = My.Settings("Redaction_OutputExtension")
                If v IsNot Nothing AndAlso System.Convert.ToString(v).Trim().Length > 0 Then defaultExtension = System.Convert.ToString(v)
            Catch
            End Try

            Dim defaultPreparedDisplay As System.String = If(displayOptions.Count > 0, displayOptions(0), "<<disabled>>")
            Dim transparentBoxes As System.Boolean = True
            Dim includeReasonCodes As System.Boolean = False
            Dim removeMetadata As System.Boolean = True
            Dim multipleFiles As System.Boolean = False
            Dim outSubdir As System.String = defaultSubdir
            Dim outExtension As System.String = defaultExtension
            Dim manualInstruction As System.String = defaultManual

            ' 3) Collect parameters in ONE form
            Dim p0 As SLib.InputParameter = New SLib.InputParameter("Prepared instruction set", defaultPreparedDisplay)
            If displayOptions.Count > 0 Then
                p0.Options = New System.Collections.Generic.List(Of System.String)(displayOptions)
            End If
            Dim p1 As SLib.InputParameter
            If displayOptions.Count = 0 Then
                p1 = New SLib.InputParameter("Redaction instructions", manualInstruction)
            Else
                p1 = New SLib.InputParameter("Manual instruction (overrides)", manualInstruction)
            End If

            Dim p2 As SLib.InputParameter = New SLib.InputParameter("Transparent redaction boxes", transparentBoxes)
            Dim p3 As SLib.InputParameter = New SLib.InputParameter("Include reason codes", includeReasonCodes)
            Dim p4 As SLib.InputParameter = New SLib.InputParameter("Remove metadata from PDFs", removeMetadata)
            Dim p5 As SLib.InputParameter = New SLib.InputParameter("Multiple files", multipleFiles)
            Dim p6 As SLib.InputParameter = New SLib.InputParameter("Output subdirectory", outSubdir)
            Dim p7 As SLib.InputParameter = New SLib.InputParameter("Output filename extension", outExtension)

            ' Offer secondary model toggle only if configured
            Dim hasSecondary As Boolean = (Not String.IsNullOrWhiteSpace(INI_AlternateModelPath)) OrElse INI_SecondAPI
            Dim p8 As SLib.InputParameter = Nothing
            If hasSecondary Then
                do2ndModel = False
                If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                    p8 = New SLib.InputParameter("Use a secondary model", do2ndModel)
                Else
                    p8 = New SLib.InputParameter("Use the secondary model", do2ndModel)
                End If
            End If

            Dim params() As SLib.InputParameter =
                If(hasSecondary,
                   New SLib.InputParameter() {p0, p1, p2, p3, p4, p5, p6, p7, p8},
                   New SLib.InputParameter() {p0, p1, p2, p3, p4, p5, p6, p7})

            If ShowCustomVariableInputForm("Please set the redaction parameters:", AN & " Prepare PDF Redactions", params) = False Then
                Return
            End If

            ' Read back values
            Dim chosenPreparedDisplay As System.String = System.Convert.ToString(params(0).Value)
            manualInstruction = System.Convert.ToString(params(1).Value)
            Try : transparentBoxes = System.Convert.ToBoolean(params(2).Value) : Catch : transparentBoxes = True : End Try
            Try : includeReasonCodes = System.Convert.ToBoolean(params(3).Value) : Catch : includeReasonCodes = False : End Try
            Try : removeMetadata = System.Convert.ToBoolean(params(4).Value) : Catch : removeMetadata = True : End Try
            Try : multipleFiles = System.Convert.ToBoolean(params(5).Value) : Catch : multipleFiles = False : End Try
            outSubdir = System.Convert.ToString(params(6).Value)
            outExtension = System.Convert.ToString(params(7).Value)
            If hasSecondary Then
                Dim secVal = params(8).Value
                If TypeOf secVal Is Boolean Then do2ndModel = CBool(secVal) Else do2ndModel = False
            End If

            ' Persist selected manual/subdir/extension in My.Settings (best-effort)
            Try : My.Settings("Redaction_ManualInstruction") = manualInstruction : Catch : End Try
            Try : My.Settings("Redaction_OutputSubdir") = outSubdir : Catch : End Try
            Try : My.Settings("Redaction_OutputExtension") = outExtension : Catch : End Try
            Try : My.Settings.Save() : Catch : End Try

            ' 4) Resolve final instruction (manual takes precedence)
            Dim preparedInstruction As System.String = Nothing
            If chosenPreparedDisplay IsNot Nothing AndAlso displayToInstruction.TryGetValue(chosenPreparedDisplay, preparedInstruction) = False Then
                preparedInstruction = Nothing
            End If
            Dim effectiveInstruction As System.String = If(Not System.String.IsNullOrWhiteSpace(manualInstruction), manualInstruction, preparedInstruction)

            If System.String.IsNullOrWhiteSpace(effectiveInstruction) Then
                ShowCustomMessageBox("No redaction instruction was provided. Enter a manual instruction or choose a prepared one.")
                Return
            End If

            ' 4.5) Switch to secondary model if requested
            If hasSecondary AndAlso do2ndModel Then
                If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                    If Not ShowModelSelection(_context, INI_AlternateModelPath) Then
                        originalConfigLoaded = False
                        ShowCustomMessageBox("The secondary model could not be loaded - aborting.")
                        Return
                    Else
                        useSecondApi = True
                    End If
                ElseIf INI_SecondAPI Then
                    useSecondApi = True
                End If
            End If

            If multipleFiles = False Then
                ' === SINGLE-FILE MODE (Drag & Drop picker) ===
                Try
                    DragDropFormLabel = "PDF file (.pdf)"
                    DragDropFormFilter = "Supported Files|*.pdf"

                    Dim filePath As System.String = GetFileName()

                    DragDropFormLabel = ""
                    DragDropFormFilter = ""

                    If System.String.IsNullOrWhiteSpace(filePath) Then
                        ShowCustomMessageBox("No file has been selected - will abort.")
                        Return
                    End If
                    If Not System.IO.File.Exists(filePath) Then
                        ShowCustomMessageBox("The selected file does not exist - will abort.")
                        Return
                    End If
                    If Not System.String.Equals(System.IO.Path.GetExtension(filePath), ".pdf", System.StringComparison.OrdinalIgnoreCase) Then
                        ShowCustomMessageBox("The selected file is not a PDF - will abort.")
                        Return
                    End If

                    Await PdfRedactionService.RunPdfRedactionAsync(
                        filePath,
                        effectiveInstruction,
                        transparentBoxes,
                        includeReasonCodes,
                        removeMetadata,
                        outSubdir,
                        outExtension,
                        False,
                        useSecondApi)

                Catch ex As System.Exception
                    ShowCustomMessageBox("File selection failed: " & ex.Message)
                End Try

                DragDropFormLabel = ""
                DragDropFormFilter = ""

            Else
                ' === MULTI-FILE MODE ===
                Dim selectedFolder As System.String = Nothing
                Try
                    Using dlg As New FolderBrowserDialog()
                        dlg.Description = "Select the folder containing the PDF files to redact"
                        dlg.ShowNewFolderButton = False
                        Dim dr = dlg.ShowDialog()
                        If dr <> DialogResult.OK OrElse System.String.IsNullOrWhiteSpace(dlg.SelectedPath) Then
                            ShowCustomMessageBox("No folder has been selected - will abort.")
                            Return
                        End If
                        selectedFolder = dlg.SelectedPath
                    End Using
                Catch ex As System.Exception
                    ShowCustomMessageBox("Folder selection failed: " & ex.Message)
                    Return
                End Try

                Dim pdfFiles As System.String()
                Try
                    pdfFiles = System.IO.Directory.GetFiles(selectedFolder, "*.pdf", SearchOption.TopDirectoryOnly)
                Catch ex As System.Exception
                    ShowCustomMessageBox("Failed to enumerate PDFs: " & ex.Message)
                    Return
                End Try

                If pdfFiles Is Nothing OrElse pdfFiles.Length = 0 Then
                    ShowCustomMessageBox("Selected folder contains no PDF files - nothing to do.")
                    Return
                End If

                ' Progress bar setup
                ShowProgressBarInSeparateThread(AN & " Prepare PDF Redactions", "Preparing redactions...")
                ProgressBarModule.CancelOperation = False

                GlobalProgressMax = pdfFiles.Length
                GlobalProgressValue = 0
                GlobalProgressLabel = "Starting..."

                Dim processed As System.Int32 = 0
                Dim errors As New System.Collections.Generic.List(Of System.String)()
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

                    Try
                        Await PdfRedactionService.RunPdfRedactionAsync(
                            inPath,
                            effectiveInstruction,
                            transparentBoxes,
                            includeReasonCodes,
                            removeMetadata,
                            outSubdir,
                            outExtension,
                            True,
                            useSecondApi)
                        processed += 1
                    Catch ex As System.Exception
                        errors.Add(fileNameOnly & ": " & ex.Message)
                    End Try
                Next

                ' Close progress bar
                ProgressBarModule.CancelOperation = True

                ' Summary
                Dim msg As System.String
                If cancelled Then
                    msg = "Operation aborted by user." & vbCrLf
                Else
                    msg = "Redaction preparation completed. Make sure you check and finalize them before using the PDF (otherwise they can be removed)." & vbCrLf
                End If
                msg &= "Processed: " & processed.ToString() & " of " & pdfFiles.Length.ToString() & " file(s)." & If(outSubdir <> "", vbCrLf &
                       "Output subfolder: " & outSubdir, "")
                If errors.Count > 0 Then
                    msg &= vbCrLf & vbCrLf & "Errors (copied into the clipboard):" & vbCrLf & String.Join(vbCrLf, errors)
                    SLib.PutInClipboard(msg)
                End If
                ShowCustomMessageBox(msg)
            End If

        Catch ex As System.Exception
            ShowCustomMessageBox("Prepare PDF Redactions failed: " & ex.Message)
        Finally
            If do2ndModel AndAlso originalConfigLoaded Then
                RestoreDefaults(_context, originalConfig)
                originalConfigLoaded = False
            End If
        End Try
    End Sub

    ''' <summary>
    ''' Finalizes PDF redactions by flattening annotations into rasterized images.
    ''' Converts annotation boxes to burned-in black rectangles; preserves sticky notes.
    ''' Supports single-file and multi-file modes.
    ''' </summary>
    Public Async Sub FlattenRedactedPDF()
        If INILoadFail() Then Return

        Try
            ' Defaults from My.Settings (Flatten_*), with safe fallbacks
            Dim defaultSubdir As System.String = ""
            Dim defaultExtension As System.String = ""
            Try
                Dim v = My.Settings("Flatten_OutputSubdir")
                If v IsNot Nothing AndAlso System.Convert.ToString(v).Trim().Length > 0 Then defaultSubdir = System.Convert.ToString(v)
            Catch
            End Try
            Try
                Dim v = My.Settings("Flatten_OutputExtension")
                If v IsNot Nothing AndAlso System.Convert.ToString(v).Trim().Length > 0 Then defaultExtension = System.Convert.ToString(v)
            Catch
            End Try

            ' Parameter window (first param: Multiple files)
            Dim multipleFiles As System.Boolean = False       ' default: not checked
            Dim includeReasonCodes As System.Boolean = False  ' default: No
            Dim dpi As System.Int32 = 300                     ' default: 300
            Dim outSubdir As System.String = defaultSubdir
            Dim outExtension As System.String = defaultExtension

            Dim p0 As SLib.InputParameter = New SLib.InputParameter("Multiple files", multipleFiles)
            Dim p1 As SLib.InputParameter = New SLib.InputParameter("Include reason codes", includeReasonCodes)
            Dim p2 As SLib.InputParameter = New SLib.InputParameter("Output DPI", dpi)
            Dim p3 As SLib.InputParameter = New SLib.InputParameter("Output subdirectory", outSubdir)
            Dim p4 As SLib.InputParameter = New SLib.InputParameter("Output filename extension", outExtension)

            Dim prm() As SLib.InputParameter = {p0, p1, p2, p3, p4}
            If ShowCustomVariableInputForm("Please set the finalization parameters:", AN & " Finalize PDF(s)", prm) = False Then
                Return
            End If

            ' Read back
            Try
                Dim v0 = prm(0).Value
                If TypeOf v0 Is System.Boolean Then multipleFiles = System.Convert.ToBoolean(v0) Else multipleFiles = False
            Catch
                multipleFiles = False
            End Try
            Try : includeReasonCodes = System.Convert.ToBoolean(prm(1).Value) : Catch : includeReasonCodes = False : End Try
            Try
                Dim t As System.Int32 = System.Convert.ToInt32(prm(2).Value)
                dpi = If(t > 0, t, 300)
            Catch
                dpi = 300
            End Try
            outSubdir = System.Convert.ToString(prm(3).Value)
            outExtension = System.Convert.ToString(prm(4).Value)

            ' Apply safe defaults depending on mode (if user left fields empty)
            If multipleFiles AndAlso System.String.IsNullOrWhiteSpace(outSubdir) Then
                outSubdir = "Final"
            End If
            If (Not multipleFiles) AndAlso System.String.IsNullOrWhiteSpace(outSubdir) AndAlso System.String.IsNullOrWhiteSpace(outExtension) Then
                outExtension = "_final"
            End If

            ' Persist Flatten_* settings
            Try : My.Settings("Flatten_OutputSubdir") = outSubdir : Catch : End Try
            Try : My.Settings("Flatten_OutputExtension") = outExtension : Catch : End Try
            Try : My.Settings.Save() : Catch : End Try

            If multipleFiles Then
                ' === MULTI-FILE MODE ===
                Dim selectedFolder As System.String = Nothing
                Try
                    Using dlg As New System.Windows.Forms.FolderBrowserDialog()
                        dlg.Description = "Select the folder containing the PDF files to finalize"
                        dlg.ShowNewFolderButton = False
                        Dim dr As System.Windows.Forms.DialogResult = dlg.ShowDialog()
                        If dr <> System.Windows.Forms.DialogResult.OK OrElse System.String.IsNullOrWhiteSpace(dlg.SelectedPath) Then
                            ShowCustomMessageBox("No folder has been selected - will abort.")
                            Return
                        End If
                        selectedFolder = dlg.SelectedPath
                    End Using
                Catch ex As System.Exception
                    ShowCustomMessageBox("Folder selection failed: " & ex.Message)
                    Return
                End Try

                Dim pdfFiles As System.String()
                Try
                    pdfFiles = System.IO.Directory.GetFiles(selectedFolder, "*.pdf", System.IO.SearchOption.TopDirectoryOnly)
                Catch ex As System.Exception
                    ShowCustomMessageBox("Failed to enumerate PDFs: " & ex.Message)
                    Return
                End Try
                If pdfFiles Is Nothing OrElse pdfFiles.Length = 0 Then
                    ShowCustomMessageBox("Selected folder contains no PDF files - nothing to do.")
                    Return
                End If

                ' Compute output base directory
                Dim outputBaseDir As System.String = selectedFolder
                If Not System.String.IsNullOrWhiteSpace(outSubdir) Then
                    outputBaseDir = System.IO.Path.Combine(selectedFolder, outSubdir)
                    Try
                        System.IO.Directory.CreateDirectory(outputBaseDir)
                    Catch ex As System.Exception
                        ShowCustomMessageBox("Cannot create output folder '" & outputBaseDir & "': " & ex.Message)
                        Return
                    End Try
                End If

                ' Extra safety: if both subdir and extension empty, set extension
                If System.String.IsNullOrWhiteSpace(outSubdir) AndAlso System.String.IsNullOrWhiteSpace(outExtension) Then
                    outExtension = "_final"
                End If

                ' Progress bar
                ShowProgressBarInSeparateThread(AN & " PDF Finalize", "Finalizing PDFs...")
                ProgressBarModule.CancelOperation = False

                GlobalProgressMax = pdfFiles.Length
                GlobalProgressValue = 0
                GlobalProgressLabel = "Starting..."

                Dim processed As System.Int32 = 0
                Dim errors As New System.Collections.Generic.List(Of System.String)()
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

                    Dim nameOnly As System.String = System.IO.Path.GetFileNameWithoutExtension(inPath)
                    Dim extOnly As System.String = System.IO.Path.GetExtension(inPath)
                    Dim outName As System.String = nameOnly & outExtension & extOnly
                    Dim outPath As System.String = System.IO.Path.Combine(outputBaseDir, outName)

                    Try
                        Await System.Threading.Tasks.Task.Run(Sub() PdfRedactionService.BurnInPdfToImageOnly(inPath, outPath, includeReasonCodes, dpi))
                        processed += 1
                    Catch ex As System.Exception
                        errors.Add(fileNameOnly & ": " & ex.Message)
                    End Try
                Next

                ' Close progress bar
                ProgressBarModule.CancelOperation = True

                ' Summary
                Dim msg As System.String
                If cancelled Then
                    msg = "Operation aborted by user." & vbCrLf
                Else
                    msg = "Finalization completed." & vbCrLf
                End If
                msg &= "Processed: " & processed.ToString() & " of " & pdfFiles.Length.ToString() & " file(s)." & vbCrLf &
                       "Output folder: " & outputBaseDir
                If errors.Count > 0 Then
                    msg &= vbCrLf & vbCrLf & "Errors (copied into the clipboard):" & vbCrLf & System.String.Join(vbCrLf, errors)
                    SLib.PutInClipboard(msg)
                End If
                ShowCustomMessageBox(msg)

            Else
                ' === SINGLE-FILE MODE (Drag & Drop picker) ===
                Try
                    DragDropFormLabel = "PDF file (.pdf)"
                    DragDropFormFilter = "Supported Files|*.pdf"

                    Dim filePath As System.String = GetFileName()

                    DragDropFormLabel = ""
                    DragDropFormFilter = ""

                    If System.String.IsNullOrWhiteSpace(filePath) Then
                        ShowCustomMessageBox("No file has been selected - will abort.")
                        Return
                    End If
                    If Not System.IO.File.Exists(filePath) Then
                        ShowCustomMessageBox("The selected file does not exist - will abort.")
                        Return
                    End If
                    If Not System.String.Equals(System.IO.Path.GetExtension(filePath), ".pdf", System.StringComparison.OrdinalIgnoreCase) Then
                        ShowCustomMessageBox("The selected file is not a PDF - will abort.")
                        Return
                    End If

                    ' Compute output path (subdir + extension)
                    Dim dir As System.String = System.IO.Path.GetDirectoryName(filePath)
                    Dim name As System.String = System.IO.Path.GetFileNameWithoutExtension(filePath)
                    Dim ext As System.String = System.IO.Path.GetExtension(filePath)
                    Dim outDir As System.String = If(System.String.IsNullOrWhiteSpace(outSubdir), dir, System.IO.Path.Combine(dir, outSubdir))
                    If Not System.String.IsNullOrWhiteSpace(outSubdir) Then
                        Try
                            System.IO.Directory.CreateDirectory(outDir)
                        Catch ex As System.Exception
                            ShowCustomMessageBox("Cannot create output folder '" & outDir & "': " & ex.Message)
                            Return
                        End Try
                    End If
                    ' Extra safety: if both subdir and extension empty, set extension
                    If System.String.IsNullOrWhiteSpace(outSubdir) AndAlso System.String.IsNullOrWhiteSpace(outExtension) Then
                        outExtension = "_final"
                    End If
                    Dim outPath As System.String = System.IO.Path.Combine(outDir, name & outExtension & ext)

                    Dim splash As New SLib.SplashScreen($"Finalizing PDF ...")
                    splash.Show()
                    splash.Refresh()

                    Try
                        Await System.Threading.Tasks.Task.Run(Sub() PdfRedactionService.BurnInPdfToImageOnly(filePath, outPath, includeReasonCodes, dpi))
                        splash.Close()
                        ShowCustomMessageBox("Finalized PDF created:" & vbCrLf & outPath)
                    Catch ex As System.Exception
                        splash.Close()
                        ShowCustomMessageBox("Finalization failed: " & ex.Message)
                    End Try
                Catch ex As System.Exception
                    DragDropFormLabel = ""
                    DragDropFormFilter = ""
                    ShowCustomMessageBox("File selection failed: " & ex.Message)
                End Try
            End If

        Catch ex As System.Exception
            ShowCustomMessageBox("Finalize PDF operation failed: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Service class providing PDF redaction operations: text extraction, LLM-based analysis,
    ''' annotation creation, and finalization via rasterization.
    ''' </summary>
    Public Class PdfRedactionService

        ' Padding for redaction boxes
        Private Const RedactionPadPts As Double = 2.5
        Private Const RedactionSidePad As Double = 2.0

        ' Subdirectory and extension defaults
        Private Const SubDirRedacted As String = "Redacted"
        Private Const RedactedExtension As String = "_redacted"

        Private Class TextPosition
            Public Property PageNumber As System.Int32
            Public Property X As System.Double
            Public Property Y As System.Double
            Public Property Width As System.Double
            Public Property Height As System.Double
            Public Property PageHeight As System.Double
        End Class

        Private Class RedactionRangeDto
            <Newtonsoft.Json.JsonProperty("start")>
            Public Property Start As System.Int32

            <Newtonsoft.Json.JsonProperty("end")>
            Public Property [End] As System.Int32
        End Class

        Private Class RedactionItemDto
            <Newtonsoft.Json.JsonProperty("id")>
            Public Property Id As System.String

            <Newtonsoft.Json.JsonProperty("reason")>
            Public Property Reason As System.String

            <Newtonsoft.Json.JsonProperty("exact_text")>
            Public Property ExactText As System.String

            <Newtonsoft.Json.JsonProperty("ranges")>
            Public Property Ranges As System.Collections.Generic.List(Of RedactionRangeDto)
        End Class

        Private Class RedactionResponseDto
            <Newtonsoft.Json.JsonProperty("redactions")>
            Public Property Redactions As System.Collections.Generic.List(Of RedactionItemDto)
        End Class

        Private Class RedactionRectangle
            Public Property PageNumber As System.Int32
            Public Property X As System.Double
            Public Property Y As System.Double
            Public Property Width As System.Double
            Public Property Height As System.Double
            Public Property ReasonCode As System.String
        End Class

        ''' <summary>
        ''' Orchestrates PDF redaction: extracts text, queries LLM, matches text occurrences,
        ''' builds rectangles, creates annotated PDF with optional metadata removal.
        ''' </summary>
        ''' <param name="filename">Input PDF path</param>
        ''' <param name="instruction">Redaction instruction for LLM</param>
        ''' <param name="transparentboxes">True for red transparent boxes; False for black opaque</param>
        ''' <param name="reasoncodes">Include reason codes in /Contents field</param>
        ''' <param name="removemetadata">Strip PDF Info dictionary and XMP metadata</param>
        ''' <param name="subdirectory">Output subdirectory (optional)</param>
        ''' <param name="extension">Filename extension suffix (optional)</param>
        ''' <param name="silent">Suppress message boxes; throw exceptions instead</param>
        ''' <param name="useSecondApi">Use secondary API/model if configured</param>

        Public Shared Async Function RunPdfRedactionAsync(filename As String, instruction As String, transparentboxes As Boolean,
                                                          reasoncodes As Boolean, removemetadata As Boolean,
                                                          subdirectory As String, extension As String, silent As Boolean,
                                                          Optional useSecondApi As Boolean = False) As System.Threading.Tasks.Task

            Try

                Dim fullText As System.String = Nothing
                Dim positions As System.Collections.Generic.List(Of TextPosition) = Nothing

                Try
                    ExtractTextAndPositions(filename, fullText, positions)
                Catch ex As System.Exception
                    Dim errMsg As String = $"Failed to extract text: {ex.Message}"
                    If Not silent Then
                        ShowCustomMessageBox(errMsg)
                        Return
                    End If
                    Throw New Exception(errMsg, ex)
                End Try

                If fullText Is Nothing OrElse positions Is Nothing OrElse positions.Count = 0 Then
                    Dim errMsg As String = $"The PDF appears to contain no extractable text - run OCR first using a PDF application."
                    If Not silent Then
                        ShowCustomMessageBox(errMsg)
                        Return
                    End If
                    Throw New Exception(errMsg)
                End If

                Dim llmText As System.String = "<TEXTTOPROCESS>" & fullText & "</TEXTTOPROCESS>"

                Dim llmJson As System.String = Nothing

                Globals.ThisAddIn.OtherPrompt = instruction

                llmJson = Await LLM(Globals.ThisAddIn.InterpolateAtRuntime(SP_Redact), llmText, "", "", 0, useSecondApi)

                llmJson = WebAgentInterpreter.SanitizeLlmResult(llmJson)

                If System.String.IsNullOrWhiteSpace(llmJson) Then
                    Dim errMsg As String = $"The AI returned an empty response (may be an AI error - you may want to retry)."
                    If Not silent Then
                        ShowCustomMessageBox(errMsg)
                        Return
                    End If
                    Throw New Exception(errMsg)
                End If

                Dim response As RedactionResponseDto = Nothing
                Try
                    response = Newtonsoft.Json.JsonConvert.DeserializeObject(Of RedactionResponseDto)(llmJson)
                Catch ex As System.Exception
                    Dim errMsg As String = $"Could not parse the model's JSON response: " & ex.Message
                    If Not silent Then
                        ShowCustomMessageBox(errMsg)
                        Return
                    End If
                    Throw New Exception(errMsg)
                End Try

                If (response Is Nothing OrElse response.Redactions Is Nothing OrElse response.Redactions.Count = 0) And Not silent Then
                    ShowCustomMessageBox($"Nothing to redact found.")
                    Return
                End If

                ' Convert exact_text to ranges by finding occurrences
                For Each item As RedactionItemDto In response.Redactions
                    If Not String.IsNullOrEmpty(item.ExactText) Then
                        item.Ranges = FindTextOccurrences(fullText, item.ExactText)
                    End If
                Next

                ' Filter out items with no valid ranges
                response.Redactions = response.Redactions.Where(Function(r) r.Ranges IsNot Nothing AndAlso r.Ranges.Count > 0).ToList()

                If response.Redactions.Count = 0 Then
                    Dim errMsg As String = $"No matching text found for the identified redactions (may be an AI error - you may want to retry)."
                    If Not silent Then
                        ShowCustomMessageBox(errMsg)
                        Return
                    End If
                    Throw New Exception(errMsg)
                End If

                Dim rectangles As System.Collections.Generic.List(Of RedactionRectangle) =
                        BuildRedactionRectangles(response, positions, fullText.Length)

                If rectangles Is Nothing OrElse rectangles.Count = 0 Then
                    Dim errMsg As String = $"No valid redaction ranges found in the AI's response (may be an AI error - you may want to retry)."
                    If Not silent Then
                        ShowCustomMessageBox(errMsg)
                        Return
                    End If
                    Throw New Exception(errMsg)
                End If

                Dim baseOutputPath As System.String = ""
                Dim OutputPath As System.String = ""

                Dim baseDir As System.String = System.IO.Path.GetDirectoryName(filename)
                Dim subdirDefined As System.Boolean = Not String.IsNullOrWhiteSpace(subdirectory)

                If subdirDefined Then
                    Dim subdirName As System.String = If(String.IsNullOrWhiteSpace(subdirectory), SubDirRedacted, subdirectory.Trim())
                    baseDir = System.IO.Path.Combine(baseDir, subdirName)
                    If Not System.IO.Directory.Exists(baseDir) Then
                        System.IO.Directory.CreateDirectory(baseDir)
                    End If
                End If

                Dim nameNoExt As System.String = System.IO.Path.GetFileNameWithoutExtension(filename)

                Dim suffix As System.String = ""
                If Not String.IsNullOrWhiteSpace(extension) Then
                    suffix = extension
                ElseIf Not subdirDefined Then
                    ' If neither subdirectory nor extension is defined, use the default suffix.
                    suffix = RedactedExtension
                End If

                baseOutputPath = System.IO.Path.Combine(baseDir, nameNoExt)
                OutputPath = baseOutputPath & suffix & ".pdf"

                Try
                    CreateRedactedPdf(filename, OutputPath, rectangles, reasoncodes, transparentboxes, True, removemetadata)
                Catch ex As System.Exception
                    Dim errMsg As String = "Error while creating redacted PDF: " & ex.Message
                    If Not silent Then
                        ShowCustomMessageBox(errMsg)
                        Return
                    End If
                    Throw New Exception(errMsg, ex)
                End Try

                If Not silent Then
                    If String.IsNullOrWhiteSpace(OutputPath) Then
                        ShowCustomMessageBox("Redaction completed successfully. Make sure you check and finalize them before using the PDF (otherwise they can be removed).")
                    Else
                        ShowCustomMessageBox("Redaction completed successfully. Make sure you check and finalize them before using the PDF (otherwise they can be removed). " & System.Environment.NewLine & System.Environment.NewLine &
                         "Output: " & OutputPath)
                    End If

                End If

            Catch ex As System.Exception
                Dim errMsg As String = ("Error in RunPdfRedactionAsync: " & ex.Message)
                If Not silent Then
                    ShowCustomMessageBox(errMsg)
                    Return
                End If
                Throw New Exception(errMsg, ex)
            End Try
        End Function


        ''' <summary>
        ''' Finds all occurrences of searchText in fullText. First attempts exact match,
        ''' then falls back to normalized (whitespace-collapsed) matching with position mapping.
        ''' </summary>
        ''' <param name="fullText">Full extracted PDF text</param>
        ''' <param name="searchText">Text snippet to locate</param>
        ''' <returns>List of RedactionRangeDto with start/end indices</returns>

        Private Shared Function FindTextOccurrences(fullText As String, searchText As String) As List(Of RedactionRangeDto)
            Dim ranges As New List(Of RedactionRangeDto)()

            If String.IsNullOrEmpty(searchText) Then
                Return ranges
            End If

            ' First try exact match
            Dim index As Integer = 0
            While index < fullText.Length
                index = fullText.IndexOf(searchText, index, StringComparison.Ordinal)
                If index = -1 Then
                    Exit While
                End If

                Dim range As New RedactionRangeDto()
                range.Start = index
                range.[End] = index + searchText.Length
                ranges.Add(range)

                index += searchText.Length
            End While

            ' If no exact matches found, try with normalized spacing
            If ranges.Count = 0 Then
                ' Normalize both texts by collapsing multiple spaces to single space
                Dim normalizedFull As String = System.Text.RegularExpressions.Regex.Replace(fullText, "\s+", " ")
                Dim normalizedSearch As String = System.Text.RegularExpressions.Regex.Replace(searchText, "\s+", " ")

                ' Build mapping from normalized position back to original position
                Dim normToOrigMap As New List(Of Integer)()
                Dim origIdx As Integer = 0

                For normIdx As Integer = 0 To normalizedFull.Length - 1
                    ' Skip extra whitespace in original
                    While origIdx < fullText.Length AndAlso Char.IsWhiteSpace(fullText(origIdx)) AndAlso
                  (normIdx = 0 OrElse Not Char.IsWhiteSpace(normalizedFull(normIdx - 1)) OrElse
                   Not Char.IsWhiteSpace(fullText(origIdx - 1)))
                        origIdx += 1
                    End While

                    If origIdx < fullText.Length Then
                        normToOrigMap.Add(origIdx)
                        origIdx += 1
                    End If
                Next

                ' Search in normalized text
                index = 0
                While index < normalizedFull.Length
                    index = normalizedFull.IndexOf(normalizedSearch, index, StringComparison.Ordinal)
                    If index = -1 Then
                        Exit While
                    End If

                    ' Map back to original positions
                    If index < normToOrigMap.Count AndAlso index + normalizedSearch.Length <= normToOrigMap.Count Then
                        Dim origStart As Integer = normToOrigMap(index)
                        ' Find the actual end position in the original text
                        Dim origEnd As Integer = origStart + searchText.Length

                        ' Adjust end position to match actual text
                        Dim testText As String = fullText.Substring(origStart, System.Math.Min(searchText.Length + 10, fullText.Length - origStart))
                        Dim actualMatch As String = System.Text.RegularExpressions.Regex.Replace(testText, "\s+", " ")

                        If actualMatch.StartsWith(normalizedSearch, StringComparison.Ordinal) Then
                            ' Find where normalized search actually ends in original
                            Dim charCount As Integer = 0
                            Dim pos As Integer = origStart
                            For Each c In normalizedSearch
                                If Not Char.IsWhiteSpace(c) Then
                                    ' Find next non-whitespace character in original
                                    While pos < fullText.Length AndAlso Char.IsWhiteSpace(fullText(pos))
                                        pos += 1
                                    End While
                                    pos += 1
                                Else
                                    ' Skip whitespace
                                    If pos < fullText.Length AndAlso Char.IsWhiteSpace(fullText(pos)) Then
                                        While pos < fullText.Length AndAlso Char.IsWhiteSpace(fullText(pos))
                                            pos += 1
                                        End While
                                    End If
                                End If
                            Next

                            Dim range As New RedactionRangeDto()
                            range.Start = origStart
                            range.[End] = pos
                            ranges.Add(range)
                        End If
                    End If

                    index += normalizedSearch.Length
                End While
            End If

            Return ranges
        End Function

        ''' <summary>
        ''' Extracts text and character-level positions from PDF using PdfPig.
        ''' Preserves word order, line breaks, and spatial layout; inserts spaces between words and lines.
        ''' </summary>
        ''' <param name="pdfPath">Input PDF path</param>
        ''' <param name="fullText">Output: concatenated text from all pages</param>
        ''' <param name="positions">Output: List of TextPosition for each character</param>

        Private Shared Sub ExtractTextAndPositions(pdfPath As System.String,
                           ByRef fullText As System.String,
                           ByRef positions As System.Collections.Generic.List(Of TextPosition))

            Dim sb As New System.Text.StringBuilder()
            positions = New System.Collections.Generic.List(Of TextPosition)()

            Using document As PDFP.PdfDocument = PDFP.PdfDocument.Open(pdfPath)
                Dim pageCount As System.Int32 = document.NumberOfPages

                For pageNumber As System.Int32 = 1 To pageCount
                    Dim page As PDFP.Content.Page = document.GetPage(pageNumber)
                    Dim pageHeight As System.Double = page.Height

                    ' Get words and sort them by position
                    Dim words = page.GetWords().OrderBy(Function(w) w.BoundingBox.Bottom).ThenBy(Function(w) w.BoundingBox.Left).ToList()

                    Dim previousWord As PDFP.Content.Word = Nothing
                    Dim currentLineY As Double = -1
                    Dim lineThreshold As Double = 2.0 ' Tolerance for same line detection

                    For Each word As PDFP.Content.Word In words
                        Dim wordText As System.String = word.Text
                        If System.String.IsNullOrEmpty(wordText) Then
                            Continue For
                        End If

                        ' Check if we need to add a space before this word
                        If previousWord IsNot Nothing Then
                            Dim sameLine As Boolean = System.Math.Abs(word.BoundingBox.Bottom - previousWord.BoundingBox.Bottom) < lineThreshold

                            If sameLine Then
                                ' Add space between words on the same line
                                Dim spaceWidth As Double = word.BoundingBox.Left - previousWord.BoundingBox.Right

                                ' Only add space if there's a gap between words
                                If spaceWidth > 0 Then
                                    sb.Append(" ")

                                    ' Add position for the space
                                    Dim spacePos As New TextPosition()
                                    spacePos.PageNumber = pageNumber
                                    spacePos.X = previousWord.BoundingBox.Right
                                    spacePos.Y = word.BoundingBox.Bottom
                                    spacePos.Width = spaceWidth
                                    spacePos.Height = word.BoundingBox.Height
                                    spacePos.PageHeight = pageHeight
                                    positions.Add(spacePos)
                                End If
                            Else
                                ' Different line - add space as line separator
                                sb.Append(" ")

                                ' Add dummy position for line break space
                                Dim spacePos As New TextPosition()
                                spacePos.PageNumber = pageNumber
                                spacePos.X = 0
                                spacePos.Y = 0
                                spacePos.Width = 0
                                spacePos.Height = 0
                                spacePos.PageHeight = pageHeight
                                positions.Add(spacePos)
                            End If
                        End If

                        ' Store the word's bounding box
                        Dim bbox As UglyToad.PdfPig.Core.PdfRectangle = word.BoundingBox

                        ' For each character in the word, store the position
                        Dim charWidth As System.Double = bbox.Width / wordText.Length

                        For i As System.Int32 = 0 To wordText.Length - 1
                            sb.Append(wordText(i))

                            Dim pos As New TextPosition()
                            pos.PageNumber = pageNumber
                            pos.X = bbox.Left + (i * charWidth)
                            pos.Y = bbox.Bottom
                            pos.Width = charWidth
                            pos.Height = bbox.Height
                            pos.PageHeight = pageHeight

                            positions.Add(pos)
                        Next

                        previousWord = word
                        currentLineY = word.BoundingBox.Bottom
                    Next

                    ' Add space after each page except the last
                    If pageNumber < pageCount Then
                        sb.Append(" ")

                        ' Add dummy position for the page break
                        Dim spacePos As New TextPosition()
                        spacePos.PageNumber = pageNumber
                        spacePos.X = 0
                        spacePos.Y = 0
                        spacePos.Width = 0
                        spacePos.Height = 0
                        spacePos.PageHeight = pageHeight
                        positions.Add(spacePos)
                    End If
                Next
            End Using

            fullText = sb.ToString()
        End Sub

        ''' <summary>
        ''' Converts LLM response ranges to RedactionRectangle objects with PDF coordinates.
        ''' Groups positions by page/line, applies padding, transforms coordinates from PdfPig to PdfSharp.
        ''' </summary>
        Private Shared Function BuildRedactionRectangles(response As RedactionResponseDto,
                                         positions As System.Collections.Generic.List(Of TextPosition),
                                         llmTextLength As System.Int32) _
                                         As System.Collections.Generic.List(Of RedactionRectangle)
            Dim rectangles As New System.Collections.Generic.List(Of RedactionRectangle)()

            If response Is Nothing OrElse response.Redactions Is Nothing Then
                Return rectangles
            End If

            Dim maxValidIndex As System.Int32 = System.Math.Min(llmTextLength, positions.Count)

            For Each item As RedactionItemDto In response.Redactions
                If item Is Nothing OrElse item.Ranges Is Nothing Then
                    Continue For
                End If

                For Each r As RedactionRangeDto In item.Ranges
                    If r Is Nothing Then
                        Continue For
                    End If

                    ' Validate range
                    If r.Start < 0 OrElse r.[End] <= r.Start OrElse r.[End] > maxValidIndex Then
                        Continue For
                    End If

                    ' Group positions by page and line (approximate Y coordinate)
                    Dim pageGroups As New Dictionary(Of Integer, List(Of TextPosition))()

                    For i As System.Int32 = r.Start To r.[End] - 1
                        Dim pos As TextPosition = positions(i)
                        If pos IsNot Nothing AndAlso pos.Width > 0 Then  ' Skip dummy positions
                            If Not pageGroups.ContainsKey(pos.PageNumber) Then
                                pageGroups(pos.PageNumber) = New List(Of TextPosition)()
                            End If
                            pageGroups(pos.PageNumber).Add(pos)
                        End If
                    Next

                    ' Create rectangles for each page
                    For Each kvp In pageGroups
                        Dim pageNumber As System.Int32 = kvp.Key
                        Dim pagePositions As List(Of TextPosition) = kvp.Value

                        If pagePositions.Count = 0 Then
                            Continue For
                        End If

                        ' Group by line (positions with similar Y coordinates)
                        Dim lineGroups As New Dictionary(Of Integer, List(Of TextPosition))()
                        For Each pos In pagePositions
                            ' Round Y to nearest 5 points to group same line
                            Dim lineKey As Integer = CInt(System.Math.Round(pos.Y / 5.0) * 5)
                            If Not lineGroups.ContainsKey(lineKey) Then
                                lineGroups(lineKey) = New List(Of TextPosition)()
                            End If
                            lineGroups(lineKey).Add(pos)
                        Next

                        ' Create a rectangle for each line
                        For Each lineKvp In lineGroups
                            Dim linePositions As List(Of TextPosition) = lineKvp.Value

                            Dim minX As System.Double = linePositions.Min(Function(p) p.X)
                            Dim maxX As System.Double = linePositions.Max(Function(p) p.X + p.Width)
                            Dim avgY As System.Double = linePositions.Average(Function(p) p.Y)
                            Dim maxHeight As System.Double = linePositions.Max(Function(p) p.Height)

                            Dim firstPos As TextPosition = linePositions.First()
                            Dim pageHeight As System.Double = firstPos.PageHeight

                            ' Convert from PdfPig coordinates (bottom-left) to PdfSharp (top-left)
                            Dim rect As New RedactionRectangle()
                            rect.PageNumber = pageNumber
                            rect.X = minX
                            rect.Y = pageHeight - (avgY + maxHeight)  ' Flip Y and account for height
                            rect.Width = maxX - minX
                            rect.Height = maxHeight

                            ' Vertical padding
                            Dim newY As Double = System.Math.Max(0, rect.Y - RedactionPadPts)
                            Dim bottom As Double = rect.Y + rect.Height + RedactionPadPts
                            Dim clampedBottom As Double = System.Math.Min(pageHeight, bottom)
                            rect.Y = newY
                            rect.Height = System.Math.Max(0, clampedBottom - newY)

                            ' Horizontal padding (left/right) using RedactionSidePad
                            Dim newX As Double = System.Math.Max(0, rect.X - RedactionSidePad)
                            Dim right As Double = rect.X + rect.Width + RedactionSidePad
                            rect.X = newX
                            rect.Width = System.Math.Max(0, right - newX)

                            ' Attach reason code/label from the current item
                            rect.ReasonCode = If(Not String.IsNullOrWhiteSpace(item.Reason),
                                                 item.Reason,
                                                 item.Id)

                            rectangles.Add(rect)
                        Next
                    Next
                Next
            Next

            Return rectangles
        End Function

        ''' <summary>
        ''' Creates redacted PDF by adding Square annotations to input PDF.
        ''' Annotations can be removable (default) or burned-in via graphics.
        ''' </summary>
        Private Shared Sub CreateRedactedPdf(inputPath As System.String,
                                     outputPath As System.String,
                                     rectangles As System.Collections.Generic.List(Of RedactionRectangle),
                                     Optional includeReasonCode As System.Boolean = True,
                                     Optional transparent As System.Boolean = True,
                                     Optional makeRemovable As System.Boolean = True,
                                     Optional removeMetadata As System.Boolean = False)

            Dim document As PdfSharp.Pdf.PdfDocument =
        PdfSharp.Pdf.IO.PdfReader.Open(inputPath, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Modify)

            Dim rectsByPage = rectangles.GroupBy(Function(r) r.PageNumber)

            For Each pageGroup In rectsByPage
                Dim pageIndex As Integer = pageGroup.Key - 1
                If pageIndex < 0 OrElse pageIndex >= document.Pages.Count Then Continue For

                Dim page As PdfSharp.Pdf.PdfPage = document.Pages(pageIndex)
                Dim pageHeightPts As Double = page.Height

                If makeRemovable Then
                    For Each r In pageGroup
                        Dim yBottom As Double = pageHeightPts - (r.Y + r.Height)
                        Dim yTop As Double = pageHeightPts - r.Y

                        Dim rectArr As New PdfSharp.Pdf.PdfArray(document)
                        rectArr.Elements.Add(New PdfSharp.Pdf.PdfReal(r.X))
                        rectArr.Elements.Add(New PdfSharp.Pdf.PdfReal(yBottom))
                        rectArr.Elements.Add(New PdfSharp.Pdf.PdfReal(r.X + r.Width))
                        rectArr.Elements.Add(New PdfSharp.Pdf.PdfReal(yTop))

                        Dim ann As New PdfSharp.Pdf.PdfDictionary(document)
                        ann.Elements.SetName("/Type", "/Annot")
                        ann.Elements.SetName("/Subtype", "/Square")
                        ann.Elements("/Rect") = rectArr
                        ann.Elements.SetInteger("/F", 4)

                        If includeReasonCode AndAlso Not String.IsNullOrWhiteSpace(r.ReasonCode) Then
                            ann.Elements.SetString("/Contents", r.ReasonCode)
                        End If

                        Dim bs As New PdfSharp.Pdf.PdfDictionary(document)
                        If transparent Then
                            bs.Elements.SetInteger("/W", 1)
                            ann.Elements("/BS") = bs
                            ann.Elements("/C") = MakePdfRgbArray(document, 1.0, 0.0, 0.0)
                        Else
                            bs.Elements.SetInteger("/W", 0)
                            ann.Elements("/BS") = bs
                            ann.Elements("/C") = MakePdfRgbArray(document, 0.0, 0.0, 0.0)
                            ann.Elements("/IC") = MakePdfRgbArray(document, 0.0, 0.0, 0.0)
                        End If

                        Dim annots As PdfSharp.Pdf.PdfArray = page.Elements.GetArray("/Annots")
                        If annots Is Nothing Then
                            annots = New PdfSharp.Pdf.PdfArray(document)
                            page.Elements("/Annots") = annots
                        End If
                        document.Internals.AddObject(ann)
                        annots.Elements.Add(ann)
                    Next
                Else
                    Using gfx As PdfSharp.Drawing.XGraphics =
                PdfSharp.Drawing.XGraphics.FromPdfPage(page, PdfSharp.Drawing.XGraphicsPdfPageOptions.Append)
                        For Each r In pageGroup
                            If transparent Then
                                gfx.DrawRectangle(PdfSharp.Drawing.XPens.Red, r.X, r.Y, r.Width, r.Height)
                            Else
                                gfx.DrawRectangle(PdfSharp.Drawing.XBrushes.Black, r.X, r.Y, r.Width, r.Height)
                            End If
                        Next
                    End Using
                End If
            Next

            If removeMetadata Then
                StripPdfMetadata(document)
            End If

            document.Save(outputPath)
            document.Close()
        End Sub

        ''' <summary>
        ''' Removes Info dictionary entries and XMP metadata from PDF catalog and pages.
        ''' Best-effort; never throws exceptions.
        ''' </summary>
        Private Shared Sub StripPdfMetadata(doc As PdfSharp.Pdf.PdfDocument)
            If doc Is Nothing Then Return
            Try
                ' Clear Info dictionary keys
                Dim info = doc.Info
                If info IsNot Nothing AndAlso info.Elements IsNot Nothing Then
                    Dim keys As New System.Collections.Generic.List(Of String)
                    For Each k As String In info.Elements.Keys
                        keys.Add(k)
                    Next
                    For Each k In keys
                        info.Elements.Remove(k)
                    Next
                End If

                ' Remove catalog-level XMP metadata
                Dim catalog As PdfSharp.Pdf.PdfDictionary = doc.Internals.Catalog
                If catalog IsNot Nothing Then
                    If catalog.Elements.ContainsKey("/Metadata") Then
                        catalog.Elements.Remove("/Metadata")
                    End If
                End If

                ' Remove page-level XMP metadata if any
                For Each p As PdfSharp.Pdf.PdfPage In doc.Pages
                    If p.Elements.ContainsKey("/Metadata") Then
                        p.Elements.Remove("/Metadata")
                    End If
                Next
            Catch
                ' best-effort; never fail redaction on metadata errors
            End Try
        End Sub

        ' Helper to build a PDF array [R G B]
        Private Shared Function MakePdfRgbArray(doc As PdfSharp.Pdf.PdfDocument, r As Double, g As Double, b As Double) As PdfSharp.Pdf.PdfArray
            Dim arr As New PdfSharp.Pdf.PdfArray(doc)
            arr.Elements.Add(New PdfSharp.Pdf.PdfReal(r))
            arr.Elements.Add(New PdfSharp.Pdf.PdfReal(g))
            arr.Elements.Add(New PdfSharp.Pdf.PdfReal(b))
            Return arr
        End Function

        <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
        Private Shared Function LoadLibrary(lpFileName As String) As IntPtr
        End Function

        <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
        Private Shared Function SetDllDirectory(lpPathName As String) As Boolean
        End Function

        Public Shared Sub EnsurePdfiumLoadedPublic()
            EnsurePdfiumLoaded()
        End Sub

        ''' <summary>
        ''' Preloads pdfium.dll from bin directory or x64 subdirectory to avoid loading failures.
        ''' Uses kernel32.dll LoadLibrary and SetDllDirectory APIs.
        ''' </summary>
        Private Shared Sub EnsurePdfiumLoaded()
            Static isLoaded As Boolean = False
            If isLoaded Then Return

            Try
                ' Get the actual bin directory (not shadow copy location)
                Dim binDir As String = AppDomain.CurrentDomain.BaseDirectory

                ' Try to load pdfium.dll from bin directory
                Dim pdfiumPath As String = System.IO.Path.Combine(binDir, "pdfium.dll")

                If System.IO.File.Exists(pdfiumPath) Then
                    ' Method 1: Direct load
                    Dim handle As IntPtr = LoadLibrary(pdfiumPath)
                    If handle <> IntPtr.Zero Then
                        isLoaded = True
                        Return
                    End If
                End If

                ' Method 2: Try adding bin directory to DLL search path
                SetDllDirectory(binDir)

                ' Method 3: Try x64 subdirectory (common for native packages)
                Dim x64Path As String = System.IO.Path.Combine(binDir, "x64", "pdfium.dll")
                If System.IO.File.Exists(x64Path) Then
                    LoadLibrary(x64Path)
                End If

                isLoaded = True

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"Failed to preload pdfium.dll: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' Finalizes PDF by burning annotations into rasterized images.
        ''' Excludes sticky notes and popup annotations; optionally renders reason code labels.
        ''' Preserves source PDF metadata.
        ''' </summary>
        ''' <param name="inputPath">Input PDF with annotations</param>
        ''' <param name="outputPath">Output rasterized PDF path</param>
        ''' <param name="Label">Render reason code labels in white text</param>
        ''' <param name="dpi">Rasterization DPI (default 300)</param>
        Public Shared Sub BurnInPdfToImageOnly(inputPath As String,
                                outputPath As String, Optional Label As Boolean = False,
                                Optional dpi As Integer = 300)

            EnsurePdfiumLoaded()

            Dim tempPath As String = System.IO.Path.GetTempFileName() & ".pdf"

            Try
                Using inputDoc As PdfSharp.Pdf.PdfDocument =
                    PdfSharp.Pdf.IO.PdfReader.Open(inputPath, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Modify)

                    For Each page As PdfSharp.Pdf.PdfPage In inputDoc.Pages
                        If page.Annotations IsNot Nothing AndAlso page.Annotations.Count > 0 Then

                            Dim toRemove As New System.Collections.Generic.List(Of PdfSharp.Pdf.Annotations.PdfAnnotation)()

                            Using gfx As PdfSharp.Drawing.XGraphics =
                                PdfSharp.Drawing.XGraphics.FromPdfPage(page, PdfSharp.Drawing.XGraphicsPdfPageOptions.Append)

                                For Each annot As PdfSharp.Pdf.Annotations.PdfAnnotation In page.Annotations
                                    ' Detect annotation subtype
                                    Dim subtypeRaw As String =
                                        annot.Elements.GetName(PdfSharp.Pdf.Annotations.PdfAnnotation.Keys.Subtype)
                                    Dim subtype As String =
                                        If(String.IsNullOrEmpty(subtypeRaw),
                                           String.Empty,
                                           subtypeRaw.TrimStart("/"c).ToLowerInvariant())

                                    ' Keep sticky notes and popups (do not burn/remove)
                                    If subtype = "text" OrElse subtype = "popup" Then
                                        Continue For
                                    End If

                                    ' Read rectangle and burn as black box
                                    Dim rect As PdfSharp.Pdf.PdfRectangle =
                                        annot.Elements.GetRectangle(PdfSharp.Pdf.Annotations.PdfAnnotation.Keys.Rect)

                                    If rect Is Nothing Then Continue For

                                    Dim x As Double = rect.X1
                                    Dim yTop As Double = rect.Y2
                                    Dim width As Double = rect.X2 - rect.X1
                                    Dim height As Double = rect.Y2 - rect.Y1
                                    Dim yDraw As Double = page.Height.Point - yTop

                                    Dim black As New PdfSharp.Drawing.XSolidBrush(PdfSharp.Drawing.XColors.Black)
                                    gfx.DrawRectangle(black, x, yDraw, width, height)

                                    ' If annotation has /Contents, render it in white inside the black box
                                    Dim contents As String = annot.Elements.GetString(PdfSharp.Pdf.Annotations.PdfAnnotation.Keys.Contents)
                                    If Label AndAlso Not String.IsNullOrEmpty(contents) Then
                                        ' Basic padding and guards for tiny boxes
                                        Dim pad As Double = System.Math.Max(1.0, System.Math.Min(6.0, height * 0.1))
                                        Dim textRect As New PdfSharp.Drawing.XRect(x + pad, yDraw + pad, System.Math.Max(0, width - 2 * pad), System.Math.Max(0, height - 2 * pad))

                                        If textRect.Width > 2 AndAlso textRect.Height > 6 Then

                                            ' Replace the direct Arial usage where you create the font for the annotation text:
                                            Dim canUseArial As Boolean = EnsurePdfSharpFontResolver()
                                            Dim family As String = If(canUseArial, "Arial", "Helvetica")

                                            Dim fittedSize As Double =
                                                PdfRedactionService.FitFontSize(gfx, contents, family, textRect, 6.0, System.Math.Max(6.0, System.Math.Min(22.0, textRect.Height * 0.9)))

                                            Dim font As New PdfSharp.Drawing.XFont(family, fittedSize, PdfRedactionService.RegularStyle())

                                            Dim white As PdfSharp.Drawing.XBrush = PdfSharp.Drawing.XBrushes.White
                                            Dim state As PdfSharp.Drawing.XGraphicsState = gfx.Save()
                                            Try
                                                gfx.IntersectClip(textRect)
                                                Dim tf As New PdfSharp.Drawing.Layout.XTextFormatter(gfx)
                                                tf.Alignment = PdfSharp.Drawing.Layout.XParagraphAlignment.Left
                                                tf.DrawString(contents, font, white, textRect, PdfSharp.Drawing.XStringFormats.TopLeft)
                                            Finally
                                                gfx.Restore(state)
                                            End Try
                                        End If
                                    End If

                                    ' We flattened this annotation
                                    toRemove.Add(annot)
                                Next
                            End Using

                            ' Remove only flattened annotations; keep others (e.g., notes)
                            For Each a In toRemove
                                page.Annotations.Remove(a)
                            Next
                        End If
                    Next

                    inputDoc.Save(tempPath)
                End Using

                ' Rasterize to new PDF with remaining annotations (e.g., note icons) rendered by Pdfium
                Using pdf As PdfiumViewer.PdfDocument = PdfiumViewer.PdfDocument.Load(tempPath)
                    Dim outDoc As New PdfSharp.Pdf.PdfDocument()

                    ' Preserve metadata from the source PDF
                    CopyPdfInfoFromPath(tempPath, outDoc)

                    For pageIndex As Integer = 0 To pdf.PageCount - 1
                        Dim sizePt As System.Drawing.SizeF = pdf.PageSizes(pageIndex)
                        Dim widthPx As Integer = CInt(System.Math.Round(sizePt.Width / 72.0 * dpi))
                        Dim heightPx As Integer = CInt(System.Math.Round(sizePt.Height / 72.0 * dpi))

                        Dim renderFlags As PdfiumViewer.PdfRenderFlags =
            PdfiumViewer.PdfRenderFlags.Annotations Or
            PdfiumViewer.PdfRenderFlags.LcdText Or
            PdfiumViewer.PdfRenderFlags.ForPrinting

                        Using rendered As System.Drawing.Image =
            pdf.Render(pageIndex, widthPx, heightPx, dpi, dpi, renderFlags)

                            Dim outPage As PdfSharp.Pdf.PdfPage = outDoc.AddPage()
                            outPage.Width = PdfSharp.Drawing.XUnit.FromPoint(sizePt.Width)
                            outPage.Height = PdfSharp.Drawing.XUnit.FromPoint(sizePt.Height)

                            Using ms As New System.IO.MemoryStream()
                                rendered.Save(ms, System.Drawing.Imaging.ImageFormat.Png)
                                ms.Position = 0
                                Using xgfx As PdfSharp.Drawing.XGraphics =
                    PdfSharp.Drawing.XGraphics.FromPdfPage(outPage)
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
            Finally
                If System.IO.File.Exists(tempPath) Then
                    System.IO.File.Delete(tempPath)
                End If
            End Try
        End Sub

        Private Shared Sub CopyPdfInfoFromPath(srcPath As String, dest As PdfSharp.Pdf.PdfDocument)
            If String.IsNullOrWhiteSpace(srcPath) OrElse dest Is Nothing Then Return
            Try
                Using src As PdfSharp.Pdf.PdfDocument =
            PdfSharp.Pdf.IO.PdfReader.Open(srcPath, PdfSharp.Pdf.IO.PdfDocumentOpenMode.InformationOnly)

                    CopyPdfInfo(src, dest)
                End Using
            Catch
                ' Ignore metadata copy failures to avoid breaking rasterization
            End Try
        End Sub

        ''' <summary>
        ''' Copies PDF Info dictionary and custom metadata from source to destination document.
        ''' Handles standard fields and custom keys; adjusts DateTime kinds to Local.
        ''' </summary>
        Private Shared Sub CopyPdfInfo(src As PdfSharp.Pdf.PdfDocument, dest As PdfSharp.Pdf.PdfDocument)
            If src Is Nothing OrElse dest Is Nothing Then Return

            Dim s = src.Info
            Dim d = dest.Info

            ' Standard fields (Producer is read-only)
            d.Title = s.Title
            d.Author = s.Author
            d.Subject = s.Subject
            d.Keywords = s.Keywords
            d.Creator = s.Creator
            If s.CreationDate <> Date.MinValue Then d.CreationDate = s.CreationDate
            If s.ModificationDate <> Date.MinValue Then d.ModificationDate = s.ModificationDate

            ' Copy custom/simple Info keys, including /Producer via Elements
            For Each key As String In s.Elements.Keys
                Select Case key
                    Case "/Title", "/Author", "/Subject", "/Keywords", "/Creator", "/CreationDate", "/ModDate"
                        ' already handled by properties above
                    Case Else
                        Dim item = s.Elements(key)
                        If TypeOf item Is PdfSharp.Pdf.PdfString Then
                            dest.Info.Elements.SetString(key, DirectCast(item, PdfSharp.Pdf.PdfString).ToString())
                        ElseIf TypeOf item Is PdfSharp.Pdf.PdfName Then
                            dest.Info.Elements.SetName(key, DirectCast(item, PdfSharp.Pdf.PdfName).Value)
                        ElseIf TypeOf item Is PdfSharp.Pdf.PdfInteger Then
                            dest.Info.Elements.SetInteger(key, DirectCast(item, PdfSharp.Pdf.PdfInteger).Value)
                        ElseIf TypeOf item Is PdfSharp.Pdf.PdfBoolean Then
                            dest.Info.Elements.SetBoolean(key, DirectCast(item, PdfSharp.Pdf.PdfBoolean).Value)
                        End If
                End Select
            Next

            Dim info = dest.Info
            If info IsNot Nothing Then
                If info.CreationDate.Kind = DateTimeKind.Utc Then
                    info.CreationDate = DateTime.SpecifyKind(info.CreationDate.ToLocalTime(), DateTimeKind.Local)
                End If
                If info.ModificationDate.Kind = DateTimeKind.Utc Then
                    info.ModificationDate = DateTime.SpecifyKind(info.ModificationDate.ToLocalTime(), DateTimeKind.Local)
                End If
                ' Alternatively, if Unspecified is to be used:
                ' info.CreationDate = DateTime.SpecifyKind(info.CreationDate, DateTimeKind.Unspecified)
                ' info.ModificationDate = DateTime.SpecifyKind(info.ModificationDate, DateTimeKind.Unspecified)
            End If
        End Sub


        Private Shared _fontResolverConfigured As Boolean = False

        ''' <summary>
        ''' Configures PdfSharp global font resolver to use Arial from Windows Fonts folder.
        ''' Falls back to Base-14 fonts if Arial unavailable.
        ''' </summary>
        ''' <returns>True if Arial configured; False if using Base-14 fallback</returns>

        Private Shared Function EnsurePdfSharpFontResolver() As Boolean
            ' If already configured by caller, keep it.
            If PdfSharp.Fonts.GlobalFontSettings.FontResolver IsNot Nothing Then
                _fontResolverConfigured = True
                Return True
            End If

            ' Try to locate Arial on Windows
            Dim fontsDir As String = System.Environment.GetFolderPath(Environment.SpecialFolder.Fonts)
            Dim arialRegular = System.IO.Path.Combine(fontsDir, "arial.ttf")
            If System.IO.File.Exists(arialRegular) Then
                PdfSharp.Fonts.GlobalFontSettings.FontResolver = New ArialFontResolver()
                _fontResolverConfigured = True
                Return True
            End If

            ' No Arial installed → use Base14 later (Helvetica) without resolver
            Return False
        End Function

        ''' <summary>
        ''' Implements PdfSharp IFontResolver for Arial font family with bold/italic variants.
        ''' Loads TTF files from Windows Fonts folder; provides fallback chain.
        ''' </summary>
        Class ArialFontResolver
            Implements PdfSharp.Fonts.IFontResolver

            Private Shared ReadOnly _fonts As New System.Collections.Generic.Dictionary(Of String, Byte())(StringComparer.OrdinalIgnoreCase)

            Shared Sub New()
                AddFont("arial#r", "arial.ttf")
                AddFont("arial#b", "arialbd.ttf")
                AddFont("arial#i", "ariali.ttf")
                AddFont("arial#bi", "arialbi.ttf")
            End Sub

            Private Shared Sub AddFont(key As String, fileName As String)
                Try
                    Dim fontsDir As String = Environment.GetFolderPath(Environment.SpecialFolder.Fonts)
                    Dim path = System.IO.Path.Combine(fontsDir, fileName)
                    If System.IO.File.Exists(path) Then
                        _fonts(key) = System.IO.File.ReadAllBytes(path)
                    End If
                Catch
                    ' Ignore; missing styles will fall back.
                End Try
            End Sub

            ' Correct interface signature: return FontResolverInfo
            Public Function ResolveTypeface(familyName As String, bold As Boolean, italic As Boolean) As PdfSharp.Fonts.FontResolverInfo _
        Implements PdfSharp.Fonts.IFontResolver.ResolveTypeface

                If String.Equals(familyName, "Arial", StringComparison.OrdinalIgnoreCase) Then
                    Dim key As String =
                If(bold AndAlso italic, "arial#bi",
                   If(bold, "arial#b",
                      If(italic, "arial#i", "arial#r")))
                    If _fonts.ContainsKey(key) Then
                        Return New PdfSharp.Fonts.FontResolverInfo(key)
                    End If
                    ' Fallback to any available Arial face
                    For Each k In New String() {"arial#r", "arial#b", "arial#i", "arial#bi"}
                        If _fonts.ContainsKey(k) Then
                            Return New PdfSharp.Fonts.FontResolverInfo(k)
                        End If
                    Next
                End If

                ' Not handled → let PDFsharp fall back (e.g., Base-14 Helvetica if used)
                Return Nothing
            End Function

            Public Function GetFont(faceName As String) As Byte() _
        Implements PdfSharp.Fonts.IFontResolver.GetFont

                Dim bytes As Byte() = Nothing
                If _fonts.TryGetValue(faceName, bytes) Then
                    Return bytes
                End If
                Return Nothing
            End Function
        End Class

        ''' <summary>
        ''' Returns XFontStyleEx.Regular enum value, with fallback to enum 0 if "Regular" not found.
        ''' </summary>
        Private Shared Function RegularStyle() As PdfSharp.Drawing.XFontStyleEx
            ' Prefer named value if present, else fallback to enum 0
            Dim t As System.Type = GetType(PdfSharp.Drawing.XFontStyleEx)
            Try
                Return CType([Enum].Parse(t, "Regular", True), PdfSharp.Drawing.XFontStyleEx)
            Catch
                Return CType([Enum].ToObject(t, 0), PdfSharp.Drawing.XFontStyleEx)
            End Try
        End Function

        ''' <summary>
        ''' Binary search for maximum font size that fits text within rectangle bounds.
        ''' Simulates word wrapping and line height.
        ''' </summary>
        Private Shared Function FitFontSize(gfx As PdfSharp.Drawing.XGraphics,
                                    text As String,
                                    family As String,
                                    rect As PdfSharp.Drawing.XRect,
                                    minSize As Double,
                                    maxSize As Double) As Double
            If String.IsNullOrEmpty(text) Then Return minSize
            Dim low As Double = minSize
            Dim high As Double = System.Math.Max(minSize, maxSize)
            Dim best As Double = low
            Dim style As PdfSharp.Drawing.XFontStyleEx = RegularStyle()

            While (high - low) > 0.5
                Dim mid = (low + high) / 2.0
                Dim f As New PdfSharp.Drawing.XFont(family, mid, style)
                If TextFits(gfx, text, f, rect) Then
                    best = mid
                    low = mid
                Else
                    high = mid
                End If
            End While

            Return System.Math.Max(minSize, System.Math.Min(best, maxSize))
        End Function

        ''' <summary>
        ''' Tests whether text fits in rectangle at given font size using word wrapping simulation.
        ''' </summary>
        Private Shared Function TextFits(gfx As PdfSharp.Drawing.XGraphics,
                                 text As String,
                                 font As PdfSharp.Drawing.XFont,
                                 rect As PdfSharp.Drawing.XRect) As Boolean
            If String.IsNullOrEmpty(text) Then Return True

            Dim words As String() = System.Text.RegularExpressions.Regex.Split(text, "\s+")
            Dim spaceW As Double = gfx.MeasureString(" ", font).Width
            Dim lineW As Double = 0
            Dim lines As Integer = 1
            Dim lineH As Double = font.Size * 1.2

            For Each w In words
                Dim wW As Double = gfx.MeasureString(w, font).Width
                If lineW > 0 AndAlso (lineW + spaceW + wW) > rect.Width Then
                    lines += 1
                    lineW = wW
                Else
                    lineW = If(lineW = 0, wW, lineW + spaceW + wW)
                End If

                If lines * lineH > rect.Height Then
                    Return False
                End If
            Next

            Return True
        End Function


    End Class
End Class
