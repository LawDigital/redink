Option Strict On
Option Explicit On

Namespace SharedLibrary
    Partial Public Class SharedMethods
        Private Const WindowsOcrPreference As System.String = "PreferWindowsOcr"
        Private Const LargeOcrPageCount As System.Int32 = 10

        Public Enum OcrRoute
            Unavailable
            Model
            Windows
            Ask
        End Enum

        Public Const AgentPdfOcrSchema As System.String =
            """ocr_engine"":{""type"":""string"",""enum"":[""auto"",""llm"",""windows"",""windows_if_large""],""description"":""PDF OCR only. Omit to honor local settings (normally LLM first). auto prefers LLM then Windows. llm and windows require that engine. windows_if_large prefers Windows from ocr_large_min_pages total pages inclusive, otherwise LLM. Automatic modes use the other available engine if needed.""}," &
            """ocr_language"":{""type"":""string"",""description"":""Optional installed Windows OCR language tag, e.g. fr-FR. Applies only if Windows OCR is used. Omit for Windows profile default. An unavailable explicit language fails with the installed tags.""}," &
            """ocr_large_min_pages"":{""type"":""integer"",""minimum"":1,""description"":""Only with windows_if_large. Default 10 total PDF pages. Only pages needing OCR are recognized.""}," &
            """ocr_batch_pages"":{""type"":""integer"",""minimum"":0,""description"":""Maximum consecutive OCR pages per LLM request. Positive values bound requests; 0 explicitly sends a whole consecutive range. With an explicit engine the default is 1. Applies only when LLM OCR is used.""}"

        Public NotInheritable Class AgentPdfOcrPolicy
            Public ReadOnly Mode As System.String
            Public ReadOnly MinimumPages As System.Int32

            Public ReadOnly LanguageTag As System.String
            Public ReadOnly BatchPages As System.Nullable(Of System.Int32)

            Private Sub New(mode As System.String, minimumPages As System.Int32, batchPages As System.Nullable(Of System.Int32), languageTag As System.String)
                Me.LanguageTag = languageTag
                Me.Mode = mode
                Me.MinimumPages = minimumPages
                Me.BatchPages = batchPages
            End Sub

            Public Shared Function Parse(arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object)) As AgentPdfOcrPolicy
                Dim raw As System.Object = Nothing
                Dim mode As System.String = "settings"
                If arguments IsNot Nothing AndAlso arguments.TryGetValue("ocr_engine", raw) Then
                    mode = System.Convert.ToString(raw, System.Globalization.CultureInfo.InvariantCulture).Trim().ToLowerInvariant()
                    Select Case mode
                        Case "auto", "llm", "windows", "windows_if_large"
                        Case Else
                            Throw New System.ArgumentException("ocr_engine must be auto, llm, windows or windows_if_large.")
                    End Select
                End If
                Dim minimumPages As System.Int32 = LargeOcrPageCount
                If arguments IsNot Nothing AndAlso arguments.TryGetValue("ocr_large_min_pages", raw) Then
                    If mode <> "windows_if_large" OrElse Not System.Int32.TryParse(System.Convert.ToString(raw, System.Globalization.CultureInfo.InvariantCulture),
                        System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, minimumPages) OrElse minimumPages < 1 Then
                        Throw New System.ArgumentException("ocr_large_min_pages requires windows_if_large and a positive integer.")
                    End If
                End If
                Dim batchPages As System.Nullable(Of System.Int32) = Nothing
                If arguments IsNot Nothing AndAlso arguments.TryGetValue("ocr_batch_pages", raw) Then
                    Dim size As System.Int32
                    If Not System.Int32.TryParse(System.Convert.ToString(raw, System.Globalization.CultureInfo.InvariantCulture),
                        System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, size) Then
                        Throw New System.ArgumentException("ocr_batch_pages must be a non-negative integer.")
                    End If
                    batchPages = size
                ElseIf mode <> "settings" Then
                    batchPages = 1
                End If
                Dim languageTag As System.String = System.String.Empty
                If arguments IsNot Nothing AndAlso arguments.TryGetValue("ocr_language", raw) Then
                    languageTag = System.Convert.ToString(raw, System.Globalization.CultureInfo.InvariantCulture).Trim()
                    If Not System.Text.RegularExpressions.Regex.IsMatch(languageTag, "\A[A-Za-z]{2,8}(?:-[A-Za-z0-9]{1,8})*\z") Then
                        Throw New System.ArgumentException("ocr_language requires a Windows OCR language tag, for example fr-FR.")
                    End If
                End If
                Return New AgentPdfOcrPolicy(mode, minimumPages, batchPages, languageTag)
            End Function

            Friend Function Resolve(context As Global.SharedLibrary.SharedLibrary.SharedContext.ISharedContext, pages As System.Int32) As OcrRoute
                Dim local As System.Boolean = WindowsOcrLanguage().Length > 0
                If Mode = "windows" Then Return If(local, OcrRoute.Windows, OcrRoute.Unavailable)
                If local AndAlso ((Mode = "settings" AndAlso PreferWindowsOcr()) OrElse
                    (Mode = "windows_if_large" AndAlso pages >= MinimumPages)) Then Return OcrRoute.Windows
                If IsLlmPdfOcrAvailable(context) Then Return OcrRoute.Model
                Return If(local AndAlso Mode <> "llm", OcrRoute.Windows, OcrRoute.Unavailable)
            End Function
        End Class

        Private Shared Function PreferWindowsOcr() As System.Boolean
            Try
                Dim value = Microsoft.Win32.Registry.GetValue(RegPath_Base.TrimEnd("\"c), WindowsOcrPreference, 0)
                Return System.Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture) = "1"
            Catch ex As System.Exception When TypeOf ex Is System.Security.SecurityException OrElse TypeOf ex Is System.UnauthorizedAccessException OrElse TypeOf ex Is System.IO.IOException
                System.Diagnostics.Debug.WriteLine("Unable to read OCR preference. Using Windows OCR when available: " & ex.Message)
                Return True
            End Try
        End Function

        Private Shared Sub SaveWindowsOcrPreference(value As System.Boolean)
            Microsoft.Win32.Registry.SetValue(RegPath_Base.TrimEnd("\"c), WindowsOcrPreference, If(value, 1, 0), Microsoft.Win32.RegistryValueKind.DWord)
        End Sub

        Private Shared Function GetOcrRoute(hasModel As System.Boolean, hasWindows As System.Boolean,
                                            preferLocal As System.Boolean, interactive As System.Boolean,
                                            pages As System.Int32) As OcrRoute
            If Not hasWindows Then Return If(hasModel, OcrRoute.Model, OcrRoute.Unavailable)
            If Not hasModel OrElse preferLocal Then Return OcrRoute.Windows
            If interactive AndAlso pages >= LargeOcrPageCount Then Return OcrRoute.Ask
            Return OcrRoute.Model
        End Function

        Private Shared Function SelectOcrRoute(path As System.String, hasModel As System.Boolean, hasWindows As System.Boolean,
                                               interactive As System.Boolean, pages As System.Int32) As OcrRoute
            Dim route = GetOcrRoute(hasModel, hasWindows, hasWindows AndAlso hasModel AndAlso PreferWindowsOcr(),
                                    interactive AndAlso Not IsHeadlessExecution, pages)
            If route <> OcrRoute.Ask Then Return route
            Dim answer = ShowCustomYesNoBox(System.IO.Path.GetFileName(path) & System.Environment.NewLine &
                $"{pages} pages need OCR. Windows OCR recognizes text locally. Model OCR can also interpret layout and tables.",
                "Use Windows OCR", "Use model OCR", AN & " OCR")
            Select Case answer
                Case 1 : Return OcrRoute.Windows
                Case 2 : Return OcrRoute.Model
                Case Else : Throw New System.OperationCanceledException("OCR engine selection was cancelled.")
            End Select
        End Function

        Public Shared Function ChoosePdfOcrRoute(path As System.String, context As Global.SharedLibrary.SharedLibrary.SharedContext.ISharedContext,
                                                   interactive As System.Boolean, pages As System.Int32) As OcrRoute
            Dim hasWindows As System.Boolean = WindowsOcrLanguage().Length > 0
            If hasWindows AndAlso PreferWindowsOcr() Then Return OcrRoute.Windows
            Return SelectOcrRoute(path, IsLlmPdfOcrAvailable(context), hasWindows, interactive, pages)
        End Function

        Public Shared Function InstalledWindowsOcrLanguages() As System.Collections.Generic.List(Of System.String)
#If PDF_LAYOUT_WINRT Then
            Try
                Return ReadInstalledWindowsOcrLanguages()
            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine("Unable to enumerate Windows OCR languages: " & ex.Message)
            End Try
#End If
            Return New System.Collections.Generic.List(Of System.String)()
        End Function

        Private Shared Function MissingWindowsOcrLanguage(tag As System.String) As System.String
            Dim installed = InstalledWindowsOcrLanguages()
            Return $"Windows OCR language '{tag}' is unavailable. Installed OCR languages: " &
                If(installed.Count = 0, "(none)", System.String.Join(", ", installed)) & "."
        End Function

        Public Shared Function DefaultWindowsOcrLanguageTag() As System.String
            Return WindowsOcrLanguage()
        End Function

        Private Shared Function WindowsOcrLanguage() As System.String
#If PDF_LAYOUT_WINRT Then
            Try
                Return ProbeWindowsOcrLanguage()
            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine("Windows OCR is unavailable in this host: " & ex.Message)
            End Try
#End If
            Return System.String.Empty
        End Function

        Public Shared Async Function TryReadOcrClipboardAsync(context As Global.SharedLibrary.SharedLibrary.SharedContext.ISharedContext) As System.Threading.Tasks.Task(Of System.String)
            If IsHeadlessExecution OrElse WindowsOcrLanguage().Length = 0 Then Return Nothing
            Dim mime As System.String = Nothing
            Dim payload As System.String = Nothing
            Dim found As System.Boolean = Await System.Threading.Tasks.Task.Run(
                Function() ClipboardHelper.TryGetClipboardObject(mime, payload)).ConfigureAwait(False)
            If Not found OrElse System.String.IsNullOrWhiteSpace(mime) Then Return Nothing
            Dim extension As System.String
            Select Case FixMimeType(mime).ToLowerInvariant()
                Case "application/pdf" : extension = ".pdf"
                Case "image/png" : extension = ".png"
                Case "image/jpeg" : extension = ".jpg"
                Case "image/bmp" : extension = ".bmp"
                Case "image/tiff" : extension = ".tiff"
                Case "image/gif" : extension = ".gif"
                Case Else : Return Nothing
            End Select
            If extension <> ".pdf" AndAlso extension <> ".tiff" AndAlso extension <> ".gif" AndAlso
                context IsNot Nothing AndAlso MediaObjectAcceptsExtension(context.INI_APICall_Object, extension) AndAlso IsPrimaryOcrModelUsable(context) AndAlso
                Not PreferWindowsOcr() Then Return Nothing

            Dim path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.Guid.NewGuid().ToString("N") & extension)
            Try
                System.IO.File.WriteAllBytes(path, System.Convert.FromBase64String(payload))
                If extension = ".pdf" Then
                    Dim read = Await ReadPdfAsTextEx(path, ReturnErrorInsteadOfEmpty:=False, DoOCR:=True, AskUser:=False,
                        context:=context, AllowOcrEngineChoice:=True).ConfigureAwait(False)
                    If Not read.ExtractionComplete.GetValueOrDefault() Then
                        ShowCustomMessageBox("Clipboard PDF extraction may be incomplete." & System.Environment.NewLine &
                            System.String.Join(System.Environment.NewLine, read.ExtractionWarnings))
                    End If
                    Return If(read.Content, System.String.Empty)
                End If
                Return Await ReadBinaryFileViaLLM(path, context, askUser:=True, taskFlag:="ImageExtraction").ConfigureAwait(False)
            Finally
                Try
                    System.IO.File.Delete(path)
                Catch ex As System.IO.IOException
                    System.Diagnostics.Debug.WriteLine("OCR clipboard cleanup: " & ex.Message)
                End Try
            End Try
        End Function

        Private Shared Function IsWindowsOcrImage(extension As System.String) As System.Boolean
            Select Case If(extension, System.String.Empty).ToLowerInvariant()
                Case ".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".gif"
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Friend Shared Function UseWindowsOcrForBackground(context As Global.SharedLibrary.SharedLibrary.SharedContext.ISharedContext,
                                                           extension As System.String, enablePdfOcr As System.Boolean) As System.Boolean
            Dim pdf As System.Boolean = (extension = ".pdf") AndAlso enablePdfOcr
            If Not pdf AndAlso Not IsWindowsOcrImage(extension) Then Return False
            If WindowsOcrLanguage().Length = 0 Then Return False
            If PreferWindowsOcr() Then Return True
            If pdf Then Return Not IsLlmPdfOcrAvailable(context)
            Return Not IsModelMediaSupported(context, extension, TaskFlagForExtension(extension))
        End Function

        Public Shared Function MayUseModelPdfOcr(context As Global.SharedLibrary.SharedLibrary.SharedContext.ISharedContext) As System.Boolean
            Return Not (PreferWindowsOcr() AndAlso WindowsOcrLanguage().Length > 0) AndAlso IsLlmPdfOcrAvailable(context)
        End Function

        Public Shared Function GetLocalOcrFingerprint() As System.String
            Return "windows-ocr-v2|language=" & WindowsOcrLanguage() & "|prefer-local=" & PreferWindowsOcr().ToString()
        End Function

#If PDF_LAYOUT_WINRT Then
        <System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)>
        Private Shared Function ProbeWindowsOcrLanguage() As System.String
            Dim engine = CreateWindowsOcrEngine()
            Return If(engine Is Nothing, System.String.Empty, engine.RecognizerLanguage.LanguageTag)
        End Function

        <System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)>
        Private Shared Function ReadInstalledWindowsOcrLanguages() As System.Collections.Generic.List(Of System.String)
            Dim tags As New System.Collections.Generic.List(Of System.String)()
            For Each language In Global.Windows.Media.Ocr.OcrEngine.AvailableRecognizerLanguages
                tags.Add(language.LanguageTag)
            Next
            tags.Sort(System.StringComparer.OrdinalIgnoreCase)
            Return tags
        End Function

        Private Shared Function CreateWindowsOcrEngine(Optional languageTag As System.String = Nothing) As Global.Windows.Media.Ocr.OcrEngine
            If Not System.String.IsNullOrEmpty(languageTag) Then
                For Each installed In Global.Windows.Media.Ocr.OcrEngine.AvailableRecognizerLanguages
                    If System.String.Equals(installed.LanguageTag, languageTag, System.StringComparison.OrdinalIgnoreCase) Then
                        Return Global.Windows.Media.Ocr.OcrEngine.TryCreateFromLanguage(installed)
                    End If
                Next
                Return Nothing
            End If
            Dim engine = Global.Windows.Media.Ocr.OcrEngine.TryCreateFromUserProfileLanguages()
            If engine IsNot Nothing Then Return engine
            For Each language In Global.Windows.Media.Ocr.OcrEngine.AvailableRecognizerLanguages
                engine = Global.Windows.Media.Ocr.OcrEngine.TryCreateFromLanguage(language)
                If engine IsNot Nothing Then Return engine
            Next
            Return Nothing
        End Function

        Private Shared Function WindowsOcrScale(width As System.Double, height As System.Double, targetScale As System.Double) As System.Double
            If width <= 0 OrElse height <= 0 Then Throw New System.IO.InvalidDataException("Empty OCR page dimensions.")
            Dim limit As System.Double = Global.Windows.Media.Ocr.OcrEngine.MaxImageDimension
            ' Bound the largest page bitmap as well as the OS limit on either edge.
            Return System.Math.Min(targetScale, System.Math.Min(limit / System.Math.Max(width, height),
                System.Math.Sqrt(12000000.0R / (width * height))))
        End Function

        ' Quality checks target fragmented alphabetic prose. Single-character CJK text and
        ' numeric tables alone do not establish that a page is incorrectly oriented.
        Private NotInheritable Class WindowsOcrReading
            Public Text As System.String
            Public WordLetters As System.Int32
            Public WordCount As System.Int32
            Public Fragmented As System.Boolean
            Public SparseFragments As System.Boolean

            Public Sub New(value As System.String)
                Text = value
                Dim lineCount As System.Int32 = 0
                Dim shortLines As System.Int32 = 0
                Dim isolatedLetters As System.Int32 = 0
                For Each line In value.Split(New System.Char() {Microsoft.VisualBasic.ChrW(10)}, System.StringSplitOptions.RemoveEmptyEntries)
                    Dim trimmed = line.Trim()
                    If trimmed.Length = 0 Then Continue For
                    lineCount += 1
                    If trimmed.Length <= 3 Then shortLines += 1
                    For Each word As System.Text.RegularExpressions.Match In AlphabeticOcrWords.Matches(trimmed)
                        If word.Length = 1 Then isolatedLetters += 1
                        If word.Length >= 3 Then
                            WordCount += 1
                            WordLetters += word.Length
                        End If
                    Next
                Next
                SparseFragments = WordLetters < 10 AndAlso isolatedLetters > 0 AndAlso shortLines * 2 >= lineCount
                Fragmented = lineCount >= 4 AndAlso shortLines * 4 >= lineCount * 3 AndAlso
                    isolatedLetters >= 2 AndAlso WordCount * 5 < lineCount
            End Sub
        End Class

        Private Shared ReadOnly AlphabeticOcrWords As New System.Text.RegularExpressions.Regex(
            "[\p{Lu}\p{Ll}\p{Lt}][\p{Lu}\p{Ll}\p{Lt}\p{M}]*",
            System.Text.RegularExpressions.RegexOptions.CultureInvariant)

        Private Shared Async Function RecognizeWindowsFrameAsync(engine As Global.Windows.Media.Ocr.OcrEngine,
                frame As Global.Windows.Graphics.Imaging.BitmapFrame, token As System.Threading.CancellationToken,
                Optional report As System.Action(Of System.String) = Nothing) As System.Threading.Tasks.Task(Of System.String)
            Dim original As New WindowsOcrReading(Await ReadWindowsOrientationAsync(engine, frame, Global.Windows.Graphics.Imaging.BitmapRotation.None, token).ConfigureAwait(False))
            If Not original.Fragmented AndAlso Not original.SparseFragments AndAlso Not System.String.IsNullOrWhiteSpace(original.Text) Then Return original.Text
            Dim runnerUp As System.Int32 = 0
            Dim best = original
            Dim bestRotation = Global.Windows.Graphics.Imaging.BitmapRotation.None
            For Each rotation In {Global.Windows.Graphics.Imaging.BitmapRotation.Clockwise90Degrees,
                                  Global.Windows.Graphics.Imaging.BitmapRotation.Clockwise270Degrees,
                                  Global.Windows.Graphics.Imaging.BitmapRotation.Clockwise180Degrees}
                token.ThrowIfCancellationRequested()
                Dim candidate As WindowsOcrReading
                Try
                    candidate = New WindowsOcrReading(Await ReadWindowsOrientationAsync(engine, frame, rotation, token).ConfigureAwait(False))
                Catch ex As System.OperationCanceledException
                    Throw
                Catch ex As System.IO.InvalidDataException
                    If report IsNot Nothing Then report("windows_ocr_orientation_failed: " & rotation.ToString() & " " & ex.Message)
                    Continue For
                End Try
                If candidate.Fragmented OrElse candidate.WordCount < 2 OrElse
                    candidate.WordLetters < System.Math.Max(10, original.WordLetters * 2) Then Continue For
                If candidate.WordLetters > best.WordLetters Then
                    runnerUp = best.WordLetters
                    best = candidate
                    bestRotation = rotation
                Else
                    runnerUp = System.Math.Max(runnerUp, candidate.WordLetters)
                End If
            Next
            If best IsNot original AndAlso best.WordLetters >= runnerUp * 1.25R Then
                If report IsNot Nothing Then report("windows_ocr_orientation: " & bestRotation.ToString())
                Return best.Text
            End If
            If original.Fragmented Then Throw New System.IO.InvalidDataException("Windows OCR returned mostly isolated characters. No unambiguous improvement was found in the tested orientations.")
            Return original.Text
        End Function

        Private Shared Async Function ReadWindowsOrientationAsync(engine As Global.Windows.Media.Ocr.OcrEngine,
                frame As Global.Windows.Graphics.Imaging.BitmapFrame, rotation As Global.Windows.Graphics.Imaging.BitmapRotation,
                token As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of System.String)
            Dim scale = WindowsOcrScale(frame.PixelWidth, frame.PixelHeight, 1.0R)
            Dim transform As New Global.Windows.Graphics.Imaging.BitmapTransform With {
                .Rotation = rotation,
                .ScaledWidth = CUInt(System.Math.Max(1, System.Math.Floor(frame.PixelWidth * scale))),
                .ScaledHeight = CUInt(System.Math.Max(1, System.Math.Floor(frame.PixelHeight * scale)))
            }
            ' Use the software bitmap format from the Windows.Media.Ocr file-image sample.
            Dim stage As System.String = "bitmap decoding"
            Try
                Using bitmap = Await System.WindowsRuntimeSystemExtensions.AsTask(frame.GetSoftwareBitmapAsync(
                        Global.Windows.Graphics.Imaging.BitmapPixelFormat.Bgra8, Global.Windows.Graphics.Imaging.BitmapAlphaMode.Premultiplied,
                        transform, Global.Windows.Graphics.Imaging.ExifOrientationMode.RespectExifOrientation,
                        Global.Windows.Graphics.Imaging.ColorManagementMode.DoNotColorManage), token).ConfigureAwait(False)
                    stage = "text recognition"
                    Dim recognition = Await System.WindowsRuntimeSystemExtensions.AsTask(engine.RecognizeAsync(bitmap), token).ConfigureAwait(False)
                    Return System.String.Join(System.Environment.NewLine, System.Linq.Enumerable.Select(recognition.Lines, Function(line) line.Text))
                End Using
            Catch ex As System.OperationCanceledException
                Throw
            Catch ex As System.Exception
                token.ThrowIfCancellationRequested()
                Throw New System.IO.InvalidDataException(
                    $"Windows OCR {stage} failed (BGRA8, premultiplied, {transform.ScaledWidth} x {transform.ScaledHeight}, HRESULT 0x{ex.HResult:X8}): {ex.Message}", ex)
            End Try
        End Function

        Private Shared Async Function WindowsPdfPageTextAsync(doc As Global.Windows.Data.Pdf.PdfDocument,
                                                              engine As Global.Windows.Media.Ocr.OcrEngine,
                                                              number As System.Int32, token As System.Threading.CancellationToken,
                                                              report As System.Action(Of System.String)) As System.Threading.Tasks.Task(Of System.String)
            Using page = doc.GetPage(CUInt(number - 1)), output As New Global.Windows.Storage.Streams.InMemoryRandomAccessStream()
                Dim scale = WindowsOcrScale(page.Size.Width, page.Size.Height, 300.0R / 96.0R)
                Dim options As New Global.Windows.Data.Pdf.PdfPageRenderOptions With {
                    .DestinationWidth = CUInt(System.Math.Max(1, System.Math.Floor(page.Size.Width * scale))),
                    .DestinationHeight = CUInt(System.Math.Max(1, System.Math.Floor(page.Size.Height * scale))),
                    .BitmapEncoderId = Global.Windows.Graphics.Imaging.BitmapEncoder.PngEncoderId
                }
                Await System.WindowsRuntimeSystemExtensions.AsTask(page.RenderToStreamAsync(output, options), token).ConfigureAwait(False)
                output.Seek(0UI)
                Dim decoder = Await System.WindowsRuntimeSystemExtensions.AsTask(Global.Windows.Graphics.Imaging.BitmapDecoder.CreateAsync(output), token).ConfigureAwait(False)
                Dim frame = Await System.WindowsRuntimeSystemExtensions.AsTask(decoder.GetFrameAsync(0UI), token).ConfigureAwait(False)
                Return Await RecognizeWindowsFrameAsync(engine, frame, token, report).ConfigureAwait(False)
            End Using
        End Function

        Private Shared Async Function ReadWindowsPdfOcrAsync(path As System.String,
                nativePages As System.Collections.Generic.IReadOnlyList(Of System.String), candidates As System.Collections.Generic.List(Of System.Int32),
                showProgress As System.Boolean, instruction As System.String,
                cancellationToken As System.Threading.CancellationToken, languageTag As System.String) As System.Threading.Tasks.Task(Of SelectivePdfOcrResult)
            Dim result As New SelectivePdfOcrResult With {.UsedWindowsOcr = True}
            Dim chunks As New System.Collections.Generic.List(Of SelectivePdfOcrRangeResult)()
            Dim dialog As OcrChunkStatusDialog = Nothing
            Try
                cancellationToken.ThrowIfCancellationRequested()
                Dim engine = CreateWindowsOcrEngine(languageTag)
                If engine Is Nothing AndAlso Not System.String.IsNullOrEmpty(languageTag) Then
                    result.LanguageError = MissingWindowsOcrLanguage(languageTag)
                    Global.SharedLibrary.Agents.TextExtractionDiagnostics.AddWarning(result.Warnings, result.LanguageError)
                    Return result
                End If
                If engine Is Nothing Then Throw New System.InvalidOperationException("Windows could not create an OCR engine.")
                Global.SharedLibrary.Agents.TextExtractionDiagnostics.AddWarning(result.Warnings, "windows_ocr_language: " & engine.RecognizerLanguage.LanguageTag)
                If Not System.String.IsNullOrWhiteSpace(instruction) Then
                    Global.SharedLibrary.Agents.TextExtractionDiagnostics.AddWarning(result.Warnings, "windows_ocr_plain_text: model-specific OCR instructions are unavailable for local recognition.")
                End If
                If showProgress Then
                    dialog = New OcrChunkStatusDialog()
                    Await dialog.ShowAsync("Windows OCR (" & engine.RecognizerLanguage.LanguageTag & ")", cancellationToken).ConfigureAwait(False)
                End If
                Using linked = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, If(dialog Is Nothing, System.Threading.CancellationToken.None, dialog.CancellationToken)),
                      input = System.IO.File.OpenRead(path), stream = System.IO.WindowsRuntimeStreamExtensions.AsRandomAccessStream(input)
                    Dim token = linked.Token
                    Dim doc = Await System.WindowsRuntimeSystemExtensions.AsTask(Global.Windows.Data.Pdf.PdfDocument.LoadFromStreamAsync(stream), token).ConfigureAwait(False)
                    If doc.PageCount <> nativePages.Count Then Throw New System.IO.InvalidDataException("The PDF readers disagree on the page count.")
                    For Each number In candidates
                        token.ThrowIfCancellationRequested()
                        If dialog IsNot Nothing Then dialog.SetPhase($"Windows OCR: page {number} of {doc.PageCount}")
                        Try
                            Dim pageText = Await WindowsPdfPageTextAsync(doc, engine, number, token,
                                Sub(note) Global.SharedLibrary.Agents.TextExtractionDiagnostics.AddWarning(result.Warnings, $"Page {number}: {note}")).ConfigureAwait(False)
                            If System.String.IsNullOrWhiteSpace(pageText) Then Throw New System.IO.InvalidDataException("No text recognized.")
                            chunks.Add(New SelectivePdfOcrRangeResult With {.StartPage = number, .EndPage = number, .Text = pageText})
                            result.ProcessedRanges.Add(New Global.SharedLibrary.Agents.TextExtractionProcessedRange With {
                                .StartPage = number, .EndPage = number, .Association = "ocr_page_result"})
                        Catch ex As System.OperationCanceledException
                            Throw
                        Catch ex As System.Exception
                            Global.SharedLibrary.Agents.TextExtractionDiagnostics.AddWarning(result.Warnings, $"windows_ocr_page_incomplete: page {number}: {ex.Message}")
                        End Try
                    Next
                    token.ThrowIfCancellationRequested()
                End Using
            Catch ex As System.OperationCanceledException
                Throw
            Catch ex As System.Exception
                Global.SharedLibrary.Agents.TextExtractionDiagnostics.AddExceptionDetails(result.Warnings, "windows_ocr_failed", ex, "local_pdf_reader")
            Finally
                If dialog IsNot Nothing Then dialog.Dispose()
            End Try
            PreserveSelectivePdfOcrProgress(result, nativePages, chunks)
            result.Success = (chunks.Count = candidates.Count)
            Return result
        End Function

        Private Shared Async Function ReadWindowsImageOrSelectModelAsync(path As System.String, hasModel As System.Boolean,
                interactive As System.Boolean, token As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of System.String)
            ' Nothing delegates to the model. A failed local read must never cause a model upload.
            Dim preferLocal = hasModel AndAlso PreferWindowsOcr()
            Dim ext = System.IO.Path.GetExtension(path).ToLowerInvariant()
            Dim canAsk = interactive AndAlso Not IsHeadlessExecution
            If hasModel AndAlso Not preferLocal AndAlso (Not canAsk OrElse (ext <> ".tif" AndAlso ext <> ".tiff" AndAlso ext <> ".gif")) Then Return Nothing
            Using input = System.IO.File.OpenRead(path), stream = System.IO.WindowsRuntimeStreamExtensions.AsRandomAccessStream(input)
                Dim decoder As Global.Windows.Graphics.Imaging.BitmapDecoder
                Try
                    decoder = Await System.WindowsRuntimeSystemExtensions.AsTask(Global.Windows.Graphics.Imaging.BitmapDecoder.CreateAsync(stream), token).ConfigureAwait(False)
                Catch ex As System.OperationCanceledException
                    Throw
                Catch ex As System.Exception When hasModel AndAlso Not preferLocal
                    System.Diagnostics.Debug.WriteLine("Windows cannot decode this image; using the configured model: " & ex.Message)
                    Return Nothing
                End Try
                If hasModel AndAlso Not preferLocal AndAlso SelectOcrRoute(path, True, True, canAsk, CInt(decoder.FrameCount)) = OcrRoute.Model Then Return Nothing
                Dim engine = CreateWindowsOcrEngine()
                If engine Is Nothing Then Throw New System.InvalidOperationException("Windows OCR became unavailable.")
                If decoder.FrameCount = 0UI Then Throw New System.IO.InvalidDataException("The image contains no frames.")
                Dim pages As New System.Collections.Generic.List(Of System.String)()
                For index As System.Int32 = 0 To CInt(decoder.FrameCount) - 1
                    token.ThrowIfCancellationRequested()
                    Dim frame = Await System.WindowsRuntimeSystemExtensions.AsTask(decoder.GetFrameAsync(CUInt(index)), token).ConfigureAwait(False)
                    Dim value = Await RecognizeWindowsFrameAsync(engine, frame, token).ConfigureAwait(False)
                    If System.String.IsNullOrWhiteSpace(value) Then Throw New System.IO.InvalidDataException($"Windows OCR recognized no text on image page {index + 1}.")
                    pages.Add(value)
                Next
                Return System.String.Join(System.Environment.NewLine & Microsoft.VisualBasic.ChrW(12) & System.Environment.NewLine, pages)
            End Using
        End Function
#End If
    End Class
End Namespace
