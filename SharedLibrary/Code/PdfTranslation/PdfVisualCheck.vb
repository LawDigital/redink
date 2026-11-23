Option Explicit On
Option Strict On
Option Infer On

Namespace Agents
    Partial Public NotInheritable Class PdfFileTranslation
        Private Shared Async Function VerifyPagePixels(source As System.String, translated As System.String, cancel As System.Action, review As System.Action(Of System.Int32, System.String), trace As System.Action(Of System.String)) As System.Threading.Tasks.Task(Of System.String)
#If PDF_LAYOUT_WINRT Then
            Try
                Return Await CompareRenderedPages(source, translated, cancel, review, trace)
            Catch ex As System.Exception When TypeOf ex Is System.Runtime.InteropServices.ExternalException OrElse
                TypeOf ex Is System.IO.IOException OrElse TypeOf ex Is System.IO.InvalidDataException OrElse
                TypeOf ex Is System.NotSupportedException OrElse TypeOf ex Is System.ArgumentException OrElse
                TypeOf ex Is System.InvalidOperationException OrElse TypeOf ex Is System.FormatException
                ' The caller has already reopened and validated the translated document.
                ' Cancellation and permission changes still propagate to the publication gate.
                cancel()
                trace("Visual check incomplete: " & ex.ToString())
                Return "Visual check incomplete. Please review the appearance of the saved PDF."
            End Try
#Else
            Return "Visual check incomplete in this build. Please review the appearance of the saved PDF."
#End If
        End Function

#If PDF_LAYOUT_WINRT Then
        Private Shared Async Function CompareRenderedPages(source As System.String, translated As System.String, cancel As System.Action, review As System.Action(Of System.Int32, System.String), trace As System.Action(Of System.String)) As System.Threading.Tasks.Task(Of System.String)
            Dim candidates As New System.Collections.Generic.List(Of System.Int32)()
            Dim rasters As New System.Collections.Generic.Dictionary(Of System.Int32, (Width As System.UInt32, Height As System.UInt32, Scale As System.Double))()
            Using oldStream = System.IO.File.OpenRead(source), newStream = System.IO.File.OpenRead(translated),
                  oldRandom = System.IO.WindowsRuntimeStreamExtensions.AsRandomAccessStream(oldStream),
                  newRandom = System.IO.WindowsRuntimeStreamExtensions.AsRandomAccessStream(newStream)
                Dim oldPdf = Await Windows.Data.Pdf.PdfDocument.LoadFromStreamAsync(oldRandom)
                Dim newPdf = Await Windows.Data.Pdf.PdfDocument.LoadFromStreamAsync(newRandom)
                Using oldText = UglyToad.PdfPig.PdfDocument.Open(source), newText = UglyToad.PdfPig.PdfDocument.Open(translated)
                    For number As System.Int32 = 1 To oldText.NumberOfPages
                        cancel()
                        Dim oldPage = oldText.GetPage(number)
                        Dim newPage = newText.GetPage(number)
                        Dim scale = System.Math.Min(2.0, 3200.0 / System.Math.Max(oldPage.Width, oldPage.Height))
                        Dim width = CUInt(System.Math.Max(1, System.Math.Ceiling(oldPage.Width * scale)))
                        Dim height = CUInt(System.Math.Max(1, System.Math.Ceiling(oldPage.Height * scale)))
                        Using oldImage = Await RenderPage(oldPdf, number, width, height), newImage = Await RenderPage(newPdf, number, width, height)
                            Using oldGraphics = System.Drawing.Graphics.FromImage(oldImage), newGraphics = System.Drawing.Graphics.FromImage(newImage)
                                For Each page In {oldPage, newPage}
                                    For Each letter In page.Letters
                                        Dim rect = letter.GlyphRectangle
                                        Dim x = CInt(System.Math.Floor(rect.Left * width / page.Width)) - 2
                                        Dim y = CInt(System.Math.Floor((page.Height - rect.Top) * height / page.Height)) - 2
                                        Dim w = CInt(System.Math.Ceiling(rect.Right * width / page.Width)) + 2 - x
                                        Dim h = CInt(System.Math.Ceiling((page.Height - rect.Bottom) * height / page.Height)) + 2 - y
                                        oldGraphics.FillRectangle(System.Drawing.Brushes.White, x, y, w, h)
                                        newGraphics.FillRectangle(System.Drawing.Brushes.White, x, y, w, h)
                                    Next
                                Next
                            End Using
                            Dim delta = MeasurePixels(oldImage, newImage, cancel)
                            If delta.Count > 0 Then
                                trace($"Page {number} masked rendering: {delta.Describe(width, height)}")
                                If delta.IsMaterial(scale) Then
                                    candidates.Add(number)
                                    rasters.Add(number, (width, height, scale))
                                End If
                            End If
                        End Using
                    Next
                End Using
            End Using
            If candidates.Count > 0 Then
                Using originalBytes As New System.IO.MemoryStream(), translatedBytes As New System.IO.MemoryStream()
                    WriteGraphicsCopy(source, candidates, originalBytes, cancel)
                    WriteGraphicsCopy(translated, candidates, translatedBytes, cancel)
                    Using oldRandom = System.IO.WindowsRuntimeStreamExtensions.AsRandomAccessStream(originalBytes),
                          newRandom = System.IO.WindowsRuntimeStreamExtensions.AsRandomAccessStream(translatedBytes)
                        Dim oldPdf = Await Windows.Data.Pdf.PdfDocument.LoadFromStreamAsync(oldRandom)
                        Dim newPdf = Await Windows.Data.Pdf.PdfDocument.LoadFromStreamAsync(newRandom)
                        For Each number In candidates
                            cancel()
                            Dim raster = rasters(number)
                            Using a = Await RenderPage(oldPdf, number, raster.Width, raster.Height), b = Await RenderPage(newPdf, number, raster.Width, raster.Height)
                                Dim delta = MeasurePixels(a, b, cancel)
                                trace($"Page {number} graphics-only rendering: {delta.Describe(raster.Width, raster.Height)}")
                                If delta.IsMaterial(raster.Scale) Then
                                    review(number, $"Page {number}: graphics differ after removing text. Inspect images and drawings on this page.")
                                End If
                            End Using
                        Next
                    End Using
                End Using
            End If
            Return "Page rendering checked outside the text."
        End Function

        Private Shared Async Function RenderPage(pdf As Windows.Data.Pdf.PdfDocument, number As System.Int32, width As System.UInt32, height As System.UInt32) As System.Threading.Tasks.Task(Of System.Drawing.Bitmap)
            Using page = pdf.GetPage(CUInt(number - 1)), stream As New Windows.Storage.Streams.InMemoryRandomAccessStream()
                Dim options As New Windows.Data.Pdf.PdfPageRenderOptions With {.DestinationWidth = width, .DestinationHeight = height, .BitmapEncoderId = Windows.Graphics.Imaging.BitmapEncoder.PngEncoderId}
                Await page.RenderToStreamAsync(stream, options)
                stream.Seek(0)
                Using bytes = System.IO.WindowsRuntimeStreamExtensions.AsStreamForRead(stream), image As New System.Drawing.Bitmap(bytes)
                    Return image.Clone(New System.Drawing.Rectangle(0, 0, image.Width, image.Height), System.Drawing.Imaging.PixelFormat.Format24bppRgb)
                End Using
            End Using
        End Function

        ' Only the private comparison copies omit text. Keep all other operators, their
        ' operands, the active clip and drawing order, using the existing stream reader.
        Private Shared Sub WriteGraphicsCopy(path As System.String, pages As System.Collections.Generic.List(Of System.Int32),
                bytes As System.IO.MemoryStream, cancel As System.Action)
            Using document = PdfSharp.Pdf.IO.PdfReader.Open(path, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Modify),
                  parsed = UglyToad.PdfPig.PdfDocument.Open(path)
                For Each number In pages
                    cancel()
                    Dim page = parsed.GetPage(number)
                    Dim target = document.Pages(number - 1)
                    Dim raw As New OperationSource(target, page)
                    Using content As New System.IO.MemoryStream()
                        For i As System.Int32 = 0 To page.Operations.Count - 1
                            If Not ShowOperator(page.Operations(i).[Operator]) Then raw.Write(i, content)
                        Next
                        SetContents(document, target, content.ToArray())
                    End Using
                Next
                document.Save(bytes, False)
                bytes.Position = 0
            End Using
        End Sub

        Private Class PixelDifference
            Public Count As System.Int32
            Public Peak As System.Int32
            Public TotalContrast As System.Int64
            Public MinX As System.Int32 = System.Int32.MaxValue
            Public MinY As System.Int32 = System.Int32.MaxValue
            Public MaxX As System.Int32
            Public MaxY As System.Int32

            Public Function IsMaterial(scale As System.Double) As System.Boolean
                ' Budget at most one square millimetre of full contrast, or four
                ' square millimetres of noticeable pixels, at the actual render scale.
                Dim pixelsPerMm = scale * 72.0 / 25.4
                Dim area = pixelsPerMm * pixelsPerMm
                Return TotalContrast > 255.0 * area OrElse Count > 4.0 * area
            End Function

            Public Function Describe(width As System.UInt32, height As System.UInt32) As System.String
                If Count = 0 Then Return "no differing pixels."
                Return $"{Count} pixels, peak RGB-channel difference {Peak}/255, bounds [{MinX},{MinY}]-[{MaxX},{MaxY}] from top left, raster {width}x{height}."
            End Function
        End Class

        Private Shared Function MeasurePixels(a As System.Drawing.Bitmap, b As System.Drawing.Bitmap, cancel As System.Action) As PixelDifference
            Dim delta As New PixelDifference()
            Dim rect As New System.Drawing.Rectangle(0, 0, a.Width, a.Height)
            Dim left = a.LockBits(rect, System.Drawing.Imaging.ImageLockMode.ReadOnly, System.Drawing.Imaging.PixelFormat.Format24bppRgb)
            Try
                Dim right = b.LockBits(rect, System.Drawing.Imaging.ImageLockMode.ReadOnly, System.Drawing.Imaging.PixelFormat.Format24bppRgb)
                Try
                    Dim rowA(a.Width * 3 - 1), rowB(a.Width * 3 - 1) As System.Byte
                    For y As System.Int32 = 0 To a.Height - 1
                        If y Mod 128 = 0 Then cancel()
                        System.Runtime.InteropServices.Marshal.Copy(System.IntPtr.Add(left.Scan0, left.Stride * y), rowA, 0, rowA.Length)
                        System.Runtime.InteropServices.Marshal.Copy(System.IntPtr.Add(right.Scan0, right.Stride * y), rowB, 0, rowB.Length)
                        For k As System.Int32 = 0 To rowA.Length - 1 Step 3
                            Dim contrast = System.Math.Max(System.Math.Abs(CInt(rowA(k)) - rowB(k)), System.Math.Max(
                                System.Math.Abs(CInt(rowA(k + 1)) - rowB(k + 1)), System.Math.Abs(CInt(rowA(k + 2)) - rowB(k + 2))))
                            If contrast <= 3 Then Continue For
                            Dim x = k \ 3
                            delta.Count += 1
                            delta.Peak = System.Math.Max(delta.Peak, contrast)
                            delta.TotalContrast += contrast
                            delta.MinX = System.Math.Min(delta.MinX, x)
                            delta.MaxX = System.Math.Max(delta.MaxX, x)
                            delta.MinY = System.Math.Min(delta.MinY, y)
                            delta.MaxY = y
                        Next
                    Next
                    Return delta
                Finally
                    b.UnlockBits(right)
                End Try
            Finally
                a.UnlockBits(left)
            End Try
        End Function
#End If
    End Class
End Namespace
