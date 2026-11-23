Option Explicit On
Option Strict Off
Option Infer On

Imports System.Linq

Namespace SharedLibrary
    Partial Public Class SharedMethods
        Private Class OverflowingText
            Public Location As System.String
            Public ShapeName As System.String
            Public Text As System.String
        End Class

        Private Shared Sub ReleaseLayoutObject(value As System.Object)
            If value IsNot Nothing AndAlso System.Runtime.InteropServices.Marshal.IsComObject(value) Then System.Runtime.InteropServices.Marshal.ReleaseComObject(value)
        End Sub

        Private Shared Sub InspectLayoutShapes(shapes As System.Object, slide As System.Boolean, location As System.String,
                                               issues As System.Collections.Generic.List(Of OverflowingText))
            For index As System.Int32 = 1 To CInt(shapes.Count)
                Dim shape As System.Object = Nothing
                Dim frame As System.Object = Nothing
                Dim textRange As System.Object = Nothing
                Dim children As System.Object = Nothing
                Try
                    shape = shapes.Item(index)
                    Dim label = location & " / " & CStr(shape.Name)
                    If CInt(shape.Type) = 6 Then ' msoGroup
                        children = shape.GroupItems
                        InspectLayoutShapes(children, slide, label, issues)
                        Continue For
                    End If
                    If slide Then
                        If CInt(shape.HasTextFrame) = 0 Then Continue For
                        frame = shape.TextFrame2
                        If CInt(frame.HasText) = 0 Then Continue For
                        textRange = frame.TextRange
                        If CDbl(textRange.BoundWidth) <= CDbl(shape.Width) - CDbl(frame.MarginLeft) - CDbl(frame.MarginRight) + 0.75 AndAlso
                           CDbl(textRange.BoundHeight) <= CDbl(shape.Height) - CDbl(frame.MarginTop) - CDbl(frame.MarginBottom) + 0.75 Then Continue For
                    Else
                        If CInt(shape.Type) <> 1 AndAlso CInt(shape.Type) <> 17 AndAlso CInt(shape.Type) <> 2 Then Continue For
                        frame = shape.TextFrame
                        If CInt(frame.HasText) = 0 OrElse CInt(frame.Overflowing) = 0 Then Continue For
                        textRange = frame.TextRange
                    End If
                    issues.Add(New OverflowingText With {.Location = label, .ShapeName = CStr(shape.Name), .Text = CStr(textRange.Text).TrimEnd(ChrW(13), ChrW(7))})
                Finally
                    ReleaseLayoutObject(children)
                    ReleaseLayoutObject(textRange)
                    ReleaseLayoutObject(frame)
                    ReleaseLayoutObject(shape)
                End Try
            Next
        End Sub

        Private Shared Function MeasureOfficeBoxes(path As System.String) As System.Collections.Generic.List(Of OverflowingText)
            Dim isSlide = System.String.Equals(System.IO.Path.GetExtension(path), ".pptx", System.StringComparison.OrdinalIgnoreCase)
            Dim issues As New System.Collections.Generic.List(Of OverflowingText)()
            Dim app As System.Object = Nothing
            Dim files As System.Object = Nothing
            Dim doc As System.Object = Nothing
            Dim slides As System.Object = Nothing
            Dim shapes As System.Object = Nothing
            Try
                app = Microsoft.VisualBasic.Interaction.CreateObject(If(isSlide, "PowerPoint.Application", "Word.Application"))
                If isSlide Then
                    files = app.Presentations
                    doc = files.Open(path, -1, 0, 0)
                    slides = doc.Slides
                    For index As System.Int32 = 1 To CInt(slides.Count)
                        Dim current As System.Object = slides.Item(index)
                        Try
                            shapes = current.Shapes
                            InspectLayoutShapes(shapes, True, "Slide " & index.ToString(), issues)
                        Finally
                            ReleaseLayoutObject(shapes)
                            shapes = Nothing
                            ReleaseLayoutObject(current)
                        End Try
                    Next
                Else
                    app.DisplayAlerts = 0
                    app.AutomationSecurity = 3
                    files = app.Documents
                    doc = files.Open(FileName:=path, ConfirmConversions:=False, ReadOnly:=True, AddToRecentFiles:=False, Visible:=False)
                    doc.Repaginate()
                    shapes = doc.Shapes
                    InspectLayoutShapes(shapes, False, "Main document", issues)
                End If
                Return issues
            Finally
                Try
                    If doc IsNot Nothing Then
                        If isSlide Then
                            doc.Close()
                        Else
                            doc.Close(0)
                        End If
                    End If
                Finally
                    ' PowerPoint can return an existing application shared with the user.
                    If app IsNot Nothing AndAlso Not isSlide Then app.Quit(0)
                    ReleaseLayoutObject(shapes)
                    ReleaseLayoutObject(slides)
                    ReleaseLayoutObject(doc)
                    ReleaseLayoutObject(files)
                    ReleaseLayoutObject(app)
                End Try
            End Try
        End Function

        ' Restrict replacement to one uniquely identifiable plain text run. All other native
        ' structures (fields, hyperlinks, mixed runs, bullets and paragraph boundaries) stay intact.
        Private Shared Function RetryableRun(part As System.Xml.XmlDocument, issue As OverflowingText) As System.Collections.Generic.List(Of System.Xml.XmlElement)
            Dim ns As New System.Xml.XmlNamespaceManager(part.NameTable)
            ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
            ns.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
            ns.AddNamespace("p", "http://schemas.openxmlformats.org/presentationml/2006/main")
            ns.AddNamespace("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")
            ns.AddNamespace("v", "urn:schemas-microsoft-com:vml")
            Dim candidates As New System.Collections.Generic.List(Of System.Xml.XmlElement)()
            For Each node As System.Xml.XmlElement In part.SelectNodes("//w:txbxContent[count(w:p)=1]/w:p[count(w:r)=1 and not(w:hyperlink) and not(w:fldSimple)]/w:r[count(*)=count(w:rPr)+1]/w:t | //p:sp/p:txBody[count(a:p)=1]/a:p[count(a:r)=1 and not(a:br) and not(a:fld)]/a:r[not(a:rPr/a:hlinkClick)]/a:t", ns)
                Dim shapeName = node.SelectSingleNode("ancestor::p:sp[1]/p:nvSpPr/p:cNvPr/@name | ancestor::wp:anchor[1]/wp:docPr/@name | ancestor::v:shape[1]/@id", ns)
                If shapeName IsNot Nothing AndAlso shapeName.Value = issue.ShapeName AndAlso node.InnerText = issue.Text Then candidates.Add(node)
            Next
            Return candidates
        End Function

        Public Shared Async Function CheckTranslationFit(path As System.String, context As SharedContext.ISharedContext,
                instruction As System.String, secondApi As System.Boolean, token As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of System.String)
            Dim staged = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(path), System.Guid.NewGuid().ToString("N") & System.IO.Path.GetExtension(path))
            Try
                token.ThrowIfCancellationRequested()
                Dim issues = MeasureOfficeBoxes(path)
                If issues.Count = 0 Then Return "No overflow detected in inspected slide shapes or main-document text boxes."
                Dim retryContext As SharedContext.ISharedContext
                If secondApi Then
                    retryContext = CreateIsolatedModelCallContext(context)
                Else
                    Dim model = ResolveIsolatedSpecialTaskModel(context, "OfflineDocs")
                    retryContext = model.Context
                    secondApi = model.UsesSecondApi
                End If
                System.IO.File.Copy(path, staged)
                Dim changed As System.Boolean = False
                Using archive = System.IO.Compression.ZipFile.Open(staged, System.IO.Compression.ZipArchiveMode.Update)
                    Dim parts As New System.Collections.Generic.Dictionary(Of System.IO.Compression.ZipArchiveEntry, System.Xml.XmlDocument)()
                    For Each entry In archive.Entries
                        If entry.FullName <> "word/document.xml" AndAlso Not System.Text.RegularExpressions.Regex.IsMatch(entry.FullName, "^ppt/slides/slide[0-9]+\.xml$") Then Continue For
                        Dim xml As New System.Xml.XmlDocument With {.PreserveWhitespace = True, .XmlResolver = Nothing}
                        Using stream = entry.Open()
                            xml.Load(stream)
                        End Using
                        parts.Add(entry, xml)
                    Next
                    Dim dirty As New System.Collections.Generic.HashSet(Of System.Xml.XmlDocument)()
                    For Each issue In issues
                        token.ThrowIfCancellationRequested()
                        Dim candidates = parts.Values.SelectMany(Function(xml) RetryableRun(xml, issue)).ToList()
                        If candidates.Count <> 1 Then Continue For
                        Dim replacement = Await LLM(retryContext, instruction & System.Environment.NewLine & context.SP_Translate_Office_Retry, issue.Text, UseSecondAPI:=secondApi, HideSplash:=True, cancellationToken:=token)
                        If System.String.IsNullOrWhiteSpace(replacement) OrElse replacement.Contains(vbCr) OrElse replacement.Contains(vbLf) Then Continue For
                        candidates(0).InnerText = replacement.Trim()
                        dirty.Add(candidates(0).OwnerDocument)
                    Next
                    For Each part In parts
                        If Not dirty.Contains(part.Value) Then Continue For
                        Using stream = part.Key.Open()
                            stream.SetLength(0)
                            part.Value.Save(stream)
                        End Using
                        changed = True
                    Next
                End Using
                If changed Then
                    Dim remaining = MeasureOfficeBoxes(staged)
                    If remaining.Count < issues.Count AndAlso remaining.All(Function(item) issues.Any(Function(before) before.Location = item.Location)) Then
                        token.ThrowIfCancellationRequested()
                        System.IO.File.Replace(staged, path, Nothing)
                        issues = remaining
                    End If
                End If
                If issues.Count = 0 Then Return "Text fitted after retry."
                Return "Review overflowing text: " & System.String.Join(" | ", issues.Select(Function(item) item.Location & " — " & item.Text))
            Catch ex As System.OperationCanceledException
                Throw
            Catch ex As System.Exception
                Return "Review the translated layout. Office could not verify the text boxes: " & ex.Message
            Finally
                If System.IO.File.Exists(staged) Then System.IO.File.Delete(staged)
            End Try
        End Function
    End Class
End Namespace
