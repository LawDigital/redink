' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: WordTools.vb
' Purpose: File-level (.docx) tools using DocumentFormat.OpenXml. Operate ONLY
'          on file paths sandboxed by PathPolicy. Never touches the currently
'          open Word document (that surface is worddoc_* tools).
'
' Tools:
'  - word_extract_text: return plain text of .docx file.
'  - word_search: substring/regex search in document; returns paragraph indices.
'  - word_write: replace/insert/append paragraph text WITHOUT markup.
'  - word_markup: same as word_write but produces tracked-change w:ins/w:del.
'  - word_comment_add: add Word comment anchored to matched substring.
'  - word_comment_list / word_comment_remove: manage comments.
'  - word_format: set paragraph style and/or run formatting on match.
'  - word_apply_template: clone template from skill references/ with substitutions.
'  - word_save_as: copy .docx to new path (workspace or Desktop).
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Text.RegularExpressions
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports W = DocumentFormat.OpenXml.Wordprocessing
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary

Namespace Agents

    Public NotInheritable Class WordTools

        Private Sub New()
        End Sub

        Public Const ToolExtract As String = "word_extract_text"
        Public Const ToolSearch As String = "word_search"
        Public Const ToolWrite As String = "word_write"
        Public Const ToolMarkup As String = "word_markup"
        Public Const ToolCommentAdd As String = "word_comment_add"
        Public Const ToolCommentList As String = "word_comment_list"
        Public Const ToolCommentRemove As String = "word_comment_remove"
        Public Const ToolFormat As String = "word_format"
        Public Const ToolApplyTemplate As String = "word_apply_template"
        Public Const ToolSaveAs As String = "word_save_as"

        Public Shared Function IsWordTool(name As String) As Boolean
            If String.IsNullOrWhiteSpace(name) Then Return False

            Select Case name
                Case ToolExtract, ToolSearch, ToolWrite, ToolMarkup,
                     ToolCommentAdd, ToolCommentList, ToolCommentRemove,
                     ToolFormat, ToolApplyTemplate, ToolSaveAs
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Public Shared Function BuildAll() As List(Of ModelConfig)
            Dim tools As New List(Of ModelConfig) From {
                BuildExtract(), BuildSearch(), BuildWrite(), BuildMarkup(),
                BuildCommentAdd(), BuildCommentList(), BuildCommentRemove(),
                BuildFormat(), BuildApplyTemplate(), BuildSaveAs()
            }

            For Each tool As ModelConfig In tools
                If tool Is Nothing Then Continue For
                Select Case tool.ToolName
                    Case ToolWrite, ToolMarkup, ToolCommentAdd, ToolCommentRemove, ToolFormat
                        ArtifactDelivery.EnableOptionalSingleFileArtifactProtocol(tool)
                End Select
            Next

            Return tools
        End Function

        ' --------------------------------------------------------------- dispatch

        Public Shared Function Execute(toolName As String, arguments As IDictionary(Of String, Object)) As String
            Try
                Dim artifactMetadata As OptionalToolArtifactMetadata = Nothing
                Dim isInPlaceArtifactTool As Boolean = False

                Select Case toolName
                    Case ToolWrite, ToolMarkup, ToolCommentAdd, ToolCommentRemove, ToolFormat
                        isInPlaceArtifactTool = True
                End Select

                If isInPlaceArtifactTool Then
                    Dim artifactFailureCode As String = ""
                    Dim artifactFailureMessage As String = ""
                    If Not ArtifactDelivery.TryPrepareOptionalToolArtifactMetadata(
                        arguments,
                        ArtifactStorageKind.Unknown,
                        artifactMetadata,
                        artifactFailureCode,
                        artifactFailureMessage) Then

                        Return Err_(artifactFailureCode, artifactFailureMessage)
                    End If
                End If

                Dim resultJson As String
                Select Case toolName
                    Case ToolExtract
                        resultJson = ExecuteExtract(arguments)
                    Case ToolSearch
                        resultJson = ExecuteSearch(arguments)
                    Case ToolWrite
                        resultJson = ExecuteWriteOrMarkup(arguments, asMarkup:=False)
                    Case ToolMarkup
                        resultJson = ExecuteWriteOrMarkup(arguments, asMarkup:=True)
                    Case ToolCommentAdd
                        resultJson = ExecuteCommentAdd(arguments)
                    Case ToolCommentList
                        resultJson = ExecuteCommentList(arguments)
                    Case ToolCommentRemove
                        resultJson = ExecuteCommentRemove(arguments)
                    Case ToolFormat
                        resultJson = ExecuteFormat(arguments)
                    Case ToolApplyTemplate
                        resultJson = ExecuteApplyTemplate(arguments)
                    Case ToolSaveAs
                        resultJson = ExecuteSaveAs(arguments)
                    Case Else
                        Return Err_("unknown_word_tool", "Unknown tool '" & toolName & "'.")
                End Select

                If isInPlaceArtifactTool AndAlso artifactMetadata IsNot Nothing Then
                    Try
                        Dim resultObject As Newtonsoft.Json.Linq.JObject = TryCast(Newtonsoft.Json.Linq.JToken.Parse(If(resultJson, "")), Newtonsoft.Json.Linq.JObject)
                        If resultObject IsNot Nothing AndAlso resultObject("error") Is Nothing Then
                            Dim physicalPath As String = If(resultObject("path")?.ToString(), "")
                            resultJson = ArtifactDelivery.AttachOptionalSingleFileArtifactToResult(resultJson, artifactMetadata, physicalPath)
                        End If
                    Catch
                    End Try
                End If

                Return resultJson
            Catch uae As UnauthorizedAccessException
                Return Err_("access_denied", uae.Message)
            Catch ex As Exception
                Return Err_("word_tool_failed", ex.Message)
            End Try
        End Function

        ' --------------------------------------------------------------- extract / search

        Private Shared Function ExecuteExtract(args As IDictionary(Of String, Object)) As String
            Dim p As String = PathPolicy.Resolve(GetStr(args, "path"), PathAccess.Read)
            If Not File.Exists(p) Then Return Err_("not_found", "File not found.")

            ' Use the rich sandboxed extractor so the returned text includes footnotes,
            ' endnotes, headers/footers, automatic paragraph/list/heading numbering,
            ' cross-references, tables, and text-box/margin text — not just body runs.
            Dim joined As String = SharedMethods.DocxTextExtractor.ReadDocxSandboxed(p)
            If joined Is Nothing Then joined = String.Empty

            If joined.StartsWith("Error:", StringComparison.Ordinal) Then
                Return Err_("extract_failed", joined.Substring("Error:".Length).Trim())
            End If

            Dim maxChars As Integer = GetInt(args, "max_chars", 0)
            Dim truncated As Boolean = False

            If maxChars > 0 AndAlso joined.Length > maxChars Then
                joined = joined.Substring(0, maxChars)
                truncated = True
            End If

            Return JsonConvert.SerializeObject(New With {
                Key .path = p,
                Key .truncated = truncated,
                Key .text = joined
            })
        End Function

        Private Shared Function ExecuteSearch(args As IDictionary(Of String, Object)) As String
            Dim p As String = PathPolicy.Resolve(GetStr(args, "path"), PathAccess.Read)
            If Not File.Exists(p) Then Return Err_("not_found", "File not found.")

            Dim query As String = GetStr(args, "query")
            If String.IsNullOrWhiteSpace(query) Then Return Err_("missing_query", "query is required.")

            Dim useRegex As Boolean = GetBool(args, "regex", False)
            Dim ignoreCase As Boolean = GetBool(args, "ignore_case", True)
            Dim maxHits As Integer = System.Math.Min(System.Math.Max(GetInt(args, "max_hits", 50), 1), 500)

            ' Fast-fail read: opening the package directly can block for minutes when the file is
            ' open/locked in Word, on a slow network share, or an un-hydrated cloud placeholder.
            ' Copy the bytes with shared read access under a bounded wait, then open from memory so
            ' a stalled file surfaces a clear error instead of hanging the agent loop.
            Dim documentBytes As Byte() = Nothing
            Dim readErrorMessage As String = Nothing
            If Not TryReadAllBytesFastFail(p, documentBytes, readErrorMessage) Then
                Return Err_("file_unavailable", readErrorMessage)
            End If

            Using doc As WordprocessingDocument = WordprocessingDocument.Open(New MemoryStream(documentBytes, writable:=False), isEditable:=False)
                Dim paragraphs As List(Of ParagraphRow) = ExtractParagraphs(doc)
                Dim hits As New List(Of Object)()
                Dim cmp As StringComparison = If(ignoreCase, StringComparison.OrdinalIgnoreCase, StringComparison.Ordinal)
                Dim rx As Regex = Nothing

                If useRegex Then
                    Dim opt As RegexOptions = RegexOptions.CultureInvariant
                    If ignoreCase Then opt = opt Or RegexOptions.IgnoreCase
                    rx = New Regex(query, opt, TimeSpan.FromSeconds(2))
                End If

                For idx As Integer = 0 To paragraphs.Count - 1
                    Dim text As String = paragraphs(idx).Text

                    If useRegex Then
                        For Each m As Match In rx.Matches(text)
                            If hits.Count >= maxHits Then Exit For
                            hits.Add(BuildHit(idx, paragraphs(idx).Story, text, m.Index, m.Length, m.Value))
                        Next
                    Else
                        Dim matches As List(Of Integer()) = FindAllInText(text, query, ignoreCase, maxHits - hits.Count)
                        For Each mm As Integer() In matches
                            hits.Add(BuildHit(idx, paragraphs(idx).Story, text, mm(0), mm(1), text.Substring(mm(0), mm(1))))
                            If hits.Count >= maxHits Then Exit For
                        Next
                    End If

                    If hits.Count >= maxHits Then Exit For
                Next

                Return JsonConvert.SerializeObject(New With {
                    Key .path = p,
                    Key .total = hits.Count,
                    Key .hits = hits
                })
            End Using
        End Function

        ' --------------------------------------------------------------- write / markup

        Private Shared Function ExecuteWriteOrMarkup(args As IDictionary(Of String, Object), asMarkup As Boolean) As String
            Dim p As String = PathPolicy.Resolve(GetStr(args, "path"), PathAccess.Write)
            If Not File.Exists(p) Then Return Err_("not_found", "File not found.")

            Dim authorRaw As String = GetStr(args, "author")
            Dim author As String =
                If(String.IsNullOrWhiteSpace(authorRaw), "Inky", authorRaw)

            Dim tasks As List(Of MarkupTask) = ParseTasks(args)
            If tasks.Count = 0 Then
                Return Err_("missing_tasks", "Provide either op/find/text or a 'tasks' array.")
            End If

            For Each task As MarkupTask In tasks
                If String.IsNullOrWhiteSpace(task.OperationId) Then
                    Return Err_(
                        "missing_operation_id",
                        "Every word_write/word_markup logical operation requires an explicit opaque operation_id. Each batched task must carry its own operation_id.")
                End If
            Next

            Dim results As New List(Of Object)()
            Dim anyApplied As Boolean = False
            Dim failedCount As Integer = 0

            Using doc As WordprocessingDocument =
                WordprocessingDocument.Open(p, isEditable:=True)

                Dim body As W.Body =
                    doc.MainDocumentPart.Document.Body

                For Each t As MarkupTask In tasks
                    Dim op As String =
                        If(String.IsNullOrWhiteSpace(t.Op), "replace", t.Op)

                    If op = "append" Then
                        For Each ln As String In SplitLines(t.Text)
                            body.AppendChild(
                                MakeMarkdownParagraph(
                                    ln,
                                    body,
                                    asMarkup,
                                    author))
                        Next

                        anyApplied = True

                        results.Add(
                            New With {
                                Key .operation_id = t.OperationId,
                                Key .op = op,
                                Key .applied = True
                            })

                        Continue For
                    End If

                    If String.IsNullOrWhiteSpace(t.Find) Then
                        failedCount += 1

                        results.Add(
                            New With {
                                Key .operation_id = t.OperationId,
                                Key .op = op,
                                Key .applied = False,
                                Key .reason = "missing_find",
                                Key .error = "missing_find"
                            })

                        Continue For
                    End If

                    If ContainsParagraphBreak(t.Find) Then
                        failedCount += 1

                        results.Add(
                            New With {
                                Key .operation_id = t.OperationId,
                                Key .op = op,
                                Key .find = t.Find,
                                Key .applied = False,
                                Key .reason = "multi_paragraph_find",
                                Key .error = "multi_paragraph_find",
                                Key .message =
                                    "'find' spans more than one paragraph. Split into one task per paragraph; to merge two paragraphs, replace the first and use op 'delete_paragraph' on the second."
                            })

                        Continue For
                    End If

                    Dim count As Integer = 0

                    For Each sp As StoryParagraph In EnumerateStoryParagraphs(doc)
                        Dim para As W.Paragraph = sp.Para
                        Dim done As Boolean = False

                        If op = "delete_paragraph" Then
                            Dim mStart As Integer
                            Dim mLen As Integer

                            If Not TryFindInText(
                                GetParagraphText(para),
                                t.Find,
                                mStart,
                                mLen) Then

                                Continue For
                            End If

                            MarkParagraphDeleted(
                                para,
                                asMarkup,
                                author,
                                sp.Scope)

                            done = True
                        Else
                            done =
                                ApplyMarkupOpToParagraphSurgical(
                                    para,
                                    t.Find,
                                    t.Text,
                                    op,
                                    asMarkup,
                                    author,
                                    sp.Scope)
                        End If

                        If done Then
                            count += 1
                            If t.OnlyFirst Then Exit For
                        End If
                    Next

                    If count > 0 Then
                        anyApplied = True

                        results.Add(
                            New With {
                                Key .operation_id = t.OperationId,
                                Key .op = op,
                                Key .find = t.Find,
                                Key .applied = True,
                                Key .matches = count
                            })
                    Else
                        failedCount += 1

                        results.Add(
                            New With {
                                Key .operation_id = t.OperationId,
                                Key .op = op,
                                Key .find = t.Find,
                                Key .applied = False,
                                Key .matches = 0,
                                Key .reason = "no_match",
                                Key .suggestions =
                                    SuggestClosestParagraphsAll(
                                        doc,
                                        t.Find,
                                        3)
                            })
                    End If
                Next

                If anyApplied Then
                    doc.MainDocumentPart.Document.Save()

                    If doc.MainDocumentPart.FootnotesPart IsNot Nothing AndAlso
                       doc.MainDocumentPart.FootnotesPart.Footnotes IsNot Nothing Then

                        doc.MainDocumentPart.FootnotesPart.Footnotes.Save()
                    End If

                    If doc.MainDocumentPart.EndnotesPart IsNot Nothing AndAlso
                       doc.MainDocumentPart.EndnotesPart.Endnotes IsNot Nothing Then

                        doc.MainDocumentPart.EndnotesPart.Endnotes.Save()
                    End If
                End If
            End Using

            Dim appliedCount As Integer =
                tasks.Count - failedCount

            Dim status As String

            If failedCount = 0 Then
                status = "complete"
            ElseIf anyApplied Then
                status = "partial"
            Else
                status = "none"
            End If

            Return JsonConvert.SerializeObject(
                New With {
                    Key .path = p,
                    Key .markup = asMarkup,
                    Key .status = status,
                    Key .applied_count = appliedCount,
                    Key .failed_count = failedCount,
                    Key .tasks = results,
                    Key .hint =
                        If(
                            failedCount = 0,
                            Nothing,
                            "Tool call completed with " &
                            failedCount.ToString() &
                            " unresolved task(s). Inspect each failed tasks[].reason/error. " &
                            "For no_match, re-read the CURRENT text and retry only the affected operation using the same operation_id. " &
                            "For missing_find or multi_paragraph_find, correct the task shape before retrying. " &
                            "Do not exceed the per-operation retry cap.")
                })
        End Function

        ''' <summary>
        ''' Returns True when <paramref name="value"/> contains an explicit paragraph
        ''' break, i.e. text that would necessarily straddle more than one W.Paragraph.
        ''' </summary>
        Private Shared Function ContainsParagraphBreak(value As String) As Boolean
            If String.IsNullOrEmpty(value) Then Return False
            Return value.IndexOfAny(New Char() {ChrW(10), ChrW(13)}) >= 0
        End Function


        ' =============================================================================
        ' Surgical, structure-preserving markup engine (file-level).
        '
        ' Unlike ApplyTextOpToParagraph, this NEVER rebuilds a paragraph from plain text.
        ' The paragraph is flattened into an ordered atom stream:
        '   - TextChar atoms: one per character, carrying a cloned rPr and existing
        '     ins/del state, and
        '   - Opaque atoms: any non-text inline content (footnote/endnote references,
        '     fields, drawings/images, tabs, breaks, comment marks, bookmarks, symbols),
        '     cloned verbatim and re-emitted so it is never lost.
        ' Only the matched span is diffed and rewritten. Word-level diff is used by
        ' default; heavily rewritten spans collapse to a single delete+insert using the
        ' shared materialRewrite* thresholds. Inserted text and opaque atoms are pinned to
        ' the formatting/position of the nearest unchanged word.
        ' =============================================================================

        Private Enum AtomKind
            TextChar
            Opaque
        End Enum

        Private NotInheritable Class ParaAtom
            Public Kind As AtomKind
            Public Ch As Char
            Public RunProps As W.RunProperties
            Public OpaqueNode As OpenXmlElement
            Public InExistingIns As Boolean
            Public InExistingDel As Boolean
        End Class

        Private Structure MarkupTask
            Public OperationId As String
            Public Op As String
            Public Find As String
            Public Text As String
            Public OnlyFirst As Boolean
        End Structure

        Private Shared Function CloneRPr(rPr As W.RunProperties) As W.RunProperties
            If rPr Is Nothing Then Return Nothing
            Return CType(rPr.CloneNode(True), W.RunProperties)
        End Function

        Private Shared Sub FlattenParagraph(container As OpenXmlElement,
                                            atoms As List(Of ParaAtom),
                                            inIns As Boolean,
                                            inDel As Boolean)
            For Each child As OpenXmlElement In container.ChildElements
                If TypeOf child Is W.ParagraphProperties Then
                    Continue For
                ElseIf TypeOf child Is W.InsertedRun Then
                    FlattenParagraph(child, atoms, True, inDel)
                ElseIf TypeOf child Is W.DeletedRun Then
                    FlattenParagraph(child, atoms, inIns, True)
                ElseIf TypeOf child Is W.Run Then
                    FlattenRun(DirectCast(child, W.Run), atoms, inIns, inDel)
                ElseIf ContainerHasEditableText(child) Then
                    ' Some editable text lives inside container elements that are direct
                    ' children of the paragraph (e.g. hyperlinks, smart tags, structured-
                    ' document-tag content, custom-XML wrappers). Treating the whole
                    ' container as opaque hides its text from the matcher, so 'find' can
                    ' never anchor there — this is why footnote/endnote text wrapped in such
                    ' a container was reported as no_match even though the text was present.
                    ' Recurse so the inner runs become matchable; genuinely inline objects
                    ' (note references, fields without text, drawings, breaks, symbols,
                    ' bookmarks, comment marks) have no w:t and stay opaque below.
                    FlattenParagraph(child, atoms, inIns, inDel)
                Else
                    atoms.Add(New ParaAtom With {
                        .Kind = AtomKind.Opaque,
                        .OpaqueNode = CType(child.CloneNode(True), OpenXmlElement),
                        .InExistingIns = inIns,
                        .InExistingDel = inDel
                    })
                End If
            Next
        End Sub

        ' True when the element carries editable text (a w:t somewhere inside), i.e. it is a
        ' text-bearing container we can safely recurse into rather than treat as one opaque
        ' blob. Elements without any w:t (drawings, note references, empty fields, bookmarks,
        ' comment range markers) return False and remain opaque.
        Private Shared Function ContainerHasEditableText(el As OpenXmlElement) As Boolean
            If el Is Nothing Then Return False
            Return el.Descendants(Of W.Text)().Any()
        End Function

        Private Shared Sub FlattenRun(run As W.Run,
                                      atoms As List(Of ParaAtom),
                                      inIns As Boolean,
                                      inDel As Boolean)
            Dim rPr As W.RunProperties = run.RunProperties

            Dim hasText As Boolean = run.Elements(Of W.Text)().Any()
            Dim hasSpecial As Boolean = run.ChildElements.Any(
                Function(c) Not (TypeOf c Is W.RunProperties OrElse TypeOf c Is W.Text))

            ' Pure special run (footnote/endnote reference alone, drawing, field char, deleted
            ' text, etc.): no editable text, so keep the whole run opaque and formatting intact.
            If hasSpecial AndAlso Not hasText Then
                atoms.Add(New ParaAtom With {
                    .Kind = AtomKind.Opaque,
                    .OpaqueNode = CType(run.CloneNode(True), OpenXmlElement),
                    .InExistingIns = inIns,
                    .InExistingDel = inDel
                })
                Return
            End If

            ' Pure text run: emit one TextChar atom per character.
            If Not hasSpecial Then
                For Each t As W.Text In run.Elements(Of W.Text)()
                    Dim s As String = t.Text
                    If s Is Nothing Then Continue For
                    For Each ch As Char In s
                        atoms.Add(New ParaAtom With {
                            .Kind = AtomKind.TextChar,
                            .Ch = ch,
                            .RunProps = CloneRPr(rPr),
                            .InExistingIns = inIns,
                            .InExistingDel = inDel
                        })
                    Next
                Next
                Return
            End If

            ' Mixed run: text AND special children in the same run. This is how many
            ' footnote/endnote styles store content (e.g. w:footnoteRef + w:tab + w:t together).
            ' Treating the whole run as opaque would hide its text from the matcher even though
            ' the text is clearly present. Walk the children in order, emitting TextChar atoms
            ' for w:t and wrapping every other child in its own run (carrying the original rPr)
            ' as an opaque atom, so the text becomes matchable and the reference mark / tab /
            ' formatting are preserved verbatim when the paragraph is rebuilt.
            For Each child As OpenXmlElement In run.ChildElements
                If TypeOf child Is W.RunProperties Then
                    Continue For
                ElseIf TypeOf child Is W.Text Then
                    Dim s As String = DirectCast(child, W.Text).Text
                    If s Is Nothing Then Continue For
                    For Each ch As Char In s
                        atoms.Add(New ParaAtom With {
                            .Kind = AtomKind.TextChar,
                            .Ch = ch,
                            .RunProps = CloneRPr(rPr),
                            .InExistingIns = inIns,
                            .InExistingDel = inDel
                        })
                    Next
                Else
                    Dim wrap As New W.Run()
                    If rPr IsNot Nothing Then wrap.AppendChild(CloneRPr(rPr))
                    wrap.AppendChild(CType(child.CloneNode(True), OpenXmlElement))
                    atoms.Add(New ParaAtom With {
                        .Kind = AtomKind.Opaque,
                        .OpaqueNode = wrap,
                        .InExistingIns = inIns,
                        .InExistingDel = inDel
                    })
                End If
            Next
        End Sub

        ' Coalescing emitter. Mode: 1 = plain, 2 = inserted (w:ins), 3 = deleted (w:del).
        Private NotInheritable Class MarkupEmitter
            Public ReadOnly Children As New List(Of OpenXmlElement)()
            Private ReadOnly _scope As OpenXmlElement
            Private ReadOnly _author As String
            Private ReadOnly _sb As New StringBuilder()
            Private _mode As Integer = 0
            Private _rPr As W.RunProperties = Nothing
            Private _rPrXml As String = Nothing

            Public Sub New(scope As OpenXmlElement, author As String)
                _scope = scope
                _author = author
            End Sub

            Public Sub AddChar(ch As Char, rPr As W.RunProperties, mode As Integer)
                Dim xml As String = If(rPr Is Nothing, String.Empty, rPr.OuterXml)
                If _sb.Length > 0 AndAlso (_mode <> mode OrElse Not String.Equals(_rPrXml, xml, StringComparison.Ordinal)) Then
                    Flush()
                End If
                _mode = mode
                _rPr = rPr
                _rPrXml = xml
                _sb.Append(ch)
            End Sub

            Public Sub AddText(text As String, rPr As W.RunProperties, mode As Integer)
                If String.IsNullOrEmpty(text) Then Return
                For Each ch As Char In text
                    AddChar(ch, rPr, mode)
                Next
            End Sub

            Public Sub AddOpaque(node As OpenXmlElement, inIns As Boolean, inDel As Boolean)
                Flush()
                If node Is Nothing Then Return
                If inDel Then
                    Dim d As New W.DeletedRun() With {.Id = NextChangeId(_scope).ToString(), .Author = _author, .Date = DateTime.UtcNow}
                    d.AppendChild(node)
                    Children.Add(d)
                ElseIf inIns Then
                    Dim ins As New W.InsertedRun() With {.Id = NextChangeId(_scope).ToString(), .Author = _author, .Date = DateTime.UtcNow}
                    ins.AppendChild(node)
                    Children.Add(ins)
                Else
                    Children.Add(node)
                End If
            End Sub

            Public Sub Flush()
                If _sb.Length = 0 Then Return
                Dim text As String = _sb.ToString()
                _sb.Clear()

                Dim run As New W.Run()
                If _rPr IsNot Nothing Then run.AppendChild(CType(_rPr.CloneNode(True), W.RunProperties))

                If _mode = 3 Then
                    run.AppendChild(New W.DeletedText(text) With {.Space = SpaceProcessingModeValues.Preserve})
                    Dim d As New W.DeletedRun() With {.Id = NextChangeId(_scope).ToString(), .Author = _author, .Date = DateTime.UtcNow}
                    d.AppendChild(run)
                    Children.Add(d)
                ElseIf _mode = 2 Then
                    run.AppendChild(New W.Text(text) With {.Space = SpaceProcessingModeValues.Preserve})
                    Dim ins As New W.InsertedRun() With {.Id = NextChangeId(_scope).ToString(), .Author = _author, .Date = DateTime.UtcNow}
                    ins.AppendChild(run)
                    Children.Add(ins)
                Else
                    run.AppendChild(New W.Text(text) With {.Space = SpaceProcessingModeValues.Preserve})
                    Children.Add(run)
                End If

                _mode = 0
                _rPr = Nothing
                _rPrXml = Nothing
            End Sub
        End Class

        Private Shared Sub EmitOriginalRange(em As MarkupEmitter, atoms As List(Of ParaAtom), fromIdx As Integer, toIdx As Integer)
            If fromIdx > toIdx Then Return
            For ai As Integer = fromIdx To toIdx
                Dim a As ParaAtom = atoms(ai)
                If a.Kind = AtomKind.TextChar Then
                    Dim mode As Integer = If(a.InExistingDel, 3, If(a.InExistingIns, 2, 1))
                    em.AddChar(a.Ch, a.RunProps, mode)
                Else
                    em.AddOpaque(a.OpaqueNode, a.InExistingIns, a.InExistingDel)
                End If
            Next
        End Sub

        ' Locates 'find' in the paragraph's final-view text projection (opaque atoms and
        ' existing deletions ignored) and returns the first/last atom indices of the match.
        Private Shared Function LocateFindInAtoms(atoms As List(Of ParaAtom),
                                                  find As String,
                                                  ByRef firstAtom As Integer,
                                                  ByRef lastAtom As Integer) As Boolean
            firstAtom = -1
            lastAtom = -1

            Dim finalCharAtom As New List(Of Integer)()
            Dim sbProj As New StringBuilder()
            For i As Integer = 0 To atoms.Count - 1
                Dim a As ParaAtom = atoms(i)
                If a.Kind = AtomKind.TextChar AndAlso Not a.InExistingDel Then
                    finalCharAtom.Add(i)
                    sbProj.Append(a.Ch)
                End If
            Next
            Dim projection As String = sbProj.ToString()

            Dim hit As List(Of Integer()) = FindAllInText(projection, find, False, 1)
            If hit.Count = 0 Then hit = FindAllInText(projection, find, True, 1)
            If hit.Count = 0 Then Return False

            Dim projStart As Integer = hit(0)(0)
            Dim projLen As Integer = hit(0)(1)
            If projLen <= 0 Then Return False

            firstAtom = finalCharAtom(projStart)
            lastAtom = finalCharAtom(projStart + projLen - 1)
            Return True
        End Function

        Private Shared Function ApplyMarkupOpToParagraphSurgical(para As W.Paragraph,
                                                                 find As String,
                                                                 newText As String,
                                                                 op As String,
                                                                 asMarkup As Boolean,
                                                                 author As String,
                                                                 scope As OpenXmlElement) As Boolean
            Dim atoms As New List(Of ParaAtom)()
            FlattenParagraph(para, atoms, False, False)

            ' Final-view projection: characters that are visible in Word (text chars not
            ' inside an existing deletion). Opaque atoms (footnotes/fields/...) contribute
            ' nothing, so 'find' matches irrespective of them.
            Dim finalCharAtom As New List(Of Integer)()
            Dim sbProj As New StringBuilder()
            For i As Integer = 0 To atoms.Count - 1
                Dim a As ParaAtom = atoms(i)
                If a.Kind = AtomKind.TextChar AndAlso Not a.InExistingDel Then
                    finalCharAtom.Add(i)
                    sbProj.Append(a.Ch)
                End If
            Next
            Dim projection As String = sbProj.ToString()

            Dim hit As List(Of Integer()) = FindAllInText(projection, find, False, 1)
            If hit.Count = 0 Then hit = FindAllInText(projection, find, True, 1)
            If hit.Count = 0 Then Return False

            Dim projStart As Integer = hit(0)(0)
            Dim projLen As Integer = hit(0)(1)
            If projLen <= 0 Then Return False
            Dim projEnd As Integer = projStart + projLen

            Dim firstAtom As Integer = finalCharAtom(projStart)
            Dim lastAtom As Integer = finalCharAtom(projEnd - 1)

            Dim midText As String = projection.Substring(projStart, projLen)

            ' Build per-mid-character labels (1 = keep, 3 = delete) and insert buffers.
            Dim labels() As Integer = New Integer(midText.Length - 1) {}
            Dim insertsAt As New Dictionary(Of Integer, String)()

            Select Case op
                Case "insert_before"
                    For k As Integer = 0 To midText.Length - 1
                        labels(k) = 1
                    Next
                    insertsAt(0) = newText
                Case "insert_after"
                    For k As Integer = 0 To midText.Length - 1
                        labels(k) = 1
                    Next
                    insertsAt(midText.Length) = newText
                Case Else ' replace
                    Dim diffs As List(Of SharedMethods.Diff) = ComputeMarkupDiff(midText, newText)
                    Dim pos As Integer = 0
                    For Each d As SharedMethods.Diff In diffs
                        Select Case d.Op
                            Case SharedMethods.Diff.Operation.Equal
                                For k As Integer = 0 To d.Text.Length - 1
                                    labels(pos) = 1
                                    pos += 1
                                Next
                            Case SharedMethods.Diff.Operation.Delete
                                For k As Integer = 0 To d.Text.Length - 1
                                    labels(pos) = 3
                                    pos += 1
                                Next
                            Case SharedMethods.Diff.Operation.Insert
                                Dim cur As String = Nothing
                                insertsAt.TryGetValue(pos, cur)
                                insertsAt(pos) = If(cur, String.Empty) & d.Text
                        End Select
                    Next
            End Select

            Dim em As New MarkupEmitter(scope, author)

            EmitOriginalRange(em, atoms, 0, firstAtom - 1)

            Dim localIndex As Integer = 0
            For ai As Integer = firstAtom To lastAtom
                Dim a As ParaAtom = atoms(ai)

                If a.Kind = AtomKind.TextChar AndAlso Not a.InExistingDel Then
                    Dim pending As String = Nothing
                    If insertsAt.TryGetValue(localIndex, pending) Then
                        ' Pin inserted text formatting to the following (nearest) word.
                        em.AddText(pending, a.RunProps, If(asMarkup, 2, 1))
                    End If

                    If labels(localIndex) = 3 Then
                        If asMarkup Then em.AddChar(a.Ch, a.RunProps, 3)  ' write mode drops deleted text
                    Else
                        em.AddChar(a.Ch, a.RunProps, If(a.InExistingIns AndAlso asMarkup, 2, 1))
                    End If

                    localIndex += 1
                ElseIf a.Kind = AtomKind.TextChar Then
                    em.AddChar(a.Ch, a.RunProps, 3)  ' preserve pre-existing deletion
                Else
                    em.AddOpaque(a.OpaqueNode, a.InExistingIns, a.InExistingDel)
                End If
            Next

            Dim tail As String = Nothing
            If insertsAt.TryGetValue(midText.Length, tail) Then
                em.AddText(tail, atoms(lastAtom).RunProps, If(asMarkup, 2, 1))
            End If

            em.Flush()
            EmitOriginalRange(em, atoms, lastAtom + 1, atoms.Count - 1)
            em.Flush()

            Dim pPr As W.ParagraphProperties = para.Elements(Of W.ParagraphProperties)().FirstOrDefault()
            para.RemoveAllChildren()
            If pPr IsNot Nothing Then para.AppendChild(CType(pPr.CloneNode(True), W.ParagraphProperties))
            For Each c As OpenXmlElement In em.Children
                para.AppendChild(c)
            Next

            Return True
        End Function

        Private Shared Sub MarkParagraphDeleted(para As W.Paragraph, asMarkup As Boolean, author As String, scope As OpenXmlElement)
            If Not asMarkup Then
                para.Remove()
                Return
            End If

            Dim atoms As New List(Of ParaAtom)()
            FlattenParagraph(para, atoms, False, False)

            Dim em As New MarkupEmitter(scope, author)
            For Each a As ParaAtom In atoms
                If a.Kind = AtomKind.TextChar Then
                    em.AddChar(a.Ch, a.RunProps, 3)
                Else
                    em.AddOpaque(a.OpaqueNode, a.InExistingIns, True)
                End If
            Next
            em.Flush()

            Dim pPr As W.ParagraphProperties = para.Elements(Of W.ParagraphProperties)().FirstOrDefault()
            para.RemoveAllChildren()
            If pPr Is Nothing Then pPr = New W.ParagraphProperties()

            ' Mark the paragraph mark itself as deleted so accepting the change merges this
            ' paragraph into the next one.
            Dim mrp As W.ParagraphMarkRunProperties = pPr.Elements(Of W.ParagraphMarkRunProperties)().FirstOrDefault()
            If mrp Is Nothing Then
                mrp = New W.ParagraphMarkRunProperties()
                pPr.AppendChild(mrp)
            End If
            If mrp.Elements(Of W.Deleted)().FirstOrDefault() Is Nothing Then
                mrp.PrependChild(New W.Deleted() With {.Id = NextChangeId(scope).ToString(), .Author = author, .Date = DateTime.UtcNow})
            End If

            para.AppendChild(pPr)
            For Each c As OpenXmlElement In em.Children
                para.AppendChild(c)
            Next
        End Sub

        ' ---------------------------------------------------- diff (word-level / text-level)

        Private Shared Function ComputeMarkupDiff(oldText As String, newText As String) As List(Of SharedMethods.Diff)
            Dim result As New List(Of SharedMethods.Diff)()
            Dim o As String = If(oldText, String.Empty)
            Dim n As String = If(newText, String.Empty)

            If o.Length = 0 AndAlso n.Length = 0 Then Return result
            If o.Length = 0 Then
                result.Add(New SharedMethods.Diff(SharedMethods.Diff.Operation.Insert, n))
                Return result
            End If
            If n.Length = 0 Then
                result.Add(New SharedMethods.Diff(SharedMethods.Diff.Operation.Delete, o))
                Return result
            End If

            If ShouldCollapseToTextLevel(o, n) Then
                result.Add(New SharedMethods.Diff(SharedMethods.Diff.Operation.Delete, o))
                result.Add(New SharedMethods.Diff(SharedMethods.Diff.Operation.Insert, n))
                Return result
            End If

            Dim t1 As List(Of String) = TokenizeForDiff(o)
            Dim t2 As List(Of String) = TokenizeForDiff(n)
            Dim j1 As String = String.Join(vbLf, t1)
            Dim j2 As String = String.Join(vbLf, t2)

            Dim builder As New DiffPlex.DiffBuilder.InlineDiffBuilder(New DiffPlex.Differ())
            Dim model As DiffPlex.DiffBuilder.Model.DiffPaneModel = builder.BuildDiffModel(j1, j2)

            For Each line As DiffPlex.DiffBuilder.Model.DiffPiece In model.Lines
                Dim txt As String = If(line.Text, String.Empty)
                If txt.Length = 0 Then Continue For

                Dim opv As SharedMethods.Diff.Operation
                Select Case line.Type
                    Case DiffPlex.DiffBuilder.Model.ChangeType.Inserted
                        opv = SharedMethods.Diff.Operation.Insert
                    Case DiffPlex.DiffBuilder.Model.ChangeType.Deleted
                        opv = SharedMethods.Diff.Operation.Delete
                    Case Else
                        opv = SharedMethods.Diff.Operation.Equal
                End Select

                If result.Count > 0 AndAlso result(result.Count - 1).Op = opv Then
                    result(result.Count - 1).Text &= txt
                Else
                    result.Add(New SharedMethods.Diff(opv, txt))
                End If
            Next

            Return result
        End Function

        Private Shared Function TokenizeForDiff(text As String) As List(Of String)
            Dim r As New List(Of String)()
            If String.IsNullOrEmpty(text) Then Return r
            For Each m As Match In Regex.Matches(text, "[\p{L}\p{M}\p{N}_]+|\s+|[^\s]", RegexOptions.Singleline)
                If m.Length > 0 Then r.Add(m.Value)
            Next
            Return r
        End Function

        Private Shared Function ShouldCollapseToTextLevel(oldText As String, newText As String) As Boolean
            If String.IsNullOrWhiteSpace(oldText) OrElse String.IsNullOrWhiteSpace(newText) Then Return False

            Dim a As List(Of String) = ComparableTokens(oldText)
            Dim b As List(Of String) = ComparableTokens(newText)
            If a.Count < 7 OrElse b.Count < 7 Then Return False

            Dim lcs As Integer = TokenLcs(a, b)
            Dim similarity As Double = (2.0 * lcs) / (a.Count + b.Count)
            Dim changed As Integer = (a.Count - lcs) + (b.Count - lcs)
            Dim ratio As Double = changed / (a.Count + b.Count)

            Return changed >= SharedMethods.materialRewriteMinimumChangedTokens AndAlso
                   similarity < SharedMethods.materialRewriteSimilarityThreshold AndAlso
                   ratio >= SharedMethods.materialRewriteChangedTokenRatioThreshold
        End Function

        Private Shared Function ComparableTokens(value As String) As List(Of String)
            Dim r As New List(Of String)()
            If String.IsNullOrEmpty(value) Then Return r
            For Each m As Match In Regex.Matches(value, "[\p{L}\p{M}\p{N}_]+")
                r.Add(m.Value.ToLowerInvariant())
            Next
            Return r
        End Function

        Private Shared Function TokenLcs(a As List(Of String), b As List(Of String)) As Integer
            If a.Count = 0 OrElse b.Count = 0 Then Return 0
            Dim prev(b.Count) As Integer
            Dim cur(b.Count) As Integer
            For i As Integer = 1 To a.Count
                For j As Integer = 1 To b.Count
                    If String.Equals(a(i - 1), b(j - 1), StringComparison.Ordinal) Then
                        cur(j) = prev(j - 1) + 1
                    Else
                        cur(j) = System.Math.Max(prev(j), cur(j - 1))
                    End If
                Next
                System.Array.Copy(cur, prev, cur.Length)
                System.Array.Clear(cur, 0, cur.Length)
            Next
            Return prev(b.Count)
        End Function

        ' ---------------------------------------------------- task parsing

        Private Shared Function ParseTasks(args As IDictionary(Of String, Object)) As List(Of MarkupTask)
            Dim result As New List(Of MarkupTask)()
            Dim token As JToken = Nothing

            If args IsNot Nothing AndAlso
               args.ContainsKey("tasks") AndAlso
               args("tasks") IsNot Nothing Then

                Try
                    token = JToken.FromObject(args("tasks"))
                Catch
                End Try
            End If

            If token IsNot Nothing AndAlso token.Type = JTokenType.Array Then
                For Each it As JToken In CType(token, JArray)
                    result.Add(New MarkupTask With {
                        .OperationId = JStr(it, "operation_id"),
                        .Op = NormOp(JStr(it, "op")),
                        .Find = JStr(it, "find"),
                        .Text = JStr(it, "text"),
                        .OnlyFirst = JBool(it, "only_first", True)
                    })
                Next
            Else
                result.Add(New MarkupTask With {
                    .OperationId = GetStr(args, "operation_id"),
                    .Op = NormOp(GetStr(args, "op")),
                    .Find = GetStr(args, "find"),
                    .Text = GetStr(args, "text"),
                    .OnlyFirst = GetBool(args, "only_first", True)
                })
            End If

            Return result
        End Function

        ' Returns the paragraph texts most similar to an unmatched 'find', so the model can
        ' re-anchor a failed task against the current document instead of guessing.
        Private Shared Function SuggestClosestParagraphs(body As W.Body, find As String, maxItems As Integer) As List(Of String)
            Dim suggestions As New List(Of String)()
            If body Is Nothing OrElse String.IsNullOrWhiteSpace(find) Then Return suggestions

            Dim needle As List(Of String) = ComparableTokens(find)
            If needle.Count = 0 Then Return suggestions

            Dim scored As New List(Of KeyValuePair(Of Double, String))()
            For Each para As W.Paragraph In body.Descendants(Of W.Paragraph)()
                Dim pt As String = GetParagraphText(para)
                If String.IsNullOrWhiteSpace(pt) Then Continue For

                Dim hay As List(Of String) = ComparableTokens(pt)
                If hay.Count = 0 Then Continue For

                Dim lcs As Integer = TokenLcs(needle, hay)
                Dim score As Double = (2.0 * lcs) / (needle.Count + hay.Count)
                If score <= 0 Then Continue For

                Dim snippet As String = pt.Trim()
                If snippet.Length > 200 Then snippet = snippet.Substring(0, 200) & "…"
                scored.Add(New KeyValuePair(Of Double, String)(score, snippet))
            Next

            scored.Sort(Function(a, b) b.Key.CompareTo(a.Key))
            For i As Integer = 0 To System.Math.Min(maxItems, scored.Count) - 1
                suggestions.Add(scored(i).Value)
            Next
            Return suggestions
        End Function

        ' Like SuggestClosestParagraphs but scans body, footnotes and endnotes so a failed
        ' footnote/endnote anchor can still be re-anchored against the current document.
        Private Shared Function SuggestClosestParagraphsAll(doc As WordprocessingDocument, find As String, maxItems As Integer) As List(Of String)
            Dim suggestions As New List(Of String)()
            If doc Is Nothing OrElse String.IsNullOrWhiteSpace(find) Then Return suggestions

            Dim needle As List(Of String) = ComparableTokens(find)
            If needle.Count = 0 Then Return suggestions

            Dim scored As New List(Of KeyValuePair(Of Double, String))()
            For Each sp As StoryParagraph In EnumerateStoryParagraphs(doc)
                Dim pt As String = GetParagraphText(sp.Para)
                If String.IsNullOrWhiteSpace(pt) Then Continue For

                Dim hay As List(Of String) = ComparableTokens(pt)
                If hay.Count = 0 Then Continue For

                Dim lcs As Integer = TokenLcs(needle, hay)
                Dim score As Double = (2.0 * lcs) / (needle.Count + hay.Count)
                If score <= 0 Then Continue For

                Dim snippet As String = pt.Trim()
                If snippet.Length > 200 Then snippet = snippet.Substring(0, 200) & "…"
                scored.Add(New KeyValuePair(Of Double, String)(score, snippet))
            Next

            scored.Sort(Function(a, b) b.Key.CompareTo(a.Key))
            For i As Integer = 0 To System.Math.Min(maxItems, scored.Count) - 1
                suggestions.Add(scored(i).Value)
            Next
            Return suggestions
        End Function

        Private Shared Function NormOp(op As String) As String
            If String.IsNullOrWhiteSpace(op) Then Return "replace"
            Return op.Trim().ToLowerInvariant()
        End Function

        Private Shared Function JStr(t As JToken, name As String) As String
            If t Is Nothing Then Return String.Empty
            Dim v As JToken = t(name)
            If v Is Nothing OrElse v.Type = JTokenType.Null Then Return String.Empty
            Return v.ToString()
        End Function

        Private Shared Function JBool(t As JToken, name As String, defaultValue As Boolean) As Boolean
            If t Is Nothing Then Return defaultValue
            Dim v As JToken = t(name)
            If v Is Nothing OrElse v.Type = JTokenType.Null Then Return defaultValue
            Select Case v.ToString().Trim().ToLowerInvariant()
                Case "true", "1", "yes" : Return True
                Case "false", "0", "no" : Return False
                Case Else : Return defaultValue
            End Select
        End Function

        ' --------------------------------------------------------------- comments

        Private Shared Function ExecuteCommentAdd(args As IDictionary(Of String, Object)) As String
            Dim p As String = PathPolicy.Resolve(GetStr(args, "path"), PathAccess.Write)
            If Not File.Exists(p) Then Return Err_("not_found", "File not found.")

            Dim defAuthor As String =
                If(String.IsNullOrWhiteSpace(GetStr(args, "author")),
                   "Inky",
                   GetStr(args, "author"))

            Dim defInitials As String =
                If(String.IsNullOrWhiteSpace(GetStr(args, "initials")),
                   "I",
                   GetStr(args, "initials"))

            Dim tasks As List(Of CommentTask) = ParseCommentTasks(args)
            If tasks.Count = 0 Then
                Return Err_("missing_tasks", "Provide either find/text or a 'tasks' array.")
            End If

            For Each task As CommentTask In tasks
                If String.IsNullOrWhiteSpace(task.OperationId) Then
                    Return Err_(
                        "missing_operation_id",
                        "Every word_comment_add logical operation requires an explicit opaque operation_id. Each batched task must carry its own operation_id.")
                End If
            Next

            Dim results As New List(Of Object)()
            Dim anyApplied As Boolean = False
            Dim failedCount As Integer = 0

            Using doc As WordprocessingDocument = WordprocessingDocument.Open(p, isEditable:=True)
                Dim main As MainDocumentPart = doc.MainDocumentPart
                Dim body As W.Body = main.Document.Body
                Dim commentsPart As WordprocessingCommentsPart =
                    main.WordprocessingCommentsPart

                If commentsPart Is Nothing Then
                    commentsPart = main.AddNewPart(Of WordprocessingCommentsPart)()
                    commentsPart.Comments = New W.Comments()
                End If

                For Each t As CommentTask In tasks
                    If String.IsNullOrWhiteSpace(t.Find) Then
                        failedCount += 1

                        results.Add(New With {
                            Key .operation_id = t.OperationId,
                            Key .find = t.Find,
                            Key .applied = False,
                            Key .reason = "missing_find"
                        })

                        Continue For
                    End If

                    If String.IsNullOrWhiteSpace(t.Text) Then
                        failedCount += 1

                        results.Add(New With {
                            Key .operation_id = t.OperationId,
                            Key .find = t.Find,
                            Key .applied = False,
                            Key .reason = "missing_text"
                        })

                        Continue For
                    End If

                    If ContainsParagraphBreak(t.Find) Then
                        failedCount += 1

                        results.Add(New With {
                            Key .operation_id = t.OperationId,
                            Key .find = t.Find,
                            Key .applied = False,
                            Key .reason = "multi_paragraph_find"
                        })

                        Continue For
                    End If

                    Dim newId As Integer = NextCommentId(commentsPart.Comments)
                    Dim attached As Boolean = False

                    For Each para As W.Paragraph In body.Descendants(Of W.Paragraph)().ToList()
                        If Not LocateFindProbe(para, t.Find) Then Continue For

                        If AttachCommentSurgical(para, t.Find, newId, body) Then
                            attached = True
                            Exit For
                        End If
                    Next

                    If attached Then
                        Dim cmt As New W.Comment() With {
                            .Id = newId.ToString(),
                            .Author = If(
                                String.IsNullOrWhiteSpace(t.Author),
                                defAuthor,
                                t.Author),
                            .Initials = If(
                                String.IsNullOrWhiteSpace(t.Initials),
                                defInitials,
                                t.Initials),
                            .Date = DateTime.UtcNow
                        }

                        cmt.AppendChild(
                            New W.Paragraph(
                                New W.Run(
                                    New W.Text(t.Text) With {
                                        .Space = SpaceProcessingModeValues.Preserve
                                    })))

                        commentsPart.Comments.AppendChild(cmt)
                        anyApplied = True

                        results.Add(New With {
                            Key .operation_id = t.OperationId,
                            Key .find = t.Find,
                            Key .applied = True,
                            Key .comment_id = newId
                        })
                    Else
                        failedCount += 1

                        results.Add(New With {
                            Key .operation_id = t.OperationId,
                            Key .find = t.Find,
                            Key .applied = False,
                            Key .reason = "no_match",
                            Key .suggestions = SuggestClosestParagraphs(body, t.Find, 3)
                        })
                    End If
                Next

                If anyApplied Then
                    commentsPart.Comments.Save()
                    main.Document.Save()
                End If
            End Using

            Dim status As String

            If failedCount = 0 Then
                status = "complete"
            ElseIf anyApplied Then
                status = "partial"
            Else
                status = "none"
            End If

            Return JsonConvert.SerializeObject(New With {
                Key .path = p,
                Key .status = status,
                Key .applied_count = tasks.Count - failedCount,
                Key .failed_count = failedCount,
                Key .tasks = results,
                Key .hint = If(
                    failedCount = 0,
                    Nothing,
                    "Tool call succeeded. " & failedCount.ToString() &
                    " comment(s) found no matching anchor (see tasks[].reason='no_match' and tasks[].suggestions). " &
                    "Re-read the document for the CURRENT text and retry only the failed 'find' values. Do not treat this as blocked.")
            })
        End Function

        Private Structure CommentTask
            Public OperationId As String
            Public Find As String
            Public Text As String
            Public Author As String
            Public Initials As String
        End Structure

        Private Shared Function ParseCommentTasks(args As IDictionary(Of String, Object)) As List(Of CommentTask)
            Dim result As New List(Of CommentTask)()
            Dim token As JToken = Nothing

            If args IsNot Nothing AndAlso
               args.ContainsKey("tasks") AndAlso
               args("tasks") IsNot Nothing Then

                Try
                    token = JToken.FromObject(args("tasks"))
                Catch
                End Try
            End If

            If token IsNot Nothing AndAlso token.Type = JTokenType.Array Then
                For Each it As JToken In CType(token, JArray)
                    result.Add(New CommentTask With {
                        .OperationId = JStr(it, "operation_id"),
                        .Find = JStr(it, "find"),
                        .Text = JStr(it, "text"),
                        .Author = JStr(it, "author"),
                        .Initials = JStr(it, "initials")
                    })
                Next
            Else
                result.Add(New CommentTask With {
                    .OperationId = GetStr(args, "operation_id"),
                    .Find = GetStr(args, "find"),
                    .Text = GetStr(args, "text"),
                    .Author = GetStr(args, "author"),
                    .Initials = GetStr(args, "initials")
                })
            End If

            Return result
        End Function

        ' Lightweight pre-check so we only run the surgical attach on a paragraph that matches.
        Private Shared Function LocateFindProbe(para As W.Paragraph, find As String) As Boolean
            Dim mStart As Integer
            Dim mLen As Integer
            Return TryFindInText(GetParagraphText(para), find, mStart, mLen)
        End Function

        Private Shared Function ExecuteCommentList(args As IDictionary(Of String, Object)) As String
            Dim p As String = PathPolicy.Resolve(GetStr(args, "path"), PathAccess.Read)
            If Not File.Exists(p) Then Return Err_("not_found", "File not found.")

            Using doc As WordprocessingDocument = WordprocessingDocument.Open(p, isEditable:=False)
                Dim part As WordprocessingCommentsPart = doc.MainDocumentPart.WordprocessingCommentsPart
                Dim list As New List(Of Object)()

                If part IsNot Nothing AndAlso part.Comments IsNot Nothing Then
                    For Each c As W.Comment In part.Comments.Elements(Of W.Comment)()
                        list.Add(New With {
                            Key .id = If(c.Id Is Nothing, Nothing, c.Id.Value),
                            Key .author = If(c.Author Is Nothing, Nothing, c.Author.Value),
                            Key .initials = If(c.Initials Is Nothing, Nothing, c.Initials.Value),
                            Key .date = If(c.Date Is Nothing, CType(Nothing, Nullable(Of DateTime)), c.Date.Value),
                            Key .text = c.InnerText
                        })
                    Next
                End If

                Return JsonConvert.SerializeObject(New With {
                    Key .path = p,
                    Key .comments = list
                })
            End Using
        End Function

        Private Shared Function ExecuteCommentRemove(args As IDictionary(Of String, Object)) As String
            Dim p As String = PathPolicy.Resolve(GetStr(args, "path"), PathAccess.Write)
            If Not File.Exists(p) Then Return Err_("not_found", "File not found.")

            Dim id As String = GetStr(args, "id")
            If String.IsNullOrWhiteSpace(id) Then Return Err_("missing_id", "id is required.")

            Using doc As WordprocessingDocument = WordprocessingDocument.Open(p, isEditable:=True)
                Dim main As MainDocumentPart = doc.MainDocumentPart
                Dim part As WordprocessingCommentsPart = main.WordprocessingCommentsPart

                If part Is Nothing OrElse part.Comments Is Nothing Then
                    Return Err_("not_found", "No comments part.")
                End If

                Dim cmt As W.Comment = part.Comments.Elements(Of W.Comment)().
                    FirstOrDefault(Function(c) c.Id IsNot Nothing AndAlso c.Id.Value = id)

                If cmt Is Nothing Then
                    Return Err_("not_found", "No comment with id '" & id & "'.")
                End If

                cmt.Remove()

                For Each n As W.CommentRangeStart In main.Document.Body.Descendants(Of W.CommentRangeStart)().
                    Where(Function(x) x.Id IsNot Nothing AndAlso x.Id.Value = id).ToList()
                    n.Remove()
                Next

                For Each n As W.CommentRangeEnd In main.Document.Body.Descendants(Of W.CommentRangeEnd)().
                    Where(Function(x) x.Id IsNot Nothing AndAlso x.Id.Value = id).ToList()
                    n.Remove()
                Next

                For Each n As W.CommentReference In main.Document.Body.Descendants(Of W.CommentReference)().
                    Where(Function(x) x.Id IsNot Nothing AndAlso x.Id.Value = id).ToList()
                    If n.Parent IsNot Nothing Then
                        n.Parent.Remove()
                    End If
                Next

                part.Comments.Save()
                main.Document.Save()

                Return JsonConvert.SerializeObject(New With {
                    Key .path = p,
                    Key .removed_id = id
                })
            End Using
        End Function

        ' --------------------------------------------------------------- format

        Private Shared Function ExecuteFormat(args As IDictionary(Of String, Object)) As String
            Dim p As String = PathPolicy.Resolve(GetStr(args, "path"), PathAccess.Write)
            If Not File.Exists(p) Then Return Err_("not_found", "File not found.")

            Dim find As String = GetStr(args, "find")
            Dim styleId As String = GetStr(args, "style") ' W.Paragraph style id (e.g. "Heading1")
            Dim bold As Boolean? = GetNullableBool(args, "bold")
            Dim italic As Boolean? = GetNullableBool(args, "italic")
            Dim underline As Boolean? = GetNullableBool(args, "underline")
            Dim sizePt As Integer = GetInt(args, "size", 0)
            Dim color As String = GetStr(args, "color") ' "RRGGBB" hex
            Dim align As String = GetStr(args, "align").ToLowerInvariant() ' left|center|right|justify

            If String.IsNullOrWhiteSpace(find) Then
                Return Err_("missing_find", "find is required (use empty W.Paragraph match by passing the W.Paragraph's text).")
            End If

            Using doc As WordprocessingDocument = WordprocessingDocument.Open(p, isEditable:=True)
                Dim mutated As Integer = 0

                ' Body, footnotes and endnotes are all formatted so the match is styled
                ' wherever it appears in the document.
                For Each sp As StoryParagraph In EnumerateStoryParagraphs(doc)
                    Dim para As W.Paragraph = sp.Para
                    Dim pt As String = GetParagraphText(para)
                    Dim fmtStart As Integer
                    Dim fmtLen As Integer
                    If Not TryFindInText(pt, find, fmtStart, fmtLen) Then Continue For

                    Dim pPr As W.ParagraphProperties = para.Elements(Of W.ParagraphProperties)().FirstOrDefault()
                    If pPr Is Nothing Then
                        pPr = New W.ParagraphProperties()
                        para.InsertAt(Of W.ParagraphProperties)(pPr, 0)
                    End If

                    If Not String.IsNullOrWhiteSpace(styleId) Then
                        Dim psid As W.ParagraphStyleId = pPr.Elements(Of W.ParagraphStyleId)().FirstOrDefault()
                        If psid Is Nothing Then
                            pPr.AppendChild(New W.ParagraphStyleId() With {.Val = styleId})
                        Else
                            psid.Val = styleId
                        End If
                    End If

                    Select Case align
                        Case "left", "center", "right", "justify"
                            Dim ja As W.Justification = pPr.Elements(Of W.Justification)().FirstOrDefault()
                            If ja Is Nothing Then
                                ja = New W.Justification()
                                pPr.AppendChild(ja)
                            End If
                            ja.Val = AlignmentFromString(align)
                    End Select

                    For Each run As W.Run In para.Elements(Of W.Run)()
                        Dim rPr As W.RunProperties = run.RunProperties
                        If rPr Is Nothing Then
                            rPr = New W.RunProperties()
                            run.InsertAt(Of W.RunProperties)(rPr, 0)
                        End If

                        If bold.HasValue Then
                            SetBool(Of W.Bold)(rPr, bold.Value)
                        End If

                        If italic.HasValue Then
                            SetBool(Of W.Italic)(rPr, italic.Value)
                        End If

                        If underline.HasValue Then
                            Dim u As W.Underline = rPr.Elements(Of W.Underline)().FirstOrDefault()
                            If underline.Value Then
                                If u Is Nothing Then
                                    rPr.AppendChild(New W.Underline() With {.Val = W.UnderlineValues.Single})
                                End If
                            ElseIf u IsNot Nothing Then
                                u.Remove()
                            End If
                        End If

                        If sizePt > 0 Then
                            Dim sz As W.FontSize = rPr.Elements(Of W.FontSize)().FirstOrDefault()
                            If sz Is Nothing Then
                                sz = New W.FontSize()
                                rPr.AppendChild(sz)
                            End If
                            sz.Val = (sizePt * 2).ToString() ' half-points
                        End If

                        If Not String.IsNullOrWhiteSpace(color) Then
                            Dim cc As W.Color = rPr.Elements(Of W.Color)().FirstOrDefault()
                            If cc Is Nothing Then
                                cc = New W.Color()
                                rPr.AppendChild(cc)
                            End If
                            cc.Val = color.TrimStart("#"c)
                        End If
                    Next

                    mutated += 1
                Next

                If mutated = 0 Then
                    Return Err_("no_match", "No W.Paragraph contained the 'find' text.")
                End If

                doc.MainDocumentPart.Document.Save()
                If doc.MainDocumentPart.FootnotesPart IsNot Nothing AndAlso doc.MainDocumentPart.FootnotesPart.Footnotes IsNot Nothing Then
                    doc.MainDocumentPart.FootnotesPart.Footnotes.Save()
                End If
                If doc.MainDocumentPart.EndnotesPart IsNot Nothing AndAlso doc.MainDocumentPart.EndnotesPart.Endnotes IsNot Nothing Then
                    doc.MainDocumentPart.EndnotesPart.Endnotes.Save()
                End If

                Return JsonConvert.SerializeObject(New With {
                    Key .path = p,
                    Key .paragraphs_changed = mutated
                })
            End Using
        End Function

        ' --------------------------------------------------------------- template / save_as

        Private Shared Function ExecuteApplyTemplate(args As IDictionary(Of String, Object)) As String
            Dim skillName As String = GetStr(args, "skill")
            Dim relTemplate As String = GetStr(args, "template") ' relative to the skill's references/
            Dim outName As String = If(GetStr(args, "output_name"), "from_template.docx")
            Dim subsToken As JToken = Nothing

            If args IsNot Nothing AndAlso args.ContainsKey("substitutions") Then
                Try
                    subsToken = JToken.FromObject(args("substitutions"))
                Catch
                End Try
            End If

            If String.IsNullOrWhiteSpace(skillName) OrElse String.IsNullOrWhiteSpace(relTemplate) Then
                Return Err_("missing_args", "Both 'skill' and 'template' are required.")
            End If

            Dim sk = AgentResources.FindSkill(skillName)
            If sk Is Nothing OrElse String.IsNullOrWhiteSpace(sk.ReferencesDir) OrElse Not Directory.Exists(sk.ReferencesDir) Then
                Return Err_("skill_not_found", "Skill or references/ directory not found.")
            End If

            Dim src As String = Path.GetFullPath(Path.Combine(sk.ReferencesDir, relTemplate))
            If Not src.StartsWith(Path.GetFullPath(sk.ReferencesDir), StringComparison.OrdinalIgnoreCase) Then
                Return Err_("path_escape", "Template path escapes references/.")
            End If

            If Not File.Exists(src) Then Return Err_("not_found", "Template not found: " & src)

            Dim artifactMetadata As OptionalToolArtifactMetadata = Nothing
            Dim artifactFailureCode As String = ""
            Dim artifactFailureMessage As String = ""

            If Not ArtifactDelivery.TryPrepareOptionalToolArtifactMetadata(
                args,
                ArtifactStorageKind.Unknown,
                artifactMetadata,
                artifactFailureCode,
                artifactFailureMessage) Then

                Return Err_(artifactFailureCode, artifactFailureMessage)
            End If

            Dim dst As String = PathPolicy.NewWritablePath(outName)
            File.Copy(src, dst, overwrite:=False)

            Dim placeholders As New Dictionary(Of String, String)(StringComparer.Ordinal)
            If subsToken IsNot Nothing AndAlso subsToken.Type = JTokenType.Object Then
                For Each kv As JProperty In CType(subsToken, JObject).Properties()
                    placeholders("{{" & kv.Name & "}}") = If(kv.Value Is Nothing, "", kv.Value.ToString())
                Next
            End If

            If placeholders.Count > 0 Then
                Using doc As WordprocessingDocument = WordprocessingDocument.Open(dst, isEditable:=True)
                    For Each para As W.Paragraph In doc.MainDocumentPart.Document.Body.Descendants(Of W.Paragraph)().ToList()
                        Dim pt As String = GetParagraphText(para)
                        Dim changed As Boolean = False
                        Dim newText As String = pt

                        For Each kv As KeyValuePair(Of String, String) In placeholders
                            If newText.IndexOf(kv.Key, StringComparison.Ordinal) >= 0 Then
                                newText = newText.Replace(kv.Key, kv.Value)
                                changed = True
                            End If
                        Next

                        If changed Then
                            Dim pPr As W.ParagraphProperties = para.Elements(Of W.ParagraphProperties)().FirstOrDefault()
                            para.RemoveAllChildren()

                            If pPr IsNot Nothing Then
                                para.AppendChild(CType(pPr.CloneNode(True), W.ParagraphProperties))
                            End If

                            para.AppendChild(MakeRun(newText))
                        End If
                    Next

                    doc.MainDocumentPart.Document.Save()
                End Using
            End If

            If artifactMetadata Is Nothing Then
                Return JsonConvert.SerializeObject(New With {
                    Key .path = dst,
                    Key .template = src,
                    Key .substitutions = placeholders.Count
                })
            End If

            Return JsonConvert.SerializeObject(New With {
                Key .path = dst,
                Key .template = src,
                Key .substitutions = placeholders.Count,
                Key .created = True,
                Key .produces_user_deliverable = artifactMetadata.ProducesUserDeliverable,
                Key .produces_intermediate_data = artifactMetadata.ProducesIntermediateData,
                Key .artifacts = New System.Object() {artifactMetadata.BuildArtifact(dst)}
            })
        End Function

        Private Shared Function ExecuteSaveAs(args As IDictionary(Of String, Object)) As String
            Dim artifactId As String =
                GetStr(args, "artifact_id").Trim()

            Dim logicalDeliverableId As String =
                GetStr(args, "logical_deliverable_id").Trim()

            Dim outputSlotId As String =
                GetStr(args, "output_slot_id").Trim()

            If artifactId = "" OrElse logicalDeliverableId = "" OrElse outputSlotId = "" Then
                Return Err_(
                    "missing_artifact_identity",
                    "word_save_as requires explicit artifact_id, logical_deliverable_id, and output_slot_id values.")
            End If

            Dim expectedToken As JToken = Nothing

            If args IsNot Nothing AndAlso
               args.ContainsKey("expected_artifacts") AndAlso
               args("expected_artifacts") IsNot Nothing Then

                Try
                    expectedToken = JToken.FromObject(args("expected_artifacts"))
                Catch
                End Try
            End If

            If expectedToken Is Nothing OrElse
               expectedToken.Type <> JTokenType.Array OrElse
               DirectCast(expectedToken, JArray).Count = 0 Then

                Return Err_(
                    "missing_expected_artifacts",
                    "word_save_as requires expected_artifacts containing the complete expected final output-slot set.")
            End If

            Dim currentSlotDeclared As Boolean = False

            For Each expectedItem As JToken In DirectCast(expectedToken, JArray)
                Dim expectedObject As JObject = TryCast(expectedItem, JObject)
                If expectedObject Is Nothing Then
                    Return Err_(
                        "invalid_expected_artifacts",
                        "Every expected_artifacts item must contain logical_deliverable_id and output_slot_id.")
                End If

                Dim expectedLogicalId As String =
                    If(expectedObject.Value(Of String)("logical_deliverable_id"), "").Trim()

                Dim expectedSlotId As String =
                    If(expectedObject.Value(Of String)("output_slot_id"), "").Trim()

                If expectedLogicalId = "" OrElse expectedSlotId = "" Then
                    Return Err_(
                        "invalid_expected_artifacts",
                        "Every expected_artifacts item must contain non-empty logical_deliverable_id and output_slot_id.")
                End If

                If String.Equals(
                    expectedLogicalId,
                    logicalDeliverableId,
                    StringComparison.Ordinal) AndAlso
                   String.Equals(
                    expectedSlotId,
                    outputSlotId,
                    StringComparison.Ordinal) Then

                    currentSlotDeclared = True
                End If
            Next

            If Not currentSlotDeclared Then
                Return Err_(
                    "current_output_slot_not_expected",
                    "The current logical_deliverable_id/output_slot_id pair must appear in expected_artifacts.")
            End If

            Dim src As String = PathPolicy.Resolve(GetStr(args, "source"), PathAccess.Read)
            If Not File.Exists(src) Then Return Err_("not_found", "Source not found.")

            Dim outName As String =
                If(GetStr(args, "output_name"), Path.GetFileName(src))

            Dim dst As String = PathPolicy.NewWritablePath(outName)
            File.Copy(src, dst, overwrite:=False)

            Dim artifact As New JObject From {
                {"artifact_id", GetStr(args, "artifact_id")},
                {"logical_deliverable_id", GetStr(args, "logical_deliverable_id")},
                {"output_slot_id", GetStr(args, "output_slot_id")},
                {"path", dst},
                {"state", "final"},
                {"delivery_intent", "deliver_to_user"},
                {"storage_kind", GetStr(args, "storage_kind")},
                {"supersedes_artifact_id", GetStr(args, "supersedes_artifact_id")}
            }

            Return JsonConvert.SerializeObject(New With {
                Key .source = src,
                Key .path = dst,
                Key .created = True,
                Key .saved = True,
                Key .produces_user_deliverable = True,
                Key .artifacts = New Object() {artifact}
            })
        End Function

        ' --------------------------------------------------------------- OOXML helpers

        Private Class ParagraphRow
            Public Index As Integer
            Public Text As String
            Public Story As String
        End Class

        ' A paragraph paired with the story it lives in ("body" | "footnote" | "endnote")
        ' and the story root element used as the change-id scope for tracked changes.
        Private Structure StoryParagraph
            Public Para As W.Paragraph
            Public Scope As OpenXmlElement
            Public Story As String
        End Structure

        ' Enumerates every editable paragraph across the main body, the footnotes part and
        ' the endnotes part. Headers/footers are intentionally excluded (separate stories).
        Private Shared Function EnumerateStoryParagraphs(doc As WordprocessingDocument) As List(Of StoryParagraph)
            Dim rows As New List(Of StoryParagraph)()

            Dim main As MainDocumentPart = doc.MainDocumentPart
            If main Is Nothing Then Return rows

            Dim body As W.Body = main.Document.Body
            If body IsNot Nothing Then
                For Each para As W.Paragraph In body.Descendants(Of W.Paragraph)().ToList()
                    rows.Add(New StoryParagraph With {.Para = para, .Scope = body, .Story = "body"})
                Next
            End If

            If main.FootnotesPart IsNot Nothing AndAlso main.FootnotesPart.Footnotes IsNot Nothing Then
                Dim fn As W.Footnotes = main.FootnotesPart.Footnotes
                For Each para As W.Paragraph In fn.Descendants(Of W.Paragraph)().ToList()
                    rows.Add(New StoryParagraph With {.Para = para, .Scope = fn, .Story = "footnote"})
                Next
            End If

            If main.EndnotesPart IsNot Nothing AndAlso main.EndnotesPart.Endnotes IsNot Nothing Then
                Dim en As W.Endnotes = main.EndnotesPart.Endnotes
                For Each para As W.Paragraph In en.Descendants(Of W.Paragraph)().ToList()
                    rows.Add(New StoryParagraph With {.Para = para, .Scope = en, .Story = "endnote"})
                Next
            End If

            Return rows
        End Function

        Private Shared Function ExtractParagraphs(doc As WordprocessingDocument) As List(Of ParagraphRow)
            Dim output As New List(Of ParagraphRow)()
            Dim i As Integer = 0

            For Each sp As StoryParagraph In EnumerateStoryParagraphs(doc)
                output.Add(New ParagraphRow With {
                    .Index = i,
                    .Text = GetParagraphText(sp.Para),
                    .Story = sp.Story
                })
                i += 1
            Next

            Return output
        End Function

        Private Shared Function GetParagraphText(p As W.Paragraph) As String
            Dim sb As New StringBuilder()

            For Each el As OpenXmlElement In p.Descendants()
                If TypeOf el Is W.Text Then
                    sb.Append(DirectCast(el, W.Text).Text)
                ElseIf TypeOf el Is W.TabChar Then
                    sb.Append(vbTab)
                ElseIf TypeOf el Is W.Break OrElse TypeOf el Is W.CarriageReturn Then
                    sb.Append(vbLf)
                End If
            Next

            Return sb.ToString()
        End Function

        ' --------------------------------------------------------------- resilient matching
        '
        ' Model-supplied 'find'/'query' rarely byte-matches the paragraph text: runs get
        ' split, whitespace collapses, NBSP/smart quotes/dashes differ. These helpers build a
        ' whitespace- and punctuation-normalized projection of the text while keeping a map
        ' back to the ORIGINAL character offsets, so callers still operate on real positions.

        Private Shared Function MapCharForMatch(c As Char) As Char
            If Char.IsWhiteSpace(c) OrElse c = ChrW(&HA0) OrElse c = ChrW(&H202F) Then
                Return " "c
            End If

            Select Case c
                Case ChrW(&H2018), ChrW(&H2019), ChrW(&H201A), ChrW(&H2032), "`"c, "´"c
                    Return "'"c
                Case ChrW(&H201C), ChrW(&H201D), ChrW(&H201E), ChrW(&H2033)
                    Return """"c
                Case ChrW(&H2013), ChrW(&H2014), ChrW(&H2212)
                    Return "-"c
                Case Else
                    Return c
            End Select
        End Function

        ' Builds a normalized string. When map isNot Nothing, map(i) is the original index of
        ' normalized char i. Consecutive whitespace collapses to a single space.
        Private Shared Function BuildNormalized(text As String,
                                                ignoreCase As Boolean,
                                                Optional ByRef map As List(Of Integer) = Nothing) As String
            Dim sb As New StringBuilder()
            Dim wantMap As Boolean = (map IsNot Nothing)
            Dim prevWasSpace As Boolean = False

            If text Is Nothing Then Return String.Empty

            For i As Integer = 0 To text.Length - 1
                Dim mapped As Char = MapCharForMatch(text(i))

                If mapped = " "c Then
                    If prevWasSpace Then Continue For
                    prevWasSpace = True
                Else
                    prevWasSpace = False
                    If ignoreCase Then mapped = Char.ToLowerInvariant(mapped)
                End If

                sb.Append(mapped)
                If wantMap Then map.Add(i)
            Next

            Return sb.ToString()
        End Function

        ' Returns (origStart, origLength) pairs of every match of 'needle' inside 'haystack'.
        Private Shared Function FindAllInText(haystack As String,
                                              needle As String,
                                              ignoreCase As Boolean,
                                              maxHits As Integer) As List(Of Integer())
            Dim results As New List(Of Integer())()
            If String.IsNullOrEmpty(haystack) OrElse String.IsNullOrEmpty(needle) Then Return results

            Dim map As New List(Of Integer)()
            Dim normHay As String = BuildNormalized(haystack, ignoreCase, map)
            Dim normNeedle As String = BuildNormalized(needle, ignoreCase).Trim()
            If normNeedle.Length = 0 Then Return results

            Dim pos As Integer = 0
            While pos <= normHay.Length - normNeedle.Length
                Dim f As Integer = normHay.IndexOf(normNeedle, pos, StringComparison.Ordinal)
                If f < 0 Then Exit While

                Dim origStart As Integer = map(f)
                Dim origEnd As Integer = map(f + normNeedle.Length - 1)
                results.Add(New Integer() {origStart, origEnd - origStart + 1})

                If maxHits > 0 AndAlso results.Count >= maxHits Then Exit While
                pos = f + normNeedle.Length
            End While

            Return results
        End Function

        ' First match; tries case-sensitive first, then case-insensitive fallback.
        Private Shared Function TryFindInText(haystack As String,
                                              needle As String,
                                              ByRef matchStart As Integer,
                                              ByRef matchLength As Integer) As Boolean
            matchStart = -1
            matchLength = 0

            Dim hit As List(Of Integer()) = FindAllInText(haystack, needle, False, 1)
            If hit.Count = 0 Then
                hit = FindAllInText(haystack, needle, True, 1)
            End If

            If hit.Count = 0 Then Return False

            matchStart = hit(0)(0)
            matchLength = hit(0)(1)
            Return True
        End Function

        Private Shared Function MakeRun(text As String, Optional deletedText As Boolean = False) As W.Run
            Dim r As New W.Run()

            If deletedText Then
                Dim dt As New W.DeletedText(text) With {.Space = SpaceProcessingModeValues.Preserve}
                r.AppendChild(dt)
            Else
                r.AppendChild(New W.Text(text) With {.Space = SpaceProcessingModeValues.Preserve})
            End If

            Return r
        End Function

        Private Shared Function SplitLines(text As String) As String()
            If text Is Nothing Then Return New String() {}
            Return text.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Split(CChar(vbLf))
        End Function

        Private Shared Sub AppendPlainRun(para As W.Paragraph, text As String)
            If Not String.IsNullOrEmpty(text) Then
                para.AppendChild(MakeRun(text))
            End If
        End Sub

        Private Shared Function MakeFormattedRun(text As String, bold As Boolean, italic As Boolean) As W.Run
            Dim r As New W.Run()

            If bold OrElse italic Then
                Dim rPr As New W.RunProperties()
                If bold Then rPr.AppendChild(New W.Bold())
                If italic Then rPr.AppendChild(New W.Italic())
                r.AppendChild(rPr)
            End If

            r.AppendChild(New W.Text(text) With {.Space = SpaceProcessingModeValues.Preserve})
            Return r
        End Function

        ' Splits a single line into runs, honouring **bold**/__bold__ and *italic*/_italic_.
        Private Shared Function ParseInlineRuns(text As String) As List(Of W.Run)
            Dim runs As New List(Of W.Run)()

            If String.IsNullOrEmpty(text) Then
                runs.Add(MakeFormattedRun(String.Empty, False, False))
                Return runs
            End If

            Dim rx As New Regex("(\*\*|__)(.+?)\1|(\*|_)(.+?)\3")
            Dim pos As Integer = 0

            For Each m As Match In rx.Matches(text)
                If m.Index > pos Then
                    runs.Add(MakeFormattedRun(text.Substring(pos, m.Index - pos), False, False))
                End If

                If m.Groups(2).Success Then
                    runs.Add(MakeFormattedRun(m.Groups(2).Value, True, False))
                Else
                    runs.Add(MakeFormattedRun(m.Groups(4).Value, False, True))
                End If

                pos = m.Index + m.Length
            Next

            If pos < text.Length Then
                runs.Add(MakeFormattedRun(text.Substring(pos), False, False))
            End If

            If runs.Count = 0 Then
                runs.Add(MakeFormattedRun(text, False, False))
            End If

            Return runs
        End Function

        ' Detects a block-level Markdown marker and returns the remaining content.
        Private Shared Function DetectBlock(line As String, ByRef styleId As String) As String
            styleId = Nothing
            If line Is Nothing Then Return String.Empty

            Dim heading As Match = Regex.Match(line, "^(#{1,6})\s+(.*)$")
            If heading.Success Then
                styleId = "Heading" & heading.Groups(1).Value.Length.ToString()
                Return heading.Groups(2).Value
            End If

            Dim bullet As Match = Regex.Match(line, "^\s*[-*+]\s+(.*)$")
            If bullet.Success Then
                styleId = "ListParagraph"
                Return bullet.Groups(1).Value
            End If

            Dim numbered As Match = Regex.Match(line, "^\s*\d+[.)]\s+(.*)$")
            If numbered.Success Then
                styleId = "ListParagraph"
                Return numbered.Groups(1).Value
            End If

            Return line
        End Function

        Private Shared Sub AppendInlineRunsWrapped(para As W.Paragraph,
                                                   text As String,
                                                   asMarkup As Boolean,
                                                   author As String,
                                                   body As W.Body)
            Dim runs As List(Of W.Run) = ParseInlineRuns(text)

            If asMarkup Then
                Dim ins As New W.InsertedRun() With {
                    .Id = NextChangeId(body).ToString(),
                    .Author = author,
                    .Date = DateTime.UtcNow
                }
                For Each r As W.Run In runs
                    ins.AppendChild(r)
                Next
                para.AppendChild(ins)
            Else
                For Each r As W.Run In runs
                    para.AppendChild(r)
                Next
            End If
        End Sub

        Private Shared Function MakeMarkdownParagraph(line As String,
                                                      body As W.Body,
                                                      asMarkup As Boolean,
                                                      author As String) As W.Paragraph
            Dim para As New W.Paragraph()
            Dim styleId As String = Nothing
            Dim content As String = DetectBlock(line, styleId)

            If styleId IsNot Nothing Then
                Dim pPr As New W.ParagraphProperties()
                pPr.AppendChild(New W.ParagraphStyleId() With {.Val = styleId})
                para.AppendChild(pPr)
            End If

            AppendInlineRunsWrapped(para, content, asMarkup, author, body)
            Return para
        End Function

        ' Appends Markdown content that may span several lines. The first line is written into
        ' 'para'; each subsequent line becomes a new sibling W.Paragraph inserted right after the
        ' previous one. Returns the W.Paragraph that received the last line (for trailing text).
        Private Shared Function AppendMarkdownContent(para As W.Paragraph,
                                                      content As String,
                                                      asMarkup As Boolean,
                                                      author As String,
                                                      body As W.Body,
                                                      allowBlockOnFirst As Boolean) As W.Paragraph
            If content Is Nothing Then Return para

            Dim lines As String() = SplitLines(content)
            If lines.Length = 0 Then Return para

            Dim current As W.Paragraph = para

            For i As Integer = 0 To lines.Length - 1
                If i = 0 Then
                    Dim styleId As String = Nothing
                    Dim text0 As String = DetectBlock(lines(0), styleId)

                    If allowBlockOnFirst AndAlso styleId IsNot Nothing Then
                        Dim pPr As W.ParagraphProperties = current.Elements(Of W.ParagraphProperties)().FirstOrDefault()
                        If pPr Is Nothing Then
                            pPr = current.PrependChild(New W.ParagraphProperties())
                        End If
                        pPr.RemoveAllChildren(Of W.ParagraphStyleId)()
                        pPr.PrependChild(New W.ParagraphStyleId() With {.Val = styleId})
                        AppendInlineRunsWrapped(current, text0, asMarkup, author, body)
                    Else
                        AppendInlineRunsWrapped(current, lines(0), asMarkup, author, body)
                    End If
                Else
                    Dim newPara As W.Paragraph = MakeMarkdownParagraph(lines(i), body, asMarkup, author)
                    current.InsertAfterSelf(newPara)
                    current = newPara
                End If
            Next

            Return current
        End Function

        Private Shared Function NextChangeId(scope As OpenXmlElement) As Integer
            Dim maxId As Integer = 0

            For Each n As W.InsertedRun In scope.Descendants(Of W.InsertedRun)()
                Dim v As Integer
                If n.Id IsNot Nothing AndAlso Integer.TryParse(n.Id.Value, v) AndAlso v > maxId Then
                    maxId = v
                End If
            Next

            For Each n As W.DeletedRun In scope.Descendants(Of W.DeletedRun)()
                Dim v As Integer
                If n.Id IsNot Nothing AndAlso Integer.TryParse(n.Id.Value, v) AndAlso v > maxId Then
                    maxId = v
                End If
            Next

            Return maxId + 1
        End Function

        Private Shared Function NextCommentId(comments As W.Comments) As Integer
            Dim maxId As Integer = 0

            For Each c As W.Comment In comments.Elements(Of W.Comment)()
                Dim v As Integer
                If c.Id IsNot Nothing AndAlso Integer.TryParse(c.Id.Value, v) AndAlso v > maxId Then
                    maxId = v
                End If
            Next

            Return maxId + 1
        End Function

        ' Content-preserving comment anchoring: wraps the matched span in CommentRangeStart/End
        ' plus a reference run, WITHOUT rebuilding the paragraph from plain text. Footnotes,
        ' fields, images, existing tracked changes and run formatting are retained.
        Private Shared Function AttachCommentSurgical(para As W.Paragraph, find As String, commentId As Integer, body As W.Body) As Boolean
            Dim atoms As New List(Of ParaAtom)()
            FlattenParagraph(para, atoms, False, False)

            Dim firstAtom As Integer
            Dim lastAtom As Integer
            If Not LocateFindInAtoms(atoms, find, firstAtom, lastAtom) Then Return False

            Dim em As New MarkupEmitter(body, "Inky")

            EmitOriginalRange(em, atoms, 0, firstAtom - 1)
            em.Flush()
            em.Children.Add(New W.CommentRangeStart() With {.Id = commentId.ToString()})

            EmitOriginalRange(em, atoms, firstAtom, lastAtom)
            em.Flush()
            em.Children.Add(New W.CommentRangeEnd() With {.Id = commentId.ToString()})

            Dim refRun As New W.Run()
            refRun.AppendChild(New W.CommentReference() With {.Id = commentId.ToString()})
            em.Children.Add(refRun)

            EmitOriginalRange(em, atoms, lastAtom + 1, atoms.Count - 1)
            em.Flush()

            Dim pPr As W.ParagraphProperties = para.Elements(Of W.ParagraphProperties)().FirstOrDefault()
            para.RemoveAllChildren()
            If pPr IsNot Nothing Then para.AppendChild(CType(pPr.CloneNode(True), W.ParagraphProperties))
            For Each c As OpenXmlElement In em.Children
                para.AppendChild(c)
            Next

            Return True
        End Function

        Private Shared Function AlignmentFromString(s As String) As W.JustificationValues
            Select Case s
                Case "center"
                    Return W.JustificationValues.Center
                Case "right"
                    Return W.JustificationValues.Right
                Case "justify"
                    Return W.JustificationValues.Both
                Case Else
                    Return W.JustificationValues.Left
            End Select
        End Function

        Private Shared Sub SetBool(Of T As {OpenXmlElement, New})(rPr As W.RunProperties, value As Boolean)
            Dim existing As T = rPr.Elements(Of T)().FirstOrDefault()

            If value Then
                If existing Is Nothing Then
                    rPr.AppendChild(New T())
                End If
            ElseIf existing IsNot Nothing Then
                existing.Remove()
            End If
        End Sub

        Private Shared Function BuildHit(paragraphIndex As Integer, story As String, paraText As String, index As Integer, length As Integer, match As String) As Object
            Dim winStart As Integer = System.Math.Max(0, index - 40)
            Dim winEnd As Integer = System.Math.Min(paraText.Length, index + length + 40)
            Dim ctx As String = paraText.Substring(winStart, winEnd - winStart)

            Return New With {
                Key .paragraph_index = paragraphIndex,
                Key .story = story,
                Key .index_in_paragraph = index,
                Key .length = length,
                Key .match = match,
                Key .context = ctx
            }
        End Function

        ' Reads a file's bytes with shared read/write access under a bounded wait so a document
        ' that is open/locked in Word, on a slow/unavailable network share, or an un-hydrated
        ' cloud placeholder fails fast with a clear message instead of blocking for minutes.
        Private Shared Function TryReadAllBytesFastFail(path As String,
                                                        ByRef bytes As Byte(),
                                                        ByRef errorMessage As String) As Boolean
            bytes = Nothing
            errorMessage = Nothing

            Try
                Dim attrs As System.IO.FileAttributes = File.GetAttributes(path)
                If (attrs And System.IO.FileAttributes.Offline) <> 0 Then
                    errorMessage = "The document is not available locally (offline/cloud placeholder). Open or download it first, then retry."
                    Return False
                End If
            Catch
            End Try

            ' Copy ByRef parameter to a local before using it inside the lambda.
            Dim localPath As String = path
            Dim readTask As System.Threading.Tasks.Task(Of Byte()) =
                System.Threading.Tasks.Task.Run(
                    Function() As Byte()
                        Using fs As New FileStream(localPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                            Using buffer As New MemoryStream()
                                fs.CopyTo(buffer)
                                Return buffer.ToArray()
                            End Using
                        End Using
                    End Function)

            If Not readTask.Wait(TimeSpan.FromSeconds(15)) Then
                errorMessage = "Reading the document timed out (it may be locked by another application or stored on a slow/unavailable location)."
                Return False
            End If

            If readTask.IsFaulted Then
                Dim ex As System.Exception = readTask.Exception
                If ex IsNot Nothing AndAlso ex.InnerException IsNot Nothing Then ex = ex.InnerException
                errorMessage = "Could not read the document: " & If(ex Is Nothing, "unknown error", ex.Message)
                Return False
            End If

            bytes = readTask.Result
            Return True
        End Function

        Private Shared Function Err_(code As String, message As String) As String
            Return JsonConvert.SerializeObject(New With {
                Key .error = code,
                Key .message = message
            })
        End Function

        ' --------------------------------------------------------------- argument helpers

        Private Shared Function GetStr(args As IDictionary(Of String, Object), name As String) As String
            If args Is Nothing Then Return ""

            Dim v As Object = Nothing
            If Not args.TryGetValue(name, v) OrElse v Is Nothing Then Return ""

            Return System.Convert.ToString(v)
        End Function

        Private Shared Function GetInt(args As IDictionary(Of String, Object), name As String, defaultValue As Integer) As Integer
            If args Is Nothing Then Return defaultValue

            Dim v As Object = Nothing
            If Not args.TryGetValue(name, v) OrElse v Is Nothing Then Return defaultValue

            Try
                Return System.Convert.ToInt32(v)
            Catch
                Dim n As Integer
                If Integer.TryParse(System.Convert.ToString(v), n) Then Return n
                Return defaultValue
            End Try
        End Function

        Private Shared Function GetBool(args As IDictionary(Of String, Object), name As String, defaultValue As Boolean) As Boolean
            Dim nb As Boolean? = GetNullableBool(args, name)
            If nb.HasValue Then Return nb.Value
            Return defaultValue
        End Function

        Private Shared Function GetNullableBool(args As IDictionary(Of String, Object), name As String) As Boolean?
            If args Is Nothing Then Return Nothing

            Dim v As Object = Nothing
            If Not args.TryGetValue(name, v) OrElse v Is Nothing Then Return Nothing

            Try
                Return System.Convert.ToBoolean(v)
            Catch
                Select Case System.Convert.ToString(v).Trim().ToLowerInvariant()
                    Case "true", "1", "yes"
                        Return True
                    Case "false", "0", "no"
                        Return False
                    Case Else
                        Return Nothing
                End Select
            End Try
        End Function

        ' --------------------------------------------------------------- factories

        Private Shared Function BuildExtract() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolExtract,
                .Tool = True,
                .ToolPriority = 880,
                .ToolErrorHandling = "skip",
                .ModelDescription = "Word (extract text)",
                .ToolDefinition = "{""name"":""" & ToolExtract & """,""description"":""Extract plain text from a .docx file, including the main body, footnotes, endnotes, headers/footers, tables, numbering, and text-box/margin text when present."",""parameters"":{""type"":""object"",""properties"":{""path"":{""type"":""string""},""max_chars"":{""type"":""integer"",""description"":""Optional cap on returned text length.""}},""required"":[""path""]}}",
                .ToolInstructionsPrompt = ToolExtract & ": Extract plain text from a .docx file, including footnotes and endnotes."
            }
        End Function

        Private Shared Function BuildSearch() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolSearch,
                .Tool = True,
                .ToolPriority = 881,
                .ToolErrorHandling = "skip",
                .ModelDescription = "Word (search)",
                .ToolDefinition = "{""name"":""" & ToolSearch & """,""description"":""Search a .docx for a substring or regex across the main body, footnotes, and endnotes. Returns W.Paragraph index, story (body|footnote|endnote), and a small context window per hit."",""parameters"":{""type"":""object"",""properties"":{""path"":{""type"":""string""},""query"":{""type"":""string""},""regex"":{""type"":""boolean""},""ignore_case"":{""type"":""boolean""},""max_hits"":{""type"":""integer""}},""required"":[""path"",""query""]}}",
                .ToolInstructionsPrompt = ToolSearch & ": Find text inside a .docx file, including footnotes and endnotes."
            }
        End Function

        Private Shared Function BuildWrite() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolWrite,
                .Tool = True,
                .ToolPriority = 882,
                .ToolErrorHandling = "skip",
                .ModelDescription = "Word (write, no markup)",
                .CapabilityTags = "docx_edit",
                .ToolDefinition =
                    "{""name"":""" & ToolWrite & """," &
                    """description"":""Modify text in a .docx WITHOUT tracked changes, in the main body AND in footnotes and endnotes, preserving fields, images, comments and run formatting. Ops: replace | insert_before | insert_after | append (append targets the main body) | delete_paragraph. 'find' must not span a paragraph break; to merge two paragraphs, replace the first and delete_paragraph the second. Pass multiple edits at once via 'tasks'.""," &
                    """parameters"":{""type"":""object"",""properties"":{" &
                    """path"":{""type"":""string""}," &
                    """operation_id"":{""type"":""string"",""description"":""Stable caller-supplied logical operation id. Reuse exactly for retries of the same logical operation.""}," &
                    """op"":{""type"":""string"",""enum"":[""replace"",""insert_before"",""insert_after"",""append"",""delete_paragraph""]}," &
                    """find"":{""type"":""string""}," &
                    """text"":{""type"":""string""}," &
                    """only_first"":{""type"":""boolean"",""description"":""Default true.""}," &
                    """tasks"":{""type"":""array"",""description"":""Batch of edits applied in order; each may match text produced by earlier tasks."",""items"":{""type"":""object"",""properties"":{" &
                    """operation_id"":{""type"":""string"",""description"":""Stable caller-supplied logical operation id for this task.""}," &
                    """op"":{""type"":""string"",""enum"":[""replace"",""insert_before"",""insert_after"",""append"",""delete_paragraph""]}," &
                    """find"":{""type"":""string""}," &
                    """text"":{""type"":""string""}," &
                    """only_first"":{""type"":""boolean""}},""required"":[""operation_id""]}}}," &
                    """required"":[""path""]}}",
                .ToolInstructionsPrompt =
                    ToolWrite & ": Edit a .docx without revision marks. Same behavior as word_markup but without tracked changes. " &
                    "Batch related edits in one call via 'tasks'. Every logical operation MUST have an explicit opaque operation_id; each batched task needs its own operation_id. Preserve it unchanged, including retries. " &
                    "The result 'status' may be complete, partial, or none: partial/none is NOT a block or failure. " &
                    "When status is partial/none, re-read the document to get the CURRENT text and retry ONLY the failed tasks[].find values; then report completion normally. " &
                    "Prefer the Outlook and Autopilot tools (like process_word_document) when they can accomplish the task; only fall back to word_* tools when those tools are not suitable, or when a skill or the user explicitly asks to use word_* tools."
            }
        End Function

        Private Shared Function BuildMarkup() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolMarkup,
                .Tool = True,
                .ToolPriority = 883,
                .ToolErrorHandling = "skip",
                .ModelDescription = "Word (markup / tracked changes)",
                .CapabilityTags = "docx_edit",
                .ToolDefinition =
                    "{""name"":""" & ToolMarkup & """," &
                    """description"":""Modify text in a .docx using tracked changes (Word revision marks), in the main body AND in footnotes and endnotes, preserving fields, images, comments, existing tracked changes and run formatting. Only inserted/deleted words are marked. Pass multiple edits at once via 'tasks'.""," &
                    """parameters"":{""type"":""object"",""properties"":{" &
                    """path"":{""type"":""string""}," &
                    """operation_id"":{""type"":""string"",""description"":""Stable caller-supplied logical operation id. Reuse exactly for retries of the same logical operation.""}," &
                    """op"":{""type"":""string"",""enum"":[""replace"",""insert_before"",""insert_after"",""append"",""delete_paragraph""]}," &
                    """find"":{""type"":""string""}," &
                    """text"":{""type"":""string""}," &
                    """author"":{""type"":""string""}," &
                    """only_first"":{""type"":""boolean""}," &
                    """tasks"":{""type"":""array"",""description"":""Batch of edits applied in order; each may match text produced by earlier tasks."",""items"":{""type"":""object"",""properties"":{" &
                    """operation_id"":{""type"":""string"",""description"":""Stable caller-supplied logical operation id for this task.""}," &
                    """op"":{""type"":""string"",""enum"":[""replace"",""insert_before"",""insert_after"",""append"",""delete_paragraph""]}," &
                    """find"":{""type"":""string""}," &
                    """text"":{""type"":""string""}," &
                    """only_first"":{""type"":""boolean""}},""required"":[""operation_id""]}}}," &
                    """required"":[""path""]}}",
                .ToolInstructionsPrompt =
                    ToolMarkup & ": Edit a .docx with revision marks (tracked changes). " &
                    "Batch related edits in one call via 'tasks'. Every logical operation MUST have an explicit opaque operation_id; each batched task needs its own operation_id. Preserve it unchanged, including retries. " &
                    "The result 'status' may be complete, partial, or none: partial/none is NOT a block or failure. " &
                    "When status is partial/none, re-read the document to get the CURRENT text and retry ONLY the failed tasks[].find values; then report completion normally. " &
                    "Prefer the Outlook and Autopilot tools when they can accomplish the task; only fall back to word_* tools when those tools are not suitable, or when a skill or the user explicitly asks to use word_* tools."
            }
        End Function

        Private Shared Function BuildCommentAdd() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolCommentAdd,
                .Tool = True,
                .ToolPriority = 884,
                .ToolErrorHandling = "skip",
                .ModelDescription = "Word (comment add)",
                .CapabilityTags = "docx_edit",
                .ToolDefinition =
                    "{""name"":""" & ToolCommentAdd & """," &
                    """description"":""Attach Word comment(s) to matched text, preserving footnotes, fields, images and formatting. 'find' uses format/whitespace-insensitive matching within a single paragraph. Add many comments at once via 'tasks'. Result reports status with per-task applied flag and suggestions.""," &
                    """parameters"":{""type"":""object"",""properties"":{" &
                    """path"":{""type"":""string""}," &
                    """operation_id"":{""type"":""string"",""description"":""Stable caller-supplied logical operation id. Reuse exactly for retries of the same logical operation.""}," &
                    """find"":{""type"":""string""}," &
                    """text"":{""type"":""string""}," &
                    """author"":{""type"":""string""}," &
                    """initials"":{""type"":""string""}," &
                    """tasks"":{""type"":""array"",""items"":{""type"":""object"",""properties"":{" &
                    """operation_id"":{""type"":""string"",""description"":""Stable caller-supplied logical operation id for this task.""}," &
                    """find"":{""type"":""string""}," &
                    """text"":{""type"":""string""}," &
                    """author"":{""type"":""string""}," &
                    """initials"":{""type"":""string""}},""required"":[""operation_id""]}}}," &
                    """required"":[""path""]}}",
                .ToolInstructionsPrompt =
                    ToolCommentAdd & ": Add Word bubble comment(s) to matched span(s). " &
                    "Add many at once via 'tasks'. Every logical comment operation MUST have an explicit opaque operation_id; each batched task needs its own operation_id. Preserve it unchanged, including retries. " &
                    "The result 'status' may be complete, partial, or none: partial/none is NOT a block or failure. " &
                    "When status is partial/none, re-read the document for the CURRENT text and retry ONLY the failed tasks[].find values; then report completion normally. " &
                    "Prefer the Outlook and Autopilot tools when they can accomplish the task; only fall back to word_* tools when those tools are not suitable, or when a skill or the user explicitly asks to use word_* tools."
            }
        End Function

        Private Shared Function BuildCommentList() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolCommentList,
                .Tool = True,
                .ToolPriority = 885,
                .ToolErrorHandling = "skip",
                .ModelDescription = "Word (comments list)",
                .ToolDefinition = "{""name"":""" & ToolCommentList & """,""description"":""List all comments in a .docx with id, author, initials, date and inner text."",""parameters"":{""type"":""object"",""properties"":{""path"":{""type"":""string""}},""required"":[""path""]}}",
                .ToolInstructionsPrompt = ToolCommentList & ": List Word comments in a .docx."
            }
        End Function

        Private Shared Function BuildCommentRemove() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolCommentRemove,
                .Tool = True,
                .ToolPriority = 886,
                .ToolErrorHandling = "skip",
                .ModelDescription = "Word (comment remove)",
                .ToolDefinition = "{""name"":""" & ToolCommentRemove & """,""description"":""Remove a Word comment by id (also strips its range markers and reference)."",""parameters"":{""type"":""object"",""properties"":{""path"":{""type"":""string""},""id"":{""type"":""string""}},""required"":[""path"",""id""]}}",
                .ToolInstructionsPrompt = ToolCommentRemove & ": Remove a Word comment by id."
            }
        End Function

        Private Shared Function BuildFormat() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolFormat,
                .Tool = True,
                .ToolPriority = 887,
                .ToolErrorHandling = "skip",
                .ModelDescription = "Word (format)",
                                .ToolDefinition = "{""name"":""" & ToolFormat & """,""description"":""Apply W.Paragraph style and/or W.Run formatting to every W.Paragraph containing 'find', in the main body AND in footnotes and endnotes. Available: style (Word style id, e.g. 'Heading1'), bold, italic, underline, size (pt), color (RRGGBB), align (left|center|right|justify)."",""parameters"":{""type"":""object"",""properties"":{""path"":{""type"":""string""},""find"":{""type"":""string""},""style"":{""type"":""string""},""bold"":{""type"":""boolean""},""italic"":{""type"":""boolean""},""underline"":{""type"":""boolean""},""size"":{""type"":""integer""},""color"":{""type"":""string""},""align"":{""type"":""string"",""enum"":[""left"",""center"",""right"",""justify""]}},""required"":[""path"",""find""]}}",
                .ToolInstructionsPrompt = ToolFormat & ": Apply W.Paragraph/W.Run formatting (style, bold/italic/underline, size, color, alignment)."
            }
        End Function

        Private Shared Function BuildApplyTemplate() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolApplyTemplate,
                .Tool = True,
                .ToolPriority = 888,
                .ToolErrorHandling = "skip",
                .ModelDescription = "Word (apply template)",
                .CapabilityTags = "artifact_generation",
                .ToolDefinition = "{""name"":""" & ToolApplyTemplate & """,""description"":""Clone a .docx template from a skill's references/ directory to a new file in the default writable root (the connected workspace, otherwise the current session's staging/working area) and substitute {{placeholders}} from the 'substitutions' object. Optional explicit artifact fields can register the resulting single file without changing legacy calls."",""parameters"":{""type"":""object"",""properties"":{""skill"":{""type"":""string"",""description"":""Skill name.""},""template"":{""type"":""string"",""description"":""Path relative to the skill's references/ directory.""},""output_name"":{""type"":""string"",""description"":""Suggested output filename (default 'from_template.docx').""},""substitutions"":{""type"":""object"",""description"":""Object of {placeholderName: value}; each key K becomes the literal '{{K}}' in the template.""},""artifact_id"":{""type"":""string""},""logical_deliverable_id"":{""type"":""string""},""output_slot_id"":{""type"":""string""},""supersedes_artifact_id"":{""type"":""string""},""artifact_state"":{""type"":""string"",""enum"":[""working"",""intermediate"",""final""]},""artifact_delivery_intent"":{""type"":""string"",""enum"":[""none"",""deliver_to_user"",""persist_only"",""deliver_and_persist""]},""storage_kind"":{""type"":""string"",""enum"":[""session_staging"",""connected_workspace"",""host_managed"",""unknown""]},""expected_artifacts"":{""type"":""array"",""items"":{""type"":""object"",""properties"":{""logical_deliverable_id"":{""type"":""string""},""output_slot_id"":{""type"":""string""}},""required"":[""logical_deliverable_id"",""output_slot_id""]}}},""required"":[""skill"",""template""]}}",
                .ToolInstructionsPrompt = ToolApplyTemplate & ": Instantiate a Word template from a skill's references/ directory. For an explicit artifact, preserve the caller-supplied opaque IDs and declare artifact_state/artifact_delivery_intent; a user-facing Final also requires the complete expected_artifacts set."
            }
        End Function

        Private Shared Function BuildSaveAs() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolSaveAs,
                .Tool = True,
                .ToolPriority = 889,
                .ToolErrorHandling = "skip",
                .ModelDescription = "Word (save as)",
                .ToolDefinition =
                    "{""name"":""" & ToolSaveAs & """," &
                    """description"":""Copy a .docx to a new path inside the default writable root. artifact_id, logical_deliverable_id, output_slot_id, and expected_artifacts are required explicit orchestration fields. supersedes_artifact_id and storage_kind are optional explicit metadata; no artifact identity is inferred from filenames or paths.""," &
                    """parameters"":{""type"":""object"",""properties"":{" &
                    """source"":{""type"":""string""}," &
                    """output_name"":{""type"":""string""}," &
                    """artifact_id"":{""type"":""string"",""description"":""Opaque caller-supplied id for this physical artifact.""}," &
                    """logical_deliverable_id"":{""type"":""string"",""description"":""Opaque caller-supplied logical deliverable id. No host inference is performed.""}," &
                    """output_slot_id"":{""type"":""string"",""description"":""Opaque caller-supplied output slot id within the logical deliverable.""}," &
                    """supersedes_artifact_id"":{""type"":""string"",""description"":""Optional exact artifact id superseded by this output.""}," &
                    """storage_kind"":{""type"":""string"",""enum"":[""session_staging"",""connected_workspace"",""host_managed"",""unknown""]}," &
                    """expected_artifacts"":{""type"":""array"",""description"":""Required complete explicit list of final output slots expected for this task. Completion is blocked until every listed slot has a current final artifact."",""items"":{""type"":""object"",""properties"":{" &
                    """logical_deliverable_id"":{""type"":""string""}," &
                    """output_slot_id"":{""type"":""string""}},""required"":[""logical_deliverable_id"",""output_slot_id""]}}}," &
                    """required"":[""source"",""artifact_id"",""logical_deliverable_id"",""output_slot_id"",""expected_artifacts""]}}",
                .ToolInstructionsPrompt =
                    ToolSaveAs & ": Copy a .docx to a new path in the writable root. " &
                    "artifact_id, logical_deliverable_id, output_slot_id, and expected_artifacts are required. expected_artifacts MUST declare the complete expected final output-slot set for this task, including the current output slot. " &
                    "Preserve artifact_id, logical_deliverable_id, output_slot_id, supersedes_artifact_id, storage_kind, and expected_artifacts exactly. Never derive or change artifact identity from the output filename."
            }
        End Function

    End Class

End Namespace
