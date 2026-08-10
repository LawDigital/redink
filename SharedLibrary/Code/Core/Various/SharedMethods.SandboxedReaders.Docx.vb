' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SharedMethods.SandboxedReaders.vb
' Purpose: Provides sandboxed, dependency-free text extraction for DOCX files
'          by parsing the underlying OpenXML package. This approach avoids
'          COM interop, enabling safe text extraction without requiring
'          Microsoft Word to be installed.
'
' Architecture:
'  1. Unpack: The .docx file is treated as a ZIP archive and its contents are
'     extracted to a temporary directory.
'  2. Parse Core Content: The primary content is read from `word/document.xml`.
'     An `XmlDocument` is used to traverse the XML structure.
'  3. Namespace Handling: A `XmlNamespaceManager` is configured with the "w"
'     prefix for the main WordprocessingML schema to enable reliable XPath queries.
'  4. Text Extraction:
'     - Paragraphs (`<w:p>`): Text is extracted from text runs (`<w:r>`) and
'       their child text nodes (`<w:t>`). Special elements like tabs and breaks
'       are handled.
'     - Tables (`<w:tbl>`): Tables are recursively processed, with rows and cells
'       identified. Metadata such as column spans and vertical merges are
'       included in the output for better contextual understanding by LLMs.
'  5. Auxiliary Content (Optional): If enabled, text from headers, footers,
'     footnotes, and endnotes is extracted from their respective XML parts
'     (e.g., `header1.xml`, `footnotes.xml`).
'  6. Cleanup: The temporary directory is deleted in a `Finally` block to
'     ensure no artifacts are left behind.
'
' Output Format:
'  - The function returns a single string. Paragraphs are separated by newlines.
'  - Tables are formatted with descriptive labels like "[Table 1]", "Row 1, Cell 1",
'    etc., to preserve structure in the plain text output.
' =============================================================================

Option Explicit On
Option Strict On
Option Infer On

Namespace SharedLibrary
    Partial Public Class SharedMethods


        ''' <summary>
        ''' Dependency-free DOCX-to-text / Markdown extractor.
        '''
        ''' Features:
        ''' - Word automatic paragraph/list/heading numbering from numbering.xml and styles.xml.
        ''' - Upper/lower Roman, upper/lower letters, decimal, decimalZero and bullets.
        ''' - Stateful multilevel numbering with restarts and start overrides.
        ''' - Forward and backward bookmark/REF cross-reference resolution.
        ''' - Complex fields and simple fields, preserving cached results for unsupported fields.
        ''' - Tables and nested tables.
        ''' - Headers, footers, footnotes and endnotes.
        ''' - Text contained in Word/VML text boxes, emitted as margin text.
        ''' - Section line-numbering settings (exact rendered line numbers require a layout engine).
        ''' </summary>
        Public NotInheritable Class DocxTextExtractor

            Private Const SB_WordNs As System.String = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

            ''' <summary>
            ''' Includes headers, footers, footnotes and endnotes.
            ''' Kept under the same name used by the original implementation.
            ''' </summary>
            Public Shared Property DocxIncludeHeaderFooterFootnotes As System.Boolean = True

            ''' <summary>
            ''' Includes text found in w:txbxContent, which commonly contains margin numbers or margin notes.
            ''' </summary>
            Public Shared Property DocxIncludeMarginText As System.Boolean = True

            ''' <summary>
            ''' Emits the section line-numbering configuration. Exact visual line numbers cannot be
            ''' reconstructed reliably without Word-compatible pagination and layout.
            ''' </summary>
            Public Shared Property DocxIncludeLineNumberSettings As System.Boolean = True

            ''' <summary>
            ''' Adds stable [H0001], [H0002], ... anchors before heading paragraphs.
            ''' References still resolve correctly when this is False.
            ''' </summary>
            Public Shared Property DocxIncludeReferenceAnchors As System.Boolean = False

            ''' <summary>
            ''' Uses a field's cached Word result when the field type or target cannot be resolved.
            ''' </summary>
            Public Shared Property DocxUseCachedFieldResultWhenUnresolved As System.Boolean = True

            Private Sub New()
            End Sub

            ''' <summary>
            ''' Converts a DOCX file directly to a UTF-8 text file without a byte-order mark.
            ''' Returns the output path on success or an Error: message on failure.
            ''' </summary>
            Public Shared Function WriteDocxTextFile(
                docxPath As System.String,
                textFilePath As System.String,
                Optional returnMarkdown As System.Boolean = False
            ) As System.String
                If System.String.IsNullOrWhiteSpace(textFilePath) Then
                    Return "Error: Output text-file path is empty."
                End If

                Try
                    Dim extractedText As System.String = ReadDocxSandboxed(docxPath, returnMarkdown)
                    If extractedText.StartsWith("Error:", System.StringComparison.Ordinal) Then
                        Return extractedText
                    End If

                    Dim outputDirectory As System.String = System.IO.Path.GetDirectoryName(textFilePath)
                    If Not System.String.IsNullOrWhiteSpace(outputDirectory) AndAlso Not System.IO.Directory.Exists(outputDirectory) Then
                        System.IO.Directory.CreateDirectory(outputDirectory)
                    End If

                    System.IO.File.WriteAllText(
                        textFilePath,
                        extractedText,
                        New System.Text.UTF8Encoding(False)
                    )

                    Return textFilePath
                Catch ex As System.Exception
                    Return "Error writing text file: " & ex.Message
                End Try
            End Function

            ''' <summary>
            ''' Reads a DOCX file and returns either the established plain-text representation
            ''' or, when returnMarkdown is True, a Markdown representation built from the same
            ''' sandboxed OpenXML analysis model. The default False value preserves all existing callers.
            ''' </summary>
            Public Shared Function ReadDocxSandboxed(
                docxPath As System.String,
                Optional returnMarkdown As System.Boolean = False
            ) As System.String
                If System.String.IsNullOrWhiteSpace(docxPath) OrElse Not System.IO.File.Exists(docxPath) Then
                    Return "Error: File not found."
                End If

                Dim tempDirectory As System.String =
                    System.IO.Path.Combine(
                        System.IO.Path.GetTempPath(),
                        "ri_docx_" & System.Guid.NewGuid().ToString("N")
                    )

                Try
                    System.IO.Compression.ZipFile.ExtractToDirectory(docxPath, tempDirectory)

                    Dim wordDirectory As System.String = System.IO.Path.Combine(tempDirectory, "word")
                    Dim documentXmlPath As System.String = System.IO.Path.Combine(wordDirectory, "document.xml")

                    If Not System.IO.File.Exists(documentXmlPath) Then
                        Return "Error: Not a valid .docx file (missing word/document.xml)."
                    End If

                    Dim context As New ExtractionContext()
                    context.WordDirectory = wordDirectory
                    context.StyleModel = LoadStyleModel(wordDirectory)
                    context.NumberingModel = LoadNumberingModel(wordDirectory)

                    Dim mainDocument As System.Xml.XmlDocument = LoadXmlDocument(documentXmlPath)
                    Dim mainNamespaceManager As System.Xml.XmlNamespaceManager = CreateNamespaceManager(mainDocument)
                    Dim bodyNode As System.Xml.XmlNode = mainDocument.SelectSingleNode("//w:body", mainNamespaceManager)

                    If bodyNode Is Nothing Then
                        Return "Error: Not a valid .docx file (missing document body)."
                    End If

                    Dim bodyStory As New StorySection()
                    bodyStory.Label = System.String.Empty
                    bodyStory.Blocks = AnalyseBlockChildren(
                        bodyNode,
                        mainNamespaceManager,
                        context,
                        New NumberingState(),
                        "MainDocument",
                        New TableNumberState()
                    )

                    If DocxIncludeLineNumberSettings Then
                        context.LineNumberSettings.AddRange(ReadLineNumberSettings(bodyNode, mainNamespaceManager))
                    End If

                    Dim supplementaryStories As New System.Collections.Generic.List(Of StorySection)()
                    Dim noteStories As New System.Collections.Generic.List(Of NoteSection)()

                    If DocxIncludeHeaderFooterFootnotes AndAlso System.IO.Directory.Exists(wordDirectory) Then
                        supplementaryStories.AddRange(
                            AnalyseStoryFiles(wordDirectory, "header*.xml", "Header", context)
                        )
                        supplementaryStories.AddRange(
                            AnalyseStoryFiles(wordDirectory, "footer*.xml", "Footer", context)
                        )
                        noteStories.AddRange(
                            AnalyseNoteFile(wordDirectory, "footnotes.xml", "footnote", "Footnote", context)
                        )
                        noteStories.AddRange(
                            AnalyseNoteFile(wordDirectory, "endnotes.xml", "endnote", "Endnote", context)
                        )
                    End If

                    AssignHeadingAnchors(context)

                    Dim output As New System.Text.StringBuilder(8192)
                    Dim fieldState As New FieldEvaluationState()

                    If returnMarkdown Then
                        If DocxIncludeLineNumberSettings AndAlso context.LineNumberSettings.Count > 0 Then
                            output.AppendLine("## Line numbering settings")
                            output.AppendLine()
                            For Each settingText As System.String In context.LineNumberSettings
                                output.AppendLine("- " & EscapeMarkdownInline(settingText))
                            Next
                            output.AppendLine()
                        End If

                        RenderBlocksMarkdown(bodyStory.Blocks, context, fieldState, output, 0)

                        For Each story As StorySection In supplementaryStories
                            Dim storyBuilder As New System.Text.StringBuilder()
                            RenderBlocksMarkdown(story.Blocks, context, New FieldEvaluationState(), storyBuilder, 0)

                            Dim storyText As System.String = storyBuilder.ToString().Trim()
                            If storyText.Length > 0 Then
                                AppendMarkdownBlankLine(output)
                                output.AppendLine("## " & EscapeMarkdownInline(story.Label))
                                output.AppendLine()
                                output.AppendLine(storyText)
                            End If
                        Next

                        RenderNoteSectionsMarkdown(noteStories, context, output)
                    Else
                        ' IMPORTANT: the established plain-text path is deliberately unchanged.
                        If DocxIncludeLineNumberSettings AndAlso context.LineNumberSettings.Count > 0 Then
                            output.AppendLine("--- Line numbering settings ---")
                            For Each settingText As System.String In context.LineNumberSettings
                                output.AppendLine(settingText)
                            Next
                            output.AppendLine()
                        End If

                        RenderBlocks(bodyStory.Blocks, context, fieldState, output, 0)

                        For Each story As StorySection In supplementaryStories
                            Dim storyBuilder As New System.Text.StringBuilder()
                            RenderBlocks(story.Blocks, context, New FieldEvaluationState(), storyBuilder, 0)

                            Dim storyText As System.String = storyBuilder.ToString().Trim()
                            If storyText.Length > 0 Then
                                output.AppendLine()
                                output.AppendLine("--- " & story.Label & " ---")
                                output.AppendLine(storyText)
                            End If
                        Next

                        RenderNoteSections(noteStories, context, output)
                    End If

                    Dim result As System.String = output.ToString().TrimEnd()

                    If System.String.IsNullOrWhiteSpace(result) Then
                        Return "Error: No text content found in .docx."
                    End If

                    Return result

                Catch ex As System.Exception
                    Return "Error reading .docx: " & ex.Message

                Finally
                    Try
                        If System.IO.Directory.Exists(tempDirectory) Then
                            System.IO.Directory.Delete(tempDirectory, True)
                        End If
                    Catch ex As System.Exception
                        ' Cleanup failure must not replace the extraction result.
                    End Try
                End Try
            End Function

            Private Shared Function LoadXmlDocument(filePath As System.String) As System.Xml.XmlDocument
                Dim document As New System.Xml.XmlDocument()
                document.PreserveWhitespace = True
                document.Load(filePath)
                Return document
            End Function

            Private Shared Function CreateNamespaceManager(document As System.Xml.XmlDocument) As System.Xml.XmlNamespaceManager
                Dim namespaceManager As New System.Xml.XmlNamespaceManager(document.NameTable)
                namespaceManager.AddNamespace("w", SB_WordNs)
                Return namespaceManager
            End Function

#Region "Analysis models"

            Private Enum DocumentBlockKind
                Paragraph
                Table
            End Enum

            Private NotInheritable Class DocumentBlock
                Public Property Kind As DocumentBlockKind
                Public Property Paragraph As ParagraphInfo
                Public Property Table As TableInfo
            End Class

            Private NotInheritable Class ParagraphInfo
                Public Sub New()
                    Me.Tokens = New System.Collections.Generic.List(Of InlineToken)()
                    Me.BookmarkNames = New System.Collections.Generic.List(Of System.String)()
                    Me.MarginParagraphs = New System.Collections.Generic.List(Of ParagraphInfo)()
                    Me.NumberText = System.String.Empty
                    Me.ListNumberFormat = System.String.Empty
                    Me.CachedPlainText = System.String.Empty
                    Me.StoryName = System.String.Empty
                    Me.AnchorText = System.String.Empty
                End Sub

                Public Property Tokens As System.Collections.Generic.List(Of InlineToken)
                Public Property BookmarkNames As System.Collections.Generic.List(Of System.String)
                Public Property MarginParagraphs As System.Collections.Generic.List(Of ParagraphInfo)
                Public Property NumberText As System.String
                Public Property ListLevel As System.Nullable(Of System.Int32)
                Public Property ListNumberFormat As System.String
                Public Property CachedPlainText As System.String
                Public Property StoryName As System.String
                Public Property SequenceIndex As System.Int32
                Public Property HeadingLevel As System.Nullable(Of System.Int32)
                Public Property AnchorText As System.String
                Public Property HasAutomaticNumber As System.Boolean
            End Class

            Private NotInheritable Class TableInfo
                Public Sub New()
                    Me.Rows = New System.Collections.Generic.List(Of TableRowInfo)()
                    Me.DisplayNumber = System.String.Empty
                End Sub

                Public Property DisplayNumber As System.String
                Public Property Rows As System.Collections.Generic.List(Of TableRowInfo)
            End Class

            Private NotInheritable Class TableRowInfo
                Public Sub New()
                    Me.Cells = New System.Collections.Generic.List(Of TableCellInfo)()
                End Sub

                Public Property Cells As System.Collections.Generic.List(Of TableCellInfo)
            End Class

            Private NotInheritable Class TableCellInfo
                Public Sub New()
                    Me.Blocks = New System.Collections.Generic.List(Of DocumentBlock)()
                    Me.GridSpan = 1
                    Me.VerticalMerge = System.String.Empty
                End Sub

                Public Property Blocks As System.Collections.Generic.List(Of DocumentBlock)
                Public Property GridSpan As System.Int32
                Public Property VerticalMerge As System.String
            End Class

            Private NotInheritable Class StorySection
                Public Sub New()
                    Me.Label = System.String.Empty
                    Me.Blocks = New System.Collections.Generic.List(Of DocumentBlock)()
                End Sub

                Public Property Label As System.String
                Public Property Blocks As System.Collections.Generic.List(Of DocumentBlock)
            End Class

            Private NotInheritable Class NoteSection
                Public Sub New()
                    Me.NoteId = System.String.Empty
                    Me.Label = System.String.Empty
                    Me.Blocks = New System.Collections.Generic.List(Of DocumentBlock)()
                End Sub

                Public Property NoteId As System.String
                Public Property Label As System.String
                Public Property Blocks As System.Collections.Generic.List(Of DocumentBlock)
            End Class

            Private NotInheritable Class ExtractionContext
                Public Sub New()
                    Me.WordDirectory = System.String.Empty
                    Me.NumberingModel = New NumberingModel()
                    Me.StyleModel = New StyleModel()
                    Me.BookmarkTargets = New System.Collections.Generic.Dictionary(Of System.String, ParagraphInfo)(System.StringComparer.OrdinalIgnoreCase)
                    Me.AllParagraphs = New System.Collections.Generic.List(Of ParagraphInfo)()
                    Me.LineNumberSettings = New System.Collections.Generic.List(Of System.String)()
                End Sub

                Public Property WordDirectory As System.String
                Public Property NumberingModel As NumberingModel
                Public Property StyleModel As StyleModel
                Public Property BookmarkTargets As System.Collections.Generic.Dictionary(Of System.String, ParagraphInfo)
                Public Property AllParagraphs As System.Collections.Generic.List(Of ParagraphInfo)
                Public Property LineNumberSettings As System.Collections.Generic.List(Of System.String)
            End Class

            Private NotInheritable Class TableNumberState
                Private _nextTopLevelTable As System.Int32

                Public Function GetNextTopLevelNumber() As System.String
                    _nextTopLevelTable += 1
                    Return _nextTopLevelTable.ToString(System.Globalization.CultureInfo.InvariantCulture)
                End Function
            End Class

#End Region

#Region "Styles and numbering models"

            Private NotInheritable Class StyleModel
                Public Sub New()
                    Me.Styles = New System.Collections.Generic.Dictionary(Of System.String, StyleDefinition)(System.StringComparer.OrdinalIgnoreCase)
                    Me.EffectiveStyles = New System.Collections.Generic.Dictionary(Of System.String, EffectiveStyleProperties)(System.StringComparer.OrdinalIgnoreCase)
                End Sub

                Public Property Styles As System.Collections.Generic.Dictionary(Of System.String, StyleDefinition)
                Public Property EffectiveStyles As System.Collections.Generic.Dictionary(Of System.String, EffectiveStyleProperties)
            End Class

            Private NotInheritable Class StyleDefinition
                Public Sub New()
                    Me.StyleId = System.String.Empty
                    Me.StyleName = System.String.Empty
                    Me.BasedOnStyleId = System.String.Empty
                End Sub

                Public Property StyleId As System.String
                Public Property StyleName As System.String
                Public Property BasedOnStyleId As System.String
                Public Property NumberingId As System.Nullable(Of System.Int32)
                Public Property NumberingLevel As System.Nullable(Of System.Int32)
                Public Property OutlineLevel As System.Nullable(Of System.Int32)
                Public Property NumberingCancelled As System.Boolean
            End Class

            Private NotInheritable Class EffectiveStyleProperties
                Public Sub New()
                    Me.StyleId = System.String.Empty
                    Me.StyleName = System.String.Empty
                End Sub

                Public Property StyleId As System.String
                Public Property StyleName As System.String
                Public Property NumberingId As System.Nullable(Of System.Int32)
                Public Property NumberingLevel As System.Nullable(Of System.Int32)
                Public Property OutlineLevel As System.Nullable(Of System.Int32)
                Public Property NumberingCancelled As System.Boolean
            End Class

            Private NotInheritable Class NumberingModel
                Public Sub New()
                    Me.AbstractNumbers = New System.Collections.Generic.Dictionary(Of System.Int32, AbstractNumberDefinition)()
                    Me.Numbers = New System.Collections.Generic.Dictionary(Of System.Int32, NumberDefinition)()
                End Sub

                Public Property AbstractNumbers As System.Collections.Generic.Dictionary(Of System.Int32, AbstractNumberDefinition)
                Public Property Numbers As System.Collections.Generic.Dictionary(Of System.Int32, NumberDefinition)
            End Class

            Private NotInheritable Class AbstractNumberDefinition
                Public Sub New()
                    Me.Levels = New System.Collections.Generic.Dictionary(Of System.Int32, NumberingLevelDefinition)()
                    Me.NumberStyleLink = System.String.Empty
                End Sub

                Public Property AbstractNumberId As System.Int32
                Public Property Levels As System.Collections.Generic.Dictionary(Of System.Int32, NumberingLevelDefinition)
                Public Property NumberStyleLink As System.String
            End Class

            Private NotInheritable Class NumberDefinition
                Public Sub New()
                    Me.LevelOverrides = New System.Collections.Generic.Dictionary(Of System.Int32, NumberingLevelOverride)()
                End Sub

                Public Property NumberId As System.Int32
                Public Property AbstractNumberId As System.Int32
                Public Property LevelOverrides As System.Collections.Generic.Dictionary(Of System.Int32, NumberingLevelOverride)
            End Class

            Private NotInheritable Class NumberingLevelOverride
                Public Property StartOverride As System.Nullable(Of System.Int32)
                Public Property LevelDefinition As NumberingLevelDefinition
            End Class

            Private NotInheritable Class NumberingLevelDefinition
                Public Sub New()
                    Me.StartValue = 1
                    Me.NumberFormat = "decimal"
                    Me.LevelText = System.String.Empty
                    Me.Suffix = "tab"
                    Me.ParagraphStyleId = System.String.Empty
                End Sub

                Public Property Level As System.Int32
                Public Property StartValue As System.Int32
                Public Property NumberFormat As System.String
                Public Property LevelText As System.String
                Public Property Suffix As System.String
                Public Property ParagraphStyleId As System.String
                Public Property RestartAfterLevel As System.Nullable(Of System.Int32)
                Public Property IsLegalNumbering As System.Boolean

                Public Function CloneDefinition() As NumberingLevelDefinition
                    Dim clone As New NumberingLevelDefinition()
                    clone.Level = Me.Level
                    clone.StartValue = Me.StartValue
                    clone.NumberFormat = Me.NumberFormat
                    clone.LevelText = Me.LevelText
                    clone.Suffix = Me.Suffix
                    clone.ParagraphStyleId = Me.ParagraphStyleId
                    clone.RestartAfterLevel = Me.RestartAfterLevel
                    clone.IsLegalNumbering = Me.IsLegalNumbering
                    Return clone
                End Function
            End Class

            Private NotInheritable Class ParagraphNumberingProperties
                Public Property NumberingId As System.Int32
                Public Property Level As System.Int32
            End Class

            Private NotInheritable Class NumberingCounterSet
                Public Sub New()
                    Me.Values = New System.Int32(8) {}
                    Me.Initialized = New System.Boolean(8) {}
                End Sub

                Public Property Values As System.Int32()
                Public Property Initialized As System.Boolean()
            End Class

            Private NotInheritable Class NumberingState
                Public Sub New()
                    Me.Counters = New System.Collections.Generic.Dictionary(Of System.Int32, NumberingCounterSet)()
                End Sub

                Public Property Counters As System.Collections.Generic.Dictionary(Of System.Int32, NumberingCounterSet)
            End Class

            Private Shared Function LoadStyleModel(wordDirectory As System.String) As StyleModel
                Dim model As New StyleModel()
                Dim stylesPath As System.String = System.IO.Path.Combine(wordDirectory, "styles.xml")

                If Not System.IO.File.Exists(stylesPath) Then
                    Return model
                End If

                Try
                    Dim stylesDocument As System.Xml.XmlDocument = LoadXmlDocument(stylesPath)
                    Dim namespaceManager As System.Xml.XmlNamespaceManager = CreateNamespaceManager(stylesDocument)
                    Dim styleNodes As System.Xml.XmlNodeList = stylesDocument.SelectNodes("//w:style", namespaceManager)

                    If styleNodes Is Nothing Then
                        Return model
                    End If

                    For Each styleNode As System.Xml.XmlNode In styleNodes
                        Dim styleId As System.String = GetWordAttributeValue(styleNode, "styleId")
                        If System.String.IsNullOrWhiteSpace(styleId) Then
                            Continue For
                        End If

                        Dim definition As New StyleDefinition()
                        definition.StyleId = styleId

                        Dim styleNameNode As System.Xml.XmlNode = styleNode.SelectSingleNode("w:name", namespaceManager)
                        If styleNameNode IsNot Nothing Then
                            definition.StyleName = GetWordAttributeValue(styleNameNode, "val")
                        End If

                        Dim basedOnNode As System.Xml.XmlNode = styleNode.SelectSingleNode("w:basedOn", namespaceManager)
                        If basedOnNode IsNot Nothing Then
                            definition.BasedOnStyleId = GetWordAttributeValue(basedOnNode, "val")
                        End If

                        Dim numIdNode As System.Xml.XmlNode = styleNode.SelectSingleNode("w:pPr/w:numPr/w:numId", namespaceManager)
                        If numIdNode IsNot Nothing Then
                            Dim numIdValue As System.Int32
                            If System.Int32.TryParse(GetWordAttributeValue(numIdNode, "val"), numIdValue) Then
                                If numIdValue = 0 Then
                                    definition.NumberingCancelled = True
                                Else
                                    definition.NumberingId = numIdValue
                                End If
                            End If
                        End If

                        Dim levelNode As System.Xml.XmlNode = styleNode.SelectSingleNode("w:pPr/w:numPr/w:ilvl", namespaceManager)
                        If levelNode IsNot Nothing Then
                            Dim levelValue As System.Int32
                            If System.Int32.TryParse(GetWordAttributeValue(levelNode, "val"), levelValue) Then
                                definition.NumberingLevel = ClampNumberingLevel(levelValue)
                            End If
                        End If

                        Dim outlineNode As System.Xml.XmlNode = styleNode.SelectSingleNode("w:pPr/w:outlineLvl", namespaceManager)
                        If outlineNode IsNot Nothing Then
                            Dim outlineValue As System.Int32
                            If System.Int32.TryParse(GetWordAttributeValue(outlineNode, "val"), outlineValue) Then
                                definition.OutlineLevel = ClampNumberingLevel(outlineValue)
                            End If
                        End If

                        model.Styles(styleId) = definition
                    Next
                Catch ex As System.Exception
                    ' Styles are optional. Extraction continues without inherited styles.
                End Try

                Return model
            End Function

            Private Shared Function LoadNumberingModel(wordDirectory As System.String) As NumberingModel
                Dim model As New NumberingModel()
                Dim numberingPath As System.String = System.IO.Path.Combine(wordDirectory, "numbering.xml")

                If Not System.IO.File.Exists(numberingPath) Then
                    Return model
                End If

                Try
                    Dim numberingDocument As System.Xml.XmlDocument = LoadXmlDocument(numberingPath)
                    Dim namespaceManager As System.Xml.XmlNamespaceManager = CreateNamespaceManager(numberingDocument)

                    Dim abstractNodes As System.Xml.XmlNodeList = numberingDocument.SelectNodes("//w:abstractNum", namespaceManager)
                    If abstractNodes IsNot Nothing Then
                        For Each abstractNode As System.Xml.XmlNode In abstractNodes
                            Dim abstractId As System.Int32
                            If Not System.Int32.TryParse(GetWordAttributeValue(abstractNode, "abstractNumId"), abstractId) Then
                                Continue For
                            End If

                            Dim abstractDefinition As New AbstractNumberDefinition()
                            abstractDefinition.AbstractNumberId = abstractId

                            Dim numberStyleLinkNode As System.Xml.XmlNode = abstractNode.SelectSingleNode("w:numStyleLink", namespaceManager)
                            If numberStyleLinkNode IsNot Nothing Then
                                abstractDefinition.NumberStyleLink = GetWordAttributeValue(numberStyleLinkNode, "val")
                            End If

                            Dim levelNodes As System.Xml.XmlNodeList = abstractNode.SelectNodes("w:lvl", namespaceManager)
                            If levelNodes IsNot Nothing Then
                                For Each levelNode As System.Xml.XmlNode In levelNodes
                                    Dim levelDefinition As NumberingLevelDefinition = ParseNumberingLevel(levelNode, namespaceManager)
                                    abstractDefinition.Levels(levelDefinition.Level) = levelDefinition
                                Next
                            End If

                            model.AbstractNumbers(abstractId) = abstractDefinition
                        Next
                    End If

                    Dim numberNodes As System.Xml.XmlNodeList = numberingDocument.SelectNodes("//w:num", namespaceManager)
                    If numberNodes IsNot Nothing Then
                        For Each numberNode As System.Xml.XmlNode In numberNodes
                            Dim numberId As System.Int32
                            If Not System.Int32.TryParse(GetWordAttributeValue(numberNode, "numId"), numberId) Then
                                Continue For
                            End If

                            Dim abstractIdNode As System.Xml.XmlNode = numberNode.SelectSingleNode("w:abstractNumId", namespaceManager)
                            If abstractIdNode Is Nothing Then
                                Continue For
                            End If

                            Dim abstractId As System.Int32
                            If Not System.Int32.TryParse(GetWordAttributeValue(abstractIdNode, "val"), abstractId) Then
                                Continue For
                            End If

                            Dim numberDefinition As New NumberDefinition()
                            numberDefinition.NumberId = numberId
                            numberDefinition.AbstractNumberId = abstractId

                            Dim overrideNodes As System.Xml.XmlNodeList = numberNode.SelectNodes("w:lvlOverride", namespaceManager)
                            If overrideNodes IsNot Nothing Then
                                For Each overrideNode As System.Xml.XmlNode In overrideNodes
                                    Dim levelValue As System.Int32
                                    If Not System.Int32.TryParse(GetWordAttributeValue(overrideNode, "ilvl"), levelValue) Then
                                        Continue For
                                    End If
                                    levelValue = ClampNumberingLevel(levelValue)

                                    Dim levelOverride As New NumberingLevelOverride()
                                    Dim startOverrideNode As System.Xml.XmlNode = overrideNode.SelectSingleNode("w:startOverride", namespaceManager)
                                    If startOverrideNode IsNot Nothing Then
                                        Dim startValue As System.Int32
                                        If System.Int32.TryParse(GetWordAttributeValue(startOverrideNode, "val"), startValue) Then
                                            levelOverride.StartOverride = startValue
                                        End If
                                    End If

                                    Dim overriddenLevelNode As System.Xml.XmlNode = overrideNode.SelectSingleNode("w:lvl", namespaceManager)
                                    If overriddenLevelNode IsNot Nothing Then
                                        levelOverride.LevelDefinition = ParseNumberingLevel(overriddenLevelNode, namespaceManager)
                                        levelOverride.LevelDefinition.Level = levelValue
                                    End If

                                    numberDefinition.LevelOverrides(levelValue) = levelOverride
                                Next
                            End If

                            model.Numbers(numberId) = numberDefinition
                        Next
                    End If
                Catch ex As System.Exception
                    ' Numbering is optional. Extraction continues with visible text only.
                End Try

                Return model
            End Function

            Private Shared Function ParseNumberingLevel(
                levelNode As System.Xml.XmlNode,
                namespaceManager As System.Xml.XmlNamespaceManager
            ) As NumberingLevelDefinition

                Dim definition As New NumberingLevelDefinition()
                Dim levelValue As System.Int32
                If System.Int32.TryParse(GetWordAttributeValue(levelNode, "ilvl"), levelValue) Then
                    definition.Level = ClampNumberingLevel(levelValue)
                End If

                Dim startNode As System.Xml.XmlNode = levelNode.SelectSingleNode("w:start", namespaceManager)
                If startNode IsNot Nothing Then
                    Dim startValue As System.Int32
                    If System.Int32.TryParse(GetWordAttributeValue(startNode, "val"), startValue) Then
                        definition.StartValue = startValue
                    End If
                End If

                Dim formatNode As System.Xml.XmlNode = levelNode.SelectSingleNode("w:numFmt", namespaceManager)
                If formatNode IsNot Nothing Then
                    Dim formatValue As System.String = GetWordAttributeValue(formatNode, "val")
                    If Not System.String.IsNullOrWhiteSpace(formatValue) Then
                        definition.NumberFormat = formatValue
                    End If
                End If

                Dim levelTextNode As System.Xml.XmlNode = levelNode.SelectSingleNode("w:lvlText", namespaceManager)
                If levelTextNode IsNot Nothing Then
                    definition.LevelText = GetWordAttributeValue(levelTextNode, "val")
                End If

                If System.String.IsNullOrEmpty(definition.LevelText) Then
                    definition.LevelText = "%" & (definition.Level + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) & "."
                End If

                Dim suffixNode As System.Xml.XmlNode = levelNode.SelectSingleNode("w:suff", namespaceManager)
                If suffixNode IsNot Nothing Then
                    Dim suffixValue As System.String = GetWordAttributeValue(suffixNode, "val")
                    If Not System.String.IsNullOrWhiteSpace(suffixValue) Then
                        definition.Suffix = suffixValue
                    End If
                End If

                Dim styleNode As System.Xml.XmlNode = levelNode.SelectSingleNode("w:pStyle", namespaceManager)
                If styleNode IsNot Nothing Then
                    definition.ParagraphStyleId = GetWordAttributeValue(styleNode, "val")
                End If

                Dim restartNode As System.Xml.XmlNode = levelNode.SelectSingleNode("w:lvlRestart", namespaceManager)
                If restartNode IsNot Nothing Then
                    Dim restartValue As System.Int32
                    If System.Int32.TryParse(GetWordAttributeValue(restartNode, "val"), restartValue) Then
                        definition.RestartAfterLevel = restartValue
                    End If
                End If

                definition.IsLegalNumbering = levelNode.SelectSingleNode("w:isLgl", namespaceManager) IsNot Nothing
                Return definition
            End Function

            Private Shared Function ResolveEffectiveStyle(
                styleId As System.String,
                styleModel As StyleModel,
                visiting As System.Collections.Generic.HashSet(Of System.String)
            ) As EffectiveStyleProperties

                If System.String.IsNullOrWhiteSpace(styleId) Then
                    Return New EffectiveStyleProperties()
                End If

                Dim cached As EffectiveStyleProperties = Nothing
                If styleModel.EffectiveStyles.TryGetValue(styleId, cached) Then
                    Return cached
                End If

                If visiting.Contains(styleId) Then
                    Return New EffectiveStyleProperties()
                End If
                visiting.Add(styleId)

                Dim definition As StyleDefinition = Nothing
                If Not styleModel.Styles.TryGetValue(styleId, definition) Then
                    visiting.Remove(styleId)
                    Return New EffectiveStyleProperties()
                End If

                Dim effective As New EffectiveStyleProperties()
                effective.StyleId = definition.StyleId
                effective.StyleName = definition.StyleName

                If Not System.String.IsNullOrWhiteSpace(definition.BasedOnStyleId) Then
                    Dim baseEffective As EffectiveStyleProperties = ResolveEffectiveStyle(definition.BasedOnStyleId, styleModel, visiting)
                    effective.NumberingId = baseEffective.NumberingId
                    effective.NumberingLevel = baseEffective.NumberingLevel
                    effective.OutlineLevel = baseEffective.OutlineLevel
                    effective.NumberingCancelled = baseEffective.NumberingCancelled
                    If System.String.IsNullOrWhiteSpace(effective.StyleName) Then
                        effective.StyleName = baseEffective.StyleName
                    End If
                End If

                If definition.NumberingCancelled Then
                    effective.NumberingCancelled = True
                    effective.NumberingId = Nothing
                    effective.NumberingLevel = Nothing
                Else
                    If definition.NumberingId.HasValue Then
                        effective.NumberingId = definition.NumberingId
                        effective.NumberingCancelled = False
                    End If
                    If definition.NumberingLevel.HasValue Then
                        effective.NumberingLevel = definition.NumberingLevel
                    End If
                End If

                If definition.OutlineLevel.HasValue Then
                    effective.OutlineLevel = definition.OutlineLevel
                End If

                styleModel.EffectiveStyles(styleId) = effective
                visiting.Remove(styleId)
                Return effective
            End Function

            Private Shared Function GetParagraphStyleId(
                paragraphNode As System.Xml.XmlNode,
                namespaceManager As System.Xml.XmlNamespaceManager
            ) As System.String

                Dim styleNode As System.Xml.XmlNode = paragraphNode.SelectSingleNode("w:pPr/w:pStyle", namespaceManager)
                If styleNode Is Nothing Then
                    Return System.String.Empty
                End If
                Return GetWordAttributeValue(styleNode, "val")
            End Function

            Private Shared Function ResolveParagraphNumbering(
                paragraphNode As System.Xml.XmlNode,
                namespaceManager As System.Xml.XmlNamespaceManager,
                context As ExtractionContext
            ) As ParagraphNumberingProperties

                Dim styleId As System.String = GetParagraphStyleId(paragraphNode, namespaceManager)
                Dim effectiveStyle As EffectiveStyleProperties = ResolveEffectiveStyle(
                    styleId,
                    context.StyleModel,
                    New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
                )

                Dim directNumberingId As System.Nullable(Of System.Int32) = Nothing
                Dim directLevel As System.Nullable(Of System.Int32) = Nothing
                Dim directCancellation As System.Boolean = False

                Dim numIdNode As System.Xml.XmlNode = paragraphNode.SelectSingleNode("w:pPr/w:numPr/w:numId", namespaceManager)
                If numIdNode IsNot Nothing Then
                    Dim numIdValue As System.Int32
                    If System.Int32.TryParse(GetWordAttributeValue(numIdNode, "val"), numIdValue) Then
                        If numIdValue = 0 Then
                            directCancellation = True
                        Else
                            directNumberingId = numIdValue
                        End If
                    End If
                End If

                If directCancellation Then
                    Return Nothing
                End If

                Dim levelNode As System.Xml.XmlNode = paragraphNode.SelectSingleNode("w:pPr/w:numPr/w:ilvl", namespaceManager)
                If levelNode IsNot Nothing Then
                    Dim levelValue As System.Int32
                    If System.Int32.TryParse(GetWordAttributeValue(levelNode, "val"), levelValue) Then
                        directLevel = ClampNumberingLevel(levelValue)
                    End If
                End If

                Dim numberingId As System.Nullable(Of System.Int32) = directNumberingId
                If Not numberingId.HasValue AndAlso Not effectiveStyle.NumberingCancelled Then
                    numberingId = effectiveStyle.NumberingId
                End If

                Dim styleLinkedNumbering As ParagraphNumberingProperties = Nothing
                If Not numberingId.HasValue AndAlso Not effectiveStyle.NumberingCancelled Then
                    styleLinkedNumbering = FindNumberingForParagraphStyle(styleId, context)
                    If styleLinkedNumbering IsNot Nothing Then
                        numberingId = styleLinkedNumbering.NumberingId
                    End If
                End If

                If Not numberingId.HasValue Then
                    Return Nothing
                End If

                Dim level As System.Nullable(Of System.Int32) = directLevel
                If Not level.HasValue Then
                    level = effectiveStyle.NumberingLevel
                End If
                If Not level.HasValue AndAlso styleLinkedNumbering IsNot Nothing Then
                    level = styleLinkedNumbering.Level
                End If
                If Not level.HasValue Then
                    level = InferNumberingLevelFromStyle(numberingId.Value, styleId, context)
                End If
                If Not level.HasValue Then
                    level = 0
                End If

                If Not context.NumberingModel.Numbers.ContainsKey(numberingId.Value) Then
                    Return Nothing
                End If

                Dim result As New ParagraphNumberingProperties()
                result.NumberingId = numberingId.Value
                result.Level = ClampNumberingLevel(level.Value)
                Return result
            End Function


            Private Shared Function FindNumberingForParagraphStyle(
                styleId As System.String,
                context As ExtractionContext
            ) As ParagraphNumberingProperties

                If System.String.IsNullOrWhiteSpace(styleId) Then
                    Return Nothing
                End If

                Dim selectedNumberId As System.Nullable(Of System.Int32) = Nothing
                Dim selectedLevel As System.Int32 = 0

                For Each numberPair As System.Collections.Generic.KeyValuePair(Of System.Int32, NumberDefinition) In context.NumberingModel.Numbers
                    Dim abstractDefinition As AbstractNumberDefinition = Nothing
                    If Not context.NumberingModel.AbstractNumbers.TryGetValue(numberPair.Value.AbstractNumberId, abstractDefinition) Then
                        Continue For
                    End If

                    For Each levelPair As System.Collections.Generic.KeyValuePair(Of System.Int32, NumberingLevelDefinition) In abstractDefinition.Levels
                        If System.String.Equals(levelPair.Value.ParagraphStyleId, styleId, System.StringComparison.OrdinalIgnoreCase) Then
                            If Not selectedNumberId.HasValue OrElse numberPair.Key < selectedNumberId.Value Then
                                selectedNumberId = numberPair.Key
                                selectedLevel = levelPair.Key
                            End If
                        End If
                    Next
                Next

                If Not selectedNumberId.HasValue Then
                    Return Nothing
                End If

                Dim result As New ParagraphNumberingProperties()
                result.NumberingId = selectedNumberId.Value
                result.Level = ClampNumberingLevel(selectedLevel)
                Return result
            End Function

            Private Shared Function InferNumberingLevelFromStyle(
                numberingId As System.Int32,
                styleId As System.String,
                context As ExtractionContext
            ) As System.Nullable(Of System.Int32)

                If System.String.IsNullOrWhiteSpace(styleId) Then
                    Return Nothing
                End If

                Dim numberDefinition As NumberDefinition = Nothing
                If Not context.NumberingModel.Numbers.TryGetValue(numberingId, numberDefinition) Then
                    Return Nothing
                End If

                Dim abstractDefinition As AbstractNumberDefinition = Nothing
                If Not context.NumberingModel.AbstractNumbers.TryGetValue(numberDefinition.AbstractNumberId, abstractDefinition) Then
                    Return Nothing
                End If

                For Each pair As System.Collections.Generic.KeyValuePair(Of System.Int32, NumberingLevelDefinition) In abstractDefinition.Levels
                    If System.String.Equals(pair.Value.ParagraphStyleId, styleId, System.StringComparison.OrdinalIgnoreCase) Then
                        Return pair.Key
                    End If
                Next

                Return Nothing
            End Function

            Private Shared Function ResolveHeadingLevel(
                paragraphNode As System.Xml.XmlNode,
                namespaceManager As System.Xml.XmlNamespaceManager,
                context As ExtractionContext
            ) As System.Nullable(Of System.Int32)

                Dim directOutlineNode As System.Xml.XmlNode = paragraphNode.SelectSingleNode("w:pPr/w:outlineLvl", namespaceManager)
                If directOutlineNode IsNot Nothing Then
                    Dim outlineValue As System.Int32
                    If System.Int32.TryParse(GetWordAttributeValue(directOutlineNode, "val"), outlineValue) AndAlso outlineValue >= 0 AndAlso outlineValue <= 8 Then
                        Return outlineValue + 1
                    End If
                End If

                Dim styleId As System.String = GetParagraphStyleId(paragraphNode, namespaceManager)
                Dim effectiveStyle As EffectiveStyleProperties = ResolveEffectiveStyle(
                    styleId,
                    context.StyleModel,
                    New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
                )

                If effectiveStyle.OutlineLevel.HasValue Then
                    Return effectiveStyle.OutlineLevel.Value + 1
                End If

                Dim candidate As System.String = (effectiveStyle.StyleName & " " & styleId).Trim()
                Dim match As System.Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(
                    candidate,
                    "(?:Heading|Überschrift|Ueberschrift)\s*([1-9])",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.CultureInvariant
                )

                If match.Success Then
                    Dim headingLevel As System.Int32
                    If System.Int32.TryParse(match.Groups(1).Value, headingLevel) Then
                        Return headingLevel
                    End If
                End If

                Return Nothing
            End Function

            Private Shared Function GetEffectiveLevelDefinition(
                numberingId As System.Int32,
                level As System.Int32,
                context As ExtractionContext,
                visitedNumberIds As System.Collections.Generic.HashSet(Of System.Int32)
            ) As NumberingLevelDefinition

                If visitedNumberIds.Contains(numberingId) Then
                    Return Nothing
                End If
                visitedNumberIds.Add(numberingId)

                Dim numberDefinition As NumberDefinition = Nothing
                If Not context.NumberingModel.Numbers.TryGetValue(numberingId, numberDefinition) Then
                    Return Nothing
                End If

                Dim levelOverride As NumberingLevelOverride = Nothing
                If numberDefinition.LevelOverrides.TryGetValue(level, levelOverride) AndAlso levelOverride.LevelDefinition IsNot Nothing Then
                    Dim overridden As NumberingLevelDefinition = levelOverride.LevelDefinition.CloneDefinition()
                    If levelOverride.StartOverride.HasValue Then
                        overridden.StartValue = levelOverride.StartOverride.Value
                    End If
                    Return overridden
                End If

                Dim abstractDefinition As AbstractNumberDefinition = Nothing
                If Not context.NumberingModel.AbstractNumbers.TryGetValue(numberDefinition.AbstractNumberId, abstractDefinition) Then
                    Return Nothing
                End If

                Dim result As NumberingLevelDefinition = Nothing
                If abstractDefinition.Levels.TryGetValue(level, result) Then
                    result = result.CloneDefinition()
                    If levelOverride IsNot Nothing AndAlso levelOverride.StartOverride.HasValue Then
                        result.StartValue = levelOverride.StartOverride.Value
                    End If
                    Return result
                End If

                If Not System.String.IsNullOrWhiteSpace(abstractDefinition.NumberStyleLink) Then
                    Dim linkedStyle As EffectiveStyleProperties = ResolveEffectiveStyle(
                        abstractDefinition.NumberStyleLink,
                        context.StyleModel,
                        New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
                    )
                    If linkedStyle.NumberingId.HasValue Then
                        Return GetEffectiveLevelDefinition(linkedStyle.NumberingId.Value, level, context, visitedNumberIds)
                    End If
                End If

                Return Nothing
            End Function

            Private Shared Function AdvanceAndFormatNumber(
                numbering As ParagraphNumberingProperties,
                context As ExtractionContext,
                state As NumberingState
            ) As System.String

                Dim levelDefinition As NumberingLevelDefinition = GetEffectiveLevelDefinition(
                    numbering.NumberingId,
                    numbering.Level,
                    context,
                    New System.Collections.Generic.HashSet(Of System.Int32)()
                )

                If levelDefinition Is Nothing Then
                    Return System.String.Empty
                End If

                Dim counterSet As NumberingCounterSet = Nothing
                If Not state.Counters.TryGetValue(numbering.NumberingId, counterSet) Then
                    counterSet = New NumberingCounterSet()
                    state.Counters(numbering.NumberingId) = counterSet
                End If

                For lowerLevel As System.Int32 = numbering.Level + 1 To 8
                    counterSet.Initialized(lowerLevel) = False
                    counterSet.Values(lowerLevel) = 0
                Next

                If Not counterSet.Initialized(numbering.Level) Then
                    counterSet.Values(numbering.Level) = levelDefinition.StartValue
                    counterSet.Initialized(numbering.Level) = True
                Else
                    counterSet.Values(numbering.Level) += 1
                End If

                Dim levelText As System.String = levelDefinition.LevelText
                If System.String.Equals(levelDefinition.NumberFormat, "none", System.StringComparison.OrdinalIgnoreCase) Then
                    Return System.String.Empty
                End If

                If System.String.Equals(levelDefinition.NumberFormat, "bullet", System.StringComparison.OrdinalIgnoreCase) Then
                    Return NormalizeBulletText(levelText)
                End If

                For placeholderLevel As System.Int32 = 0 To 8
                    Dim placeholder As System.String = "%" & (placeholderLevel + 1).ToString(System.Globalization.CultureInfo.InvariantCulture)
                    If levelText.IndexOf(placeholder, System.StringComparison.Ordinal) < 0 Then
                        Continue For
                    End If

                    Dim placeholderDefinition As NumberingLevelDefinition = GetEffectiveLevelDefinition(
                        numbering.NumberingId,
                        placeholderLevel,
                        context,
                        New System.Collections.Generic.HashSet(Of System.Int32)()
                    )

                    If placeholderDefinition Is Nothing Then
                        placeholderDefinition = New NumberingLevelDefinition()
                        placeholderDefinition.Level = placeholderLevel
                    End If

                    If Not counterSet.Initialized(placeholderLevel) Then
                        counterSet.Values(placeholderLevel) = placeholderDefinition.StartValue
                        counterSet.Initialized(placeholderLevel) = True
                    End If

                    Dim numberFormat As System.String = placeholderDefinition.NumberFormat
                    If levelDefinition.IsLegalNumbering Then
                        numberFormat = "decimal"
                    End If

                    Dim replacement As System.String = FormatListNumber(counterSet.Values(placeholderLevel), numberFormat)
                    levelText = levelText.Replace(placeholder, replacement)
                Next

                Return levelText.Trim()
            End Function

            Private Shared Function FormatListNumber(value As System.Int32, numberFormat As System.String) As System.String
                Select Case numberFormat.ToLowerInvariant()
                    Case "upperroman"
                        Return ToRoman(value).ToUpperInvariant()
                    Case "lowerroman"
                        Return ToRoman(value).ToLowerInvariant()
                    Case "upperletter"
                        Return ToAlphabetic(value).ToUpperInvariant()
                    Case "lowerletter"
                        Return ToAlphabetic(value).ToLowerInvariant()
                    Case "decimalzero"
                        Return value.ToString("00", System.Globalization.CultureInfo.InvariantCulture)
                    Case "ordinal"
                        Return ToOrdinal(value)
                    Case "bullet"
                        Return "•"
                    Case "none"
                        Return System.String.Empty
                    Case Else
                        Return value.ToString(System.Globalization.CultureInfo.InvariantCulture)
                End Select
            End Function

            Private Shared Function ToRoman(value As System.Int32) As System.String
                If value <= 0 Then
                    Return value.ToString(System.Globalization.CultureInfo.InvariantCulture)
                End If

                Dim values As System.Int32() = {1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1}
                Dim symbols As System.String() = {"M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"}
                Dim builder As New System.Text.StringBuilder()
                Dim remaining As System.Int32 = value

                For index As System.Int32 = 0 To values.Length - 1
                    While remaining >= values(index)
                        builder.Append(symbols(index))
                        remaining -= values(index)
                    End While
                Next

                Return builder.ToString()
            End Function

            Private Shared Function ToAlphabetic(value As System.Int32) As System.String
                If value <= 0 Then
                    Return value.ToString(System.Globalization.CultureInfo.InvariantCulture)
                End If

                Dim builder As New System.Text.StringBuilder()
                Dim remaining As System.Int32 = value
                While remaining > 0
                    remaining -= 1
                    builder.Insert(0, System.Convert.ToChar(System.Convert.ToInt32("A"c) + (remaining Mod 26)))
                    remaining \= 26
                End While
                Return builder.ToString()
            End Function

            Private Shared Function ToOrdinal(value As System.Int32) As System.String
                Dim absoluteValue As System.Int32 = System.Math.Abs(value)
                Dim lastTwoDigits As System.Int32 = absoluteValue Mod 100
                Dim suffix As System.String

                If lastTwoDigits >= 11 AndAlso lastTwoDigits <= 13 Then
                    suffix = "th"
                Else
                    Select Case absoluteValue Mod 10
                        Case 1
                            suffix = "st"
                        Case 2
                            suffix = "nd"
                        Case 3
                            suffix = "rd"
                        Case Else
                            suffix = "th"
                    End Select
                End If

                Return value.ToString(System.Globalization.CultureInfo.InvariantCulture) & suffix
            End Function

            Private Shared Function NormalizeBulletText(value As System.String) As System.String
                If System.String.IsNullOrEmpty(value) Then
                    Return "•"
                End If

                Dim result As System.String = value
                result = result.Replace(System.Convert.ToChar(&HF0B7), "•"c)
                result = result.Replace(System.Convert.ToChar(&HF0A7), "▪"c)
                result = result.Replace(System.Convert.ToChar(&HF0D8), "➢"c)
                result = result.Replace(System.Convert.ToChar(&HF0FC), "✓"c)
                Return result.Trim()
            End Function

            Private Shared Function ClampNumberingLevel(value As System.Int32) As System.Int32
                If value < 0 Then
                    Return 0
                End If
                If value > 8 Then
                    Return 8
                End If
                Return value
            End Function

#End Region

#Region "Block and paragraph analysis"

            Private Shared Function AnalyseBlockChildren(
                parentNode As System.Xml.XmlNode,
                namespaceManager As System.Xml.XmlNamespaceManager,
                context As ExtractionContext,
                numberingState As NumberingState,
                storyName As System.String,
                tableNumberState As TableNumberState
            ) As System.Collections.Generic.List(Of DocumentBlock)

                Dim blocks As New System.Collections.Generic.List(Of DocumentBlock)()

                For Each childNode As System.Xml.XmlNode In parentNode.ChildNodes
                    If childNode.NamespaceURI <> SB_WordNs Then
                        Continue For
                    End If

                    Select Case childNode.LocalName
                        Case "p"
                            Dim paragraph As ParagraphInfo = AnalyseParagraph(
                                childNode,
                                namespaceManager,
                                context,
                                numberingState,
                                storyName,
                                True
                            )
                            Dim paragraphBlock As New DocumentBlock()
                            paragraphBlock.Kind = DocumentBlockKind.Paragraph
                            paragraphBlock.Paragraph = paragraph
                            blocks.Add(paragraphBlock)

                        Case "tbl"
                            Dim tableBlock As New DocumentBlock()
                            tableBlock.Kind = DocumentBlockKind.Table
                            tableBlock.Table = AnalyseTable(
                                childNode,
                                namespaceManager,
                                context,
                                numberingState,
                                storyName,
                                tableNumberState.GetNextTopLevelNumber(),
                                tableNumberState
                            )
                            blocks.Add(tableBlock)

                        Case "sdt", "customXml", "ins", "moveTo", "moveFrom"
                            Dim contentNode As System.Xml.XmlNode = childNode.SelectSingleNode("w:sdtContent", namespaceManager)
                            If contentNode Is Nothing Then
                                contentNode = childNode
                            End If
                            blocks.AddRange(
                                AnalyseBlockChildren(
                                    contentNode,
                                    namespaceManager,
                                    context,
                                    numberingState,
                                    storyName,
                                    tableNumberState
                                )
                            )
                    End Select
                Next

                Return blocks
            End Function

            Private Shared Function AnalyseTable(
                tableNode As System.Xml.XmlNode,
                namespaceManager As System.Xml.XmlNamespaceManager,
                context As ExtractionContext,
                numberingState As NumberingState,
                storyName As System.String,
                displayNumber As System.String,
                tableNumberState As TableNumberState
            ) As TableInfo

                Dim table As New TableInfo()
                table.DisplayNumber = displayNumber

                Dim rowNodes As System.Xml.XmlNodeList = tableNode.SelectNodes("w:tr", namespaceManager)
                If rowNodes Is Nothing Then
                    Return table
                End If

                Dim nestedTableIndex As System.Int32 = 0

                For Each rowNode As System.Xml.XmlNode In rowNodes
                    Dim row As New TableRowInfo()
                    Dim cellNodes As System.Xml.XmlNodeList = rowNode.SelectNodes("w:tc", namespaceManager)

                    If cellNodes IsNot Nothing Then
                        For Each cellNode As System.Xml.XmlNode In cellNodes
                            Dim cell As New TableCellInfo()
                            cell.GridSpan = GetDocxGridSpan(cellNode, namespaceManager)
                            cell.VerticalMerge = GetDocxVerticalMerge(cellNode, namespaceManager)

                            For Each childNode As System.Xml.XmlNode In cellNode.ChildNodes
                                If childNode.NamespaceURI <> SB_WordNs Then
                                    Continue For
                                End If

                                Select Case childNode.LocalName
                                    Case "p"
                                        Dim paragraphBlock As New DocumentBlock()
                                        paragraphBlock.Kind = DocumentBlockKind.Paragraph
                                        paragraphBlock.Paragraph = AnalyseParagraph(
                                            childNode,
                                            namespaceManager,
                                            context,
                                            numberingState,
                                            storyName,
                                            True
                                        )
                                        cell.Blocks.Add(paragraphBlock)

                                    Case "tbl"
                                        nestedTableIndex += 1
                                        Dim nestedBlock As New DocumentBlock()
                                        nestedBlock.Kind = DocumentBlockKind.Table
                                        nestedBlock.Table = AnalyseTable(
                                            childNode,
                                            namespaceManager,
                                            context,
                                            numberingState,
                                            storyName,
                                            displayNumber & "." & nestedTableIndex.ToString(System.Globalization.CultureInfo.InvariantCulture),
                                            tableNumberState
                                        )
                                        cell.Blocks.Add(nestedBlock)

                                    Case "sdt", "customXml", "ins", "moveTo", "moveFrom"
                                        Dim contentNode As System.Xml.XmlNode = childNode.SelectSingleNode("w:sdtContent", namespaceManager)
                                        If contentNode Is Nothing Then
                                            contentNode = childNode
                                        End If
                                        cell.Blocks.AddRange(
                                            AnalyseBlockChildren(
                                                contentNode,
                                                namespaceManager,
                                                context,
                                                numberingState,
                                                storyName,
                                                tableNumberState
                                            )
                                        )
                                End Select
                            Next

                            row.Cells.Add(cell)
                        Next
                    End If

                    table.Rows.Add(row)
                Next

                Return table
            End Function

            Private Shared Function AnalyseParagraph(
                paragraphNode As System.Xml.XmlNode,
                namespaceManager As System.Xml.XmlNamespaceManager,
                context As ExtractionContext,
                numberingState As NumberingState,
                storyName As System.String,
                registerGlobally As System.Boolean
            ) As ParagraphInfo

                Dim paragraph As New ParagraphInfo()
                paragraph.StoryName = storyName
                paragraph.SequenceIndex = context.AllParagraphs.Count + 1
                paragraph.HeadingLevel = ResolveHeadingLevel(paragraphNode, namespaceManager, context)

                Dim numbering As ParagraphNumberingProperties = ResolveParagraphNumbering(paragraphNode, namespaceManager, context)
                If numbering IsNot Nothing Then
                    paragraph.ListLevel = numbering.Level
                    Dim effectiveListDefinition As NumberingLevelDefinition = GetEffectiveLevelDefinition(
                        numbering.NumberingId,
                        numbering.Level,
                        context,
                        New System.Collections.Generic.HashSet(Of System.Int32)()
                    )
                    If effectiveListDefinition IsNot Nothing Then
                        paragraph.ListNumberFormat = effectiveListDefinition.NumberFormat
                    End If

                    paragraph.NumberText = AdvanceAndFormatNumber(numbering, context, numberingState)
                    paragraph.HasAutomaticNumber = Not System.String.IsNullOrWhiteSpace(paragraph.NumberText)
                End If

                paragraph.Tokens = ParseInlineTokens(paragraphNode, namespaceManager)
                paragraph.CachedPlainText = RenderTokensUsingCachedResults(paragraph.Tokens).Trim()

                Dim bookmarkNodes As System.Xml.XmlNodeList = paragraphNode.SelectNodes(
                    ".//w:bookmarkStart[not(ancestor::w:txbxContent)]",
                    namespaceManager
                )

                If bookmarkNodes IsNot Nothing Then
                    For Each bookmarkNode As System.Xml.XmlNode In bookmarkNodes
                        Dim bookmarkName As System.String = GetWordAttributeValue(bookmarkNode, "name")
                        If Not System.String.IsNullOrWhiteSpace(bookmarkName) AndAlso Not bookmarkName.StartsWith("_GoBack", System.StringComparison.OrdinalIgnoreCase) Then
                            If Not paragraph.BookmarkNames.Contains(bookmarkName) Then
                                paragraph.BookmarkNames.Add(bookmarkName)
                            End If
                        End If
                    Next
                End If

                If DocxIncludeMarginText Then
                    Dim textBoxNodes As System.Xml.XmlNodeList = paragraphNode.SelectNodes(".//w:txbxContent", namespaceManager)
                    If textBoxNodes IsNot Nothing Then
                        For Each textBoxNode As System.Xml.XmlNode In textBoxNodes
                            Dim marginNumberingState As New NumberingState()
                            Dim marginParagraphNodes As System.Xml.XmlNodeList = textBoxNode.SelectNodes(".//w:p[count(ancestor::w:txbxContent)=1]", namespaceManager)
                            If marginParagraphNodes IsNot Nothing Then
                                For Each marginParagraphNode As System.Xml.XmlNode In marginParagraphNodes
                                    paragraph.MarginParagraphs.Add(
                                        AnalyseParagraph(
                                            marginParagraphNode,
                                            namespaceManager,
                                            context,
                                            marginNumberingState,
                                            storyName & ":Margin",
                                            False
                                        )
                                    )
                                Next
                            End If
                        Next
                    End If
                End If

                If registerGlobally Then
                    context.AllParagraphs.Add(paragraph)
                End If

                For Each bookmarkName As System.String In paragraph.BookmarkNames
                    context.BookmarkTargets(bookmarkName) = paragraph
                Next

                Return paragraph
            End Function

            Private Shared Function AnalyseStoryFiles(
                wordDirectory As System.String,
                filePattern As System.String,
                sectionLabel As System.String,
                context As ExtractionContext
            ) As System.Collections.Generic.List(Of StorySection)

                Dim stories As New System.Collections.Generic.List(Of StorySection)()

                Try
                    Dim files As System.String() = System.IO.Directory.GetFiles(wordDirectory, filePattern)
                    System.Array.Sort(files, System.StringComparer.OrdinalIgnoreCase)

                    For Each filePath As System.String In files
                        Dim document As System.Xml.XmlDocument = LoadXmlDocument(filePath)
                        Dim namespaceManager As System.Xml.XmlNamespaceManager = CreateNamespaceManager(document)
                        Dim root As System.Xml.XmlNode = document.DocumentElement
                        If root Is Nothing Then
                            Continue For
                        End If

                        Dim fileLabel As System.String = System.IO.Path.GetFileNameWithoutExtension(filePath)
                        Dim numberPart As System.String = GetTrailingDigits(fileLabel)

                        Dim story As New StorySection()
                        story.Label = sectionLabel
                        If numberPart.Length > 0 Then
                            story.Label &= " " & numberPart
                        End If

                        story.Blocks = AnalyseBlockChildren(
                            root,
                            namespaceManager,
                            context,
                            New NumberingState(),
                            story.Label,
                            New TableNumberState()
                        )
                        stories.Add(story)
                    Next
                Catch ex As System.Exception
                    ' Supplementary parts are optional.
                End Try

                Return stories
            End Function

            Private Shared Function AnalyseNoteFile(
                wordDirectory As System.String,
                fileName As System.String,
                noteElementLocalName As System.String,
                sectionLabel As System.String,
                context As ExtractionContext
            ) As System.Collections.Generic.List(Of NoteSection)

                Dim notes As New System.Collections.Generic.List(Of NoteSection)()
                Dim filePath As System.String = System.IO.Path.Combine(wordDirectory, fileName)

                If Not System.IO.File.Exists(filePath) Then
                    Return notes
                End If

                Try
                    Dim document As System.Xml.XmlDocument = LoadXmlDocument(filePath)
                    Dim namespaceManager As System.Xml.XmlNamespaceManager = CreateNamespaceManager(document)
                    Dim noteNodes As System.Xml.XmlNodeList = document.SelectNodes("//w:" & noteElementLocalName, namespaceManager)

                    If noteNodes Is Nothing Then
                        Return notes
                    End If

                    For Each noteNode As System.Xml.XmlNode In noteNodes
                        Dim noteId As System.String = GetWordAttributeValue(noteNode, "id")
                        If noteId = "0" OrElse noteId = "-1" Then
                            Continue For
                        End If

                        Dim note As New NoteSection()
                        note.NoteId = If(System.String.IsNullOrWhiteSpace(noteId), "?", noteId)
                        note.Label = sectionLabel
                        note.Blocks = AnalyseBlockChildren(
                            noteNode,
                            namespaceManager,
                            context,
                            New NumberingState(),
                            sectionLabel & " " & note.NoteId,
                            New TableNumberState()
                        )
                        notes.Add(note)
                    Next
                Catch ex As System.Exception
                    ' Notes are optional.
                End Try

                Return notes
            End Function

            Private Shared Function GetTrailingDigits(value As System.String) As System.String
                Dim result As System.String = System.String.Empty
                For index As System.Int32 = value.Length - 1 To 0 Step -1
                    If System.Char.IsDigit(value(index)) Then
                        result = value(index) & result
                    Else
                        Exit For
                    End If
                Next
                Return result
            End Function

            Private Shared Sub AssignHeadingAnchors(context As ExtractionContext)
                Dim headingIndex As System.Int32 = 0
                For Each paragraph As ParagraphInfo In context.AllParagraphs
                    If paragraph.HeadingLevel.HasValue Then
                        headingIndex += 1
                        paragraph.AnchorText = "H" & headingIndex.ToString("0000", System.Globalization.CultureInfo.InvariantCulture)
                    End If
                Next
            End Sub

#End Region

#Region "Inline text and fields"

            Private Enum InlineTokenKind
                Text
                Field
            End Enum

            Private NotInheritable Class InlineToken
                Public Sub New()
                    Me.Text = System.String.Empty
                    Me.Instruction = System.String.Empty
                    Me.CachedResult = System.String.Empty
                End Sub

                Public Property Kind As InlineTokenKind
                Public Property Text As System.String
                Public Property Instruction As System.String
                Public Property CachedResult As System.String
            End Class

            Private NotInheritable Class FieldAccumulator
                Public Sub New()
                    Me.Instruction = New System.Text.StringBuilder()
                    Me.CachedResult = New System.Text.StringBuilder()
                End Sub

                Public Property Instruction As System.Text.StringBuilder
                Public Property CachedResult As System.Text.StringBuilder
                Public Property InResult As System.Boolean
            End Class

            Private NotInheritable Class InlineParseContext
                Public Sub New()
                    Me.Tokens = New System.Collections.Generic.List(Of InlineToken)()
                    Me.FieldStack = New System.Collections.Generic.Stack(Of FieldAccumulator)()
                End Sub

                Public Property Tokens As System.Collections.Generic.List(Of InlineToken)
                Public Property FieldStack As System.Collections.Generic.Stack(Of FieldAccumulator)
            End Class

            Private Shared Function ParseInlineTokens(
                paragraphNode As System.Xml.XmlNode,
                namespaceManager As System.Xml.XmlNamespaceManager
            ) As System.Collections.Generic.List(Of InlineToken)

                Dim parseContext As New InlineParseContext()
                ProcessInlineChildren(paragraphNode, namespaceManager, parseContext)

                While parseContext.FieldStack.Count > 0
                    CompleteCurrentField(parseContext)
                End While

                Return MergeAdjacentTextTokens(parseContext.Tokens)
            End Function

            Private Shared Sub ProcessInlineChildren(
                parentNode As System.Xml.XmlNode,
                namespaceManager As System.Xml.XmlNamespaceManager,
                parseContext As InlineParseContext
            )
                For Each childNode As System.Xml.XmlNode In parentNode.ChildNodes
                    If childNode.NamespaceURI <> SB_WordNs Then
                        Continue For
                    End If

                    If childNode.LocalName = "txbxContent" Then
                        Continue For
                    End If

                    Select Case childNode.LocalName
                        Case "fldSimple"
                            ProcessSimpleField(childNode, namespaceManager, parseContext)

                        Case "r"
                            ProcessRun(childNode, parseContext)

                        Case "hyperlink"
                            If HasFieldDescendant(childNode, namespaceManager) Then
                                ProcessInlineChildren(childNode, namespaceManager, parseContext)
                            Else
                                Dim anchor As System.String = GetWordAttributeValue(childNode, "anchor")
                                If Not System.String.IsNullOrWhiteSpace(anchor) Then
                                    Dim token As New InlineToken()
                                    token.Kind = InlineTokenKind.Field
                                    token.Instruction = "HYPERLINK \l """ & anchor & """"
                                    token.CachedResult = ExtractVisibleText(childNode, namespaceManager)
                                    AddInlineToken(parseContext, token)
                                Else
                                    ProcessInlineChildren(childNode, namespaceManager, parseContext)
                                End If
                            End If

                        Case "sdt", "sdtContent", "smartTag", "customXml", "ins", "moveTo", "moveFrom"
                            ProcessInlineChildren(childNode, namespaceManager, parseContext)

                        Case "del"
                            ' Deleted text is intentionally not emitted.

                        Case Else
                            ProcessInlineChildren(childNode, namespaceManager, parseContext)
                    End Select
                Next
            End Sub

            Private Shared Function HasFieldDescendant(
                node As System.Xml.XmlNode,
                namespaceManager As System.Xml.XmlNamespaceManager
            ) As System.Boolean
                Return node.SelectSingleNode(".//w:fldChar | .//w:fldSimple | .//w:instrText", namespaceManager) IsNot Nothing
            End Function

            Private Shared Sub ProcessSimpleField(
                fieldNode As System.Xml.XmlNode,
                namespaceManager As System.Xml.XmlNamespaceManager,
                parseContext As InlineParseContext
            )
                Dim token As New InlineToken()
                token.Kind = InlineTokenKind.Field
                token.Instruction = GetWordAttributeValue(fieldNode, "instr")
                token.CachedResult = ExtractVisibleText(fieldNode, namespaceManager)
                AddInlineToken(parseContext, token)
            End Sub

            Private Shared Sub ProcessRun(runNode As System.Xml.XmlNode, parseContext As InlineParseContext)
                For Each runChild As System.Xml.XmlNode In runNode.ChildNodes
                    If runChild.NamespaceURI <> SB_WordNs Then
                        Continue For
                    End If

                    Select Case runChild.LocalName
                        Case "fldChar"
                            Dim fieldType As System.String = GetWordAttributeValue(runChild, "fldCharType")
                            Select Case fieldType.ToLowerInvariant()
                                Case "begin"
                                    parseContext.FieldStack.Push(New FieldAccumulator())
                                Case "separate"
                                    If parseContext.FieldStack.Count > 0 Then
                                        parseContext.FieldStack.Peek().InResult = True
                                    End If
                                Case "end"
                                    CompleteCurrentField(parseContext)
                            End Select

                        Case "instrText"
                            If parseContext.FieldStack.Count > 0 Then
                                parseContext.FieldStack.Peek().Instruction.Append(runChild.InnerText)
                            End If

                        Case "t", "delText"
                            AppendInlineText(parseContext, runChild.InnerText)

                        Case "tab", "ptab"
                            AppendInlineText(parseContext, vbTab)

                        Case "br", "cr"
                            AppendInlineText(parseContext, System.Environment.NewLine)

                        Case "noBreakHyphen"
                            AppendInlineText(parseContext, System.Convert.ToChar(&H2011).ToString())

                        Case "softHyphen"
                            AppendInlineText(parseContext, System.Convert.ToChar(&HAD).ToString())

                        Case "footnoteReference"
                            If DocxIncludeHeaderFooterFootnotes Then
                                Dim noteId As System.String = GetWordAttributeValue(runChild, "id")
                                If Not System.String.IsNullOrWhiteSpace(noteId) AndAlso noteId <> "0" AndAlso noteId <> "-1" Then
                                    AppendInlineText(parseContext, " [Footnote " & noteId & "]")
                                End If
                            End If

                        Case "endnoteReference"
                            If DocxIncludeHeaderFooterFootnotes Then
                                Dim noteId As System.String = GetWordAttributeValue(runChild, "id")
                                If Not System.String.IsNullOrWhiteSpace(noteId) AndAlso noteId <> "0" AndAlso noteId <> "-1" Then
                                    AppendInlineText(parseContext, " [Endnote " & noteId & "]")
                                End If
                            End If

                        Case "sym"
                            AppendInlineText(parseContext, ReadSymbolCharacter(runChild))
                    End Select
                Next
            End Sub

            Private Shared Function ReadSymbolCharacter(symbolNode As System.Xml.XmlNode) As System.String
                Dim characterValue As System.String = GetWordAttributeValue(symbolNode, "char")
                If System.String.IsNullOrWhiteSpace(characterValue) Then
                    Return System.String.Empty
                End If

                Dim codePoint As System.Int32
                If System.Int32.TryParse(
                    characterValue,
                    System.Globalization.NumberStyles.HexNumber,
                    System.Globalization.CultureInfo.InvariantCulture,
                    codePoint
                ) Then
                    Try
                        Return System.Char.ConvertFromUtf32(codePoint)
                    Catch ex As System.Exception
                        Return System.String.Empty
                    End Try
                End If

                Return System.String.Empty
            End Function

            Private Shared Sub AppendInlineText(parseContext As InlineParseContext, value As System.String)
                If System.String.IsNullOrEmpty(value) Then
                    Return
                End If

                If parseContext.FieldStack.Count > 0 Then
                    Dim currentField As FieldAccumulator = parseContext.FieldStack.Peek()
                    If currentField.InResult Then
                        currentField.CachedResult.Append(value)
                    End If
                    Return
                End If

                Dim token As New InlineToken()
                token.Kind = InlineTokenKind.Text
                token.Text = value
                AddInlineToken(parseContext, token)
            End Sub

            Private Shared Sub CompleteCurrentField(parseContext As InlineParseContext)
                If parseContext.FieldStack.Count = 0 Then
                    Return
                End If

                Dim completed As FieldAccumulator = parseContext.FieldStack.Pop()
                Dim token As New InlineToken()
                token.Kind = InlineTokenKind.Field
                token.Instruction = completed.Instruction.ToString().Trim()
                token.CachedResult = completed.CachedResult.ToString()

                If parseContext.FieldStack.Count > 0 Then
                    Dim parent As FieldAccumulator = parseContext.FieldStack.Peek()
                    If parent.InResult Then
                        parent.CachedResult.Append(token.CachedResult)
                    End If
                Else
                    AddInlineToken(parseContext, token)
                End If
            End Sub

            Private Shared Sub AddInlineToken(parseContext As InlineParseContext, token As InlineToken)
                If parseContext.FieldStack.Count > 0 Then
                    Dim parent As FieldAccumulator = parseContext.FieldStack.Peek()
                    If parent.InResult Then
                        If token.Kind = InlineTokenKind.Text Then
                            parent.CachedResult.Append(token.Text)
                        Else
                            parent.CachedResult.Append(token.CachedResult)
                        End If
                    End If
                Else
                    parseContext.Tokens.Add(token)
                End If
            End Sub

            Private Shared Function MergeAdjacentTextTokens(
                tokens As System.Collections.Generic.List(Of InlineToken)
            ) As System.Collections.Generic.List(Of InlineToken)

                Dim result As New System.Collections.Generic.List(Of InlineToken)()
                For Each token As InlineToken In tokens
                    If token.Kind = InlineTokenKind.Text AndAlso result.Count > 0 AndAlso result(result.Count - 1).Kind = InlineTokenKind.Text Then
                        result(result.Count - 1).Text &= token.Text
                    Else
                        result.Add(token)
                    End If
                Next
                Return result
            End Function

            Private Shared Function ExtractVisibleText(
                node As System.Xml.XmlNode,
                namespaceManager As System.Xml.XmlNamespaceManager
            ) As System.String

                Dim builder As New System.Text.StringBuilder()
                ExtractVisibleTextRecursive(node, namespaceManager, builder)
                Return builder.ToString()
            End Function

            Private Shared Sub ExtractVisibleTextRecursive(
                node As System.Xml.XmlNode,
                namespaceManager As System.Xml.XmlNamespaceManager,
                builder As System.Text.StringBuilder
            )
                For Each childNode As System.Xml.XmlNode In node.ChildNodes
                    If childNode.NamespaceURI <> SB_WordNs Then
                        Continue For
                    End If

                    If childNode.LocalName = "txbxContent" Then
                        Continue For
                    End If

                    Select Case childNode.LocalName
                        Case "t", "delText"
                            builder.Append(childNode.InnerText)
                        Case "tab", "ptab"
                            builder.Append(vbTab)
                        Case "br", "cr"
                            builder.AppendLine()
                        Case "noBreakHyphen"
                            builder.Append(System.Convert.ToChar(&H2011))
                        Case "softHyphen"
                            builder.Append(System.Convert.ToChar(&HAD))
                        Case "footnoteReference"
                            Dim noteId As System.String = GetWordAttributeValue(childNode, "id")
                            If noteId <> "0" AndAlso noteId <> "-1" AndAlso noteId.Length > 0 Then
                                builder.Append(" [Footnote " & noteId & "]")
                            End If
                        Case "endnoteReference"
                            Dim noteId As System.String = GetWordAttributeValue(childNode, "id")
                            If noteId <> "0" AndAlso noteId <> "-1" AndAlso noteId.Length > 0 Then
                                builder.Append(" [Endnote " & noteId & "]")
                            End If
                        Case "sym"
                            builder.Append(ReadSymbolCharacter(childNode))
                        Case Else
                            ExtractVisibleTextRecursive(childNode, namespaceManager, builder)
                    End Select
                Next
            End Sub

            Private Shared Function RenderTokensUsingCachedResults(
                tokens As System.Collections.Generic.List(Of InlineToken)
            ) As System.String

                Dim builder As New System.Text.StringBuilder()
                For Each token As InlineToken In tokens
                    If token.Kind = InlineTokenKind.Text Then
                        builder.Append(token.Text)
                    Else
                        builder.Append(token.CachedResult)
                    End If
                Next
                Return builder.ToString()
            End Function

            Private NotInheritable Class FieldEvaluationState
                Public Sub New()
                    Me.SequenceCounters = New System.Collections.Generic.Dictionary(Of System.String, System.Int32)(System.StringComparer.OrdinalIgnoreCase)
                End Sub

                Public Property SequenceCounters As System.Collections.Generic.Dictionary(Of System.String, System.Int32)
            End Class

            Private Shared Function RenderParagraphText(
                paragraph As ParagraphInfo,
                context As ExtractionContext,
                fieldState As FieldEvaluationState
            ) As System.String

                Dim builder As New System.Text.StringBuilder()
                For Each token As InlineToken In paragraph.Tokens
                    If token.Kind = InlineTokenKind.Text Then
                        builder.Append(token.Text)
                    Else
                        builder.Append(EvaluateField(token, paragraph, context, fieldState, False))
                    End If
                Next

                Dim plainText As System.String = builder.ToString().Trim()
                Dim displayText As System.String = CombineNumberAndText(paragraph.NumberText, plainText)

                If DocxIncludeReferenceAnchors AndAlso Not System.String.IsNullOrWhiteSpace(paragraph.AnchorText) Then
                    displayText = "[" & paragraph.AnchorText & "] " & displayText
                End If

                Return displayText.Trim()
            End Function

            Private Shared Function EvaluateField(
                token As InlineToken,
                containingParagraph As ParagraphInfo,
                context As ExtractionContext,
                fieldState As FieldEvaluationState,
                forReferenceText As System.Boolean
            ) As System.String

                Dim instruction As System.String = CollapseWhitespace(token.Instruction).Trim()
                If instruction.Length = 0 Then
                    Return token.CachedResult
                End If

                Dim parts As System.Collections.Generic.List(Of System.String) = TokenizeFieldInstruction(instruction)
                If parts.Count = 0 Then
                    Return token.CachedResult
                End If

                Dim command As System.String = parts(0).ToUpperInvariant()

                Select Case command
                    Case "REF"
                        Return EvaluateReferenceField(parts, token.CachedResult, context)

                    Case "HYPERLINK"
                        Return EvaluateHyperlinkField(parts, token.CachedResult, context)

                    Case "SEQ"
                        If forReferenceText Then
                            Return GetCachedOrFallback(token.CachedResult, instruction)
                        End If
                        Return EvaluateSequenceField(parts, token.CachedResult, fieldState)

                    Case "NOTEREF"
                        Return EvaluateReferenceField(parts, token.CachedResult, context)

                    Case "PAGEREF", "PAGE", "NUMPAGES", "SECTIONPAGES", "SECTION", "STYLEREF", "LISTNUM", "AUTONUM", "AUTONUMLGL", "AUTONUMOUT"
                        Return GetCachedOrFallback(token.CachedResult, instruction)

                    Case Else
                        Return GetCachedOrFallback(token.CachedResult, instruction)
                End Select
            End Function

            Private Shared Function EvaluateReferenceField(
                parts As System.Collections.Generic.List(Of System.String),
                cachedResult As System.String,
                context As ExtractionContext
            ) As System.String

                Dim bookmarkName As System.String = GetFirstFieldArgument(parts, 1)
                If System.String.IsNullOrWhiteSpace(bookmarkName) Then
                    Return GetCachedOrFallback(cachedResult, System.String.Join(" ", parts))
                End If

                Dim target As ParagraphInfo = Nothing
                If Not context.BookmarkTargets.TryGetValue(bookmarkName, target) Then
                    Return GetCachedOrFallback(cachedResult, "REF " & bookmarkName)
                End If

                Dim numberOnly As System.Boolean = HasAnyFieldSwitch(parts, "\n", "\r", "\w")
                If numberOnly Then
                    If Not System.String.IsNullOrWhiteSpace(target.NumberText) Then
                        Return target.NumberText
                    End If
                    Return GetParagraphReferenceText(target, context)
                End If

                Dim targetText As System.String = GetParagraphReferenceText(target, context)
                If targetText.Length > 0 Then
                    Return targetText
                End If

                Return GetCachedOrFallback(cachedResult, "REF " & bookmarkName)
            End Function

            Private Shared Function EvaluateHyperlinkField(
                parts As System.Collections.Generic.List(Of System.String),
                cachedResult As System.String,
                context As ExtractionContext
            ) As System.String

                Dim localTarget As System.String = GetSwitchArgument(parts, "\l")
                If System.String.IsNullOrWhiteSpace(localTarget) Then
                    Return cachedResult
                End If

                Dim target As ParagraphInfo = Nothing
                If Not context.BookmarkTargets.TryGetValue(localTarget, target) Then
                    Return cachedResult
                End If

                Dim targetText As System.String = GetParagraphReferenceText(target, context)
                Dim cachedTrimmed As System.String = cachedResult.Trim()

                If cachedTrimmed.Length = 0 Then
                    Return targetText
                End If

                Dim targetPlain As System.String = target.CachedPlainText.Trim()
                If System.String.Equals(cachedTrimmed, targetPlain, System.StringComparison.OrdinalIgnoreCase) OrElse
                   System.String.Equals(cachedTrimmed, target.NumberText, System.StringComparison.OrdinalIgnoreCase) Then
                    Return targetText
                End If

                Return cachedResult
            End Function

            Private Shared Function EvaluateSequenceField(
                parts As System.Collections.Generic.List(Of System.String),
                cachedResult As System.String,
                fieldState As FieldEvaluationState
            ) As System.String

                Dim identifier As System.String = GetFirstFieldArgument(parts, 1)
                If System.String.IsNullOrWhiteSpace(identifier) Then
                    Return GetCachedOrFallback(cachedResult, System.String.Join(" ", parts))
                End If

                Dim currentValue As System.Int32 = 0
                fieldState.SequenceCounters.TryGetValue(identifier, currentValue)

                Dim resetValueText As System.String = GetSwitchArgument(parts, "\r")
                Dim repeatCurrent As System.Boolean = HasFieldSwitch(parts, "\c")

                If Not System.String.IsNullOrWhiteSpace(resetValueText) Then
                    Dim resetValue As System.Int32
                    If System.Int32.TryParse(resetValueText, resetValue) Then
                        currentValue = resetValue
                    Else
                        currentValue += 1
                    End If
                ElseIf repeatCurrent Then
                    If currentValue = 0 Then
                        currentValue = 1
                    End If
                Else
                    currentValue += 1
                End If

                fieldState.SequenceCounters(identifier) = currentValue

                Dim formatSwitch As System.String = GetSwitchArgument(parts, "\*")
                If System.String.IsNullOrWhiteSpace(formatSwitch) Then
                    Return currentValue.ToString(System.Globalization.CultureInfo.InvariantCulture)
                End If

                If System.String.Equals(formatSwitch, "ROMAN", System.StringComparison.Ordinal) Then
                    Return ToRoman(currentValue).ToUpperInvariant()
                End If
                If System.String.Equals(formatSwitch, "roman", System.StringComparison.Ordinal) Then
                    Return ToRoman(currentValue).ToLowerInvariant()
                End If
                If System.String.Equals(formatSwitch, "ALPHABETIC", System.StringComparison.Ordinal) Then
                    Return ToAlphabetic(currentValue).ToUpperInvariant()
                End If
                If System.String.Equals(formatSwitch, "alphabetic", System.StringComparison.Ordinal) Then
                    Return ToAlphabetic(currentValue).ToLowerInvariant()
                End If

                Return currentValue.ToString(System.Globalization.CultureInfo.InvariantCulture)
            End Function

            Private Shared Function GetParagraphReferenceText(
                paragraph As ParagraphInfo,
                context As ExtractionContext
            ) As System.String

                ' Use the target paragraph's cached field results here. This prevents circular REF
                ' fields from recursing indefinitely, while the calculated automatic paragraph
                ' number is still taken from the current numbering state.
                Dim targetText As System.String = RenderTokensUsingCachedResults(paragraph.Tokens).Trim()
                Return CombineNumberAndText(paragraph.NumberText, targetText)
            End Function

            Private Shared Function GetCachedOrFallback(cachedResult As System.String, instruction As System.String) As System.String
                If DocxUseCachedFieldResultWhenUnresolved AndAlso Not System.String.IsNullOrWhiteSpace(cachedResult) Then
                    Return cachedResult
                End If
                Return "[Field: " & instruction & "]"
            End Function

            Private Shared Function TokenizeFieldInstruction(instruction As System.String) As System.Collections.Generic.List(Of System.String)
                Dim result As New System.Collections.Generic.List(Of System.String)()
                Dim current As New System.Text.StringBuilder()
                Dim inQuotes As System.Boolean = False
                Dim quoteCharacter As System.Char = System.Convert.ToChar(0)

                For index As System.Int32 = 0 To instruction.Length - 1
                    Dim character As System.Char = instruction(index)

                    If character = """"c OrElse character = "'"c Then
                        If inQuotes AndAlso character = quoteCharacter Then
                            inQuotes = False
                        ElseIf Not inQuotes Then
                            inQuotes = True
                            quoteCharacter = character
                        Else
                            current.Append(character)
                        End If
                    ElseIf System.Char.IsWhiteSpace(character) AndAlso Not inQuotes Then
                        If current.Length > 0 Then
                            result.Add(current.ToString())
                            current.Clear()
                        End If
                    Else
                        current.Append(character)
                    End If
                Next

                If current.Length > 0 Then
                    result.Add(current.ToString())
                End If

                Return result
            End Function

            Private Shared Function GetFirstFieldArgument(
                parts As System.Collections.Generic.List(Of System.String),
                startIndex As System.Int32
            ) As System.String
                For index As System.Int32 = startIndex To parts.Count - 1
                    If Not parts(index).StartsWith("\", System.StringComparison.Ordinal) Then
                        Return parts(index)
                    End If
                Next
                Return System.String.Empty
            End Function

            Private Shared Function HasAnyFieldSwitch(
                parts As System.Collections.Generic.List(Of System.String),
                ParamArray switches As System.String()
            ) As System.Boolean
                For Each switchValue As System.String In switches
                    If HasFieldSwitch(parts, switchValue) Then
                        Return True
                    End If
                Next
                Return False
            End Function

            Private Shared Function HasFieldSwitch(
                parts As System.Collections.Generic.List(Of System.String),
                switchValue As System.String
            ) As System.Boolean
                For Each part As System.String In parts
                    If System.String.Equals(part, switchValue, System.StringComparison.OrdinalIgnoreCase) Then
                        Return True
                    End If
                Next
                Return False
            End Function

            Private Shared Function GetSwitchArgument(
                parts As System.Collections.Generic.List(Of System.String),
                switchValue As System.String
            ) As System.String
                For index As System.Int32 = 0 To parts.Count - 2
                    If System.String.Equals(parts(index), switchValue, System.StringComparison.OrdinalIgnoreCase) Then
                        Return parts(index + 1)
                    End If
                Next
                Return System.String.Empty
            End Function

            Private Shared Function CollapseWhitespace(value As System.String) As System.String
                Return System.Text.RegularExpressions.Regex.Replace(value, "\s+", " ")
            End Function

            Private Shared Function CombineNumberAndText(numberText As System.String, plainText As System.String) As System.String
                Dim trimmedNumber As System.String = numberText.Trim()
                Dim trimmedText As System.String = plainText.Trim()

                If trimmedNumber.Length = 0 Then
                    Return trimmedText
                End If
                If trimmedText.Length = 0 Then
                    Return trimmedNumber
                End If

                If StartsWithEquivalentNumber(trimmedText, trimmedNumber) Then
                    Return trimmedText
                End If

                Return trimmedNumber & " " & trimmedText
            End Function

            Private Shared Function StartsWithEquivalentNumber(text As System.String, numberText As System.String) As System.Boolean
                If text.StartsWith(numberText, System.StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If

                Dim normalizedText As System.String = System.Text.RegularExpressions.Regex.Replace(text, "\s+", " ").Trim()
                Dim normalizedNumber As System.String = System.Text.RegularExpressions.Regex.Replace(numberText, "\s+", " ").Trim()
                Return normalizedText.StartsWith(normalizedNumber, System.StringComparison.OrdinalIgnoreCase)
            End Function

#End Region

#Region "Rendering"

            Private Shared Sub RenderBlocks(
                blocks As System.Collections.Generic.List(Of DocumentBlock),
                context As ExtractionContext,
                fieldState As FieldEvaluationState,
                output As System.Text.StringBuilder,
                nestingLevel As System.Int32
            )
                For Each block As DocumentBlock In blocks
                    Select Case block.Kind
                        Case DocumentBlockKind.Paragraph
                            Dim paragraphText As System.String = RenderParagraphText(block.Paragraph, context, fieldState)
                            If paragraphText.Length > 0 Then
                                output.AppendLine(paragraphText)
                            Else
                                output.AppendLine()
                            End If

                            If DocxIncludeMarginText AndAlso block.Paragraph.MarginParagraphs.Count > 0 Then
                                For Each marginParagraph As ParagraphInfo In block.Paragraph.MarginParagraphs
                                    Dim marginText As System.String = RenderParagraphText(
                                        marginParagraph,
                                        context,
                                        New FieldEvaluationState()
                                    )
                                    If Not System.String.IsNullOrWhiteSpace(marginText) Then
                                        output.AppendLine("[Margin] " & marginText)
                                    End If
                                Next
                            End If

                        Case DocumentBlockKind.Table
                            RenderTable(block.Table, context, fieldState, output, nestingLevel)
                    End Select
                Next
            End Sub

            Private Shared Sub RenderTable(
                table As TableInfo,
                context As ExtractionContext,
                fieldState As FieldEvaluationState,
                output As System.Text.StringBuilder,
                nestingLevel As System.Int32
            )
                Dim indent As System.String = New System.String(" "c, nestingLevel * 2)
                output.AppendLine()
                output.AppendLine(indent & "[Table " & table.DisplayNumber & "]")

                If table.Rows.Count = 0 Then
                    output.AppendLine(indent & "[Empty table]")
                    output.AppendLine(indent & "[/Table " & table.DisplayNumber & "]")
                    output.AppendLine()
                    Return
                End If

                For rowIndex As System.Int32 = 0 To table.Rows.Count - 1
                    Dim row As TableRowInfo = table.Rows(rowIndex)
                    If row.Cells.Count = 0 Then
                        output.AppendLine(
                            indent & "Row " &
                            (rowIndex + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) &
                            ": [empty row]"
                        )
                        Continue For
                    End If

                    Dim visualColumnIndex As System.Int32 = 1
                    For cellIndex As System.Int32 = 0 To row.Cells.Count - 1
                        Dim cell As TableCellInfo = row.Cells(cellIndex)
                        Dim cellBuilder As New System.Text.StringBuilder()
                        RenderBlocks(cell.Blocks, context, fieldState, cellBuilder, nestingLevel + 1)
                        Dim cellText As System.String = cellBuilder.ToString().Trim()
                        If cellText.Length = 0 Then
                            cellText = "[empty]"
                        End If

                        Dim startColumn As System.Int32 = visualColumnIndex
                        Dim endColumn As System.Int32 = visualColumnIndex + cell.GridSpan - 1
                        Dim label As System.String =
                            indent &
                            "Row " & (rowIndex + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) &
                            ", Cell " & (cellIndex + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) &
                            ", Column " & startColumn.ToString(System.Globalization.CultureInfo.InvariantCulture)

                        If cell.GridSpan > 1 Then
                            label &=
                                "-" & endColumn.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                                " spanning " & cell.GridSpan.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                                " columns"
                        End If

                        If Not System.String.IsNullOrWhiteSpace(cell.VerticalMerge) Then
                            If System.String.Equals(cell.VerticalMerge, "restart", System.StringComparison.OrdinalIgnoreCase) Then
                                label &= ", starts vertical merge"
                            Else
                                label &= ", continues vertical merge from row above"
                            End If
                        End If

                        output.AppendLine(label & ": " & cellText)
                        visualColumnIndex += cell.GridSpan
                    Next
                Next

                output.AppendLine(indent & "[/Table " & table.DisplayNumber & "]")
                output.AppendLine()
            End Sub

            ' -----------------------------------------------------------------------------
            ' Markdown rendering
            ' -----------------------------------------------------------------------------

            Private Shared Sub RenderBlocksMarkdown(
                blocks As System.Collections.Generic.List(Of DocumentBlock),
                context As ExtractionContext,
                fieldState As FieldEvaluationState,
                output As System.Text.StringBuilder,
                nestingLevel As System.Int32
            )
                Dim previousWasList As System.Boolean = False

                For Each block As DocumentBlock In blocks
                    Select Case block.Kind
                        Case DocumentBlockKind.Paragraph
                            Dim paragraph As ParagraphInfo = block.Paragraph
                            Dim paragraphText As System.String = RenderParagraphText(paragraph, context, fieldState)

                            If System.String.IsNullOrWhiteSpace(paragraphText) Then
                                If Not previousWasList Then
                                    AppendMarkdownBlankLine(output)
                                End If
                                previousWasList = False
                                Continue For
                            End If

                            If paragraph.HeadingLevel.HasValue Then
                                AppendMarkdownBlankLine(output)
                                Dim headingLevel As System.Int32 = System.Math.Max(1, System.Math.Min(6, paragraph.HeadingLevel.Value))
                                output.AppendLine(New System.String("#"c, headingLevel) & " " & EscapeMarkdownInline(paragraphText))
                                output.AppendLine()
                                previousWasList = False
                            ElseIf IsMarkdownListParagraph(paragraph) Then
                                Dim listText As System.String = RenderMarkdownListItemText(paragraph, context, fieldState)
                                Dim listIndent As System.Int32 = If(paragraph.ListLevel.HasValue, paragraph.ListLevel.Value, 0)
                                listIndent += nestingLevel
                                output.Append(New System.String(" "c, System.Math.Max(0, listIndent) * 4))
                                output.Append(GetMarkdownListMarker(paragraph))
                                output.AppendLine(EscapeMarkdownInline(listText))
                                previousWasList = True
                            Else
                                If previousWasList Then
                                    output.AppendLine()
                                End If
                                output.AppendLine(EscapeMarkdownInline(paragraphText).Replace(System.Environment.NewLine, "  " & System.Environment.NewLine))
                                output.AppendLine()
                                previousWasList = False
                            End If

                            If DocxIncludeMarginText AndAlso paragraph.MarginParagraphs.Count > 0 Then
                                If previousWasList Then
                                    output.AppendLine()
                                    previousWasList = False
                                End If
                                For Each marginParagraph As ParagraphInfo In paragraph.MarginParagraphs
                                    Dim marginText As System.String = RenderParagraphText(
                                        marginParagraph,
                                        context,
                                        New FieldEvaluationState()
                                    )
                                    If Not System.String.IsNullOrWhiteSpace(marginText) Then
                                        output.AppendLine("> **Margin:** " & EscapeMarkdownInline(marginText))
                                        output.AppendLine(">")
                                    End If
                                Next
                            End If

                        Case DocumentBlockKind.Table
                            If previousWasList Then
                                output.AppendLine()
                                previousWasList = False
                            End If
                            RenderTableMarkdown(block.Table, context, fieldState, output, nestingLevel)
                    End Select
                Next
            End Sub

            Private Shared Function IsMarkdownListParagraph(paragraph As ParagraphInfo) As System.Boolean
                Return paragraph IsNot Nothing AndAlso
                    paragraph.HasAutomaticNumber AndAlso
                    paragraph.ListLevel.HasValue AndAlso
                    Not paragraph.HeadingLevel.HasValue
            End Function

            Private Shared Function RenderMarkdownListItemText(
                paragraph As ParagraphInfo,
                context As ExtractionContext,
                fieldState As FieldEvaluationState
            ) As System.String
                Dim builder As New System.Text.StringBuilder()
                For Each token As InlineToken In paragraph.Tokens
                    If token.Kind = InlineTokenKind.Text Then
                        builder.Append(token.Text)
                    Else
                        builder.Append(EvaluateField(token, paragraph, context, fieldState, False))
                    End If
                Next

                Dim plainText As System.String = builder.ToString().Trim()
                If StartsWithEquivalentNumber(plainText, paragraph.NumberText) Then
                    Dim numberText As System.String = paragraph.NumberText.Trim()
                    If plainText.Length > numberText.Length Then
                        plainText = plainText.Substring(numberText.Length).TrimStart()
                    End If
                End If
                Return plainText
            End Function

            Private Shared Function GetMarkdownListMarker(paragraph As ParagraphInfo) As System.String
                Dim numberFormat As System.String = If(paragraph.ListNumberFormat, System.String.Empty).ToLowerInvariant()
                If numberFormat = "bullet" OrElse LooksLikeBullet(paragraph.NumberText) Then
                    Return "- "
                End If

                Select Case numberFormat
                    Case "decimal", "decimalzero", "ordinal"
                        Return "1. "
                    Case Else
                        Dim visibleNumber As System.String = paragraph.NumberText.Trim()
                        If System.String.IsNullOrWhiteSpace(visibleNumber) Then
                            Return "1. "
                        End If
                        Return visibleNumber & " "
                End Select
            End Function

            Private Shared Function LooksLikeBullet(value As System.String) As System.Boolean
                If System.String.IsNullOrWhiteSpace(value) Then
                    Return False
                End If
                Dim firstCharacter As System.Char = value.Trim()(0)
                Return "•▪▫◦●○*+‣⁃▸-".IndexOf(firstCharacter) >= 0
            End Function

            Private Shared Sub RenderTableMarkdown(
                table As TableInfo,
                context As ExtractionContext,
                fieldState As FieldEvaluationState,
                output As System.Text.StringBuilder,
                nestingLevel As System.Int32
            )
                AppendMarkdownBlankLine(output)

                If table.Rows.Count = 0 Then
                    output.AppendLine("*Empty table*")
                    output.AppendLine()
                    Return
                End If

                If CanRenderAsPipeTable(table) Then
                    RenderPipeTableMarkdown(table, context, fieldState, output, nestingLevel)
                Else
                    RenderHtmlTableMarkdown(table, context, fieldState, output, nestingLevel)
                End If

                output.AppendLine()
            End Sub

            Private Shared Function CanRenderAsPipeTable(table As TableInfo) As System.Boolean
                If table Is Nothing OrElse table.Rows.Count = 0 Then
                    Return False
                End If

                Dim expectedCells As System.Int32 = -1
                For Each row As TableRowInfo In table.Rows
                    If expectedCells < 0 Then
                        expectedCells = row.Cells.Count
                    ElseIf row.Cells.Count <> expectedCells Then
                        Return False
                    End If

                    For Each cell As TableCellInfo In row.Cells
                        If cell.GridSpan <> 1 OrElse Not System.String.IsNullOrWhiteSpace(cell.VerticalMerge) Then
                            Return False
                        End If
                        For Each block As DocumentBlock In cell.Blocks
                            If block.Kind = DocumentBlockKind.Table Then
                                Return False
                            End If
                        Next
                    Next
                Next

                Return expectedCells > 0
            End Function

            Private Shared Sub RenderPipeTableMarkdown(
                table As TableInfo,
                context As ExtractionContext,
                fieldState As FieldEvaluationState,
                output As System.Text.StringBuilder,
                nestingLevel As System.Int32
            )
                Dim columnCount As System.Int32 = table.Rows(0).Cells.Count
                Dim renderedRows As New System.Collections.Generic.List(Of System.Collections.Generic.List(Of System.String))()

                For Each row As TableRowInfo In table.Rows
                    Dim renderedRow As New System.Collections.Generic.List(Of System.String)()
                    For Each cell As TableCellInfo In row.Cells
                        renderedRow.Add(RenderMarkdownTableCell(cell, context, fieldState, nestingLevel + 1))
                    Next
                    renderedRows.Add(renderedRow)
                Next

                output.Append("|")
                For columnIndex As System.Int32 = 0 To columnCount - 1
                    output.Append(" " & renderedRows(0)(columnIndex) & " |")
                Next
                output.AppendLine()

                output.Append("|")
                For columnIndex As System.Int32 = 0 To columnCount - 1
                    output.Append(" --- |")
                Next
                output.AppendLine()

                For rowIndex As System.Int32 = 1 To renderedRows.Count - 1
                    output.Append("|")
                    For columnIndex As System.Int32 = 0 To columnCount - 1
                        output.Append(" " & renderedRows(rowIndex)(columnIndex) & " |")
                    Next
                    output.AppendLine()
                Next
            End Sub

            Private Shared Sub RenderHtmlTableMarkdown(
                table As TableInfo,
                context As ExtractionContext,
                fieldState As FieldEvaluationState,
                output As System.Text.StringBuilder,
                nestingLevel As System.Int32
            )
                output.AppendLine("<table>")
                For Each row As TableRowInfo In table.Rows
                    output.AppendLine("  <tr>")
                    For Each cell As TableCellInfo In row.Cells
                        Dim attributes As New System.Text.StringBuilder()
                        If cell.GridSpan > 1 Then
                            attributes.Append(" colspan=""")
                            attributes.Append(cell.GridSpan.ToString(System.Globalization.CultureInfo.InvariantCulture))
                            attributes.Append("""")
                        End If
                        If System.String.Equals(cell.VerticalMerge, "restart", System.StringComparison.OrdinalIgnoreCase) Then
                            Dim rowSpan As System.Int32 = CountVerticalMergeSpan(table, row, cell)
                            If rowSpan > 1 Then
                                attributes.Append(" rowspan=""")
                                attributes.Append(rowSpan.ToString(System.Globalization.CultureInfo.InvariantCulture))
                                attributes.Append("""")
                            End If
                        ElseIf System.String.Equals(cell.VerticalMerge, "continue", System.StringComparison.OrdinalIgnoreCase) Then
                            Continue For
                        End If

                        Dim cellText As System.String = RenderMarkdownTableCell(cell, context, fieldState, nestingLevel + 1)
                        output.AppendLine("    <td" & attributes.ToString() & ">" & EscapeHtml(cellText) & "</td>")
                    Next
                    output.AppendLine("  </tr>")
                Next
                output.AppendLine("</table>")
            End Sub

            Private Shared Function CountVerticalMergeSpan(
                table As TableInfo,
                startRow As TableRowInfo,
                startCell As TableCellInfo
            ) As System.Int32
                Dim startRowIndex As System.Int32 = table.Rows.IndexOf(startRow)
                Dim cellIndex As System.Int32 = startRow.Cells.IndexOf(startCell)
                If startRowIndex < 0 OrElse cellIndex < 0 Then
                    Return 1
                End If

                Dim span As System.Int32 = 1
                For rowIndex As System.Int32 = startRowIndex + 1 To table.Rows.Count - 1
                    If cellIndex >= table.Rows(rowIndex).Cells.Count Then
                        Exit For
                    End If
                    Dim candidate As TableCellInfo = table.Rows(rowIndex).Cells(cellIndex)
                    If Not System.String.Equals(candidate.VerticalMerge, "continue", System.StringComparison.OrdinalIgnoreCase) Then
                        Exit For
                    End If
                    span += 1
                Next
                Return span
            End Function

            Private Shared Function RenderMarkdownTableCell(
                cell As TableCellInfo,
                context As ExtractionContext,
                fieldState As FieldEvaluationState,
                nestingLevel As System.Int32
            ) As System.String
                Dim builder As New System.Text.StringBuilder()
                RenderBlocksMarkdown(cell.Blocks, context, fieldState, builder, nestingLevel)
                Dim text As System.String = builder.ToString().Trim()
                text = text.Replace("|", "\|")
                text = System.Text.RegularExpressions.Regex.Replace(text, "\r?\n\s*\r?\n", "<br><br>")
                text = text.Replace(System.Environment.NewLine, "<br>")
                Return text
            End Function

            Private Shared Sub RenderNoteSectionsMarkdown(
                notes As System.Collections.Generic.List(Of NoteSection),
                context As ExtractionContext,
                output As System.Text.StringBuilder
            )
                If notes.Count = 0 Then
                    Return
                End If

                Dim grouped As New System.Collections.Generic.Dictionary(Of System.String, System.Collections.Generic.List(Of NoteSection))(System.StringComparer.OrdinalIgnoreCase)
                For Each note As NoteSection In notes
                    Dim group As System.Collections.Generic.List(Of NoteSection) = Nothing
                    If Not grouped.TryGetValue(note.Label, group) Then
                        group = New System.Collections.Generic.List(Of NoteSection)()
                        grouped(note.Label) = group
                    End If
                    group.Add(note)
                Next

                For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, System.Collections.Generic.List(Of NoteSection)) In grouped
                    AppendMarkdownBlankLine(output)
                    output.AppendLine("## " & EscapeMarkdownInline(pair.Key & "s"))
                    output.AppendLine()

                    For Each note As NoteSection In pair.Value
                        Dim noteBuilder As New System.Text.StringBuilder()
                        RenderBlocksMarkdown(note.Blocks, context, New FieldEvaluationState(), noteBuilder, 0)
                        Dim noteText As System.String = noteBuilder.ToString().Trim()
                        If noteText.Length > 0 Then
                            noteText = RemoveLeadingNoteReferenceMarker(noteText, pair.Key, note.NoteId)
                            noteText = System.Text.RegularExpressions.Regex.Replace(noteText, "\r?\n", " ").Trim()
                            output.AppendLine("- **" & EscapeMarkdownInline(pair.Key & " " & note.NoteId) & ":** " & noteText)
                        End If
                    Next
                    output.AppendLine()
                Next
            End Sub

            Private Shared Sub AppendMarkdownBlankLine(output As System.Text.StringBuilder)
                If output.Length = 0 Then
                    Return
                End If

                Dim current As System.String = output.ToString()
                If current.EndsWith(System.Environment.NewLine & System.Environment.NewLine, System.StringComparison.Ordinal) Then
                    Return
                End If
                If current.EndsWith(System.Environment.NewLine, System.StringComparison.Ordinal) Then
                    output.AppendLine()
                Else
                    output.AppendLine()
                    output.AppendLine()
                End If
            End Sub

            Private Shared Function EscapeMarkdownInline(value As System.String) As System.String
                If System.String.IsNullOrEmpty(value) Then
                    Return System.String.Empty
                End If

                Dim result As System.String = value
                result = result.Replace("\", "\\")
                result = result.Replace("`", "\`")
                result = result.Replace("*", "\*")
                result = result.Replace("_", "\_")
                result = result.Replace("[", "\[")
                result = result.Replace("]", "\]")
                result = result.Replace("|", "\|")
                Return result
            End Function

            Private Shared Function EscapeHtml(value As System.String) As System.String
                If System.String.IsNullOrEmpty(value) Then
                    Return System.String.Empty
                End If
                Return value.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("""", "&quot;")
            End Function

            Private Shared Sub RenderNoteSections(
                notes As System.Collections.Generic.List(Of NoteSection),
                context As ExtractionContext,
                output As System.Text.StringBuilder
            )
                If notes.Count = 0 Then
                    Return
                End If

                Dim grouped As New System.Collections.Generic.Dictionary(Of System.String, System.Collections.Generic.List(Of NoteSection))(System.StringComparer.OrdinalIgnoreCase)
                For Each note As NoteSection In notes
                    Dim group As System.Collections.Generic.List(Of NoteSection) = Nothing
                    If Not grouped.TryGetValue(note.Label, group) Then
                        group = New System.Collections.Generic.List(Of NoteSection)()
                        grouped(note.Label) = group
                    End If
                    group.Add(note)
                Next

                For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, System.Collections.Generic.List(Of NoteSection)) In grouped
                    output.AppendLine()
                    output.AppendLine("--- " & pair.Key & "s ---")

                    For Each note As NoteSection In pair.Value
                        Dim noteBuilder As New System.Text.StringBuilder()
                        RenderBlocks(note.Blocks, context, New FieldEvaluationState(), noteBuilder, 0)
                        Dim noteText As System.String = noteBuilder.ToString().Trim()
                        If noteText.Length > 0 Then
                            noteText = RemoveLeadingNoteReferenceMarker(noteText, pair.Key, note.NoteId)
                            output.AppendLine("[" & pair.Key & " " & note.NoteId & "] " & noteText)
                        End If
                    Next
                Next
            End Sub

            Private Shared Function RemoveLeadingNoteReferenceMarker(
                noteText As System.String,
                label As System.String,
                noteId As System.String
            ) As System.String
                Dim marker As System.String = "[" & label & " " & noteId & "]"
                Dim trimmed As System.String = noteText.TrimStart()
                If trimmed.StartsWith(marker, System.StringComparison.OrdinalIgnoreCase) Then
                    Return trimmed.Substring(marker.Length).TrimStart()
                End If
                Return noteText
            End Function

#End Region

#Region "Line numbers, tables and XML helpers"

            Private Shared Function ReadLineNumberSettings(
                bodyNode As System.Xml.XmlNode,
                namespaceManager As System.Xml.XmlNamespaceManager
            ) As System.Collections.Generic.List(Of System.String)

                Dim result As New System.Collections.Generic.List(Of System.String)()
                Dim lineNumberNodes As System.Xml.XmlNodeList = bodyNode.SelectNodes(".//w:sectPr/w:lnNumType", namespaceManager)
                If lineNumberNodes Is Nothing Then
                    Return result
                End If

                Dim sectionIndex As System.Int32 = 0
                For Each lineNumberNode As System.Xml.XmlNode In lineNumberNodes
                    sectionIndex += 1
                    Dim countBy As System.String = GetWordAttributeValue(lineNumberNode, "countBy")
                    Dim startValue As System.String = GetWordAttributeValue(lineNumberNode, "start")
                    Dim distance As System.String = GetWordAttributeValue(lineNumberNode, "distance")
                    Dim restart As System.String = GetWordAttributeValue(lineNumberNode, "restart")

                    Dim builder As New System.Text.StringBuilder()
                    builder.Append("[Section ")
                    builder.Append(sectionIndex.ToString(System.Globalization.CultureInfo.InvariantCulture))
                    builder.Append(" line numbering")

                    If countBy.Length > 0 Then
                        builder.Append(", every ")
                        builder.Append(countBy)
                        builder.Append(" line(s)")
                    End If
                    If startValue.Length > 0 Then
                        builder.Append(", starts at ")
                        builder.Append(startValue)
                    End If
                    If restart.Length > 0 Then
                        builder.Append(", restart=")
                        builder.Append(restart)
                    End If
                    If distance.Length > 0 Then
                        builder.Append(", distance=")
                        builder.Append(distance)
                        builder.Append(" twentieth(s) of a point")
                    End If
                    builder.Append("]")

                    result.Add(builder.ToString())
                Next

                Return result
            End Function

            Private Shared Function GetDocxGridSpan(
                cellNode As System.Xml.XmlNode,
                namespaceManager As System.Xml.XmlNamespaceManager
            ) As System.Int32
                Dim gridSpanNode As System.Xml.XmlNode = cellNode.SelectSingleNode("w:tcPr/w:gridSpan", namespaceManager)
                If gridSpanNode Is Nothing Then
                    Return 1
                End If

                Dim result As System.Int32
                If System.Int32.TryParse(GetWordAttributeValue(gridSpanNode, "val"), result) AndAlso result > 1 Then
                    Return result
                End If
                Return 1
            End Function

            Private Shared Function GetDocxVerticalMerge(
                cellNode As System.Xml.XmlNode,
                namespaceManager As System.Xml.XmlNamespaceManager
            ) As System.String
                Dim verticalMergeNode As System.Xml.XmlNode = cellNode.SelectSingleNode("w:tcPr/w:vMerge", namespaceManager)
                If verticalMergeNode Is Nothing Then
                    Return System.String.Empty
                End If

                Dim value As System.String = GetWordAttributeValue(verticalMergeNode, "val")
                If System.String.IsNullOrWhiteSpace(value) Then
                    Return "continue"
                End If
                Return value
            End Function

            Private Shared Function GetWordAttributeValue(
                node As System.Xml.XmlNode,
                localName As System.String
            ) As System.String
                If node Is Nothing OrElse node.Attributes Is Nothing Then
                    Return System.String.Empty
                End If

                Dim attribute As System.Xml.XmlNode = node.Attributes.GetNamedItem(localName, SB_WordNs)
                If attribute IsNot Nothing Then
                    Return attribute.Value
                End If

                attribute = node.Attributes.GetNamedItem("w:" & localName)
                If attribute IsNot Nothing Then
                    Return attribute.Value
                End If

                attribute = node.Attributes.GetNamedItem(localName)
                If attribute IsNot Nothing Then
                    Return attribute.Value
                End If

                Return System.String.Empty
            End Function

#End Region
        End Class

    End Class

End Namespace
