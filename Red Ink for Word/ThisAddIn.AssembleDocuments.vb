' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AssembleDocuments.vb
' Purpose: Assembles new documents from one or more Word templates combined with
'          user-provided facts and instructions, using a two-phase Plan & Execute
'          approach that preserves template formatting through OpenXML node cloning.
'
' Architecture / Key Ideas:
'  - Plan & Execute: A single LLM "planning" call produces a JSON assembly plan
'    with per-paragraph actions; a series of focused "execution" calls produce
'    text for sections that need adaptation, merging, or generation.
'  - OpenXML Cloning: Every output paragraph is created by deep-cloning a w:p
'    node from a source template, then replacing only w:t text content.  This
'    guarantees that w:pPr, w:rPr, numbering, indentation, and spacing are
'    inherited from the template — identical formatting by construction.
'  - Table Handling: Tables are indexed as single content nodes.  For "copy"
'    the entire w:tbl is deep-cloned.  For "adapt" the table is cloned, the
'    LLM receives and returns structured tab/newline cell data, and rows can
'    be removed if the LLM marks them with a [REMOVE] token.
'  - Primary Template: When multiple templates are combined, one is designated
'    "primary" and its paragraph styles are used for all generated/merged text.
'  - Section-by-Section Execution: Each section in the plan is dispatched as a
'    separate LLM call.  A rolling context summary of completed sections keeps
'    the LLM coherent across boundaries without exceeding token limits.
'  - Copy-First: The plan defaults to "copy" actions (no LLM call, byte-perfect
'    formatting) and uses the LLM only where facts demand textual changes.
'  - Reuses: DOCX ZIP extraction/repacking, ExtractTranslateParagraphsFromXml,
'    DistributeProportional, SetTextNodeWithSpacePreserve, non-breaking-space
'    preservation, ProgressBarModule, and SharedMethods.LLM from the existing
'    translate/correct pipeline.
'
' Notes:
'  - Entry point: AssembleDocumentFromTemplates (called from Freestyle "Assemble:")
'  - System prompts SP_Assemble_Plan and SP_Assemble_Execute should be defined
'    in SharedMethods.Constants and exposed via ThisAddIn.Properties.  Temporary
'    default constants are included here for standalone operation.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.IO.Compression
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Word
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ' ═════════════════════════════════════════════════════════════════════════
    ' Constants
    ' ═════════════════════════════════════════════════════════════════════════

    ''' <summary>Suffix appended to the output filename.</summary>
    Private Const AssembleOutputSuffix As String = "_assembled"

    ''' <summary>Token that the LLM uses to mark a table row for removal.</summary>
    Private Const TableRowRemoveToken As String = "[REMOVE]"


    ' ═════════════════════════════════════════════════════════════════════════
    ' Data Structures
    ' ═════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Represents a loaded template with its extracted paragraph structure and XML.
    ''' </summary>
    Private Class AssemblyTemplateInfo
        Public Property Name As String
        Public Property FilePath As String
        Public Property TempDir As String
        Public Property XmlDoc As System.Xml.XmlDocument
        Public Property NsMgr As System.Xml.XmlNamespaceManager

        ''' <summary>
        ''' All direct-child content nodes of w:body in document order.
        ''' This includes w:p (paragraph), w:tbl (table), and w:sdt (structured
        ''' document tag / content control) elements.  Tables are indexed as a
        ''' single entry — deep-cloning copies all rows, cells, merges, borders,
        ''' shading, and column widths.
        ''' </summary>
        Public Property ContentNodes As List(Of System.Xml.XmlNode)

        ''' <summary>
        ''' The node type at each index: "p" for paragraph, "tbl" for table,
        ''' "sdt" for structured document tag.  Used by the planning prompt to
        ''' annotate each index so the LLM knows what it is copying.
        ''' </summary>
        Public Property ContentNodeTypes As List(Of String)

        ''' <summary>
        ''' Extracted paragraph/table text info, one entry per ContentNodes index.
        ''' For w:tbl nodes the FullText contains the concatenated cell text
        ''' (row-by-row, tab-separated) so the LLM can read table content during
        ''' planning.  For w:p nodes this is the standard TranslateParagraphInfo.
        ''' </summary>
        Public Property Paragraphs As List(Of TranslateParagraphInfo)
    End Class


    ' ═════════════════════════════════════════════════════════════════════════
    ' Entry Point
    ' ═════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Entry point for template-based document assembly.
    ''' Called from Freestyle with 'Assemble:' prefix, or directly from ribbon.
    ''' </summary>
    ''' <param name="userInstruction">The user's instruction/facts from the prompt.</param>
    ''' <param name="additionalContext">Additional context from {doc}/{dir} triggers (may be empty).</param>
    ''' <param name="useSecondAPI">Whether to use the alternate model.</param>
    Public Async Function AssembleDocumentFromTemplates(
        userInstruction As String,
        additionalContext As String,
        useSecondAPI As Boolean) As System.Threading.Tasks.Task

        Dim templatePaths As New List(Of String)()
        Dim templates As New Dictionary(Of String, AssemblyTemplateInfo)(StringComparer.OrdinalIgnoreCase)
        Dim primaryTemplateName As String = ""

        Try
            ' ═══ Step 1: Template Selection ═══

            ' First template (required)
            Dim firstPaths As List(Of String) = AssembleSelectTemplatesByBrowse("Select the PRIMARY Word template (.docx)")
            If firstPaths.Count = 0 Then Return
            templatePaths.AddRange(firstPaths)

            ' Ask if user wants additional templates
            While True
                Dim addMore As Integer = ShowCustomYesNoBox(
                    $"Selected {templatePaths.Count} template(s):" & vbCrLf &
                    String.Join(vbCrLf, templatePaths.Select(Function(p) "  • " & Path.GetFileName(p))) & vbCrLf & vbCrLf &
                    "Add more templates?",
                    "No, continue with assembly", "Yes, add another template")

                If addMore = 0 Then Return
                If addMore = 1 Then Exit While

                Dim morePaths As List(Of String) = AssembleSelectTemplatesByBrowse("Select an additional template (.docx)")
                If morePaths.Count > 0 Then templatePaths.AddRange(morePaths)
            End While

            ' ═══ Step 2: Primary Template Selection ═══

            Dim primaryIndex As Integer = 0
            If templatePaths.Count > 1 Then
                Dim items As New List(Of SelectionItem)()
                For i As Integer = 0 To templatePaths.Count - 1
                    items.Add(New SelectionItem(Path.GetFileName(templatePaths(i)), i + 1))
                Next

                Dim chosen As Integer = SelectValue(items, 1,
                    "Which template should be the PRIMARY template?" & vbCrLf &
                    "The primary template determines the output document's formatting (fonts, styles, page layout).",
                    AN & " Assembly — Primary Template")

                If chosen <= 0 Then Return
                primaryIndex = chosen - 1
            End If

            ' ═══ Step 3: Load and Analyze Templates ═══

            ProgressBarModule.GlobalProgressValue = 0
            ProgressBarModule.GlobalProgressMax = templatePaths.Count + 3 ' +3 for plan, execute, assemble
            ProgressBarModule.GlobalProgressLabel = "Loading templates..."
            ProgressBarModule.CancelOperation = False
            ProgressBarModule.ShowProgressBarInSeparateThread(AN & " Assembly", "Loading templates...")

            For i As Integer = 0 To templatePaths.Count - 1
                If ProgressBarModule.CancelOperation Then Return

                Dim tmplPath As String = templatePaths(i)
                Dim tmplName As String = Path.GetFileNameWithoutExtension(tmplPath)

                ' Ensure unique names
                If templates.ContainsKey(tmplName) Then
                    tmplName = tmplName & "_" & (i + 1).ToString()
                End If

                Dim tmplInfo As AssemblyTemplateInfo = AssembleLoadTemplate(tmplPath, tmplName)
                If tmplInfo Is Nothing Then
                    ShowCustomMessageBox($"Failed to load template: {Path.GetFileName(tmplPath)}")
                    Return
                End If

                templates(tmplName) = tmplInfo
                If i = primaryIndex Then primaryTemplateName = tmplName

                ProgressBarModule.GlobalProgressValue = i + 1
                ProgressBarModule.GlobalProgressLabel = $"Loaded template {i + 1}/{templatePaths.Count}: {Path.GetFileName(tmplPath)}"
            Next

            ' ═══ Step 4: Build Planning Prompt ═══

            If ProgressBarModule.CancelOperation Then Return
            ProgressBarModule.GlobalProgressLabel = "Planning document structure..."

            Dim planPromptBuilder As New StringBuilder()

            ' Include full text of each template with content indices
            For Each kvp In templates
                Dim tmpl As AssemblyTemplateInfo = kvp.Value
                planPromptBuilder.AppendLine($"<TEMPLATE name=""{kvp.Key}""{If(kvp.Key = primaryTemplateName, " primary=""true""", "")}>")
                For pIdx As Integer = 0 To tmpl.Paragraphs.Count - 1
                    Dim p As TranslateParagraphInfo = tmpl.Paragraphs(pIdx)
                    Dim nodeType As String = tmpl.ContentNodeTypes(pIdx)

                    ' Build annotation based on node type
                    Dim typeHint As String = ""
                    Select Case nodeType
                        Case "tbl"
                            ' Count rows and columns for the LLM
                            Dim tblNode As System.Xml.XmlNode = tmpl.ContentNodes(pIdx)
                            Dim rowCount As Integer = tblNode.SelectNodes("w:tr", tmpl.NsMgr).Count
                            Dim firstRow As System.Xml.XmlNode = tblNode.SelectSingleNode("w:tr", tmpl.NsMgr)
                            Dim colCount As Integer = 0
                            If firstRow IsNot Nothing Then
                                colCount = firstRow.SelectNodes("w:tc", tmpl.NsMgr).Count
                            End If
                            typeHint = $" [TABLE {rowCount}x{colCount}]"

                        Case "sdt"
                            typeHint = " [ContentControl]"

                        Case "p"
                            ' Check for heading style
                            Dim paraNode As System.Xml.XmlNode = tmpl.ContentNodes(pIdx)
                            Dim styleNode As System.Xml.XmlNode = paraNode.SelectSingleNode("w:pPr/w:pStyle/@w:val", tmpl.NsMgr)
                            If styleNode IsNot Nothing Then
                                Dim styleName As String = styleNode.Value
                                Dim headingLevel As Integer = AssembleGetHeadingLevel(styleName)
                                If headingLevel > 0 Then
                                    typeHint = $" [Heading{headingLevel}]"
                                End If
                            End If
                    End Select

                    Dim displayText As String = If(p.IsEmpty, "(empty)", p.FullText)
                    planPromptBuilder.AppendLine($"[{pIdx}]{typeHint} {displayText}")
                Next
                planPromptBuilder.AppendLine($"</TEMPLATE>")
                planPromptBuilder.AppendLine()
            Next

            ' Include facts/instruction
            planPromptBuilder.AppendLine("<FACTS>")
            planPromptBuilder.AppendLine(userInstruction)
            planPromptBuilder.AppendLine("</FACTS>")

            ' Include additional context if present
            If Not String.IsNullOrWhiteSpace(additionalContext) Then
                planPromptBuilder.AppendLine()
                planPromptBuilder.AppendLine("<CONTEXT>")
                planPromptBuilder.AppendLine(additionalContext)
                planPromptBuilder.AppendLine("</CONTEXT>")
            End If

            ' ═══ Step 5: Planning LLM Call ═══

            Dim planSystemPrompt As String = InterpolateAtRuntime(SP_Assemble_Plan)
            Dim planResponse As String = Await SharedMethods.LLM(
                _context, planSystemPrompt, planPromptBuilder.ToString(),
                "", "", 0, useSecondAPI, True)

            If String.IsNullOrWhiteSpace(planResponse) Then
                ShowCustomMessageBox("The LLM returned an empty response during planning. Assembly aborted.")
                Return
            End If

            ' ═══ Step 6: Parse the Assembly Plan ═══

            Dim plan As Newtonsoft.Json.Linq.JObject = Nothing
            Try
                ' Strip markdown fences if the LLM wrapped the JSON
                Dim cleanResponse As String = planResponse.Trim()
                If cleanResponse.StartsWith("```") Then
                    Dim firstNewline As Integer = cleanResponse.IndexOf(vbLf)
                    If firstNewline > 0 Then cleanResponse = cleanResponse.Substring(firstNewline + 1)
                    If cleanResponse.EndsWith("```") Then cleanResponse = cleanResponse.Substring(0, cleanResponse.Length - 3).Trim()
                End If
                plan = Newtonsoft.Json.Linq.JObject.Parse(cleanResponse)
            Catch ex As Exception
                ShowCustomMessageBox($"Failed to parse the assembly plan as JSON. The LLM response may be malformed." & vbCrLf & vbCrLf &
                                     "Response (first 500 chars):" & vbCrLf & planResponse.Substring(0, Math.Min(500, planResponse.Length)))
                Return
            End Try

            Dim outputSections As Newtonsoft.Json.Linq.JArray = plan("outputSections")
            If outputSections Is Nothing OrElse outputSections.Count = 0 Then
                ShowCustomMessageBox("The assembly plan contains no output sections. Assembly aborted.")
                Return
            End If

            ProgressBarModule.GlobalProgressValue += 1

            ' ═══ Step 7: Execute Plan — Section by Section ═══

            If ProgressBarModule.CancelOperation Then Return
            ProgressBarModule.GlobalProgressMax = outputSections.Count + 2
            ProgressBarModule.GlobalProgressValue = 0
            ProgressBarModule.GlobalProgressLabel = "Executing assembly plan..."

            ' Rolling context summary for coherence across sections
            Dim contextSummary As New StringBuilder()

            ' We will build a list of cloned content nodes in output order
            Dim outputParagraphNodes As New List(Of System.Xml.XmlNode)()

            ' Use primary template's XmlDoc as the import target
            Dim primaryTemplate As AssemblyTemplateInfo = templates(primaryTemplateName)

            For sectionIdx As Integer = 0 To outputSections.Count - 1
                If ProgressBarModule.CancelOperation Then Exit For

                Dim section As Newtonsoft.Json.Linq.JObject = CType(outputSections(sectionIdx), Newtonsoft.Json.Linq.JObject)
                Dim sectionHeading As String = If(section("heading")?.ToString(), "")
                Dim sectionTextCollector As New StringBuilder()

                ProgressBarModule.GlobalProgressLabel = $"Section {sectionIdx + 1}/{outputSections.Count}: {sectionHeading}"

                ' ── Insert heading paragraph ──
                Dim headingStyleFrom As Newtonsoft.Json.Linq.JObject = CType(section("headingStyleFrom"), Newtonsoft.Json.Linq.JObject)
                If headingStyleFrom IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(sectionHeading) Then
                    Dim hTmplName As String = If(headingStyleFrom("template")?.ToString(), primaryTemplateName)
                    Dim hParaIdx As Integer = If(headingStyleFrom("paragraphIndex")?.ToObject(Of Integer)(), 0)
                    Dim headingNode As System.Xml.XmlNode = AssembleCloneAndImport(primaryTemplate.XmlDoc, templates, hTmplName, hParaIdx)
                    If headingNode IsNot Nothing Then
                        AssembleReplaceAllText(headingNode, sectionHeading, primaryTemplate.NsMgr)
                        outputParagraphNodes.Add(headingNode)
                        sectionTextCollector.AppendLine(sectionHeading)
                    End If
                End If

                ' ── Process entries ──
                Dim entries As Newtonsoft.Json.Linq.JArray = section("entries")
                If entries Is Nothing Then Continue For

                For Each entry As Newtonsoft.Json.Linq.JObject In entries
                    If ProgressBarModule.CancelOperation Then Exit For

                    Dim action As String = If(entry("action")?.ToString(), "").ToLowerInvariant()

                    Select Case action

                        Case "copy"
                            Dim srcTmpl As String = If(entry("source")?.ToString(), primaryTemplateName)
                            Dim paraIdx As Integer = If(entry("paragraphIndex")?.ToObject(Of Integer)(), 0)
                            Dim cloned As System.Xml.XmlNode = AssembleCloneAndImport(primaryTemplate.XmlDoc, templates, srcTmpl, paraIdx)
                            If cloned IsNot Nothing Then
                                outputParagraphNodes.Add(cloned)
                                ' Collect text for context summary
                                If templates.ContainsKey(srcTmpl) AndAlso paraIdx < templates(srcTmpl).Paragraphs.Count Then
                                    sectionTextCollector.AppendLine(templates(srcTmpl).Paragraphs(paraIdx).FullText)
                                End If
                            End If

                        Case "copyrange"
                            Dim srcTmpl As String = If(entry("source")?.ToString(), primaryTemplateName)
                            Dim fromIdx As Integer = If(entry("fromIndex")?.ToObject(Of Integer)(), 0)
                            Dim toIdx As Integer = If(entry("toIndex")?.ToObject(Of Integer)(), fromIdx)
                            For pIdx As Integer = fromIdx To toIdx
                                Dim cloned As System.Xml.XmlNode = AssembleCloneAndImport(primaryTemplate.XmlDoc, templates, srcTmpl, pIdx)
                                If cloned IsNot Nothing Then
                                    outputParagraphNodes.Add(cloned)
                                    If templates.ContainsKey(srcTmpl) AndAlso pIdx < templates(srcTmpl).Paragraphs.Count Then
                                        sectionTextCollector.AppendLine(templates(srcTmpl).Paragraphs(pIdx).FullText)
                                    End If
                                End If
                            Next

                        Case "adapt"
                            Dim srcTmpl As String = If(entry("source")?.ToString(), primaryTemplateName)
                            Dim paraIdx As Integer = If(entry("paragraphIndex")?.ToObject(Of Integer)(), 0)
                            Dim instruction As String = If(entry("instruction")?.ToString(), "")

                            ' Get source text
                            Dim sourceText As String = ""
                            If templates.ContainsKey(srcTmpl) AndAlso paraIdx < templates(srcTmpl).Paragraphs.Count Then
                                sourceText = templates(srcTmpl).Paragraphs(paraIdx).FullText
                            End If

                            ' Determine node type
                            Dim srcNodeType As String = ""
                            If templates.ContainsKey(srcTmpl) AndAlso paraIdx < templates(srcTmpl).ContentNodeTypes.Count Then
                                srcNodeType = templates(srcTmpl).ContentNodeTypes(paraIdx)
                            End If

                            ' For tables, use the table-specific execution flow
                            If srcNodeType = "tbl" Then
                                Dim adaptedTableText As String = Await AssembleExecuteTableSection(
                                    sourceText, instruction, userInstruction, contextSummary.ToString(), useSecondAPI)

                                Dim cloned As System.Xml.XmlNode = AssembleCloneAndImport(primaryTemplate.XmlDoc, templates, srcTmpl, paraIdx)
                                If cloned IsNot Nothing Then
                                    If Not String.IsNullOrWhiteSpace(adaptedTableText) Then
                                        AssembleReplaceTableText(cloned, adaptedTableText, primaryTemplate.NsMgr)
                                    End If
                                    outputParagraphNodes.Add(cloned)
                                    sectionTextCollector.AppendLine(If(adaptedTableText, sourceText))
                                End If
                            Else
                                ' Paragraph: standard execution + proportional distribution
                                Dim adaptedText As String = Await AssembleExecuteSection(
                                    sourceText, instruction, userInstruction, contextSummary.ToString(), useSecondAPI)

                                Dim cloned As System.Xml.XmlNode = AssembleCloneAndImport(primaryTemplate.XmlDoc, templates, srcTmpl, paraIdx)
                                If cloned IsNot Nothing Then
                                    If Not String.IsNullOrWhiteSpace(adaptedText) Then
                                        AssembleReplaceTextPreservingRuns(cloned, adaptedText, primaryTemplate.NsMgr)
                                    End If
                                    outputParagraphNodes.Add(cloned)
                                    sectionTextCollector.AppendLine(If(adaptedText, sourceText))
                                End If
                            End If

                        Case "generate"
                            Dim instruction As String = If(entry("instruction")?.ToString(), "")
                            Dim estimatedParas As Integer = If(entry("estimatedParagraphs")?.ToObject(Of Integer)(), 1)

                            ' Determine style donor
                            Dim styleFrom As Newtonsoft.Json.Linq.JObject = CType(entry("styleFrom"), Newtonsoft.Json.Linq.JObject)
                            Dim donorTmpl As String = If(styleFrom?("template")?.ToString(), primaryTemplateName)
                            Dim donorIdx As Integer = If(styleFrom?("paragraphIndex")?.ToObject(Of Integer)(), 0)

                            ' LLM call to generate
                            Dim generatedText As String = Await AssembleExecuteSection(
                                "", instruction, userInstruction, contextSummary.ToString(), useSecondAPI)

                            If Not String.IsNullOrWhiteSpace(generatedText) Then
                                ' Split on paragraph boundaries
                                Dim genParagraphs As String() = generatedText.Split(
                                    New String() {vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)

                                For Each genPara In genParagraphs
                                    Dim trimmed As String = genPara.Trim()
                                    If String.IsNullOrWhiteSpace(trimmed) Then Continue For

                                    Dim cloned As System.Xml.XmlNode = AssembleCloneAndImport(
                                        primaryTemplate.XmlDoc, templates, donorTmpl, donorIdx)
                                    If cloned IsNot Nothing Then
                                        AssembleReplaceAllText(cloned, trimmed, primaryTemplate.NsMgr)
                                        outputParagraphNodes.Add(cloned)
                                    End If
                                Next
                                sectionTextCollector.AppendLine(generatedText)
                            End If

                        Case "merge"
                            Dim instruction As String = If(entry("instruction")?.ToString(), "")
                            Dim sources As Newtonsoft.Json.Linq.JArray = CType(entry("sources"), Newtonsoft.Json.Linq.JArray)

                            ' Collect source texts
                            Dim mergeSourceText As New StringBuilder()
                            If sources IsNot Nothing Then
                                For Each src As Newtonsoft.Json.Linq.JObject In sources
                                    Dim mTmpl As String = If(src("template")?.ToString(), primaryTemplateName)
                                    Dim mFrom As Integer = If(src("fromIndex")?.ToObject(Of Integer)(), 0)
                                    Dim mTo As Integer = If(src("toIndex")?.ToObject(Of Integer)(), mFrom)
                                    If templates.ContainsKey(mTmpl) Then
                                        mergeSourceText.AppendLine($"--- From template: {mTmpl} ---")
                                        For mIdx As Integer = mFrom To Math.Min(mTo, templates(mTmpl).Paragraphs.Count - 1)
                                            mergeSourceText.AppendLine(templates(mTmpl).Paragraphs(mIdx).FullText)
                                        Next
                                    End If
                                Next
                            End If

                            ' Style donor
                            Dim mergeStyleFrom As Newtonsoft.Json.Linq.JObject = CType(entry("styleFrom"), Newtonsoft.Json.Linq.JObject)
                            Dim mergeDonorTmpl As String = If(mergeStyleFrom?("template")?.ToString(), primaryTemplateName)
                            Dim mergeDonorIdx As Integer = If(mergeStyleFrom?("paragraphIndex")?.ToObject(Of Integer)(), 0)

                            ' LLM call to merge
                            Dim mergedText As String = Await AssembleExecuteSection(
                                mergeSourceText.ToString(), instruction, userInstruction, contextSummary.ToString(), useSecondAPI)

                            If Not String.IsNullOrWhiteSpace(mergedText) Then
                                Dim mergeParagraphs As String() = mergedText.Split(
                                    New String() {vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)

                                For Each mergePara In mergeParagraphs
                                    Dim trimmed As String = mergePara.Trim()
                                    If String.IsNullOrWhiteSpace(trimmed) Then Continue For

                                    Dim cloned As System.Xml.XmlNode = AssembleCloneAndImport(
                                        primaryTemplate.XmlDoc, templates, mergeDonorTmpl, mergeDonorIdx)
                                    If cloned IsNot Nothing Then
                                        AssembleReplaceAllText(cloned, trimmed, primaryTemplate.NsMgr)
                                        outputParagraphNodes.Add(cloned)
                                    End If
                                Next
                                sectionTextCollector.AppendLine(mergedText)
                            End If

                    End Select
                Next

                ' ── Update rolling context summary ──
                Dim sectionText As String = sectionTextCollector.ToString().Trim()
                If sectionText.Length > 0 Then
                    ' Summarize if the section is long, otherwise include as-is
                    If sectionText.Length > INI_AssembleMaxContextSummaryChars Then
                        Dim summaryPrompt As String = SP_Assemble_Summarize.
                            Replace("{SectionHeading}", sectionHeading).
                            Replace("{SectionText}", sectionText.Substring(0, Math.Min(sectionText.Length, INI_AssembleExecMaxChars)))
                        Dim summaryResult As String = Await SharedMethods.LLM(
                            _context, summaryPrompt, "", "", "", 0, useSecondAPI, True)
                        If Not String.IsNullOrWhiteSpace(summaryResult) Then
                            contextSummary.AppendLine($"[Section: {sectionHeading}] {summaryResult.Trim()}")
                        End If
                    Else
                        contextSummary.AppendLine($"[Section: {sectionHeading}] {sectionText}")
                    End If

                    ' Trim context summary if it exceeds budget
                    If contextSummary.Length > INI_AssembleMaxContextSummaryChars * 2 Then
                        Dim fullSummary As String = contextSummary.ToString()
                        contextSummary.Clear()
                        contextSummary.Append(fullSummary.Substring(fullSummary.Length - INI_AssembleMaxContextSummaryChars * 2))
                    End If
                End If

                ProgressBarModule.GlobalProgressValue = sectionIdx + 1
            Next

            If outputParagraphNodes.Count = 0 Then
                ShowCustomMessageBox("No output paragraphs were produced. Assembly aborted.")
                Return
            End If

            ' ═══ Step 8: Build Output DOCX ═══

            If ProgressBarModule.CancelOperation Then Return
            ProgressBarModule.GlobalProgressLabel = "Assembling output document..."

            Dim primaryPath As String = templatePaths(primaryIndex)
            Dim outputDir As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            Dim outputName As String = Path.GetFileNameWithoutExtension(primaryPath) & AssembleOutputSuffix & ".docx"
            Dim outputPath As String = Path.Combine(outputDir, outputName)

            ' Copy primary template as base (preserves styles.xml, numbering, theme, headers, footers, etc.)
            File.Copy(primaryPath, outputPath, overwrite:=True)

            ' Open the output DOCX and replace w:body content
            Dim outputTempDir As String = Path.Combine(Path.GetTempPath(), $"{AN2}_asm_{Guid.NewGuid():N}")
            Try
                ZipFile.ExtractToDirectory(outputPath, outputTempDir)

                Dim documentXmlPath As String = Path.Combine(outputTempDir, "word", "document.xml")
                If Not File.Exists(documentXmlPath) Then
                    ShowCustomMessageBox("Invalid DOCX structure — document.xml not found in output template.")
                    Return
                End If

                Dim outputXmlDoc As New System.Xml.XmlDocument()
                outputXmlDoc.PreserveWhitespace = True
                outputXmlDoc.Load(documentXmlPath)

                Dim outputNsMgr As New System.Xml.XmlNamespaceManager(outputXmlDoc.NameTable)
                outputNsMgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

                Dim body As System.Xml.XmlNode = outputXmlDoc.SelectSingleNode("//w:body", outputNsMgr)
                If body Is Nothing Then
                    ShowCustomMessageBox("Invalid DOCX structure — w:body not found.")
                    Return
                End If

                ' Preserve w:sectPr (page layout, headers/footers) — always last child
                Dim sectPr As System.Xml.XmlNode = body.SelectSingleNode("w:sectPr", outputNsMgr)

                ' Remove all existing content nodes (w:p, w:tbl, w:sdt) but keep sectPr
                Dim nodesToRemove As New List(Of System.Xml.XmlNode)()
                For Each child As System.Xml.XmlNode In body.ChildNodes
                    If child IsNot sectPr Then nodesToRemove.Add(child)
                Next
                For Each n In nodesToRemove
                    body.RemoveChild(n)
                Next

                ' Insert all output content nodes
                For Each pNode In outputParagraphNodes
                    ' Import into output document DOM
                    Dim imported As System.Xml.XmlNode = outputXmlDoc.ImportNode(pNode, deep:=True)
                    If sectPr IsNot Nothing Then
                        body.InsertBefore(imported, sectPr)
                    Else
                        body.AppendChild(imported)
                    End If
                Next

                ' Save and repack
                outputXmlDoc.Save(documentXmlPath)
                File.Delete(outputPath)
                ZipFile.CreateFromDirectory(outputTempDir, outputPath, CompressionLevel.Optimal, False)

            Finally
                If Directory.Exists(outputTempDir) Then
                    Try : Directory.Delete(outputTempDir, recursive:=True) : Catch : End Try
                End If
            End Try

            ' ═══ Step 9: Post-process via Word Interop ═══
            ' Open the assembled DOCX in Word and clean up empty paragraphs.
            ' Word's own rendering engine correctly identifies which paragraphs
            ' are truly empty (including those with style-based numbering,
            ' inherited bullets, etc.) — no XML guesswork needed.

            If Not ProgressBarModule.CancelOperation Then
                ProgressBarModule.GlobalProgressLabel = "Cleaning up empty paragraphs..."
                AssembleCleanupViaWord(outputPath)
            End If

            ProgressBarModule.GlobalProgressValue = ProgressBarModule.GlobalProgressMax
            ProgressBarModule.CancelOperation = True  ' Close progress bar before showing dialog

            ' ═══ Step 10: Offer Compare Version ═══

            Dim compareChoice As Integer = ShowCustomYesNoBox(
                $"Document assembled successfully:" & vbCrLf & vbCrLf &
                $"  {outputName}" & vbCrLf & vbCrLf &
                $"Output: {outputPath}" & vbCrLf & vbCrLf &
                "Would you like to create a compare version showing changes from the primary template?",
                "No, done",
                "Yes, create compare",
                AN & " Assembly")

            If compareChoice = 2 Then
                ' Create compare document: primary template (original) vs assembled output (revised)
                Dim comparePath As String = Path.Combine(
                    outputDir,
                    Path.GetFileNameWithoutExtension(primaryPath) & AssembleOutputSuffix & "_compare.docx")

                Dim compareSuccess As Boolean = AssembleCreateCompareDocument(primaryPath, outputPath, comparePath)

                If compareSuccess Then
                    ShowCustomMessageBox(
                        $"Compare document created:" & vbCrLf & vbCrLf &
                        $"  {Path.GetFileName(comparePath)}" & vbCrLf & vbCrLf &
                        $"Output: {comparePath}",
                        AN & " Assembly")
                Else
                    ShowCustomMessageBox("Failed to create compare document.", AN & " Assembly")
                End If
            End If

        Catch ex As Exception
            ShowCustomMessageBox($"Error during document assembly: {ex.Message}", AN & " Assembly Error")
        Finally
            ' Cleanup all template temp directories
            For Each kvp In templates
                If kvp.Value.TempDir IsNot Nothing AndAlso Directory.Exists(kvp.Value.TempDir) Then
                    Try : Directory.Delete(kvp.Value.TempDir, recursive:=True) : Catch : End Try
                End If
            Next
            ProgressBarModule.CancelOperation = True
        End Try
    End Function


    ' ═════════════════════════════════════════════════════════════════════════
    ' Template Loading
    ' ═════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Loads a DOCX template: extracts to temp dir, parses document.xml,
    ''' builds a flat content-node index (w:p + w:tbl + w:sdt in document order)
    ''' and extracts readable text for each node.
    ''' </summary>
    Private Function AssembleLoadTemplate(docxPath As String, name As String) As AssemblyTemplateInfo
        Dim tempDir As String = Path.Combine(Path.GetTempPath(), $"{AN2}_asmtpl_{Guid.NewGuid():N}")

        Try
            ZipFile.ExtractToDirectory(docxPath, tempDir)

            Dim documentXmlPath As String = Path.Combine(tempDir, "word", "document.xml")
            If Not File.Exists(documentXmlPath) Then Return Nothing

            Dim xmlDoc As New System.Xml.XmlDocument()
            xmlDoc.PreserveWhitespace = True
            xmlDoc.Load(documentXmlPath)

            Dim nsMgr As New System.Xml.XmlNamespaceManager(xmlDoc.NameTable)
            nsMgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

            Dim body As System.Xml.XmlNode = xmlDoc.SelectSingleNode("//w:body", nsMgr)
            If body Is Nothing Then Return Nothing

            Dim contentNodes As New List(Of System.Xml.XmlNode)()
            Dim contentNodeTypes As New List(Of String)()
            Dim paragraphs As New List(Of TranslateParagraphInfo)()
            Dim contentIndex As Integer = 0

            For Each child As System.Xml.XmlNode In body.ChildNodes
                If child.NodeType <> System.Xml.XmlNodeType.Element Then Continue For

                Select Case child.LocalName
                    Case "p"
                        contentNodes.Add(child)
                        contentNodeTypes.Add("p")
                        Dim paraInfo As TranslateParagraphInfo = AssembleExtractParagraphText(child, nsMgr, contentIndex)
                        paragraphs.Add(paraInfo)
                        contentIndex += 1

                    Case "tbl"
                        contentNodes.Add(child)
                        contentNodeTypes.Add("tbl")
                        Dim tableText As String = AssembleExtractTableText(child, nsMgr)
                        paragraphs.Add(New TranslateParagraphInfo() With {
                            .Index = contentIndex,
                            .TextRuns = New List(Of TranslateTextRunInfo)(),
                            .FullText = tableText,
                            .MarkerText = Nothing,
                            .TranslatedText = Nothing,
                            .IsEmpty = String.IsNullOrWhiteSpace(tableText)
                        })
                        contentIndex += 1

                    Case "sdt"
                        contentNodes.Add(child)
                        contentNodeTypes.Add("sdt")
                        Dim sdtText As String = AssembleExtractAllWtText(child, nsMgr)
                        paragraphs.Add(New TranslateParagraphInfo() With {
                            .Index = contentIndex,
                            .TextRuns = New List(Of TranslateTextRunInfo)(),
                            .FullText = sdtText,
                            .MarkerText = Nothing,
                            .TranslatedText = Nothing,
                            .IsEmpty = String.IsNullOrWhiteSpace(sdtText)
                        })
                        contentIndex += 1

                    Case "sectPr"
                        Continue For

                    Case Else
                        contentNodes.Add(child)
                        contentNodeTypes.Add(child.LocalName)
                        paragraphs.Add(New TranslateParagraphInfo() With {
                            .Index = contentIndex,
                            .TextRuns = New List(Of TranslateTextRunInfo)(),
                            .FullText = "",
                            .MarkerText = Nothing,
                            .TranslatedText = Nothing,
                            .IsEmpty = True
                        })
                        contentIndex += 1
                End Select
            Next

            Return New AssemblyTemplateInfo() With {
                .Name = name,
                .FilePath = docxPath,
                .TempDir = tempDir,
                .XmlDoc = xmlDoc,
                .NsMgr = nsMgr,
                .ContentNodes = contentNodes,
                .ContentNodeTypes = contentNodeTypes,
                .Paragraphs = paragraphs
            }

        Catch ex As Exception
            Debug.WriteLine($"AssembleLoadTemplate error: {ex.Message}")
            If Directory.Exists(tempDir) Then
                Try : Directory.Delete(tempDir, recursive:=True) : Catch : End Try
            End If
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Extracts text from a single w:p node for assembly indexing.
    ''' </summary>
    Private Shared Function AssembleExtractParagraphText(
        paraNode As System.Xml.XmlNode,
        nsMgr As System.Xml.XmlNamespaceManager,
        index As Integer) As TranslateParagraphInfo

        Dim paraInfo As New TranslateParagraphInfo() With {
            .Index = index,
            .TextRuns = New List(Of TranslateTextRunInfo)(),
            .MarkerText = Nothing,
            .TranslatedText = Nothing
        }

        Dim textNodes As System.Xml.XmlNodeList = paraNode.SelectNodes(".//w:t", nsMgr)
        Dim fullTextBuilder As New StringBuilder()

        For Each textNode As System.Xml.XmlNode In textNodes
            Dim text As String = textNode.InnerText
            paraInfo.TextRuns.Add(New TranslateTextRunInfo() With {
                .TextNode = textNode,
                .OriginalText = text,
                .HasNoteReferenceBefore = False
            })
            fullTextBuilder.Append(text)
        Next

        paraInfo.FullText = fullTextBuilder.ToString()
        paraInfo.IsEmpty = String.IsNullOrWhiteSpace(paraInfo.FullText)

        Return paraInfo
    End Function

    ''' <summary>
    ''' Extracts readable text from a w:tbl node as tab-separated values (TSV).
    ''' One row per line, cells separated by TAB characters.
    ''' This is the canonical format sent to and received from the LLM.
    ''' </summary>
    Private Shared Function AssembleExtractTableText(
        tblNode As System.Xml.XmlNode,
        nsMgr As System.Xml.XmlNamespaceManager) As String

        Dim sb As New StringBuilder()

        Dim rows As System.Xml.XmlNodeList = tblNode.SelectNodes("w:tr", nsMgr)
        For Each row As System.Xml.XmlNode In rows
            Dim cells As System.Xml.XmlNodeList = row.SelectNodes("w:tc", nsMgr)
            Dim cellTexts As New List(Of String)()

            For Each cell As System.Xml.XmlNode In cells
                Dim cellText As String = AssembleExtractAllWtText(cell, nsMgr)
                cellTexts.Add(If(cellText, ""))
            Next

            sb.AppendLine(String.Join(vbTab, cellTexts))
        Next

        Return sb.ToString().TrimEnd()
    End Function

    ''' <summary>
    ''' Extracts all w:t text from any XML subtree, concatenated with spaces
    ''' between paragraphs.
    ''' </summary>
    Private Shared Function AssembleExtractAllWtText(
        node As System.Xml.XmlNode,
        nsMgr As System.Xml.XmlNamespaceManager) As String

        Dim paras As System.Xml.XmlNodeList = node.SelectNodes(".//w:p", nsMgr)
        Dim parts As New List(Of String)()

        For Each p As System.Xml.XmlNode In paras
            Dim textNodes As System.Xml.XmlNodeList = p.SelectNodes(".//w:t", nsMgr)
            Dim pText As New StringBuilder()
            For Each t As System.Xml.XmlNode In textNodes
                pText.Append(t.InnerText)
            Next
            Dim s As String = pText.ToString().Trim()
            If s.Length > 0 Then parts.Add(s)
        Next

        Return String.Join(" ", parts)
    End Function



    ' ═════════════════════════════════════════════════════════════════════════
    ' LLM Execution Calls
    ' ═════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Executes a single section-level LLM call for adapt/generate/merge actions on paragraphs.
    ''' </summary>
    Private Async Function AssembleExecuteSection(
        sourceText As String,
        instruction As String,
        facts As String,
        contextSummaryText As String,
        useSecondAPI As Boolean) As System.Threading.Tasks.Task(Of String)

        Dim promptBuilder As New StringBuilder()

        If Not String.IsNullOrWhiteSpace(sourceText) Then
            promptBuilder.AppendLine("<SOURCETEXT>")
            If sourceText.Length > INI_AssembleExecMaxChars Then
                promptBuilder.AppendLine(sourceText.Substring(0, INI_AssembleExecMaxChars))
                promptBuilder.AppendLine("... (truncated)")
            Else
                promptBuilder.AppendLine(sourceText)
            End If
            promptBuilder.AppendLine("</SOURCETEXT>")
            promptBuilder.AppendLine()
        End If

        promptBuilder.AppendLine("<INSTRUCTION>")
        promptBuilder.AppendLine(instruction)
        promptBuilder.AppendLine("</INSTRUCTION>")
        promptBuilder.AppendLine()

        If Not String.IsNullOrWhiteSpace(facts) Then
            promptBuilder.AppendLine("<FACTS>")
            promptBuilder.AppendLine(facts)
            promptBuilder.AppendLine("</FACTS>")
            promptBuilder.AppendLine()
        End If

        If Not String.IsNullOrWhiteSpace(contextSummaryText) Then
            promptBuilder.AppendLine("<CONTEXT>")
            promptBuilder.AppendLine(contextSummaryText)
            promptBuilder.AppendLine("</CONTEXT>")
        End If

        Dim systemPrompt As String = InterpolateAtRuntime(SP_Assemble_Execute)

        Dim response As String = Await SharedMethods.LLM(
            _context, systemPrompt, promptBuilder.ToString(),
            "", "", 0, useSecondAPI, True)

        If String.IsNullOrWhiteSpace(response) Then Return Nothing

        ' Strip any stray XML tags the LLM may have wrapped around the response
        response = response.Trim()
        response = Regex.Replace(response, "^</?SOURCETEXT>", "", RegexOptions.IgnoreCase).Trim()
        response = Regex.Replace(response, "</?SOURCETEXT>$", "", RegexOptions.IgnoreCase).Trim()
        response = Regex.Replace(response, "^</?INSTRUCTION>", "", RegexOptions.IgnoreCase).Trim()
        response = Regex.Replace(response, "</?INSTRUCTION>$", "", RegexOptions.IgnoreCase).Trim()

        Return response
    End Function

    ''' <summary>
    ''' Executes a table-specific LLM call.  Sends the table content in TSV format
    ''' with explicit framing so the LLM returns it in the same structure.
    ''' </summary>
    Private Async Function AssembleExecuteTableSection(
        tableText As String,
        instruction As String,
        facts As String,
        contextSummaryText As String,
        useSecondAPI As Boolean) As System.Threading.Tasks.Task(Of String)

        Dim promptBuilder As New StringBuilder()

        ' Send the table with an explicit TSV framing so the LLM knows the format
        promptBuilder.AppendLine("<SOURCETEXT format=""tsv"">")
        promptBuilder.AppendLine("The following is a table in TSV (tab-separated values) format.")
        promptBuilder.AppendLine("Each line is one row. Cells within a row are separated by TAB characters.")
        promptBuilder.AppendLine("---TABLE START---")
        If tableText.Length > INI_AssembleExecMaxChars Then
            promptBuilder.AppendLine(tableText.Substring(0, INI_AssembleExecMaxChars))
            promptBuilder.AppendLine("... (truncated)")
        Else
            promptBuilder.AppendLine(tableText)
        End If
        promptBuilder.AppendLine("---TABLE END---")
        promptBuilder.AppendLine("</SOURCETEXT>")
        promptBuilder.AppendLine()

        promptBuilder.AppendLine("<INSTRUCTION>")
        promptBuilder.AppendLine(instruction)
        promptBuilder.AppendLine()
        promptBuilder.AppendLine("IMPORTANT: Return the adapted table in EXACTLY the same TSV format: one row per line, cells separated by TAB characters.")
        promptBuilder.AppendLine("Keep the same number of columns in every row.")
        promptBuilder.AppendLine("To remove a row, replace its entire line with: [REMOVE]")
        promptBuilder.AppendLine("Do NOT add any commentary, markdown, or code fences — return ONLY the TSV rows.")
        promptBuilder.AppendLine("</INSTRUCTION>")
        promptBuilder.AppendLine()

        If Not String.IsNullOrWhiteSpace(facts) Then
            promptBuilder.AppendLine("<FACTS>")
            promptBuilder.AppendLine(facts)
            promptBuilder.AppendLine("</FACTS>")
            promptBuilder.AppendLine()
        End If

        If Not String.IsNullOrWhiteSpace(contextSummaryText) Then
            promptBuilder.AppendLine("<CONTEXT>")
            promptBuilder.AppendLine(contextSummaryText)
            promptBuilder.AppendLine("</CONTEXT>")
        End If

        Dim systemPrompt As String = InterpolateAtRuntime(SP_Assemble_Execute)

        Dim response As String = Await SharedMethods.LLM(
            _context, systemPrompt, promptBuilder.ToString(),
            "", "", 0, useSecondAPI, True)

        If String.IsNullOrWhiteSpace(response) Then Return Nothing

        ' Strip any wrapping the LLM may have added
        response = response.Trim()

        ' Remove markdown code fences if present
        If response.StartsWith("```") Then
            Dim firstNewline As Integer = response.IndexOf(vbLf)
            If firstNewline > 0 Then response = response.Substring(firstNewline + 1)
            If response.EndsWith("```") Then response = response.Substring(0, response.Length - 3).Trim()
        End If

        ' Remove our framing markers if the LLM echoed them
        response = Regex.Replace(response, "^---TABLE START---\s*" & vbLf & "?", "", RegexOptions.IgnoreCase).Trim()
        response = Regex.Replace(response, vbLf & "?\s*---TABLE END---$", "", RegexOptions.IgnoreCase).Trim()
        response = Regex.Replace(response, "^</?SOURCETEXT[^>]*>", "", RegexOptions.IgnoreCase).Trim()
        response = Regex.Replace(response, "</?SOURCETEXT[^>]*>$", "", RegexOptions.IgnoreCase).Trim()

        Debug.WriteLine($"[AssembleTable] LLM response ({response.Length} chars):")
        Debug.WriteLine(response)

        Return response
    End Function


    ' ═════════════════════════════════════════════════════════════════════════
    ' OpenXML Operations
    ' ═════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Deep-clones a content node (w:p, w:tbl, or w:sdt) from a template and
    ''' imports it into a target document's DOM.
    ''' </summary>
    Private Shared Function AssembleCloneAndImport(
        targetDoc As System.Xml.XmlDocument,
        templates As Dictionary(Of String, AssemblyTemplateInfo),
        templateName As String,
        contentIndex As Integer) As System.Xml.XmlNode

        If Not templates.ContainsKey(templateName) Then Return Nothing
        Dim tmpl As AssemblyTemplateInfo = templates(templateName)
        If contentIndex < 0 OrElse contentIndex >= tmpl.ContentNodes.Count Then Return Nothing

        Dim sourceNode As System.Xml.XmlNode = tmpl.ContentNodes(contentIndex)
        Return targetDoc.ImportNode(sourceNode, deep:=True)
    End Function

    ''' <summary>
    ''' Replaces all w:t text in a paragraph with new text, distributing across existing runs.
    ''' </summary>
    Private Sub AssembleReplaceTextPreservingRuns(
        paragraphNode As System.Xml.XmlNode,
        newText As String,
        nsMgr As System.Xml.XmlNamespaceManager)

        Dim textNodes As System.Xml.XmlNodeList = paragraphNode.SelectNodes(".//w:t", nsMgr)
        If textNodes.Count = 0 Then Return

        If textNodes.Count = 1 Then
            SetTextNodeWithSpacePreserve(textNodes(0), newText)
            Return
        End If

        Dim tempPara As New TranslateParagraphInfo() With {
            .TextRuns = New List(Of TranslateTextRunInfo)(),
            .FullText = ""
        }

        Dim fullTextBuilder As New StringBuilder()
        For Each tn As System.Xml.XmlNode In textNodes
            tempPara.TextRuns.Add(New TranslateTextRunInfo() With {
                .TextNode = tn,
                .OriginalText = tn.InnerText,
                .HasNoteReferenceBefore = False
            })
            fullTextBuilder.Append(tn.InnerText)
        Next
        tempPara.FullText = fullTextBuilder.ToString()
        tempPara.IsEmpty = String.IsNullOrWhiteSpace(tempPara.FullText)

        DistributeProportional(tempPara, newText)
    End Sub

    ''' <summary>
    ''' Replaces ALL text in a paragraph by putting everything into the first w:t node
    ''' and clearing subsequent ones.
    ''' </summary>
    Private Shared Sub AssembleReplaceAllText(
        paragraphNode As System.Xml.XmlNode,
        newText As String,
        nsMgr As System.Xml.XmlNamespaceManager)

        Dim textNodes As System.Xml.XmlNodeList = paragraphNode.SelectNodes(".//w:t", nsMgr)
        If textNodes.Count = 0 Then Return

        textNodes(0).InnerText = If(newText, "")

        If newText IsNot Nothing AndAlso (newText.StartsWith(" ") OrElse newText.EndsWith(" ")) Then
            Dim spaceAttr = textNodes(0).Attributes("xml:space")
            If spaceAttr Is Nothing Then
                spaceAttr = textNodes(0).OwnerDocument.CreateAttribute(
                    "xml", "space", "http://www.w3.org/XML/1998/namespace")
                textNodes(0).Attributes.Append(spaceAttr)
            End If
            spaceAttr.Value = "preserve"
        End If

        For i As Integer = 1 To textNodes.Count - 1
            textNodes(i).InnerText = ""
        Next
    End Sub

    ''' <summary>
    ''' Replaces cell text in a cloned w:tbl node using adapted TSV text from the LLM.
    ''' 
    ''' The LLM returns text in the same row/column format as it received:
    ''' one row per line, cells separated by TAB characters.
    ''' 
    ''' Row-level removal:  A line that is exactly [REMOVE] removes the entire w:tr.
    ''' Cell-level removal: A cell whose text is [REMOVE] is cleared to empty.
    '''                     If ALL cells in a row are [REMOVE], the entire row is removed.
    ''' 
    ''' If the LLM response has fewer non-removed rows than the table, remaining
    ''' XML rows keep their original text.  Extra response rows are ignored.
    ''' </summary>
    Private Sub AssembleReplaceTableText(
        tblNode As System.Xml.XmlNode,
        adaptedText As String,
        nsMgr As System.Xml.XmlNamespaceManager)

        If String.IsNullOrWhiteSpace(adaptedText) Then Return

        ' Parse LLM response into rows
        Dim adaptedRows As String() = adaptedText.Split(New String() {vbCrLf, vbLf}, StringSplitOptions.None)

        ' Remove trailing empty lines (LLM may add a trailing newline)
        Dim rowList As New List(Of String)(adaptedRows)
        While rowList.Count > 0 AndAlso String.IsNullOrWhiteSpace(rowList(rowList.Count - 1))
            rowList.RemoveAt(rowList.Count - 1)
        End While

        Dim xmlRows As System.Xml.XmlNodeList = tblNode.SelectNodes("w:tr", nsMgr)
        Dim xmlRowList As New List(Of System.Xml.XmlNode)()
        For Each r As System.Xml.XmlNode In xmlRows
            xmlRowList.Add(r)
        Next

        Debug.WriteLine($"[AssembleTable] XML rows: {xmlRowList.Count}, LLM rows: {rowList.Count}")

        ' ── Phase 1: Process each row — update text or mark for removal ──
        Dim rowsToRemove As New List(Of System.Xml.XmlNode)()
        Dim rowIdx As Integer = 0

        For Each xmlRow In xmlRowList
            If rowIdx >= rowList.Count Then Exit For

            Dim rowText As String = rowList(rowIdx).Trim()

            ' Check for row-level removal token (entire line is [REMOVE])
            If rowText.Equals(TableRowRemoveToken, StringComparison.OrdinalIgnoreCase) Then
                rowsToRemove.Add(xmlRow)
                Debug.WriteLine($"[AssembleTable] Row {rowIdx}: MARKED FOR REMOVAL (row-level)")
                rowIdx += 1
                Continue For
            End If

            ' Split adapted row into cells
            Dim adaptedCells As String() = rowList(rowIdx).Split(New String() {vbTab}, StringSplitOptions.None)
            Dim xmlCells As System.Xml.XmlNodeList = xmlRow.SelectNodes("w:tc", nsMgr)

            Debug.WriteLine($"[AssembleTable] Row {rowIdx}: {adaptedCells.Length} adapted cells, {xmlCells.Count} XML cells")

            ' Track whether all cells in this row are [REMOVE]
            Dim allCellsRemoved As Boolean = True
            Dim anyCellRemoved As Boolean = False

            Dim cellIdx As Integer = 0
            For Each xmlCell As System.Xml.XmlNode In xmlCells
                If cellIdx >= adaptedCells.Length Then
                    ' More XML cells than adapted cells — these are untouched,
                    ' so the row is not entirely removed
                    allCellsRemoved = False
                    Exit For
                End If

                Dim newCellText As String = adaptedCells(cellIdx)
                Dim trimmedCellText As String = newCellText.Trim()

                ' Check for cell-level [REMOVE] token
                If trimmedCellText.Equals(TableRowRemoveToken, StringComparison.OrdinalIgnoreCase) Then
                    ' Clear this cell to empty
                    anyCellRemoved = True
                    Dim cellParas As System.Xml.XmlNodeList = xmlCell.SelectNodes("w:p", nsMgr)
                    For Each cp As System.Xml.XmlNode In cellParas
                        AssembleReplaceAllText(cp, "", nsMgr)
                    Next
                    Debug.WriteLine($"[AssembleTable] Row {rowIdx}, Cell {cellIdx}: CELL CLEARED ([REMOVE] in cell)")
                Else
                    allCellsRemoved = False

                    ' Normal cell text replacement
                    Dim cellParas As System.Xml.XmlNodeList = xmlCell.SelectNodes("w:p", nsMgr)
                    If cellParas.Count > 0 Then
                        AssembleReplaceCellParagraphText(cellParas(0), newCellText, nsMgr)
                        For pIdx As Integer = 1 To cellParas.Count - 1
                            AssembleReplaceAllText(cellParas(pIdx), "", nsMgr)
                        Next
                    End If
                End If

                cellIdx += 1
            Next

            ' If all cells in this row are [REMOVE], remove the entire row
            If allCellsRemoved AndAlso anyCellRemoved Then
                rowsToRemove.Add(xmlRow)
                Debug.WriteLine($"[AssembleTable] Row {rowIdx}: ALL CELLS [REMOVE] → removing entire row")
            End If

            rowIdx += 1
        Next

        ' ── Phase 2: Remove marked rows from the XML ──
        For Each rowNode In rowsToRemove
            Debug.WriteLine($"[AssembleTable] Removing w:tr node")
            tblNode.RemoveChild(rowNode)
        Next

        Debug.WriteLine($"[AssembleTable] Final table: {tblNode.SelectNodes("w:tr", nsMgr).Count} rows")
    End Sub

    ''' <summary>
    ''' Replaces text in a single table cell paragraph (w:p inside w:tc),
    ''' preserving the first run's formatting (w:rPr) and distributing new text
    ''' into existing runs via proportional distribution.
    ''' </summary>
    Private Sub AssembleReplaceCellParagraphText(
        cellPara As System.Xml.XmlNode,
        newText As String,
        nsMgr As System.Xml.XmlNamespaceManager)

        Dim textNodes As System.Xml.XmlNodeList = cellPara.SelectNodes(".//w:t", nsMgr)

        If textNodes.Count = 0 Then
            ' No w:t nodes — need to create one inside an existing or new w:r
            Dim firstRun As System.Xml.XmlNode = cellPara.SelectSingleNode("w:r", nsMgr)
            If firstRun Is Nothing Then
                ' Create a minimal w:r with a w:t
                Dim doc As System.Xml.XmlDocument = cellPara.OwnerDocument
                firstRun = doc.CreateElement("w", "r", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
                cellPara.AppendChild(firstRun)
            End If

            Dim newT As System.Xml.XmlNode = cellPara.OwnerDocument.CreateElement(
                "w", "t", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
            newT.InnerText = If(newText, "")
            firstRun.AppendChild(newT)

            If Not String.IsNullOrWhiteSpace(newText) Then
                Dim spaceAttr = cellPara.OwnerDocument.CreateAttribute(
                    "xml", "space", "http://www.w3.org/XML/1998/namespace")
                spaceAttr.Value = "preserve"
                newT.Attributes.Append(spaceAttr)
            End If
            Return
        End If

        If textNodes.Count = 1 Then
            SetTextNodeWithSpacePreserve(textNodes(0), If(newText, ""))
            Return
        End If

        ' Multiple runs: use proportional distribution
        Dim tempPara As New TranslateParagraphInfo() With {
            .TextRuns = New List(Of TranslateTextRunInfo)(),
            .FullText = ""
        }
        Dim fullTextBuilder As New StringBuilder()
        For Each tn As System.Xml.XmlNode In textNodes
            tempPara.TextRuns.Add(New TranslateTextRunInfo() With {
                .TextNode = tn,
                .OriginalText = tn.InnerText,
                .HasNoteReferenceBefore = False
            })
            fullTextBuilder.Append(tn.InnerText)
        Next
        tempPara.FullText = fullTextBuilder.ToString()
        tempPara.IsEmpty = String.IsNullOrWhiteSpace(tempPara.FullText)

        DistributeProportional(tempPara, If(newText, ""))
    End Sub


    ' ═════════════════════════════════════════════════════════════════════════
    ' UI Helpers
    ' ═════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Shows a DragDropForm for selecting a template file.
    ''' If no pre-configured assembly template paths are available, goes straight
    ''' to the drag &amp; drop dialog.  Otherwise offers a selection list first.
    ''' </summary>
    Private Function AssembleSelectTemplatesByBrowse(label As String) As List(Of String)
        Dim result As New List(Of String)()

        ' ── Try pre-configured template paths first ──
        Dim preConfigured As String = AssembleSelectFromConfiguredPaths(label)
        If preConfigured IsNot Nothing Then
            If preConfigured = "" Then
                ' User cancelled
                Return result
            End If
            result.Add(preConfigured)
            Return result
        End If

        ' ── Fall back to drag & drop browse ──
        Globals.ThisAddIn.DragDropFormLabel = label
        Globals.ThisAddIn.DragDropFormFilter = "Word Document (*.docx)|*.docx"

        Try
            Using frm As New DragDropForm(DragDropMode.FileOnly)
                Dim __safeDialogOwner1308 As System.Windows.Forms.IWin32Window = SharedLibrary.SharedLibrary.SharedMethods.ResolveSameThreadDialogOwner()
                If If(__safeDialogOwner1308 IsNot Nothing, frm.ShowDialog(__safeDialogOwner1308), frm.ShowDialog()) = DialogResult.OK Then
                    Dim selectedPath As String = frm.SelectedFilePath
                    If File.Exists(selectedPath) AndAlso
                       Path.GetExtension(selectedPath).Equals(".docx", StringComparison.OrdinalIgnoreCase) Then
                        result.Add(selectedPath)
                    End If
                End If
            End Using
        Finally
            Globals.ThisAddIn.DragDropFormLabel = ""
            Globals.ThisAddIn.DragDropFormFilter = ""
        End Try

        Return result
    End Function

    ''' <summary>
    ''' Checks INI_AssemblePath and INI_AssemblePathLocal for .docx template files.
    ''' If templates are found, shows a SelectValue dialog listing local files first
    ''' (marked "(local)"), then global files, plus a "Browse…" option at the end.
    ''' </summary>
    ''' <returns>
    ''' Full path to the selected template, empty string if user cancelled,
    ''' or Nothing if no configured paths exist / user chose "Browse…".
    ''' </returns>
    Private Function AssembleSelectFromConfiguredPaths(label As String) As String

        ' Expand and normalize configured paths
        Dim assemblePath As String = Nothing
        Dim assemblePathLocal As String = Nothing

        Try
            assemblePath = ExpandEnvironmentVariables(INI_AssemblePath)
            If Not String.IsNullOrWhiteSpace(assemblePath) AndAlso Not assemblePath.EndsWith("\", StringComparison.Ordinal) Then
                assemblePath &= "\"
            End If
        Catch
        End Try

        Try
            assemblePathLocal = ExpandEnvironmentVariables(INI_AssemblePathLocal)
            If Not String.IsNullOrWhiteSpace(assemblePathLocal) AndAlso Not assemblePathLocal.EndsWith("\", StringComparison.Ordinal) Then
                assemblePathLocal &= "\"
            End If
        Catch
        End Try

        Dim hasGlobal As Boolean = Not String.IsNullOrWhiteSpace(assemblePath) AndAlso Directory.Exists(assemblePath)
        Dim hasLocal As Boolean = Not String.IsNullOrWhiteSpace(assemblePathLocal) AndAlso Directory.Exists(assemblePathLocal)

        If Not hasGlobal AndAlso Not hasLocal Then Return Nothing

        ' Collect .docx files from both paths
        ' Value → full path mapping (SelectionItem.Value is 1-based index)
        Dim indexToPath As New Dictionary(Of Integer, String)()
        Dim items As New List(Of SelectionItem)()
        Dim idx As Integer = 1

        ' Local templates first (higher priority)
        If hasLocal Then
            Try
                Dim localFiles = Directory.GetFiles(assemblePathLocal, "*.docx", SearchOption.TopDirectoryOnly)
                Array.Sort(localFiles, StringComparer.OrdinalIgnoreCase)
                For Each f In localFiles
                    ' Skip hidden/system files (e.g. ~$lock files)
                    Dim attr As FileAttributes = File.GetAttributes(f)
                    If (attr And FileAttributes.Hidden) = FileAttributes.Hidden Then Continue For
                    If (attr And FileAttributes.System) = FileAttributes.System Then Continue For

                    Dim disp As String = Path.GetFileNameWithoutExtension(f) & " (local)"
                    items.Add(New SelectionItem(disp, idx))
                    indexToPath(idx) = f
                    idx += 1
                Next
            Catch ex As Exception
                Debug.WriteLine($"[Assemble] Error reading local path: {ex.Message}")
            End Try
        End If

        ' Global templates
        If hasGlobal Then
            Try
                Dim globalFiles = Directory.GetFiles(assemblePath, "*.docx", SearchOption.TopDirectoryOnly)
                Array.Sort(globalFiles, StringComparer.OrdinalIgnoreCase)
                For Each f In globalFiles
                    ' Skip hidden/system files (e.g. ~$lock files)
                    Dim attr As FileAttributes = File.GetAttributes(f)
                    If (attr And FileAttributes.Hidden) = FileAttributes.Hidden Then Continue For
                    If (attr And FileAttributes.System) = FileAttributes.System Then Continue For

                    Dim disp As String = Path.GetFileNameWithoutExtension(f)
                    items.Add(New SelectionItem(disp, idx))
                    indexToPath(idx) = f
                    idx += 1
                Next
            Catch ex As Exception
                Debug.WriteLine($"[Assemble] Error reading global path: {ex.Message}")
            End Try
        End If

        If items.Count = 0 Then Return Nothing

        ' Add a "Browse…" option at the end
        Const BrowseValue As Integer = -1
        items.Add(New SelectionItem("Browse for a different file…", BrowseValue))

        ' Show selection dialog
        Dim chosen As Integer = SelectValue(
            items, items(0).Value,
            label,
            AN & " Assembly — Select Template")

        ' Handle result
        If chosen = 0 Then
            Return ""  ' Cancelled (Escape / closed)
        End If

        If chosen = BrowseValue Then
            Return Nothing  ' Fall through to drag & drop
        End If

        Dim chosenPath As String = Nothing
        If indexToPath.TryGetValue(chosen, chosenPath) AndAlso
           Not String.IsNullOrWhiteSpace(chosenPath) AndAlso
           File.Exists(chosenPath) Then
            Return chosenPath
        End If

        Return Nothing
    End Function


    ' ═════════════════════════════════════════════════════════════════════════
    ' Heading Detection
    ' ═════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Determines the heading level from a Word style name.
    ''' Returns 0 if the style is not a heading.
    ''' </summary>
    Private Shared Function AssembleGetHeadingLevel(styleName As String) As Integer
        If String.IsNullOrWhiteSpace(styleName) Then Return 0

        Dim m As Match = Regex.Match(styleName, "(?:Heading|berschrift)\s*(\d+)", RegexOptions.IgnoreCase)
        If m.Success Then
            Dim level As Integer
            If Integer.TryParse(m.Groups(1).Value, level) Then Return level
        End If

        If styleName.Equals("Title", StringComparison.OrdinalIgnoreCase) Then Return 1

        Return 0
    End Function


    ' ═════════════════════════════════════════════════════════════════════════
    ' Cleaning up Word & Compare
    ' ═════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Opens the assembled DOCX in Word and removes empty paragraphs
    ''' (including orphan bullets/numbers and blank table cell lines).
    ''' Uses Word's own layout engine — no XML namespace issues.
    ''' </summary>
    Private Sub AssembleCleanupViaWord(docxPath As String)
        Dim wordApp As Word.Application = Nothing
        Dim doc As Word.Document = Nothing
        Dim prevScreenUpdating As Boolean = True
        Dim prevAlerts As Word.WdAlertLevel = Word.WdAlertLevel.wdAlertsNone

        Try
            wordApp = Globals.ThisAddIn.Application
            prevScreenUpdating = wordApp.ScreenUpdating
            prevAlerts = wordApp.DisplayAlerts
            wordApp.ScreenUpdating = False
            wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone

            doc = wordApp.Documents.Open(
                docxPath, ReadOnly:=False, Visible:=False, AddToRecentFiles:=False)

            Dim removedCount As Integer = 0

            ' Pass 1: Remove empty paragraphs in the document body
            ' Iterate backwards to avoid index shifting
            For i As Integer = doc.Paragraphs.Count To 1 Step -1
                Try
                    Dim para As Word.Paragraph = doc.Paragraphs(i)
                    Dim text As String = para.Range.Text

                    ' Word paragraph text always ends with vbCr — a truly empty
                    ' paragraph is just vbCr (length 1) or whitespace + vbCr
                    If text IsNot Nothing AndAlso text.Trim(vbCr(0), " "c, vbTab(0)).Length = 0 Then
                        para.Range.Delete()
                        removedCount += 1
                    End If
                Catch
                    ' Skip paragraphs that can't be accessed (e.g., in headers/footers)
                End Try
            Next

            ' Pass 2: Clean up empty paragraphs inside tables
            For Each tbl As Word.Table In doc.Tables
                Try
                    For Each cell As Word.Cell In tbl.Range.Cells
                        Try
                            ' Cell text ends with end-of-cell marker (Chr(13) & Chr(7))
                            Dim cellText As String = cell.Range.Text
                            If cellText IsNot Nothing Then
                                ' Strip end-of-cell marker and check what's left
                                cellText = cellText.TrimEnd(Chr(13), Chr(7), " "c, vbTab(0))
                            End If

                            ' If cell has multiple paragraphs, remove the empty ones
                            ' but always keep at least one (Word requires it)
                            Dim cellParaCount As Integer = cell.Range.Paragraphs.Count
                            If cellParaCount > 1 Then
                                For p As Integer = cellParaCount To 2 Step -1
                                    Try
                                        Dim cellPara As Word.Paragraph = cell.Range.Paragraphs(p)
                                        Dim paraText As String = cellPara.Range.Text
                                        If paraText IsNot Nothing AndAlso
                                           paraText.Trim(vbCr(0), Chr(7), " "c, vbTab(0)).Length = 0 Then
                                            cellPara.Range.Delete()
                                            removedCount += 1
                                        End If
                                    Catch
                                    End Try
                                Next
                            End If
                        Catch
                        End Try
                    Next
                Catch
                End Try
            Next

            If removedCount > 0 Then
                Debug.WriteLine($"[Assemble] Cleanup: removed {removedCount} empty paragraph(s) via Word")
            End If

            doc.Save()
            doc.Close(WdSaveOptions.wdDoNotSaveChanges)
            doc = Nothing

        Catch ex As Exception
            Debug.WriteLine($"[Assemble] Cleanup via Word failed: {ex.Message}")
        Finally
            If doc IsNot Nothing Then
                Try : doc.Close(WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
            End If
            If wordApp IsNot Nothing Then
                Try
                    wordApp.DisplayAlerts = prevAlerts
                    wordApp.ScreenUpdating = prevScreenUpdating
                Catch
                End Try
            End If
        End Try
    End Sub

    ''' <summary>
    ''' Creates a Word compare document showing differences between the
    ''' primary template (original) and the assembled output (revised).
    ''' Uses Word's built-in CompareDocuments feature with tracked changes.
    ''' </summary>
    ''' <param name="originalPath">Path to the primary template.</param>
    ''' <param name="revisedPath">Path to the assembled output document.</param>
    ''' <param name="comparePath">Path where the compare document will be saved.</param>
    ''' <returns>True if successful, False otherwise.</returns>
    Private Function AssembleCreateCompareDocument(
        originalPath As String,
        revisedPath As String,
        comparePath As String) As Boolean

        Dim wordApp As Word.Application = Nothing
        Dim originalDoc As Word.Document = Nothing
        Dim revisedDoc As Word.Document = Nothing
        Dim compareDoc As Word.Document = Nothing

        Try
            wordApp = Globals.ThisAddIn.Application
            Dim wasScreenUpdating As Boolean = wordApp.ScreenUpdating
            wordApp.ScreenUpdating = False

            ' Open both documents
            originalDoc = wordApp.Documents.Open(originalPath, ReadOnly:=True, Visible:=False, AddToRecentFiles:=False)
            revisedDoc = wordApp.Documents.Open(revisedPath, ReadOnly:=True, Visible:=False, AddToRecentFiles:=False)

            ' Create comparison document using Word's built-in compare feature
            compareDoc = wordApp.CompareDocuments(
                OriginalDocument:=originalDoc,
                RevisedDocument:=revisedDoc,
                Destination:=WdCompareDestination.wdCompareDestinationNew,
                Granularity:=WdGranularity.wdGranularityWordLevel,
                CompareFormatting:=True,
                CompareCaseChanges:=True,
                CompareWhitespace:=True,
                CompareTables:=True,
                CompareHeaders:=True,
                CompareFootnotes:=True,
                CompareTextboxes:=True,
                CompareFields:=True,
                CompareComments:=True,
                RevisedAuthor:=GetMarkupAuthorOrCurrent(wordApp),
                IgnoreAllComparisonWarnings:=True
            )

            ' Save the compare document
            compareDoc.SaveAs2(comparePath, WdSaveFormat.wdFormatXMLDocument)

            ' Close documents
            compareDoc.Close(WdSaveOptions.wdDoNotSaveChanges)
            compareDoc = Nothing

            revisedDoc.Close(WdSaveOptions.wdDoNotSaveChanges)
            revisedDoc = Nothing

            originalDoc.Close(WdSaveOptions.wdDoNotSaveChanges)
            originalDoc = Nothing

            wordApp.ScreenUpdating = wasScreenUpdating
            Return True

        Catch ex As Exception
            Debug.WriteLine($"[Assemble] CreateCompareDocument error: {ex.Message}")
            Return False
        Finally
            If compareDoc IsNot Nothing Then
                Try : compareDoc.Close(WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
            End If
            If revisedDoc IsNot Nothing Then
                Try : revisedDoc.Close(WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
            End If
            If originalDoc IsNot Nothing Then
                Try : originalDoc.Close(WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
            End If
            If wordApp IsNot Nothing Then
                Try : wordApp.ScreenUpdating = True : Catch : End Try
            End If
        End Try
    End Function

End Class
