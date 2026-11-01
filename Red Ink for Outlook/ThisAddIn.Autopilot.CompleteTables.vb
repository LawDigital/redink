' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved.
' For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.CompleteTables.vb
' Purpose:
'   Headless AutoPilot / Local Chat port of the Word add-in's table-completion
'   engine. Completes empty or incomplete Word tables, body placeholders, and
'   form fields by extracting DOCX OpenXML structure to JSON, asking the LLM
'   for structured updates, then patching the DOCX while preserving formatting.
'
' Notes:
'  - This is intentionally a mechanical port of `Red Ink for Word\ThisAddIn.CompleteTables.vb`.
'  - UI-specific Word add-in behavior was replaced with AutoPilot / Local Chat
'    logging, cancellation, and output-file handling.
'  - Core extraction, prompt, parsing, and apply logic is preserved.
'  - Supports `.docx` and `.doc` input. Legacy `.doc` is converted to `.docx`
'    in a temporary file before processing.
'  - Produces `<name>_completed.docx` and, by default, a compare document.
'
' KEY ARCHITECTURE
'   • Surface: ThisAddIn, AutoPilotTableCompletionEngine (host partial/UI).
'   • Entry points/helpers: ExecuteCompleteWordTablesTool, ProcessDocumentAsync,
'     ConvertLegacyDocToDocxAsync.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.IO.Compression
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Xml
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    Private Async Function ExecuteCompleteWordTablesTool(
            toolCall As ToolCall,
            context As ToolExecutionContext,
            ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName,
            .Timestamp = DateTime.UtcNow
        }

        Try
            Dim instruction As String = GetArgString(toolCall.Arguments, "instruction")
            If String.IsNullOrWhiteSpace(instruction) Then
                response.Success = False
                response.Response = "Missing required parameter: instruction"
                Return response
            End If

            Dim targetNames As List(Of String) = GetArgStringArray(toolCall.Arguments, "attachment_names")
            Dim useSecondApi As Boolean = GetArgBool(toolCall.Arguments, "use_second_api", False)
            Dim createCompareDocument As Boolean = GetArgBool(toolCall.Arguments, "create_compare_document", True)

            Dim toProcess As New List(Of AutoPilotAttachmentInfo)()

            If targetNames IsNot Nothing AndAlso targetNames.Count > 0 Then
                For Each name As String In targetNames
                    Dim att As AutoPilotAttachmentInfo = FindAttachment(name)
                    If att Is Nothing Then Continue For
                    If att.IsOverSizeLimit Then Continue For
                    If String.IsNullOrWhiteSpace(att.TempFilePath) OrElse Not File.Exists(att.TempFilePath) Then Continue For

                    Dim ext As String = Path.GetExtension(att.TempFilePath).ToLowerInvariant()
                    If ext = ".docx" OrElse ext = ".doc" Then
                        toProcess.Add(att)
                    End If
                Next
            ElseIf _apCurrentAttachments IsNot Nothing Then
                toProcess = _apCurrentAttachments.
                    Where(Function(a) a IsNot Nothing AndAlso
                                      Not a.IsOverSizeLimit AndAlso
                                      Not String.IsNullOrWhiteSpace(a.TempFilePath) AndAlso
                                      File.Exists(a.TempFilePath) AndAlso
                                      (a.Extension = ".docx" OrElse a.Extension = ".doc")).
                    ToList()
            End If

            If toProcess.Count = 0 Then
                response.Success = False
                response.Response = "No processable Word document attachments found."
                Return response
            End If

            Dim engine As New AutoPilotTableCompletionEngine(Me, context, ct)
            Dim resultMessages As New List(Of String)()

            For Each att As AutoPilotAttachmentInfo In toProcess
                ct.ThrowIfCancellationRequested()

                Dim baseName As String = Path.GetFileNameWithoutExtension(att.OriginalFileName)
                Dim outputName As String = baseName & "_completed.docx"
                Dim outputPath As String = Path.Combine(_apCurrentTempDir, outputName)

                Dim counter As Integer = 1
                While File.Exists(outputPath)
                    outputName = $"{baseName}_completed_{counter}.docx"
                    outputPath = Path.Combine(_apCurrentTempDir, outputName)
                    counter += 1
                End While

                context.Log($"Completing Word tables: {att.OriginalFileName}")
                ApDashboardLog($"📋 Completing tables: {att.OriginalFileName}", "step")

                Dim execResult = Await engine.ProcessDocumentAsync(
                    att.TempFilePath,
                    outputPath,
                    instruction.Trim(),
                    useSecondApi,
                    createCompareDocument)

                If execResult.Success Then
                    Dim registrationTarget As AutoPilotAttachmentInfo =
                        If(att.IsToolOutput,
                           _apCurrentAttachments.FirstOrDefault(
                               Function(a) a.OutputFiles IsNot Nothing AndAlso
                                             a.OutputFiles.Any(Function(p) Path.GetFileName(p).Equals(att.OriginalFileName, StringComparison.OrdinalIgnoreCase))),
                           att)

                    If registrationTarget Is Nothing AndAlso
                       _apCurrentAttachments IsNot Nothing AndAlso
                       _apCurrentAttachments.Count > 0 Then
                        registrationTarget = _apCurrentAttachments(0)
                    End If

                    If registrationTarget IsNot Nothing Then
                        registrationTarget.OutputFiles.Add(execResult.CompletedPath)

                        If Not String.IsNullOrWhiteSpace(execResult.ComparePath) AndAlso
                           File.Exists(execResult.ComparePath) Then
                            registrationTarget.OutputFiles.Add(execResult.ComparePath)
                        End If
                    End If

                    resultMessages.Add(execResult.Message)
                    ApDashboardLog($"✓ Completed tables: {Path.GetFileName(execResult.CompletedPath)}", "info")
                Else
                    resultMessages.Add(execResult.Message)
                    ApDashboardLog($"⚠ Table completion failed: {att.OriginalFileName}", "warn")
                End If
            Next

            response.Success = resultMessages.Any(Function(m) m.StartsWith("✓", StringComparison.Ordinal))
            response.Response = String.Join(vbCrLf, resultMessages)
            Return response

        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation was cancelled."
            response.Response = response.ErrorMessage
            Return response
        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error completing Word tables: {ex.Message}"
            Return response
        End Try
    End Function

    Private NotInheritable Class AutoPilotTableCompletionEngine

        Private ReadOnly _owner As ThisAddIn
        Private ReadOnly _context As ToolExecutionContext
        Private ReadOnly _ct As CancellationToken

        Private Const W14Ns As String = "http://schemas.microsoft.com/office/word/2010/wordml"
        Private Const WNs As String = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        Private Const TableCompleteDebugEnabled As Boolean = False
        Private Const TableCompleteDebugIncludeXml As Boolean = False
        Private Const TableCompleteDebugIncludePackageDump As Boolean = False
        Private Const TableCompleteDebugLogCellApply As Boolean = False
        Private Const TableCompleteDebugVerifySavedDocx As Boolean = False
        Private Const TableCompleteDebugVerifySavedDocxWithWordInterop As Boolean = False
        Private Const TableCompleteDebugLogFileName As String = "RI_TableCompletion_Debug.log"

        Private Shared ReadOnly TableCompleteDebugSyncRoot As New Object()

        Private Shared ReadOnly PlaceholderPattern As New Regex(
            "^\s*(\[(?!\s*\d+\s*\]$)[^\]\r\n]+\]|\{[^}\r\n]*\}|<[^>\r\n]*>|_{3,}|\.{3,}|…+|TBD|TBC|N/?A|TODO|FIXME|XXX|INSERT|ENTER)\s*$",
            RegexOptions.IgnoreCase Or RegexOptions.Compiled Or RegexOptions.CultureInvariant)

#Region "Data Classes"

        Private Class TableCellInfo
            Public Property Row As Integer
            Public Property Col As Integer
            Public Property Text As String
            Public Property GridSpan As Integer
            Public Property VMerge As String
            Public Property TextNodes As List(Of XmlNode)
            Public Property IsHeader As Boolean
            Public Property IsPlaceholder As Boolean
            Public Property WidthPct As Double
            Public Property XmlNode As XmlNode
            Public Property IsWritable As Boolean
            Public Property ProtectionReason As String
            Public Property FormFields As List(Of FormFieldInfo)
        End Class

        Private Class TableRowInfo
            Public Property RowIndex As Integer
            Public Property Cells As List(Of TableCellInfo)
            Public Property XmlNode As XmlNode
        End Class

        Private Class CompletionBatchesResult
            Public Property TableUpdates As List(Of TableUpdateResponse)
            Public Property BodySectionUpdates As List(Of BodySectionUpdateResponse)
        End Class

        Private Class TableInfo
            Public Property TableIndex As Integer
            Public Property Rows As List(Of TableRowInfo)
            Public Property XmlNode As XmlNode
            Public Property ColumnCount As Integer
            Public Property ContextBefore As String
            Public Property ColumnWidths As List(Of Double)
        End Class

        Private Class TableCellUpdate
            Public Property Row As Integer
            Public Property Col As Integer
            Public Property Text As String
            Public Property Action As String
        End Class

        Private Class TableUpdateResponse
            Public Property TableIndex As Integer
            Public Property Updates As List(Of TableCellUpdate)
            Public Property NewRows As List(Of List(Of String))
            Public Property InsertAfterRow As Integer
            Public Property FieldUpdates As List(Of FormFieldUpdate)
        End Class

        Private Enum FormFieldType
            Dropdown
            Checkbox
            TextInput
            LegacyDropdown
            LegacyCheckbox
        End Enum

        Private Class FormFieldOptionInfo
            Public Property DisplayText As String
            Public Property StoredValue As String
        End Class

        Private Class FormFieldInfo
            Public Property FieldId As String
            Public Property FieldType As FormFieldType
            Public Property Name As String
            Public Property CurrentValue As String
            Public Property CurrentOptionValue As String
            Public Property Options As List(Of FormFieldOptionInfo)
            Public Property XmlNode As XmlNode
            Public Property IsWritable As Boolean
            Public Property ProtectionReason As String
        End Class

        Private Class FormFieldUpdate
            Public Property Row As Integer
            Public Property Col As Integer
            Public Property FieldId As String
            Public Property Value As String
            Public Property OptionText As String
            Public Property OptionValue As String
            Public Property OptionIndex As Integer
        End Class

        Private Class BodySectionInfo
            Public Property SectionIndex As Integer
            Public Property HeadingText As String
            Public Property Paragraphs As List(Of BodyParagraphInfo)
            Public Property FormFields As List(Of FormFieldInfo)
        End Class

        Private Class BodyParagraphInfo
            Public Property Index As Integer
            Public Property Text As String
            Public Property IsPlaceholder As Boolean
            Public Property IsEmpty As Boolean
            Public Property TextNodes As List(Of XmlNode)
            Public Property XmlNode As XmlNode
        End Class

        Private Class BodySectionUpdateResponse
            Public Property SectionIndex As Integer
            Public Property ParagraphUpdates As List(Of BodyParagraphUpdate)
            Public Property FieldUpdates As List(Of BodyFieldUpdate)
        End Class

        Private Class BodyParagraphUpdate
            Public Property Index As Integer
            Public Property Text As String
            Public Property Action As String
        End Class

        Private Class BodyFieldUpdate
            Public Property FieldId As String
            Public Property Value As String
            Public Property OptionText As String
            Public Property OptionValue As String
            Public Property OptionIndex As Integer
        End Class

        Friend Class CompletionExecutionResult
            Public Property Success As Boolean
            Public Property CompletedPath As String
            Public Property ComparePath As String
            Public Property Message As String
        End Class

#End Region

        Public Sub New(owner As ThisAddIn, context As ToolExecutionContext, ct As CancellationToken)
            _owner = owner
            _context = context
            _ct = ct
        End Sub

#Region "Execution Pipeline"

        Public Async Function ProcessDocumentAsync(inputPath As String,
                                                   outputPath As String,
                                                   userInstructions As String,
                                                   useSecondApi As Boolean,
                                                   createCompareDocument As Boolean) As Task(Of CompletionExecutionResult)

            Dim result As New CompletionExecutionResult() With {
                .Success = False,
                .CompletedPath = outputPath,
                .ComparePath = Nothing,
                .Message = ""
            }

            Dim tempDocxPath As String = Nothing

            Try
                If Path.GetExtension(inputPath).ToLowerInvariant() = ".doc" Then
                    tempDocxPath = Path.Combine(Path.GetTempPath(), $"RI_tblconv_{Guid.NewGuid():N}.docx")
                    Dim converted As Boolean = Await ConvertLegacyDocToDocxAsync(inputPath, tempDocxPath)
                    If Not converted OrElse Not File.Exists(tempDocxPath) Then
                        result.Message = $"✗ {Path.GetFileName(inputPath)}: Failed to convert legacy .doc file to .docx."
                        Return result
                    End If
                Else
                    tempDocxPath = inputPath
                End If

                File.Copy(tempDocxPath, outputPath, overwrite:=True)

                StartTableCompletionDebugSession(inputPath, outputPath, userInstructions)

                Dim success As Boolean = Await ProcessDocxTables(outputPath, userInstructions, useSecondApi)
                If Not success Then
                    result.Message = $"✗ {Path.GetFileName(inputPath)}: Table completion failed."
                    Return result
                End If

                If createCompareDocument Then
                    Dim compareSourcePath As String = tempDocxPath
                    If String.IsNullOrWhiteSpace(compareSourcePath) OrElse Not File.Exists(compareSourcePath) Then
                        compareSourcePath = inputPath
                    End If

                    Dim comparePath As String = GetTableCompletionCompareFilePath(outputPath)
                    Dim compareCreated As Boolean =
                        _owner.CreateWordCompareDocumentForAutoPilot(compareSourcePath, outputPath, comparePath)

                    If compareCreated AndAlso File.Exists(comparePath) Then
                        result.ComparePath = comparePath
                    End If
                End If

                result.Success = True
                result.Message =
                    If(Not String.IsNullOrWhiteSpace(result.ComparePath),
                       $"✓ {Path.GetFileName(inputPath)}: Tables completed successfully. Output: {Path.GetFileName(outputPath)} + compare document.",
                       $"✓ {Path.GetFileName(inputPath)}: Tables completed successfully. Output: {Path.GetFileName(outputPath)}")
                Return result

            Catch ex As OperationCanceledException
                Throw
            Catch ex As Exception
                WriteTableCompletionDebug("ProcessDocumentAsync exception", ex.ToString())
                result.Message = $"✗ {Path.GetFileName(inputPath)}: {ex.Message}"
                Return result
            Finally
                If tempDocxPath IsNot Nothing AndAlso
                   Not tempDocxPath.Equals(inputPath, StringComparison.OrdinalIgnoreCase) AndAlso
                   File.Exists(tempDocxPath) Then
                    Try : File.Delete(tempDocxPath) : Catch : End Try
                End If
            End Try
        End Function

        Private Async Function ConvertLegacyDocToDocxAsync(inputPath As String, outputPath As String) As Task(Of Boolean)
            Return Await _owner.SwitchToUi(
                Function()
                    Dim wordApp As Microsoft.Office.Interop.Word.Application = Nothing
                    Dim doc As Microsoft.Office.Interop.Word.Document = Nothing
                    Dim weCreated As Boolean = False

                    Try
                        Try
                            wordApp = DirectCast(GetObject(, "Word.Application"), Microsoft.Office.Interop.Word.Application)
                        Catch
                            wordApp = New Microsoft.Office.Interop.Word.Application()
                            wordApp.Visible = False
                            weCreated = True
                        End Try

                        wordApp.ScreenUpdating = False
                        doc = wordApp.Documents.Open(inputPath, ReadOnly:=True, Visible:=False, AddToRecentFiles:=False)
                        doc.SaveAs2(outputPath, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument)
                        Return True

                    Catch ex As Exception
                        Debug.WriteLine("ConvertLegacyDocToDocxAsync error: " & ex.Message)
                        Return False

                    Finally
                        If doc IsNot Nothing Then
                            Try : doc.Close(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
                            Try : Runtime.InteropServices.Marshal.FinalReleaseComObject(doc) : Catch : End Try
                        End If

                        If wordApp IsNot Nothing Then
                            Try : wordApp.ScreenUpdating = True : Catch : End Try
                            If weCreated Then
                                Try : wordApp.Quit(False) : Catch : End Try
                            End If
                            Try : Runtime.InteropServices.Marshal.FinalReleaseComObject(wordApp) : Catch : End Try
                        End If
                    End Try
                End Function)
        End Function

        Private Async Function ProcessDocxTables(docxPath As String,
                                                 userInstructions As String,
                                                 Optional useSecondAPI As Boolean = False) As Task(Of Boolean)

            Dim tempDir As String = Path.Combine(Path.GetTempPath(), $"RI_tbl_{Guid.NewGuid():N}")

            Try
                _ct.ThrowIfCancellationRequested()

                ZipFile.ExtractToDirectory(docxPath, tempDir)

                If TableCompleteDebugIncludePackageDump Then DumpDocxPackageDiagnostics(docxPath)

                Dim documentXmlPath As String = Path.Combine(tempDir, "word", "document.xml")
                If Not File.Exists(documentXmlPath) Then
                    LogWarn("Invalid DOCX structure - document.xml not found.")
                    Return False
                End If

                Dim xmlDoc As New XmlDocument()
                xmlDoc.PreserveWhitespace = True
                xmlDoc.Load(documentXmlPath)

                Dim nsMgr As New XmlNamespaceManager(xmlDoc.NameTable)
                nsMgr.AddNamespace("w", WNs)
                nsMgr.AddNamespace("w14", W14Ns)

                Dim documentContext As String = ExtractDocumentBodyContext(xmlDoc, nsMgr)
                Dim tables As List(Of TableInfo) = ExtractTablesFromXml(xmlDoc, nsMgr)
                Dim bodySections As List(Of BodySectionInfo) = ExtractBodySectionsFromXml(xmlDoc, nsMgr)

                WriteTableCompletionDebug("Extracted fields", BuildFormFieldDebugReport(tables))

                If tables.Count = 0 AndAlso bodySections.Count = 0 Then
                    LogWarn("No tables or body sections found in the document.")
                    Return False
                End If

                Dim placeholderCount As Integer = 0
                Dim formFieldCount As Integer = 0

                For Each tbl In tables
                    For Each row In tbl.Rows
                        For Each cell In row.Cells
                            If cell.IsPlaceholder Then placeholderCount += 1
                            If cell.FormFields IsNot Nothing Then formFieldCount += cell.FormFields.Count
                        Next
                    Next
                Next

                For Each section In bodySections
                    If section.Paragraphs IsNot Nothing Then
                        For Each para In section.Paragraphs
                            If para.IsPlaceholder OrElse para.IsEmpty Then placeholderCount += 1
                        Next
                    End If

                    If section.FormFields IsNot Nothing Then
                        formFieldCount += section.FormFields.Count
                    End If
                Next

                Dim confirmMsg As String = $"Found {tables.Count} table(s)"
                If bodySections.Count > 0 Then confirmMsg &= $", {bodySections.Count} body section(s)"
                If placeholderCount > 0 Then confirmMsg &= $", {placeholderCount} placeholder/empty target(s)"
                If formFieldCount > 0 Then confirmMsg &= $", {formFieldCount} form field(s)"
                confirmMsg &= ". Continuing with AI completion."

                LogInfo(confirmMsg)

                Dim completionResult As CompletionBatchesResult =
                    Await ProcessCompletionBatches(tables, bodySections, userInstructions, documentContext, useSecondAPI)

                If completionResult Is Nothing Then Return False

                ApplyTableUpdates(tables, completionResult.TableUpdates, nsMgr)
                ApplyBodySectionUpdates(bodySections, completionResult.BodySectionUpdates, nsMgr)

                xmlDoc.Save(documentXmlPath)
                File.Delete(docxPath)
                ZipFile.CreateFromDirectory(tempDir, docxPath, CompressionLevel.Optimal, False)

                If TableCompleteDebugVerifySavedDocx Then
                    VerifySavedDocxCellUpdates(docxPath, completionResult.TableUpdates)
                End If

                If TableCompleteDebugVerifySavedDocxWithWordInterop Then
                    VerifySavedDocxWithWordInterop(docxPath)
                End If

                WriteTableCompletionDebug("Completion finished", $"Saved output document: {docxPath}")
                Return True

            Catch ex As OperationCanceledException
                Throw
            Catch ex As Exception
                WriteTableCompletionDebug("ProcessDocxTables exception", ex.ToString())
                Throw
            Finally
                If Directory.Exists(tempDir) Then
                    Try : Directory.Delete(tempDir, recursive:=True) : Catch : End Try
                End If
            End Try
        End Function

        Private Function GetTableCompletionCompareFilePath(completedPath As String) As String
            Dim dir As String = Path.GetDirectoryName(completedPath)
            Dim nameWithoutExt As String = Path.GetFileNameWithoutExtension(completedPath)
            Dim ext As String = Path.GetExtension(completedPath)

            If nameWithoutExt.EndsWith("_completed", StringComparison.OrdinalIgnoreCase) Then
                nameWithoutExt = nameWithoutExt.Substring(0, nameWithoutExt.Length - "_completed".Length) & "_completed_compare"
            Else
                nameWithoutExt &= "_compare"
            End If

            Return Path.Combine(dir, nameWithoutExt & ext)
        End Function

#End Region

#Region "Extraction"

        Private Function ExtractVisibleParagraphTextExcludingSdts(paragraphNode As XmlNode,
                                                                 nsMgr As XmlNamespaceManager,
                                                                 ByRef textNodes As List(Of XmlNode)) As String
            textNodes = New List(Of XmlNode)()
            If paragraphNode Is Nothing Then Return ""

            Dim sb As New StringBuilder()
            Dim allTextNodes As XmlNodeList = paragraphNode.SelectNodes(".//w:t", nsMgr)

            For Each tNode As XmlNode In allTextNodes
                Dim ancestor As XmlNode = tNode.ParentNode
                Dim insideSdt As Boolean = False

                While ancestor IsNot Nothing AndAlso ancestor IsNot paragraphNode
                    If ancestor.LocalName = "sdt" Then
                        insideSdt = True
                        Exit While
                    End If
                    ancestor = ancestor.ParentNode
                End While

                If insideSdt Then Continue For

                textNodes.Add(tNode)
                sb.Append(tNode.InnerText)
            Next

            Return sb.ToString()
        End Function

        Private Function ExtractBestContextParagraphTextFromNode(node As XmlNode,
                                                                 nsMgr As XmlNamespaceManager) As String
            If node Is Nothing Then Return ""

            Dim paragraphNodes As New List(Of XmlNode)()

            If node.LocalName = "p" Then
                paragraphNodes.Add(node)
            Else
                For Each pNode As XmlNode In node.SelectNodes(".//w:p[not(ancestor::w:tbl)]", nsMgr)
                    paragraphNodes.Add(pNode)
                Next
            End If

            For Each pNode As XmlNode In paragraphNodes
                Dim textNodes As List(Of XmlNode) = Nothing
                Dim text As String = NormalizeTextForPlaceholderDetection(ExtractVisibleParagraphTextExcludingSdts(pNode, nsMgr, textNodes))
                If Not String.IsNullOrWhiteSpace(text) AndAlso Not IsPlaceholderText(text) Then
                    Return text
                End If
            Next

            Return ""
        End Function

        Private Function ExtractCandidateBodyParagraphsFromNode(node As XmlNode,
                                                                nsMgr As XmlNamespaceManager) As List(Of BodyParagraphInfo)
            Dim results As New List(Of BodyParagraphInfo)()
            If node Is Nothing Then Return results

            Dim paragraphNodes As New List(Of XmlNode)()

            If node.LocalName = "p" Then
                paragraphNodes.Add(node)
            Else
                For Each pNode As XmlNode In node.SelectNodes(".//w:p[not(ancestor::w:tbl)]", nsMgr)
                    paragraphNodes.Add(pNode)
                Next
            End If

            Dim paragraphIndex As Integer = 0

            For Each pNode As XmlNode In paragraphNodes
                Dim textNodes As List(Of XmlNode) = Nothing
                Dim rawText As String = ExtractVisibleParagraphTextExcludingSdts(pNode, nsMgr, textNodes)
                Dim normalizedText As String = NormalizeTextForPlaceholderDetection(rawText)

                results.Add(New BodyParagraphInfo() With {
                    .Index = paragraphIndex,
                    .Text = rawText,
                    .IsPlaceholder = False,
                    .IsEmpty = String.IsNullOrWhiteSpace(normalizedText),
                    .TextNodes = If(textNodes, New List(Of XmlNode)()),
                    .XmlNode = pNode
                })

                paragraphIndex += 1
            Next

            Return results
        End Function

        Private Function ExtractFormFieldsFromNode(containerNode As XmlNode,
                                                   nsMgr As XmlNamespaceManager,
                                                   isWritable As Boolean,
                                                   protectionReason As String) As List(Of FormFieldInfo)
            Dim results As New List(Of FormFieldInfo)()
            If containerNode Is Nothing Then Return results

            Dim fieldIndex As Integer = 0
            Dim sdtNodes As New List(Of XmlNode)()

            Dim owningSdtNode As XmlNode = GetOwningSdtNodeForCell(containerNode)
            If owningSdtNode IsNot Nothing Then
                sdtNodes.Add(owningSdtNode)
            End If

            Dim descendantSdtNodes As XmlNodeList = containerNode.SelectNodes(".//w:sdt", nsMgr)
            For Each sdtNode As XmlNode In descendantSdtNodes
                If Not sdtNodes.Any(Function(existingNode) existingNode Is sdtNode) Then
                    sdtNodes.Add(sdtNode)
                End If
            Next

            For Each sdtNode As XmlNode In sdtNodes
                Dim sdtPr As XmlNode = sdtNode.SelectSingleNode("w:sdtPr", nsMgr)
                If sdtPr Is Nothing Then Continue For

                Dim fieldProtectionReason As String = ""
                Dim ff As New FormFieldInfo() With {
                    .FieldId = $"f{fieldIndex}",
                    .XmlNode = sdtNode,
                    .Options = New List(Of FormFieldOptionInfo)(),
                    .Name = "",
                    .CurrentValue = "",
                    .CurrentOptionValue = "",
                    .IsWritable = IsFormFieldWritable(sdtNode, nsMgr, fieldProtectionReason),
                    .ProtectionReason = fieldProtectionReason
                }

                Dim tagNode As XmlNode = sdtPr.SelectSingleNode("w:tag/@w:val", nsMgr)
                Dim aliasNode As XmlNode = sdtPr.SelectSingleNode("w:alias/@w:val", nsMgr)
                ff.Name = If(tagNode?.Value, If(aliasNode?.Value, ""))

                Dim showingPlcHdr As Boolean = (sdtPr.SelectSingleNode("w:showingPlcHdr", nsMgr) IsNot Nothing)

                Dim dropDownNode As XmlNode = sdtPr.SelectSingleNode("w:dropDownList", nsMgr)
                Dim comboBoxNode As XmlNode = sdtPr.SelectSingleNode("w:comboBox", nsMgr)
                Dim checkboxNode As XmlNode = sdtPr.SelectSingleNode("w14:checkbox", nsMgr)
                Dim textNode As XmlNode = sdtPr.SelectSingleNode("w:text", nsMgr)
                Dim dateNode As XmlNode = sdtPr.SelectSingleNode("w:date", nsMgr)

                If dropDownNode IsNot Nothing OrElse comboBoxNode IsNot Nothing Then
                    ff.FieldType = FormFieldType.Dropdown

                    Dim listParent As XmlNode = If(dropDownNode, comboBoxNode)
                    Dim listItems As XmlNodeList = listParent.SelectNodes("w:listItem", nsMgr)
                    For Each li As XmlNode In listItems
                        Dim displayAttr As XmlNode = li.Attributes("w:displayText")
                        Dim valueAttr As XmlNode = li.Attributes("w:value")
                        AddFormFieldOption(ff.Options, If(displayAttr?.Value, ""), If(valueAttr?.Value, ""))
                    Next

                    Dim rawDisplayValue As String = GetSdtContentText(sdtNode, nsMgr)
                    Dim lastValueAttr As XmlNode = listParent.Attributes("w:lastValue")
                    Dim rawStoredValue As String = If(lastValueAttr?.Value, "")

                    If showingPlcHdr Then
                        ff.CurrentValue = ""
                        ff.CurrentOptionValue = ""
                    Else
                        Dim currentOption As FormFieldOptionInfo = ResolveCurrentDropdownOption(ff, rawDisplayValue, rawStoredValue)
                        If currentOption IsNot Nothing Then
                            ff.CurrentValue = currentOption.DisplayText
                            ff.CurrentOptionValue = currentOption.StoredValue
                        Else
                            ff.CurrentValue = If(rawDisplayValue, "")
                            ff.CurrentOptionValue = If(rawStoredValue, "")
                        End If
                    End If

                ElseIf checkboxNode IsNot Nothing Then
                    ff.FieldType = FormFieldType.Checkbox
                    AddFormFieldOption(ff.Options, "true", "true")
                    AddFormFieldOption(ff.Options, "false", "false")

                    Dim checkedNode As XmlNode = checkboxNode.SelectSingleNode("w14:checked/@w14:val", nsMgr)
                    If checkedNode IsNot Nothing Then
                        ff.CurrentValue = If(checkedNode.Value = "1" OrElse checkedNode.Value.ToLowerInvariant() = "true", "true", "false")
                    Else
                        ff.CurrentValue = "false"
                    End If
                    ff.CurrentOptionValue = ff.CurrentValue

                ElseIf textNode IsNot Nothing OrElse dateNode IsNot Nothing Then
                    ff.FieldType = FormFieldType.TextInput
                    Dim rawValue As String = GetSdtContentText(sdtNode, nsMgr)
                    ff.CurrentValue = If(showingPlcHdr, "", rawValue)
                    ff.CurrentOptionValue = ff.CurrentValue

                Else
                    ff.FieldType = FormFieldType.TextInput
                    Dim rawValue As String = GetSdtContentText(sdtNode, nsMgr)
                    ff.CurrentValue = If(showingPlcHdr, "", rawValue)
                    ff.CurrentOptionValue = ff.CurrentValue
                End If

                results.Add(ff)
                fieldIndex += 1
            Next

            Dim fldCharNodes As XmlNodeList = containerNode.SelectNodes(".//w:r[w:fldChar[@w:fldCharType='begin']]", nsMgr)
            For Each fldRunNode As XmlNode In fldCharNodes
                Dim fldChar As XmlNode = fldRunNode.SelectSingleNode("w:fldChar", nsMgr)
                If fldChar Is Nothing Then Continue For

                Dim ffData As XmlNode = fldChar.SelectSingleNode("w:ffData", nsMgr)
                If ffData Is Nothing Then Continue For

                Dim ff As New FormFieldInfo() With {
                    .FieldId = $"f{fieldIndex}",
                    .XmlNode = fldChar,
                    .Options = New List(Of FormFieldOptionInfo)(),
                    .CurrentValue = "",
                    .CurrentOptionValue = "",
                    .IsWritable = isWritable,
                    .ProtectionReason = protectionReason
                }

                Dim nameNode As XmlNode = ffData.SelectSingleNode("w:name/@w:val", nsMgr)
                ff.Name = If(nameNode?.Value, "")

                Dim ddList As XmlNode = ffData.SelectSingleNode("w:ddList", nsMgr)
                If ddList IsNot Nothing Then
                    ff.FieldType = FormFieldType.LegacyDropdown

                    Dim listEntries As XmlNodeList = ddList.SelectNodes("w:listEntry/@w:val", nsMgr)
                    For Each entry As XmlNode In listEntries
                        AddFormFieldOption(ff.Options, entry.Value, entry.Value)
                    Next

                    Dim resultNode As XmlNode = ddList.SelectSingleNode("w:result/@w:val", nsMgr)
                    Dim selectedIdx As Integer = 0
                    If resultNode IsNot Nothing Then Integer.TryParse(resultNode.Value, selectedIdx)

                    If selectedIdx >= 0 AndAlso selectedIdx < ff.Options.Count Then
                        ff.CurrentValue = ff.Options(selectedIdx).DisplayText
                        ff.CurrentOptionValue = ff.Options(selectedIdx).StoredValue
                    End If

                    results.Add(ff)
                    fieldIndex += 1
                    Continue For
                End If

                Dim checkBox As XmlNode = ffData.SelectSingleNode("w:checkBox", nsMgr)
                If checkBox IsNot Nothing Then
                    ff.FieldType = FormFieldType.LegacyCheckbox
                    AddFormFieldOption(ff.Options, "true", "true")
                    AddFormFieldOption(ff.Options, "false", "false")

                    Dim checkedNode As XmlNode = checkBox.SelectSingleNode("w:checked/@w:val", nsMgr)
                    If checkedNode IsNot Nothing Then
                        ff.CurrentValue = If(checkedNode.Value = "1" OrElse checkedNode.Value.ToLowerInvariant() = "true", "true", "false")
                    Else
                        Dim defaultNode As XmlNode = checkBox.SelectSingleNode("w:default/@w:val", nsMgr)
                        ff.CurrentValue = If(defaultNode IsNot Nothing AndAlso (defaultNode.Value = "1" OrElse defaultNode.Value.ToLowerInvariant() = "true"), "true", "false")
                    End If

                    ff.CurrentOptionValue = ff.CurrentValue
                    results.Add(ff)
                    fieldIndex += 1
                End If
            Next

            Return results
        End Function

        Private Function TryBuildBodySection(node As XmlNode,
                                             sectionIndex As Integer,
                                             headingText As String,
                                             nsMgr As XmlNamespaceManager) As BodySectionInfo
            If node Is Nothing Then Return Nothing
            If node.LocalName = "tbl" Then Return Nothing

            Dim paragraphs As List(Of BodyParagraphInfo) = ExtractCandidateBodyParagraphsFromNode(node, nsMgr)

            Dim protectionReason As String = ""
            Dim isWritable As Boolean = IsCellWritable(node, nsMgr, protectionReason)
            Dim fields As List(Of FormFieldInfo) = ExtractFormFieldsFromNode(node, nsMgr, isWritable, protectionReason)

            If paragraphs.Count = 0 AndAlso (fields Is Nothing OrElse fields.Count = 0) Then
                Return Nothing
            End If

            Return New BodySectionInfo() With {
                .SectionIndex = sectionIndex,
                .HeadingText = If(headingText, ""),
                .Paragraphs = paragraphs,
                .FormFields = fields
            }
        End Function

        Private Function ExtractBodySectionsFromXml(xmlDoc As XmlDocument,
                                                    nsMgr As XmlNamespaceManager) As List(Of BodySectionInfo)
            Dim results As New List(Of BodySectionInfo)()
            Dim bodyNode As XmlNode = xmlDoc.SelectSingleNode("//w:body", nsMgr)
            If bodyNode Is Nothing Then Return results

            Dim sectionIndex As Integer = 0
            Dim lastContextText As String = ""

            For Each child As XmlNode In bodyNode.ChildNodes
                If child.NodeType <> XmlNodeType.Element Then Continue For
                If child.LocalName = "tbl" Then Continue For

                Dim section As BodySectionInfo = TryBuildBodySection(child, sectionIndex, lastContextText, nsMgr)
                If section IsNot Nothing Then
                    results.Add(section)
                    sectionIndex += 1
                End If

                Dim contextText As String = ExtractBestContextParagraphTextFromNode(child, nsMgr)
                If Not String.IsNullOrWhiteSpace(contextText) Then
                    lastContextText = contextText
                End If
            Next

            Return results
        End Function

        Private Shared Function NormalizeTextForPlaceholderDetection(value As String) As String
            If value Is Nothing Then Return ""

            Dim normalized As String = value.Normalize()
            normalized = normalized.Replace(ChrW(&HA0), " ")
            normalized = normalized.Replace(ChrW(&H2007), " ")
            normalized = normalized.Replace(ChrW(&H202F), " ")
            normalized = normalized.Replace(ChrW(&H200B), "")
            normalized = normalized.Replace(ChrW(&H200C), "")
            normalized = normalized.Replace(ChrW(&H200D), "")
            normalized = normalized.Replace(ChrW(&HFEFF), "")
            normalized = normalized.Replace(vbTab, " ")
            normalized = normalized.Replace(vbCr, " ")
            normalized = normalized.Replace(vbLf, " ")
            normalized = Regex.Replace(normalized, "\s+", " ").Trim()

            Return normalized
        End Function

        Private Shared Function IsPlaceholderText(value As String) As Boolean
            Dim normalized As String = NormalizeTextForPlaceholderDetection(value)
            If String.IsNullOrWhiteSpace(normalized) Then Return False
            Return PlaceholderPattern.IsMatch(normalized)
        End Function

        Private Function IsTextNodeInsideSdt(textNode As XmlNode, tcNode As XmlNode) As Boolean
            Dim current As XmlNode = textNode.ParentNode

            While current IsNot Nothing AndAlso current IsNot tcNode
                If current.LocalName = "sdt" Then Return True
                current = current.ParentNode
            End While

            Return False
        End Function

        Private Function TryGetCellOverrideWidthPct(tcNode As XmlNode,
                                                    totalGridWidth As Double,
                                                    nsMgr As XmlNamespaceManager,
                                                    ByRef widthPct As Double) As Boolean
            widthPct = 0

            If tcNode Is Nothing OrElse nsMgr Is Nothing Then Return False

            Dim tcWidthValueNode As XmlNode = tcNode.SelectSingleNode("w:tcPr/w:tcW/@w:w", nsMgr)
            If tcWidthValueNode Is Nothing Then Return False

            Dim tcWidthTypeNode As XmlNode = tcNode.SelectSingleNode("w:tcPr/w:tcW/@w:type", nsMgr)
            Dim tcWidthType As String = If(tcWidthTypeNode?.Value, "").Trim().ToLowerInvariant()

            Dim rawWidth As Double
            If Not Double.TryParse(tcWidthValueNode.Value, rawWidth) OrElse rawWidth <= 0 Then Return False

            Select Case tcWidthType
                Case "pct"
                    widthPct = Math.Round(rawWidth / 50.0R, 1)
                    Return widthPct > 0

                Case "dxa", ""
                    If totalGridWidth > 0 Then
                        widthPct = Math.Round(rawWidth / totalGridWidth * 100.0R, 1)
                        Return widthPct > 0
                    End If
            End Select

            Return False
        End Function

        Private Function GetDocumentOrderTableIndexMap(xmlDoc As XmlDocument,
                                                       nsMgr As XmlNamespaceManager) As Dictionary(Of XmlNode, Integer)
            Dim result As New Dictionary(Of XmlNode, Integer)()
            Dim tblNodes As XmlNodeList = xmlDoc.SelectNodes("//w:tbl", nsMgr)
            Dim index As Integer = 0

            For Each tblNode As XmlNode In tblNodes
                result(tblNode) = index
                index += 1
            Next

            Return result
        End Function

        Private Sub AppendBodyContextNode(node As XmlNode,
                                          nsMgr As XmlNamespaceManager,
                                          sb As StringBuilder,
                                          tableIndexMap As Dictionary(Of XmlNode, Integer))
            If node Is Nothing Then Exit Sub
            If node.NodeType <> XmlNodeType.Element Then Exit Sub

            Select Case node.LocalName
                Case "p"
                    Dim text As String = GetParagraphPlainText(node, nsMgr)
                    If Not String.IsNullOrWhiteSpace(text) Then
                        sb.AppendLine(text)
                    End If

                Case "tbl"
                    Dim tableIndex As Integer = -1
                    If tableIndexMap IsNot Nothing AndAlso tableIndexMap.ContainsKey(node) Then
                        tableIndex = tableIndexMap(node)
                    End If

                    If tableIndex >= 0 Then
                        sb.AppendLine($"[Table {tableIndex}]")
                    End If

                Case "sdt", "customXml", "smartTag", "ins", "moveTo", "hyperlink", "Choice", "Fallback"
                    For Each child As XmlNode In node.ChildNodes
                        If child.NodeType <> XmlNodeType.Element Then Continue For
                        AppendBodyContextNode(child, nsMgr, sb, tableIndexMap)
                    Next

                Case "AlternateContent"
                    Dim branch As XmlNode = Nothing

                    For Each child As XmlNode In node.ChildNodes
                        If child.NodeType <> XmlNodeType.Element Then Continue For
                        If child.LocalName = "Choice" Then
                            branch = child
                            Exit For
                        End If
                    Next

                    If branch Is Nothing Then
                        For Each child As XmlNode In node.ChildNodes
                            If child.NodeType <> XmlNodeType.Element Then Continue For
                            If child.LocalName = "Fallback" Then
                                branch = child
                                Exit For
                            End If
                        Next
                    End If

                    If branch IsNot Nothing Then
                        AppendBodyContextNode(branch, nsMgr, sb, tableIndexMap)
                    End If
            End Select
        End Sub

        Private Function GetRowStructureSignature(row As TableRowInfo) As String
            If row Is Nothing OrElse row.Cells Is Nothing Then Return ""

            Dim parts As New List(Of String)()
            parts.Add(row.Cells.Count.ToString())

            For Each cell In row.Cells.OrderBy(Function(c) c.Col)
                Dim fieldCount As Integer = If(cell.FormFields IsNot Nothing, cell.FormFields.Count, 0)
                parts.Add($"{cell.Col}:{cell.GridSpan}:{cell.VMerge}:{fieldCount}")
            Next

            Return String.Join("|", parts)
        End Function

        Private Function FindTemplateRowForInsertion(tbl As TableInfo, insertAfterRow As Integer) As TableRowInfo
            If tbl Is Nothing OrElse tbl.Rows Is Nothing OrElse tbl.Rows.Count = 0 Then Return Nothing

            Dim anchor As TableRowInfo = Nothing
            If insertAfterRow >= 0 Then
                anchor = tbl.Rows.FirstOrDefault(Function(r) r.RowIndex = insertAfterRow)
            End If
            If anchor Is Nothing Then
                anchor = tbl.Rows.LastOrDefault(Function(r) Not r.Cells.All(Function(c) c.IsHeader))
            End If
            If anchor Is Nothing Then
                anchor = tbl.Rows.Last()
            End If

            Dim targetSignature As String = GetRowStructureSignature(anchor)

            Dim candidates As IEnumerable(Of TableRowInfo) =
                tbl.Rows.
                    Where(Function(r) Not r.Cells.All(Function(c) c.IsHeader)).
                    OrderBy(Function(r) Math.Abs(r.RowIndex - anchor.RowIndex))

            Dim signatureMatch As TableRowInfo =
                candidates.FirstOrDefault(Function(r) GetRowStructureSignature(r) = targetSignature)

            If signatureMatch IsNot Nothing Then Return signatureMatch

            Dim sameCellCount As TableRowInfo =
                candidates.FirstOrDefault(Function(r) r.Cells.Count = anchor.Cells.Count)

            If sameCellCount IsNot Nothing Then Return sameCellCount

            Return candidates.LastOrDefault()
        End Function

        Private Function CloneTemplateRowBeforeApply(tbl As TableInfo, insertAfterRow As Integer) As XmlNode
            Dim templateRow As TableRowInfo = FindTemplateRowForInsertion(tbl, insertAfterRow)
            If templateRow Is Nothing OrElse templateRow.XmlNode Is Nothing Then Return Nothing
            Return templateRow.XmlNode.CloneNode(deep:=True)
        End Function

        Private Function RowContainsVerticalMerge(rowNode As XmlNode, nsMgr As XmlNamespaceManager) As Boolean
            If rowNode Is Nothing Then Return False
            Return rowNode.SelectSingleNode(".//w:tcPr/w:vMerge", nsMgr) IsNot Nothing
        End Function

        Private Function ExtractJsonArrayPayload(response As String) As String
            If String.IsNullOrWhiteSpace(response) Then Return ""

            Dim normalized As String = response.Trim()

            If normalized.StartsWith("```", StringComparison.Ordinal) Then
                Dim firstNewline As Integer = normalized.IndexOf(vbLf)
                If firstNewline >= 0 Then
                    normalized = normalized.Substring(firstNewline + 1)
                End If
                If normalized.EndsWith("```", StringComparison.Ordinal) Then
                    normalized = normalized.Substring(0, normalized.Length - 3)
                End If
                normalized = normalized.Trim()
            End If

            If normalized.StartsWith("<json>", StringComparison.OrdinalIgnoreCase) AndAlso
               normalized.EndsWith("</json>", StringComparison.OrdinalIgnoreCase) Then
                normalized = normalized.Substring(6, normalized.Length - 13).Trim()
            End If

            Dim startIdx As Integer = normalized.IndexOf("["c)
            If startIdx < 0 Then Return normalized

            Dim depth As Integer = 0
            Dim inString As Boolean = False
            Dim escapeNext As Boolean = False

            For i As Integer = startIdx To normalized.Length - 1
                Dim ch As Char = normalized(i)

                If escapeNext Then
                    escapeNext = False
                    Continue For
                End If

                If ch = "\"c AndAlso inString Then
                    escapeNext = True
                    Continue For
                End If

                If ch = """"c Then
                    inString = Not inString
                    Continue For
                End If

                If inString Then Continue For

                If ch = "["c Then
                    depth += 1
                ElseIf ch = "]"c Then
                    depth -= 1
                    If depth = 0 Then
                        Return normalized.Substring(startIdx, i - startIdx + 1).Trim()
                    End If
                End If
            Next

            Return normalized
        End Function

        Private Sub ApplySdtDefaultRunProperties(sdtNode As XmlNode, nsMgr As XmlNamespaceManager)
            If sdtNode Is Nothing Then Exit Sub

            Dim sdtPrRunProps As XmlNode = sdtNode.SelectSingleNode("w:sdtPr/w:rPr", nsMgr)
            Dim firstRun As XmlNode = sdtNode.SelectSingleNode("w:sdtContent//w:r", nsMgr)

            If firstRun Is Nothing Then Exit Sub

            Dim existingRpr As XmlNode = firstRun.SelectSingleNode("w:rPr", nsMgr)
            If existingRpr IsNot Nothing Then
                firstRun.RemoveChild(existingRpr)
            End If

            If sdtPrRunProps Is Nothing Then Exit Sub

            Dim clonedRpr As XmlNode = sdtPrRunProps.CloneNode(deep:=True)

            Dim placeholderStyles As XmlNodeList = clonedRpr.SelectNodes("w:rStyle[@w:val='PlaceholderText' or @w:val='Platzhaltertext']", nsMgr)
            For Each styleNode As XmlNode In placeholderStyles
                clonedRpr.RemoveChild(styleNode)
            Next

            firstRun.PrependChild(clonedRpr)
        End Sub

        Private Function ExtractDocumentBodyContext(xmlDoc As XmlDocument,
                                                    nsMgr As XmlNamespaceManager) As String
            Dim bodyNode As XmlNode = xmlDoc.SelectSingleNode("//w:body", nsMgr)
            If bodyNode Is Nothing Then Return ""

            Dim sb As New StringBuilder()
            Dim tableIndexMap As Dictionary(Of XmlNode, Integer) = GetDocumentOrderTableIndexMap(xmlDoc, nsMgr)

            For Each child As XmlNode In bodyNode.ChildNodes
                If child.NodeType <> XmlNodeType.Element Then Continue For
                AppendBodyContextNode(child, nsMgr, sb, tableIndexMap)

            Next

            ' This isolated file worker needs the full source; never discard middle/end sections.
            Return sb.ToString()
        End Function

        Private Function ExtractTablesFromXml(xmlDoc As XmlDocument,
                                              nsMgr As XmlNamespaceManager) As List(Of TableInfo)
            Dim tables As New List(Of TableInfo)()
            Dim tblNodes As XmlNodeList = xmlDoc.SelectNodes("//w:tbl", nsMgr)
            Dim tableIndex As Integer = 0

            For Each tblNode As XmlNode In tblNodes
                Dim tblInfo As New TableInfo() With {
                    .TableIndex = tableIndex,
                    .Rows = New List(Of TableRowInfo)(),
                    .XmlNode = tblNode,
                    .ColumnCount = 0,
                    .ContextBefore = "",
                    .ColumnWidths = New List(Of Double)()
                }

                Dim gridCols As XmlNodeList = tblNode.SelectNodes("w:tblGrid/w:gridCol", nsMgr)
                Dim rawWidths As New List(Of Double)()
                Dim totalGridWidth As Double = 0

                For Each gridCol As XmlNode In gridCols
                    Dim wAttr As XmlNode = gridCol.Attributes("w:w")
                    Dim colWidth As Double = 0
                    If wAttr IsNot Nothing Then Double.TryParse(wAttr.Value, colWidth)
                    rawWidths.Add(colWidth)
                    totalGridWidth += colWidth
                Next

                If totalGridWidth > 0 AndAlso rawWidths.Count > 0 Then
                    For Each w In rawWidths
                        tblInfo.ColumnWidths.Add(Math.Round(w / totalGridWidth * 100.0R, 1))
                    Next
                End If

                Dim prevSibling As XmlNode = tblNode.PreviousSibling
                Dim contextParts As New List(Of String)()
                Dim contextCount As Integer = 0

                While prevSibling IsNot Nothing AndAlso contextCount < 3
                    If prevSibling.NodeType = XmlNodeType.Element AndAlso prevSibling.LocalName = "p" Then
                        Dim paraText As String = GetParagraphPlainText(prevSibling, nsMgr)
                        If Not String.IsNullOrWhiteSpace(paraText) Then
                            contextParts.Insert(0, paraText)
                            contextCount += 1
                        End If
                    End If
                    prevSibling = prevSibling.PreviousSibling
                End While

                tblInfo.ContextBefore = String.Join(" | ", contextParts)

                Dim trNodes As XmlNodeList = tblNode.SelectNodes("w:tr", nsMgr)
                Dim rowIndex As Integer = 0

                For Each trNode As XmlNode In trNodes
                    Dim rowInfo As New TableRowInfo() With {
                        .RowIndex = rowIndex,
                        .Cells = New List(Of TableCellInfo)(),
                        .XmlNode = trNode
                    }

                    Dim isHeaderRow As Boolean = (trNode.SelectSingleNode("w:trPr/w:tblHeader", nsMgr) IsNot Nothing)
                    Dim effectiveTcNodes As List(Of XmlNode) = GetEffectiveRowCellNodes(trNode, nsMgr)
                    Dim colIndex As Integer = 0

                    For Each tcNode As XmlNode In effectiveTcNodes
                        Dim cellInfo As New TableCellInfo() With {
                            .Row = rowIndex,
                            .Col = colIndex,
                            .TextNodes = New List(Of XmlNode)(),
                            .IsHeader = isHeaderRow OrElse rowIndex = 0,
                            .GridSpan = 1,
                            .VMerge = "",
                            .IsPlaceholder = False,
                            .WidthPct = 0,
                            .XmlNode = tcNode,
                            .IsWritable = True,
                            .ProtectionReason = "",
                            .FormFields = New List(Of FormFieldInfo)()
                        }

                        Dim gridSpanNode As XmlNode = tcNode.SelectSingleNode("w:tcPr/w:gridSpan/@w:val", nsMgr)
                        If gridSpanNode IsNot Nothing Then
                            Dim span As Integer
                            If Integer.TryParse(gridSpanNode.Value, span) AndAlso span > 0 Then
                                cellInfo.GridSpan = span
                            End If
                        End If

                        Dim overrideWidthPct As Double = 0
                        If TryGetCellOverrideWidthPct(tcNode, totalGridWidth, nsMgr, overrideWidthPct) Then
                            cellInfo.WidthPct = overrideWidthPct
                        ElseIf tblInfo.ColumnWidths.Count > 0 Then
                            Dim widthSum As Double = 0
                            For gi As Integer = colIndex To Math.Min(colIndex + cellInfo.GridSpan - 1, tblInfo.ColumnWidths.Count - 1)
                                widthSum += tblInfo.ColumnWidths(gi)
                            Next
                            cellInfo.WidthPct = widthSum
                        End If

                        Dim vMergeNode As XmlNode = tcNode.SelectSingleNode("w:tcPr/w:vMerge", nsMgr)
                        If vMergeNode IsNot Nothing Then
                            Dim vMergeVal As XmlNode = vMergeNode.Attributes("w:val")
                            cellInfo.VMerge = If(vMergeVal IsNot Nothing AndAlso vMergeVal.Value = "restart", "restart", "continue")
                        End If

                        Dim cellProtectionReason As String = ""
                        cellInfo.IsWritable = IsCellWritable(tcNode, nsMgr, cellProtectionReason)
                        cellInfo.ProtectionReason = cellProtectionReason

                        ExtractFormFieldsFromCell(tcNode, cellInfo, nsMgr)

                        Dim textBuilder As New StringBuilder()
                        Dim allTextNodes As XmlNodeList = tcNode.SelectNodes(".//w:t", nsMgr)

                        For Each tNode As XmlNode In allTextNodes
                            If IsTextNodeInsideSdt(tNode, tcNode) Then Continue For
                            cellInfo.TextNodes.Add(tNode)
                            textBuilder.Append(tNode.InnerText)
                        Next

                        cellInfo.Text = textBuilder.ToString()
                        cellInfo.IsPlaceholder = IsPlaceholderText(cellInfo.Text)

                        rowInfo.Cells.Add(cellInfo)
                        colIndex += cellInfo.GridSpan
                    Next

                    If colIndex > tblInfo.ColumnCount Then
                        tblInfo.ColumnCount = colIndex
                    End If

                    tblInfo.Rows.Add(rowInfo)
                    rowIndex += 1
                Next

                tables.Add(tblInfo)
                tableIndex += 1
            Next

            Return tables
        End Function

        Private Function GetEffectiveRowCellNodes(trNode As XmlNode, nsMgr As XmlNamespaceManager) As List(Of XmlNode)
            Dim results As New List(Of XmlNode)()

            If trNode Is Nothing Then Return results

            Dim expectedGridColumns As Integer = GetExpectedRowGridColumnCount(trNode, nsMgr)

            For Each child As XmlNode In trNode.ChildNodes
                If child.NodeType <> XmlNodeType.Element Then Continue For

                CollectEffectiveRowCellNodesCore(child, nsMgr, results, expectedGridColumns)

                If expectedGridColumns > 0 AndAlso GetAccumulatedGridSpan(results, nsMgr) >= expectedGridColumns Then
                    Exit For
                End If
            Next

            Return results
        End Function

        Private Sub CollectEffectiveRowCellNodesCore(node As XmlNode,
                                                     nsMgr As XmlNamespaceManager,
                                                     results As List(Of XmlNode),
                                                     expectedGridColumns As Integer)
            If node Is Nothing Then Exit Sub
            If nsMgr Is Nothing Then Exit Sub
            If results Is Nothing Then Exit Sub
            If node.NodeType <> XmlNodeType.Element Then Exit Sub

            If expectedGridColumns > 0 AndAlso GetAccumulatedGridSpan(results, nsMgr) >= expectedGridColumns Then
                Exit Sub
            End If

            Select Case node.LocalName
                Case "tc"
                    results.Add(node)
                    Exit Sub

                Case "sdt"
                    Dim sdtContent As XmlNode = Nothing

                    For Each child As XmlNode In node.ChildNodes
                        If child.NodeType <> XmlNodeType.Element Then Continue For
                        If child.LocalName = "sdtContent" Then
                            sdtContent = child
                            Exit For
                        End If
                    Next

                    If sdtContent Is Nothing Then Exit Sub

                    For Each child As XmlNode In sdtContent.ChildNodes
                        If child.NodeType <> XmlNodeType.Element Then Continue For

                        CollectEffectiveRowCellNodesCore(child, nsMgr, results, expectedGridColumns)

                        If expectedGridColumns > 0 AndAlso GetAccumulatedGridSpan(results, nsMgr) >= expectedGridColumns Then
                            Exit For
                        End If
                    Next

                Case "customXml", "smartTag", "ins", "moveTo", "hyperlink", "Choice", "Fallback"
                    For Each child As XmlNode In node.ChildNodes
                        If child.NodeType <> XmlNodeType.Element Then Continue For

                        CollectEffectiveRowCellNodesCore(child, nsMgr, results, expectedGridColumns)

                        If expectedGridColumns > 0 AndAlso GetAccumulatedGridSpan(results, nsMgr) >= expectedGridColumns Then
                            Exit For
                        End If
                    Next

                Case "AlternateContent"
                    Dim branch As XmlNode = Nothing

                    For Each child As XmlNode In node.ChildNodes
                        If child.NodeType <> XmlNodeType.Element Then Continue For
                        If child.LocalName = "Choice" Then
                            branch = child
                            Exit For
                        End If
                    Next

                    If branch Is Nothing Then
                        For Each child As XmlNode In node.ChildNodes
                            If child.NodeType <> XmlNodeType.Element Then Continue For
                            If child.LocalName = "Fallback" Then
                                branch = child
                                Exit For
                            End If
                        Next
                    End If

                    If branch IsNot Nothing Then
                        CollectEffectiveRowCellNodesCore(branch, nsMgr, results, expectedGridColumns)
                    End If

                Case "del", "moveFrom"
                    Exit Sub
            End Select
        End Sub

        Private Function GetAncestorRowNode(node As XmlNode) As XmlNode
            Dim current As XmlNode = node

            While current IsNot Nothing
                If current.LocalName = "tr" Then Return current
                If current.LocalName = "tbl" OrElse current.LocalName = "body" Then Exit While
                current = current.ParentNode
            End While

            Return Nothing
        End Function

        Private Function GetExpectedRowGridColumnCount(trNode As XmlNode, nsMgr As XmlNamespaceManager) As Integer
            If trNode Is Nothing Then Return 0
            If nsMgr Is Nothing Then Return 0

            Dim current As XmlNode = trNode

            While current IsNot Nothing AndAlso current.LocalName <> "tbl"
                current = current.ParentNode
            End While

            If current Is Nothing Then Return 0

            Dim gridCols As XmlNodeList = current.SelectNodes("w:tblGrid/w:gridCol", nsMgr)
            If gridCols Is Nothing Then Return 0

            Return gridCols.Count
        End Function

        Private Function GetAccumulatedGridSpan(tcNodes As List(Of XmlNode), nsMgr As XmlNamespaceManager) As Integer
            If tcNodes Is Nothing OrElse tcNodes.Count = 0 Then Return 0

            Dim total As Integer = 0

            For Each tcNode As XmlNode In tcNodes
                total += GetGridSpanValue(tcNode, nsMgr)
            Next

            Return total
        End Function

        Private Function GetGridSpanValue(tcNode As XmlNode, nsMgr As XmlNamespaceManager) As Integer
            If tcNode Is Nothing Then Return 1
            If nsMgr Is Nothing Then Return 1

            Dim spanNode As XmlNode = tcNode.SelectSingleNode("w:tcPr/w:gridSpan/@w:val", nsMgr)
            If spanNode Is Nothing Then Return 1

            Dim span As Integer
            If Integer.TryParse(spanNode.Value, span) AndAlso span > 0 Then
                Return span
            End If

            Return 1
        End Function

        Private Function GetOwningSdtNodeForCell(tcNode As XmlNode) As XmlNode
            Dim ancestor As XmlNode = tcNode.ParentNode

            While ancestor IsNot Nothing
                If ancestor.LocalName = "sdt" Then Return ancestor
                If ancestor.LocalName = "body" Then Exit While
                ancestor = ancestor.ParentNode
            End While

            Return Nothing
        End Function

        Private Function TryGetLockedSdtReason(node As XmlNode,
                                               nsMgr As XmlNamespaceManager,
                                               ByRef reason As String) As Boolean
            Dim current As XmlNode = node

            While current IsNot Nothing
                If current.NodeType = XmlNodeType.Element AndAlso current.LocalName = "sdt" Then
                    Dim lockAttr As XmlNode = current.SelectSingleNode("w:sdtPr/w:lock/@w:val", nsMgr)
                    If lockAttr IsNot Nothing Then
                        Dim lockValue As String = lockAttr.Value.Trim().ToLowerInvariant()

                        If lockValue = "contentlocked" OrElse
                           lockValue = "sdtcontentlocked" OrElse
                           lockValue = "sdtlocked" Then
                            reason = $"SDT content is locked ({lockAttr.Value})."
                            Return True
                        End If
                    End If
                End If

                If current.LocalName = "body" Then Exit While
                current = current.ParentNode
            End While

            reason = ""
            Return False
        End Function

        Private Function IsCellWritable(tcNode As XmlNode, nsMgr As XmlNamespaceManager, ByRef reason As String) As Boolean
            Return Not TryGetLockedSdtReason(tcNode, nsMgr, reason)
        End Function

        Private Function IsFormFieldWritable(fieldNode As XmlNode, nsMgr As XmlNamespaceManager, ByRef reason As String) As Boolean
            Return Not TryGetLockedSdtReason(fieldNode, nsMgr, reason)
        End Function

        Private Sub ExtractFormFieldsFromCell(tcNode As XmlNode, cellInfo As TableCellInfo, nsMgr As XmlNamespaceManager)
            If cellInfo Is Nothing Then Exit Sub

            Dim protectionReason As String = If(cellInfo.ProtectionReason, "")
            Dim writable As Boolean = cellInfo.IsWritable

            cellInfo.FormFields = ExtractFormFieldsFromNode(tcNode, nsMgr, writable, protectionReason)
        End Sub

        Private Function GetSdtContentText(sdtNode As XmlNode, nsMgr As XmlNamespaceManager) As String
            Dim sdtContent As XmlNode = sdtNode.SelectSingleNode("w:sdtContent", nsMgr)
            If sdtContent Is Nothing Then Return ""
            Dim sb As New StringBuilder()
            For Each ct As XmlNode In sdtContent.SelectNodes(".//w:t", nsMgr)
                sb.Append(ct.InnerText)
            Next
            Return sb.ToString()
        End Function

        Private Function GetParagraphPlainText(paraNode As XmlNode, nsMgr As XmlNamespaceManager) As String
            Dim sb As New StringBuilder()
            Dim textNodes As XmlNodeList = paraNode.SelectNodes(".//w:t", nsMgr)
            For Each tNode As XmlNode In textNodes
                sb.Append(tNode.InnerText)
            Next
            Return sb.ToString()
        End Function

        Private Sub AddFormFieldOption(options As List(Of FormFieldOptionInfo), displayText As String, storedValue As String)
            Dim normalizedDisplayText As String = If(displayText, "").Trim()
            Dim normalizedStoredValue As String = If(storedValue, "").Trim()

            If String.IsNullOrWhiteSpace(normalizedDisplayText) Then normalizedDisplayText = normalizedStoredValue
            If String.IsNullOrWhiteSpace(normalizedStoredValue) Then normalizedStoredValue = normalizedDisplayText
            If String.IsNullOrWhiteSpace(normalizedDisplayText) AndAlso String.IsNullOrWhiteSpace(normalizedStoredValue) Then Exit Sub

            If options.Any(Function(o) o.DisplayText.Equals(normalizedDisplayText, StringComparison.OrdinalIgnoreCase) AndAlso
                                     o.StoredValue.Equals(normalizedStoredValue, StringComparison.OrdinalIgnoreCase)) Then
                Exit Sub
            End If

            options.Add(New FormFieldOptionInfo() With {
                .DisplayText = normalizedDisplayText,
                .StoredValue = normalizedStoredValue
            })
        End Sub

        Private Function ResolveCurrentDropdownOption(ff As FormFieldInfo, rawDisplayValue As String, rawStoredValue As String) As FormFieldOptionInfo
            If ff Is Nothing OrElse ff.Options Is Nothing OrElse ff.Options.Count = 0 Then Return Nothing

            If Not String.IsNullOrWhiteSpace(rawStoredValue) Then
                Dim optionByStoredValue As FormFieldOptionInfo =
                    ff.Options.FirstOrDefault(Function(o) o.StoredValue.Equals(rawStoredValue, StringComparison.OrdinalIgnoreCase))
                If optionByStoredValue IsNot Nothing Then Return optionByStoredValue
            End If

            If Not String.IsNullOrWhiteSpace(rawDisplayValue) Then
                Dim optionByDisplayText As FormFieldOptionInfo =
                    ff.Options.FirstOrDefault(Function(o) o.DisplayText.Equals(rawDisplayValue, StringComparison.OrdinalIgnoreCase))
                If optionByDisplayText IsNot Nothing Then Return optionByDisplayText

                Dim optionByDisplayedStoredValue As FormFieldOptionInfo =
                    ff.Options.FirstOrDefault(Function(o) o.StoredValue.Equals(rawDisplayValue, StringComparison.OrdinalIgnoreCase))
                If optionByDisplayedStoredValue IsNot Nothing Then Return optionByDisplayedStoredValue
            End If

            Return Nothing
        End Function

#End Region

#Region "JSON Serialization"

        Private Function BodySectionToJson(section As BodySectionInfo) As String
            Dim jObj As New JObject()
            jObj("sectionIndex") = section.SectionIndex

            If Not String.IsNullOrWhiteSpace(section.HeadingText) Then
                jObj("headingText") = section.HeadingText
            End If

            Dim jParagraphs As New JArray()
            If section.Paragraphs IsNot Nothing Then
                For Each para In section.Paragraphs
                    Dim jPara As New JObject()
                    jPara("index") = para.Index
                    jPara("text") = If(para.Text, "")
                    If para.IsEmpty Then jPara("empty") = True
                    jParagraphs.Add(jPara)
                Next
            End If
            jObj("paragraphs") = jParagraphs

            Dim jFields As New JArray()
            If section.FormFields IsNot Nothing Then
                For Each ff In section.FormFields
                    Dim jFF As New JObject()
                    jFF("id") = ff.FieldId
                    jFF("type") = ff.FieldType.ToString().ToLowerInvariant()
                    If Not String.IsNullOrWhiteSpace(ff.Name) Then jFF("name") = ff.Name
                    jFF("val") = If(ff.CurrentValue, "")
                    If Not String.IsNullOrWhiteSpace(ff.CurrentOptionValue) Then
                        jFF("storedVal") = ff.CurrentOptionValue
                    End If
                    If Not ff.IsWritable Then jFF("ro") = True
                    If ff.IsWritable AndAlso String.IsNullOrWhiteSpace(ff.CurrentValue) AndAlso String.IsNullOrWhiteSpace(ff.CurrentOptionValue) Then
                        jFF("empty") = True
                    End If

                    If ff.Options IsNot Nothing AndAlso ff.Options.Count > 0 Then
                        Dim jOpts As New JArray()

                        If ff.FieldType = FormFieldType.Dropdown OrElse ff.FieldType = FormFieldType.LegacyDropdown Then
                            For Each opt In ff.Options
                                Dim jOpt As New JObject()
                                jOpt("text") = If(opt.DisplayText, "")
                                jOpt("value") = If(opt.StoredValue, "")
                                jOpts.Add(jOpt)
                            Next
                        Else
                            For Each opt In ff.Options
                                jOpts.Add(If(opt.DisplayText, ""))
                            Next
                        End If

                        jFF("opts") = jOpts
                    End If

                    jFields.Add(jFF)
                Next
            End If
            jObj("fields") = jFields

            Return jObj.ToString(Newtonsoft.Json.Formatting.None)
        End Function

        Private Function BuildAllBodySectionsJson(bodySections As List(Of BodySectionInfo)) As String
            Dim jArr As New JArray()

            If bodySections Is Nothing Then Return "[]"

            For Each section In bodySections
                jArr.Add(JObject.Parse(BodySectionToJson(section)))
            Next

            Return jArr.ToString(Newtonsoft.Json.Formatting.None)
        End Function

        Private Function BuildAllTablesJson(tables As List(Of TableInfo)) As String
            Dim jArr As New JArray()

            If tables Is Nothing Then Return "[]"

            For Each tbl In tables
                jArr.Add(JObject.Parse(TableToJson(tbl)))
            Next

            Return jArr.ToString(Newtonsoft.Json.Formatting.None)
        End Function

        Private Shared Function TableToJson(tbl As TableInfo) As String
            Dim jObj As New JObject()
            jObj("tableIndex") = tbl.TableIndex

            If Not String.IsNullOrWhiteSpace(tbl.ContextBefore) Then
                jObj("contextBefore") = tbl.ContextBefore
            End If

            jObj("columns") = tbl.ColumnCount

            If tbl.ColumnWidths IsNot Nothing AndAlso tbl.ColumnWidths.Count > 0 Then
                Dim jWidths As New JArray()
                For Each w In tbl.ColumnWidths
                    jWidths.Add(w)
                Next
                jObj("colWidthsPct") = jWidths
            End If

            Dim jRows As New JArray()
            For Each row In tbl.Rows
                Dim jRow As New JObject()
                jRow("r") = row.RowIndex

                Dim jCells As New JArray()
                For Each cell In row.Cells
                    Dim jCell As New JObject()
                    Dim cellText As String = If(cell.Text, "")
                    Dim normalizedCellText As String = NormalizeTextForPlaceholderDetection(cellText)

                    jCell("c") = cell.Col
                    jCell("t") = cellText

                    If String.IsNullOrWhiteSpace(normalizedCellText) Then
                        jCell("empty") = True
                    End If

                    If cell.GridSpan > 1 Then jCell("span") = cell.GridSpan
                    If cell.VMerge <> "" Then jCell("vm") = cell.VMerge
                    If cell.IsHeader Then jCell("hdr") = True
                    If cell.WidthPct > 0 Then jCell("w") = cell.WidthPct
                    If Not cell.IsWritable Then jCell("ro") = True

                    If cell.FormFields IsNot Nothing AndAlso cell.FormFields.Count > 0 Then
                        Dim jFields As New JArray()
                        For Each ff In cell.FormFields
                            Dim jFF As New JObject()
                            jFF("id") = ff.FieldId
                            jFF("type") = ff.FieldType.ToString().ToLowerInvariant()
                            If Not String.IsNullOrWhiteSpace(ff.Name) Then jFF("name") = ff.Name
                            jFF("val") = If(ff.CurrentValue, "")
                            If Not String.IsNullOrWhiteSpace(ff.CurrentOptionValue) Then
                                jFF("storedVal") = ff.CurrentOptionValue
                            End If
                            If Not ff.IsWritable Then jFF("ro") = True
                            If ff.IsWritable AndAlso String.IsNullOrWhiteSpace(ff.CurrentValue) AndAlso String.IsNullOrWhiteSpace(ff.CurrentOptionValue) Then
                                jFF("empty") = True
                            End If

                            If ff.Options IsNot Nothing AndAlso ff.Options.Count > 0 Then
                                Dim jOpts As New JArray()

                                If ff.FieldType = FormFieldType.Dropdown OrElse ff.FieldType = FormFieldType.LegacyDropdown Then
                                    For Each opt In ff.Options
                                        Dim jOpt As New JObject()
                                        jOpt("text") = If(opt.DisplayText, "")
                                        jOpt("value") = If(opt.StoredValue, "")
                                        jOpts.Add(jOpt)
                                    Next
                                Else
                                    For Each opt In ff.Options
                                        jOpts.Add(If(opt.DisplayText, ""))
                                    Next
                                End If

                                jFF("opts") = jOpts
                            End If

                            jFields.Add(jFF)
                        Next
                        jCell("fields") = jFields
                    End If

                    jCells.Add(jCell)
                Next

                jRow("cells") = jCells
                jRows.Add(jRow)
            Next

            jObj("rows") = jRows
            Return jObj.ToString(Newtonsoft.Json.Formatting.None)
        End Function

#End Region

#Region "LLM Batching"

        Private Function BuildTableCellKey(tableIndex As Integer, rowIndex As Integer, colIndex As Integer) As String
            Return $"{tableIndex}:{rowIndex}:{colIndex}"
        End Function

        Private Function BuildTableFieldKey(tableIndex As Integer, rowIndex As Integer, colIndex As Integer, fieldId As String) As String
            Return $"{tableIndex}:{rowIndex}:{colIndex}:{fieldId}"
        End Function

        Private Function GetHandledTableCellKeys(updates As List(Of TableUpdateResponse)) As HashSet(Of String)
            Dim result As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            If updates Is Nothing Then Return result

            For Each update In updates
                If update Is Nothing OrElse update.Updates Is Nothing Then Continue For

                For Each cellUpdate In update.Updates
                    result.Add(BuildTableCellKey(update.TableIndex, cellUpdate.Row, cellUpdate.Col))
                Next
            Next

            Return result
        End Function

        Private Function GetHandledTableFieldKeys(updates As List(Of TableUpdateResponse)) As HashSet(Of String)
            Dim result As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            If updates Is Nothing Then Return result

            For Each update In updates
                If update Is Nothing OrElse update.FieldUpdates Is Nothing Then Continue For

                For Each fieldUpdate In update.FieldUpdates
                    result.Add(BuildTableFieldKey(update.TableIndex, fieldUpdate.Row, fieldUpdate.Col, fieldUpdate.FieldId))
                Next
            Next

            Return result
        End Function

        Private Function BuildUnresolvedTableTargetsJson(tableBlock As List(Of TableInfo),
                                                         updates As List(Of TableUpdateResponse)) As String
            Dim handledCellKeys As HashSet(Of String) = GetHandledTableCellKeys(updates)
            Dim handledFieldKeys As HashSet(Of String) = GetHandledTableFieldKeys(updates)
            Dim jArr As New JArray()

            If tableBlock Is Nothing Then Return "[]"

            For Each tbl In tableBlock
                Dim jTbl As New JObject()
                jTbl("tableIndex") = tbl.TableIndex

                Dim jCells As New JArray()
                Dim jFields As New JArray()

                For Each row In tbl.Rows
                    For Each cell In row.Cells
                        If cell.IsWritable Then
                            Dim normalizedText As String = NormalizeTextForPlaceholderDetection(If(cell.Text, ""))
                            Dim hasFields As Boolean = (cell.FormFields IsNot Nothing AndAlso cell.FormFields.Count > 0)

                            If String.IsNullOrWhiteSpace(normalizedText) AndAlso Not hasFields Then
                                Dim key As String = BuildTableCellKey(tbl.TableIndex, row.RowIndex, cell.Col)
                                If Not handledCellKeys.Contains(key) Then
                                    Dim jCell As New JObject()
                                    jCell("r") = row.RowIndex
                                    jCell("c") = cell.Col
                                    jCell("t") = If(cell.Text, "")
                                    If cell.WidthPct > 0 Then jCell("w") = cell.WidthPct
                                    jCells.Add(jCell)
                                End If
                            End If
                        End If

                        If cell.FormFields IsNot Nothing Then
                            For Each ff In cell.FormFields
                                If Not ff.IsWritable Then Continue For

                                Dim isEmptyField As Boolean =
                                    String.IsNullOrWhiteSpace(ff.CurrentValue) AndAlso
                                    String.IsNullOrWhiteSpace(ff.CurrentOptionValue)

                                If Not isEmptyField Then Continue For

                                Dim key As String = BuildTableFieldKey(tbl.TableIndex, row.RowIndex, cell.Col, ff.FieldId)
                                If handledFieldKeys.Contains(key) Then Continue For

                                Dim jField As New JObject()
                                jField("r") = row.RowIndex
                                jField("c") = cell.Col
                                jField("id") = ff.FieldId
                                jField("type") = ff.FieldType.ToString().ToLowerInvariant()
                                If Not String.IsNullOrWhiteSpace(ff.Name) Then jField("name") = ff.Name

                                If ff.Options IsNot Nothing AndAlso ff.Options.Count > 0 Then
                                    Dim jOpts As New JArray()
                                    If ff.FieldType = FormFieldType.Dropdown OrElse ff.FieldType = FormFieldType.LegacyDropdown Then
                                        For Each opt In ff.Options
                                            Dim jOpt As New JObject()
                                            jOpt("text") = If(opt.DisplayText, "")
                                            jOpt("value") = If(opt.StoredValue, "")
                                            jOpts.Add(jOpt)
                                        Next
                                    Else
                                        For Each opt In ff.Options
                                            jOpts.Add(If(opt.DisplayText, ""))
                                        Next
                                    End If
                                    jField("opts") = jOpts
                                End If

                                jFields.Add(jField)
                            Next
                        End If
                    Next
                Next

                If jCells.Count > 0 OrElse jFields.Count > 0 Then
                    jTbl("unresolvedCells") = jCells
                    jTbl("unresolvedFields") = jFields
                    jArr.Add(jTbl)
                End If
            Next

            Return jArr.ToString(Newtonsoft.Json.Formatting.None)
        End Function

        Private Function MergeTableUpdateResponses(primaryUpdates As List(Of TableUpdateResponse),
                                                   secondaryUpdates As List(Of TableUpdateResponse)) As List(Of TableUpdateResponse)
            Dim mergedByTable As New Dictionary(Of Integer, TableUpdateResponse)()

            Dim allUpdates As New List(Of TableUpdateResponse)()
            If primaryUpdates IsNot Nothing Then allUpdates.AddRange(primaryUpdates)
            If secondaryUpdates IsNot Nothing Then allUpdates.AddRange(secondaryUpdates)

            For Each update In allUpdates
                If update Is Nothing Then Continue For

                If Not mergedByTable.ContainsKey(update.TableIndex) Then
                    mergedByTable(update.TableIndex) = New TableUpdateResponse() With {
                        .TableIndex = update.TableIndex,
                        .Updates = New List(Of TableCellUpdate)(),
                        .FieldUpdates = New List(Of FormFieldUpdate)(),
                        .NewRows = New List(Of List(Of String))(),
                        .InsertAfterRow = update.InsertAfterRow
                    }
                End If

                Dim target As TableUpdateResponse = mergedByTable(update.TableIndex)

                If target.InsertAfterRow < 0 AndAlso update.InsertAfterRow >= 0 Then
                    target.InsertAfterRow = update.InsertAfterRow
                End If

                If update.Updates IsNot Nothing Then
                    For Each cellUpdate In update.Updates
                        Dim existing As TableCellUpdate =
                            target.Updates.FirstOrDefault(Function(u) u.Row = cellUpdate.Row AndAlso u.Col = cellUpdate.Col)

                        If existing IsNot Nothing Then
                            existing.Action = cellUpdate.Action
                            existing.Text = cellUpdate.Text
                        Else
                            target.Updates.Add(New TableCellUpdate() With {
                                .Row = cellUpdate.Row,
                                .Col = cellUpdate.Col,
                                .Action = cellUpdate.Action,
                                .Text = cellUpdate.Text
                            })
                        End If
                    Next
                End If

                If update.FieldUpdates IsNot Nothing Then
                    For Each fieldUpdate In update.FieldUpdates
                        Dim existing As FormFieldUpdate =
                            target.FieldUpdates.FirstOrDefault(Function(f) f.Row = fieldUpdate.Row AndAlso
                                                                         f.Col = fieldUpdate.Col AndAlso
                                                                         String.Equals(f.FieldId, fieldUpdate.FieldId, StringComparison.OrdinalIgnoreCase))

                        If existing IsNot Nothing Then
                            existing.Value = fieldUpdate.Value
                            existing.OptionText = fieldUpdate.OptionText
                            existing.OptionValue = fieldUpdate.OptionValue
                            existing.OptionIndex = fieldUpdate.OptionIndex
                        Else
                            target.FieldUpdates.Add(New FormFieldUpdate() With {
                                .Row = fieldUpdate.Row,
                                .Col = fieldUpdate.Col,
                                .FieldId = fieldUpdate.FieldId,
                                .Value = fieldUpdate.Value,
                                .OptionText = fieldUpdate.OptionText,
                                .OptionValue = fieldUpdate.OptionValue,
                                .OptionIndex = fieldUpdate.OptionIndex
                            })
                        End If
                    Next
                End If

                If update.NewRows IsNot Nothing Then
                    For Each newRow In update.NewRows
                        target.NewRows.Add(New List(Of String)(newRow))
                    Next
                End If
            Next

            Return mergedByTable.Values.OrderBy(Function(u) u.TableIndex).ToList()
        End Function

        Private Function BuildTableCoverageSystemPrompt() As String
            Return "You are performing a coverage pass for table completion." & vbCrLf &
                "You will receive:" & vbCrLf &
                "- CONSISTENCY SUMMARY" & vbCrLf &
                "- DOCUMENT CONTEXT" & vbCrLf &
                "- ALL TABLES" & vbCrLf &
                "- UNRESOLVED TARGETS: writable empty cells and writable empty fields that were not explicitly addressed in the first pass." & vbCrLf & vbCrLf &
                "Your task is to address every unresolved target." & vbCrLf &
                "- For each unresolved empty cell, return an update with action=""complete"", ""clear"", or ""keep""." & vbCrLf &
                "- For each unresolved empty field, return a field update." & vbCrLf &
                "- Do not leave unresolved targets unanswered." & vbCrLf &
                "- Prefer filling existing empty rows before suggesting newRows." & vbCrLf &
                "- If additional logical entries are required and existing rows are exhausted, add newRows." & vbCrLf & vbCrLf &
                "Return ONLY a JSON array in the same schema as the main table completion response."
        End Function

        Private Function BuildTableBlocks(tables As List(Of TableInfo)) As List(Of List(Of TableInfo))
            Const maxBlockChars As Integer = 16000

            Dim blocks As New List(Of List(Of TableInfo))()
            If tables Is Nothing OrElse tables.Count = 0 Then Return blocks

            Dim currentBlock As New List(Of TableInfo)()
            Dim currentLength As Integer = 2

            For Each tbl In tables
                Dim json As String = TableToJson(tbl)
                Dim additionalLength As Integer = If(currentBlock.Count = 0, 0, 1) + json.Length

                If currentBlock.Count > 0 AndAlso currentLength + additionalLength > maxBlockChars Then
                    blocks.Add(currentBlock)
                    currentBlock = New List(Of TableInfo)()
                    currentLength = 2
                End If

                currentBlock.Add(tbl)
                currentLength += additionalLength
            Next

            If currentBlock.Count > 0 Then
                blocks.Add(currentBlock)
            End If

            Return blocks
        End Function

        Private Function BuildBodySectionBlocks(bodySections As List(Of BodySectionInfo)) As List(Of List(Of BodySectionInfo))
            Const maxBlockChars As Integer = 14000

            Dim blocks As New List(Of List(Of BodySectionInfo))()
            If bodySections Is Nothing OrElse bodySections.Count = 0 Then Return blocks

            Dim currentBlock As New List(Of BodySectionInfo)()
            Dim currentLength As Integer = 2

            For Each section In bodySections
                Dim json As String = BodySectionToJson(section)
                Dim additionalLength As Integer = If(currentBlock.Count = 0, 0, 1) + json.Length

                If currentBlock.Count > 0 AndAlso currentLength + additionalLength > maxBlockChars Then
                    blocks.Add(currentBlock)
                    currentBlock = New List(Of BodySectionInfo)()
                    currentLength = 2
                End If

                currentBlock.Add(section)
                currentLength += additionalLength
            Next

            If currentBlock.Count > 0 Then
                blocks.Add(currentBlock)
            End If

            Return blocks
        End Function

        Private Async Function BuildConsistencySummaryJson(tables As List(Of TableInfo),
                                                           bodySections As List(Of BodySectionInfo),
                                                           userInstructions As String,
                                                           documentContext As String,
                                                           Optional useSecondAPI As Boolean = False) As Task(Of String)
            Dim summaryPromptBuilder As New StringBuilder()
            summaryPromptBuilder.AppendLine($"USER INSTRUCTIONS: {userInstructions}")
            summaryPromptBuilder.AppendLine()

            If Not String.IsNullOrWhiteSpace(documentContext) Then
                summaryPromptBuilder.AppendLine("[DOCUMENT CONTEXT]")
                summaryPromptBuilder.AppendLine(documentContext)
                summaryPromptBuilder.AppendLine("[/DOCUMENT CONTEXT]")
                summaryPromptBuilder.AppendLine()
            End If

            summaryPromptBuilder.AppendLine("[ALL TABLES]")
            summaryPromptBuilder.AppendLine(BuildAllTablesJson(tables))
            summaryPromptBuilder.AppendLine("[/ALL TABLES]")
            summaryPromptBuilder.AppendLine()

            summaryPromptBuilder.AppendLine("[BODY SECTIONS]")
            summaryPromptBuilder.AppendLine(BuildAllBodySectionsJson(bodySections))
            summaryPromptBuilder.AppendLine("[/BODY SECTIONS]")

            Dim summaryResponse As String = Await SharedMethods.LLM(
                _owner._context,
                _owner.InterpolateAtRuntime(BuildConsistencySummarySystemPrompt()),
                summaryPromptBuilder.ToString(),
                "", "", 0, useSecondAPI, True, cancellationToken:=_ct)

            Return NormalizeConsistencySummaryJson(summaryResponse)
        End Function

        Private Async Function ProcessCompletionBatches(tables As List(Of TableInfo),
                                                        bodySections As List(Of BodySectionInfo),
                                                        userInstructions As String,
                                                        documentContext As String,
                                                        Optional useSecondAPI As Boolean = False) As Task(Of CompletionBatchesResult)
            Dim tableBlocks As List(Of List(Of TableInfo)) = BuildTableBlocks(tables)
            Dim bodyBlocks As List(Of List(Of BodySectionInfo)) = BuildBodySectionBlocks(bodySections)

            Dim totalWorkBlocks As Integer = tableBlocks.Count + bodyBlocks.Count
            SetProgress("Building form consistency summary", 0, Math.Max(1, totalWorkBlocks + 1))

            If IsCancelled() Then Return Nothing

            Dim consistencySummaryJson As String =
                Await BuildConsistencySummaryJson(tables, bodySections, userInstructions, documentContext, useSecondAPI)

            SetProgress("Building form consistency summary", 1, Math.Max(1, totalWorkBlocks + 1))

            Dim result As New CompletionBatchesResult() With {
                .TableUpdates = New List(Of TableUpdateResponse)(),
                .BodySectionUpdates = New List(Of BodySectionUpdateResponse)()
            }

            Dim processedBlocks As Integer = 0

            For i As Integer = 0 To tableBlocks.Count - 1
                If IsCancelled() Then Return Nothing

                SetProgress($"Processing block {processedBlocks + 1} of {totalWorkBlocks} (tables)",
                            1 + processedBlocks,
                            Math.Max(1, totalWorkBlocks + 1))

                Dim blockUpdates As List(Of TableUpdateResponse) =
                    Await ProcessSingleTableBlock(tableBlocks(i), userInstructions, documentContext, consistencySummaryJson, useSecondAPI)

                If blockUpdates Is Nothing Then Return Nothing
                result.TableUpdates.AddRange(blockUpdates)

                processedBlocks += 1
            Next

            For i As Integer = 0 To bodyBlocks.Count - 1
                If IsCancelled() Then Return Nothing

                SetProgress($"Processing block {processedBlocks + 1} of {totalWorkBlocks} (body)",
                            1 + processedBlocks,
                            Math.Max(1, totalWorkBlocks + 1))

                Dim blockUpdates As List(Of BodySectionUpdateResponse) =
                    Await ProcessSingleBodyBlock(bodyBlocks(i), userInstructions, documentContext, consistencySummaryJson, useSecondAPI)

                If blockUpdates Is Nothing Then Return Nothing
                result.BodySectionUpdates.AddRange(blockUpdates)

                processedBlocks += 1
            Next

            Return result
        End Function

        Private Async Function ProcessSingleTableBlock(tableBlock As List(Of TableInfo),
                                                       userInstructions As String,
                                                       documentContext As String,
                                                       consistencySummaryJson As String,
                                                       Optional useSecondAPI As Boolean = False) As Task(Of List(Of TableUpdateResponse))
            Dim promptBuilder As New StringBuilder()
            promptBuilder.AppendLine($"USER INSTRUCTIONS: {userInstructions}")
            promptBuilder.AppendLine()

            promptBuilder.AppendLine("[CONSISTENCY SUMMARY]")
            promptBuilder.AppendLine(consistencySummaryJson)
            promptBuilder.AppendLine("[/CONSISTENCY SUMMARY]")
            promptBuilder.AppendLine()

            If Not String.IsNullOrWhiteSpace(documentContext) Then
                promptBuilder.AppendLine("[DOCUMENT CONTEXT]")
                promptBuilder.AppendLine(documentContext)
                promptBuilder.AppendLine("[/DOCUMENT CONTEXT]")
                promptBuilder.AppendLine()
            End If

            promptBuilder.AppendLine("[ALL TABLES]")
            promptBuilder.AppendLine(BuildAllTablesJson(tableBlock))
            promptBuilder.AppendLine("[/ALL TABLES]")

            Dim response As String = Await SharedMethods.LLM(
                _owner._context,
                _owner.InterpolateAtRuntime(BuildTableCompletionSystemPrompt()),
                promptBuilder.ToString(),
                "", "", 0, useSecondAPI, True, cancellationToken:=_ct)

            If String.IsNullOrWhiteSpace(response) Then
                LogWarn("LLM returned empty response for a table block.")
                Return Nothing
            End If

            Dim firstPassUpdates As List(Of TableUpdateResponse) = ParseTableCompletionResponse(response)
            If firstPassUpdates Is Nothing Then Return Nothing

            Dim unresolvedTargetsJson As String = BuildUnresolvedTableTargetsJson(tableBlock, firstPassUpdates)
            If unresolvedTargetsJson = "[]" Then
                Return firstPassUpdates
            End If

            Dim coveragePromptBuilder As New StringBuilder()
            coveragePromptBuilder.AppendLine($"USER INSTRUCTIONS: {userInstructions}")
            coveragePromptBuilder.AppendLine()

            coveragePromptBuilder.AppendLine("[CONSISTENCY SUMMARY]")
            coveragePromptBuilder.AppendLine(consistencySummaryJson)
            coveragePromptBuilder.AppendLine("[/CONSISTENCY SUMMARY]")
            coveragePromptBuilder.AppendLine()

            If Not String.IsNullOrWhiteSpace(documentContext) Then
                coveragePromptBuilder.AppendLine("[DOCUMENT CONTEXT]")
                coveragePromptBuilder.AppendLine(documentContext)
                coveragePromptBuilder.AppendLine("[/DOCUMENT CONTEXT]")
                coveragePromptBuilder.AppendLine()
            End If

            coveragePromptBuilder.AppendLine("[ALL TABLES]")
            coveragePromptBuilder.AppendLine(BuildAllTablesJson(tableBlock))
            coveragePromptBuilder.AppendLine("[/ALL TABLES]")
            coveragePromptBuilder.AppendLine()

            coveragePromptBuilder.AppendLine("[UNRESOLVED TARGETS]")
            coveragePromptBuilder.AppendLine(unresolvedTargetsJson)
            coveragePromptBuilder.AppendLine("[/UNRESOLVED TARGETS]")

            Dim coverageResponse As String = Await SharedMethods.LLM(
                _owner._context,
                _owner.InterpolateAtRuntime(BuildTableCoverageSystemPrompt()),
                coveragePromptBuilder.ToString(),
                "", "", 0, useSecondAPI, True, cancellationToken:=_ct)

            If String.IsNullOrWhiteSpace(coverageResponse) Then
                Return firstPassUpdates
            End If

            Dim secondPassUpdates As List(Of TableUpdateResponse) = ParseTableCompletionResponse(coverageResponse)
            If secondPassUpdates Is Nothing Then
                Return firstPassUpdates
            End If

            Return MergeTableUpdateResponses(firstPassUpdates, secondPassUpdates)
        End Function

        Private Async Function ProcessSingleBodyBlock(bodyBlock As List(Of BodySectionInfo),
                                                      userInstructions As String,
                                                      documentContext As String,
                                                      consistencySummaryJson As String,
                                                      Optional useSecondAPI As Boolean = False) As Task(Of List(Of BodySectionUpdateResponse))
            Dim promptBuilder As New StringBuilder()
            promptBuilder.AppendLine($"USER INSTRUCTIONS: {userInstructions}")
            promptBuilder.AppendLine()

            promptBuilder.AppendLine("[CONSISTENCY SUMMARY]")
            promptBuilder.AppendLine(consistencySummaryJson)
            promptBuilder.AppendLine("[/CONSISTENCY SUMMARY]")
            promptBuilder.AppendLine()

            If Not String.IsNullOrWhiteSpace(documentContext) Then
                promptBuilder.AppendLine("[DOCUMENT CONTEXT]")
                promptBuilder.AppendLine(documentContext)
                promptBuilder.AppendLine("[/DOCUMENT CONTEXT]")
                promptBuilder.AppendLine()
            End If

            promptBuilder.AppendLine("[BODY SECTIONS]")
            promptBuilder.AppendLine(BuildAllBodySectionsJson(bodyBlock))
            promptBuilder.AppendLine("[/BODY SECTIONS]")

            Dim response As String = Await SharedMethods.LLM(
                _owner._context,
                _owner.InterpolateAtRuntime(BuildBodyCompletionSystemPrompt()),
                promptBuilder.ToString(),
                "", "", 0, useSecondAPI, True, cancellationToken:=_ct)

            If String.IsNullOrWhiteSpace(response) Then
                LogWarn("LLM returned empty response for a body block.")
                Return Nothing
            End If

            Return ParseBodyCompletionResponse(response)
        End Function

        Private Function BuildBodyCompletionSystemPrompt() As String
            Return "You are a professional document assistant that completes and revises body-level content and form fields in a Word document." & vbCrLf &
                "You will receive:" & vbCrLf &
                "- CONSISTENCY SUMMARY" & vbCrLf &
                "- DOCUMENT CONTEXT" & vbCrLf &
                "- BODY SECTIONS: JSON sections containing nearby heading text, paragraphs, and body-level form fields." & vbCrLf & vbCrLf &
                "IMPORTANT:" & vbCrLf &
                "- Do NOT rely on heuristic placeholder flags. Decide from the actual paragraph text, nearby headings, and document context whether a paragraph contains a real value, a sample/example, instructional text, or content that should be completed/replaced." & vbCrLf &
                "- Preserve meaningful existing text unless there is a clear reason to replace it." & vbCrLf &
                "- Every writable empty paragraph and every writable empty field in the provided body sections must be explicitly addressed." & vbCrLf & vbCrLf &
                "Each body section has:" & vbCrLf &
                "- ""sectionIndex"": unique identifier" & vbCrLf &
                "- ""headingText"": nearby context heading" & vbCrLf &
                "- ""paragraphs"": array with:" & vbCrLf &
                "  - ""index"": paragraph identifier within the section" & vbCrLf &
                "  - ""text"": current plain visible text" & vbCrLf &
                "  - ""empty"": empty flag" & vbCrLf &
                "- ""fields"": array with the same field schema used for tables" & vbCrLf & vbCrLf &
                "For each paragraph update include an ""action"":" & vbCrLf &
                "- ""keep"": keep existing text unchanged" & vbCrLf &
                "- ""replace"": replace existing text" & vbCrLf &
                "- ""complete"": fill an empty/incomplete/example/sample paragraph with a final value" & vbCrLf &
                "- ""clear"": clear the paragraph text" & vbCrLf & vbCrLf &
                "Coverage requirements:" & vbCrLf &
                "- Every writable empty paragraph must appear in ""paragraphUpdates""." & vbCrLf &
                "- Every writable empty field must appear in ""fieldUpdates""." & vbCrLf &
                "- If you intentionally leave an empty paragraph unchanged, return it with action=""keep""." & vbCrLf & vbCrLf &
                "For dropdowns/legacydropdowns prefer optionValue, then optionText, then optionIndex." & vbCrLf &
                "For checkboxes use val=true/false." & vbCrLf &
                "For text inputs use val=text." & vbCrLf & vbCrLf &
                "Return ONLY a JSON array:" & vbCrLf &
                "[" & vbCrLf &
                "  {" & vbCrLf &
                "    ""sectionIndex"": <int>," & vbCrLf &
                "    ""paragraphUpdates"": [{""index"": <int>, ""action"": ""replace|complete|clear|keep"", ""text"": ""new text""}]," & vbCrLf &
                "    ""fieldUpdates"": [{""id"": ""f0"", ""val"": ""text"", ""optionText"": ""display text"", ""optionValue"": ""stored value"", ""optionIndex"": <int>}]" & vbCrLf &
                "  }" & vbCrLf &
                "]"
        End Function

        Private Function BuildConsistencySummarySystemPrompt() As String
            Return "You are a professional document analyst. " &
                "Your task is to read the full document context and all extracted tables, then produce a compact JSON summary " &
                "that captures the consistency rules the table-completion model should follow across the entire form." & vbCrLf & vbCrLf &
                "Return ONLY a single JSON object, no markdown fences, no commentary." & vbCrLf & vbCrLf &
                "The JSON object should be compact and may contain these properties:" & vbCrLf &
                "- ""documentType"": short description of the form/document type" & vbCrLf &
                "- ""theme"": short description of the project or business topic" & vbCrLf &
                "- ""terminology"": array of canonical terms that should be used consistently" & vbCrLf &
                "- ""entities"": array of important names/roles/concepts inferred from the form" & vbCrLf &
                "- ""styleRules"": array of rules about tone, wording, brevity, and formatting conventions" & vbCrLf &
                "- ""crossTableRules"": array of consistency rules that should apply across multiple tables" & vbCrLf &
                "- ""dropdownConventions"": array of preferred selection patterns for dropdowns when the form implies a consistent choice pattern" & vbCrLf &
                "- ""dateConventions"": array of rules for date style or scheduling consistency" & vbCrLf &
                "- ""numericConventions"": array of rules for numbers, priorities, scales, or scoring" & vbCrLf &
                "- ""warnings"": array of cautions about ambiguity or places where consistency matters" & vbCrLf & vbCrLf &
                "Keep the summary concise, practical, and reusable by a second model pass. " &
                "Do not include raw table data. Do not repeat the full input."
        End Function

        Private Function BuildTableCompletionSystemPrompt() As String
            Return "You are a professional document assistant that completes and revises tables across an entire form or document. " &
                "You will receive:" & vbCrLf &
                "- CONSISTENCY SUMMARY: a compact JSON summary of rules and conventions inferred from the full form." & vbCrLf &
                "- DOCUMENT CONTEXT: the full document body text with [Table N] markers showing where each table sits." & vbCrLf &
                "- ALL TABLES: a JSON array containing extracted tables in the document." & vbCrLf & vbCrLf &
                "IMPORTANT:" & vbCrLf &
                "- Do NOT rely on heuristic placeholder flags. Decide from the actual content, context, headings, and consistency rules whether a cell contains a real value, a sample/example, instructional text, or content that should be completed/replaced." & vbCrLf &
                "- Existing text may be meaningful, a sample value, or a placeholder. You must decide." & vbCrLf &
                "- Preserve meaningful existing text unless there is a clear reason to replace it." & vbCrLf &
                "- However, every writable empty cell and every writable empty form field in the provided tables must be explicitly addressed." & vbCrLf &
                "- Do not leave empty gaps between populated data rows unless the table structure clearly requires a blank separator row." & vbCrLf &
                "- Prefer filling existing empty rows before adding newRows." & vbCrLf &
                "- Add newRows when the existing visible data rows are insufficient to complete the logical sequence of entries required by the table." & vbCrLf & vbCrLf &
                "JSON SCHEMA FOR INPUT CELLS:" & vbCrLf &
                "- ""c"": column index" & vbCrLf &
                "- ""t"": current text" & vbCrLf &
                "- ""empty"": true means the cell is empty after normalization" & vbCrLf &
                "- ""span"": optional column span" & vbCrLf &
                "- ""vm"": optional vertical merge indicator" & vbCrLf &
                "- ""hdr"": header cell" & vbCrLf &
                "- ""w"": approximate width percent" & vbCrLf &
                "- ""ro"": read-only" & vbCrLf &
                "- ""fields"": optional form field array" & vbCrLf & vbCrLf &
                "ROW INDEX RULE:" & vbCrLf &
                "- Row indices in ""updates"" and ""fieldUpdates"" always refer to the ORIGINAL extracted rows shown in the input JSON for that table, before any ""newRows"" are inserted." & vbCrLf &
                "- ""insertAfterRow"" also refers to the ORIGINAL extracted row index." & vbCrLf & vbCrLf &
                "TASK:" & vbCrLf &
                "For each cell you choose to mention in ""updates"", include an ""action"":" & vbCrLf &
                "- ""keep"": keep existing text unchanged" & vbCrLf &
                "- ""replace"": replace existing text" & vbCrLf &
                "- ""complete"": fill an empty/incomplete/example/sample cell with a final value" & vbCrLf &
                "- ""clear"": clear the cell text" & vbCrLf & vbCrLf &
                "Coverage requirements:" & vbCrLf &
                "- Every writable empty cell in the provided tables must either appear in ""updates"" or be covered by a logically necessary ""newRows"" decision." & vbCrLf &
                "- Every writable empty field must appear in ""fieldUpdates""." & vbCrLf &
                "- If you intentionally leave an empty writable cell unchanged, return it in ""updates"" with action=""keep""." & vbCrLf &
                "- If you populate a later empty row in a data table, do not leave earlier writable empty data rows blank unless they are intentionally blank; in that case return them with action=""keep""." & vbCrLf & vbCrLf &
                "For dropdowns and legacydropdowns, prefer optionValue, then optionText, then optionIndex." & vbCrLf &
                "For text inputs use val." & vbCrLf &
                "For checkboxes use val=true/false." & vbCrLf &
                "Never modify anything marked ro=true." & vbCrLf & vbCrLf &
                "RESPONSE FORMAT — return ONLY a JSON array. Each element:" & vbCrLf &
                "{" & vbCrLf &
                "  ""tableIndex"": <int>," & vbCrLf &
                "  ""updates"": [{""r"": <row>, ""c"": <col>, ""action"": ""replace|complete|clear|keep"", ""t"": ""new text""}]," & vbCrLf &
                "  ""fieldUpdates"": [{""r"": <row>, ""c"": <col>, ""id"": ""f0"", ""val"": ""text"", ""optionText"": ""display text"", ""optionValue"": ""stored value"", ""optionIndex"": <int>}]," & vbCrLf &
                "  ""newRows"": [[""cell1"", ""cell2"", ...]]," & vbCrLf &
                "  ""insertAfterRow"": <int>" & vbCrLf &
                "}" & vbCrLf & vbCrLf &
                "Return ONLY the JSON array, no explanations or markdown fences." &
                " {Ignore} {Location} {CurrentDate}"
        End Function

        Private Function NormalizeConsistencySummaryJson(response As String) As String
            If System.String.IsNullOrWhiteSpace(response) Then
                Throw New System.InvalidOperationException("The consistency analysis returned no JSON object; the document was not changed.")
            End If

            Try
                Dim normalized As String = response.Trim()

                If normalized.StartsWith("```") Then
                    Dim firstNewline As Integer = normalized.IndexOf(vbLf)
                    If firstNewline > 0 Then normalized = normalized.Substring(firstNewline + 1)
                    If normalized.EndsWith("```") Then normalized = normalized.Substring(0, normalized.Length - 3)
                    normalized = normalized.Trim()
                End If

                Dim parsed As JObject = JObject.Parse(normalized)
                Dim compactJson As String = parsed.ToString(Newtonsoft.Json.Formatting.None)

                Return compactJson
            Catch ex As System.Exception
                WriteTableCompletionDebug("NormalizeConsistencySummaryJson error", ex.ToString() & vbCrLf & vbCrLf & response)
                Throw New System.InvalidOperationException("The consistency analysis returned invalid JSON; the document was not changed.", ex)
            End Try
        End Function

#End Region

#Region "Response Parsing"

        Private Function ParseBodyCompletionResponse(response As String) As List(Of BodySectionUpdateResponse)
            Dim results As New List(Of BodySectionUpdateResponse)()

            Try
                Dim payload As String = ExtractJsonArrayPayload(response)
                If String.IsNullOrWhiteSpace(payload) Then
                    Throw New InvalidOperationException("No JSON array payload found.")
                End If

                Dim parsed As JArray = JArray.Parse(payload)

                For Each item As JToken In parsed
                    If item.Type <> JTokenType.Object Then Continue For
                    Dim jObj As JObject = CType(item, JObject)

                    If jObj("sectionIndex") Is Nothing Then Continue For

                    Dim update As New BodySectionUpdateResponse() With {
                        .SectionIndex = CInt(jObj("sectionIndex")),
                        .ParagraphUpdates = New List(Of BodyParagraphUpdate)(),
                        .FieldUpdates = New List(Of BodyFieldUpdate)()
                    }

                    Dim paragraphUpdatesToken As JToken = jObj("paragraphUpdates")
                    If paragraphUpdatesToken IsNot Nothing AndAlso paragraphUpdatesToken.Type = JTokenType.Array Then
                        For Each paraItem As JToken In CType(paragraphUpdatesToken, JArray)
                            If paraItem.Type <> JTokenType.Object Then Continue For
                            Dim paraObj As JObject = CType(paraItem, JObject)

                            If paraObj("index") Is Nothing Then Continue For

                            update.ParagraphUpdates.Add(New BodyParagraphUpdate() With {
                                .Index = CInt(paraObj("index")),
                                .Action = NormalizeRequestedTextAction(If(paraObj("action")?.ToString(), "replace")),
                                .Text = If(paraObj("text")?.ToString(), "")
                            })
                        Next
                    End If

                    Dim fieldUpdatesToken As JToken = jObj("fieldUpdates")
                    If fieldUpdatesToken IsNot Nothing AndAlso fieldUpdatesToken.Type = JTokenType.Array Then
                        For Each fuItem As JToken In CType(fieldUpdatesToken, JArray)
                            If fuItem.Type <> JTokenType.Object Then Continue For
                            Dim fuObj As JObject = CType(fuItem, JObject)

                            If fuObj("id") Is Nothing Then Continue For

                            Dim fu As New BodyFieldUpdate() With {
                                .FieldId = CStr(fuObj("id")),
                                .Value = If(fuObj("val")?.ToString(), Nothing),
                                .OptionText = If(fuObj("optionText")?.ToString(), Nothing),
                                .OptionValue = If(fuObj("optionValue")?.ToString(), Nothing),
                                .OptionIndex = -1
                            }

                            If fuObj("optionIndex") IsNot Nothing Then
                                fu.OptionIndex = CInt(fuObj("optionIndex"))
                            End If

                            update.FieldUpdates.Add(fu)
                        Next
                    End If

                    results.Add(update)
                Next

            Catch ex As Exception
                LogWarn($"Failed to parse body completion response as JSON: {ex.Message}")
                Return Nothing
            End Try

            Return results
        End Function

        Private Function ParseTableCompletionResponse(response As String) As List(Of TableUpdateResponse)
            Dim results As New List(Of TableUpdateResponse)()

            Try
                Dim payload As String = ExtractJsonArrayPayload(response)
                If String.IsNullOrWhiteSpace(payload) Then
                    Throw New InvalidOperationException("No JSON array payload found.")
                End If

                Dim parsed As JArray = JArray.Parse(payload)

                For Each item As JToken In parsed
                    If item.Type <> JTokenType.Object Then Continue For
                    Dim jObj As JObject = CType(item, JObject)

                    Dim update As New TableUpdateResponse() With {
                        .Updates = New List(Of TableCellUpdate)(),
                        .NewRows = New List(Of List(Of String))(),
                        .FieldUpdates = New List(Of FormFieldUpdate)(),
                        .InsertAfterRow = -1
                    }

                    If jObj("tableIndex") IsNot Nothing Then update.TableIndex = CInt(jObj("tableIndex"))
                    If jObj("insertAfterRow") IsNot Nothing Then update.InsertAfterRow = CInt(jObj("insertAfterRow"))

                    Dim updatesToken As JToken = jObj("updates")
                    If updatesToken IsNot Nothing AndAlso updatesToken.Type = JTokenType.Array Then
                        For Each updItem As JToken In CType(updatesToken, JArray)
                            If updItem.Type <> JTokenType.Object Then Continue For
                            Dim updObj As JObject = CType(updItem, JObject)

                            If updObj("r") Is Nothing OrElse updObj("c") Is Nothing Then Continue For

                            update.Updates.Add(New TableCellUpdate() With {
                                .Row = CInt(updObj("r")),
                                .Col = CInt(updObj("c")),
                                .Action = NormalizeRequestedTextAction(If(updObj("action")?.ToString(), "replace")),
                                .Text = If(updObj("t")?.ToString(), "")
                            })
                        Next
                    End If

                    Dim fieldUpdatesToken As JToken = jObj("fieldUpdates")
                    If fieldUpdatesToken IsNot Nothing AndAlso fieldUpdatesToken.Type = JTokenType.Array Then
                        For Each fuItem As JToken In CType(fieldUpdatesToken, JArray)
                            If fuItem.Type <> JTokenType.Object Then Continue For
                            Dim fuObj As JObject = CType(fuItem, JObject)

                            If fuObj("r") Is Nothing OrElse fuObj("c") Is Nothing OrElse fuObj("id") Is Nothing Then Continue For

                            Dim fu As New FormFieldUpdate() With {
                                .Row = CInt(fuObj("r")),
                                .Col = CInt(fuObj("c")),
                                .FieldId = CStr(fuObj("id")),
                                .Value = If(fuObj("val")?.ToString(), Nothing),
                                .OptionText = If(fuObj("optionText")?.ToString(), Nothing),
                                .OptionValue = If(fuObj("optionValue")?.ToString(), Nothing),
                                .OptionIndex = -1
                            }

                            If fuObj("optionIndex") IsNot Nothing Then
                                fu.OptionIndex = CInt(fuObj("optionIndex"))
                            End If

                            update.FieldUpdates.Add(fu)
                        Next
                    End If

                    Dim newRowsToken As JToken = jObj("newRows")
                    If newRowsToken IsNot Nothing AndAlso newRowsToken.Type = JTokenType.Array Then
                        For Each rowItem As JToken In CType(newRowsToken, JArray)
                            If rowItem.Type <> JTokenType.Array Then Continue For

                            Dim rowTexts As New List(Of String)()
                            For Each cellVal As JToken In CType(rowItem, JArray)
                                rowTexts.Add(If(cellVal IsNot Nothing, cellVal.ToString(), ""))
                            Next

                            update.NewRows.Add(rowTexts)
                        Next
                    End If

                    results.Add(update)
                Next

            Catch ex As Exception
                LogWarn($"Failed to parse LLM response as JSON: {ex.Message}")
                Return Nothing
            End Try

            Return results
        End Function

#End Region

#Region "Apply Updates"

        Private Function NormalizeRequestedTextAction(action As String) As String
            Dim normalized As String = If(action, "").Trim().ToLowerInvariant()

            Select Case normalized
                Case "keep", "replace", "complete", "clear"
                    Return normalized
                Case Else
                    Return "replace"
            End Select
        End Function

        Private Function FindNonPlaceholderRunProperties(scope As XmlNode,
                                                         nsMgr As XmlNamespaceManager) As XmlNode
            If scope Is Nothing Then Return Nothing

            Dim searchScope As XmlNode = scope
            While searchScope IsNot Nothing AndAlso
                  searchScope.LocalName <> "tc" AndAlso
                  searchScope.LocalName <> "body"
                searchScope = searchScope.ParentNode
            End While
            If searchScope Is Nothing Then searchScope = scope

            Dim rPrNodes As XmlNodeList = searchScope.SelectNodes(".//w:r/w:rPr", nsMgr)
            For Each rPr As XmlNode In rPrNodes
                Dim rStyle As XmlNode = rPr.SelectSingleNode("w:rStyle/@w:val", nsMgr)
                If rStyle IsNot Nothing Then
                    Dim val As String = rStyle.Value
                    If val.Equals("PlaceholderText", StringComparison.OrdinalIgnoreCase) OrElse
                       val.Equals("Platzhaltertext", StringComparison.OrdinalIgnoreCase) Then
                        Continue For
                    End If
                End If
                Return rPr
            Next

            Return Nothing
        End Function

        Private Function CreateCleanTextRun(parentParagraph As XmlNode,
                                            text As String,
                                            nsMgr As XmlNamespaceManager) As XmlNode
            Dim ownerDoc As XmlDocument = parentParagraph.OwnerDocument
            Dim rNode As XmlElement = ownerDoc.CreateElement("w", "r", WNs)

            Dim templateRpr As XmlNode = FindNonPlaceholderRunProperties(parentParagraph, nsMgr)
            If templateRpr IsNot Nothing Then
                rNode.AppendChild(templateRpr.CloneNode(deep:=True))
            End If

            Dim tNode As XmlElement = ownerDoc.CreateElement("w", "t", WNs)
            rNode.AppendChild(tNode)
            SetTextNodeWithSpacePreserve(tNode, If(text, ""))

            Return rNode
        End Function

        Private Sub SetTextNodeWithSpacePreserve(textNode As XmlNode, text As String)
            If textNode Is Nothing Then Exit Sub

            Dim safeText As String = If(text, "")
            textNode.InnerText = safeText

            Dim xmlSpaceAttr As XmlAttribute = Nothing
            Try
                xmlSpaceAttr = CType(textNode.Attributes("xml:space"), XmlAttribute)
            Catch
            End Try

            Dim requiresPreserve As Boolean =
                safeText.Length > 0 AndAlso
                (safeText.StartsWith(" ", StringComparison.Ordinal) OrElse
                 safeText.EndsWith(" ", StringComparison.Ordinal) OrElse
                 safeText.Contains("  "))

            If requiresPreserve Then
                If xmlSpaceAttr Is Nothing Then
                    xmlSpaceAttr = textNode.OwnerDocument.CreateAttribute("xml", "space", "http://www.w3.org/XML/1998/namespace")
                    textNode.Attributes.Append(xmlSpaceAttr)
                End If

                xmlSpaceAttr.Value = "preserve"
            ElseIf xmlSpaceAttr IsNot Nothing Then
                Try
                    textNode.Attributes.Remove(xmlSpaceAttr)
                Catch
                End Try
            End If
        End Sub

        Private Sub WriteTextIntoParagraph(paragraph As BodyParagraphInfo,
                                           newText As String,
                                           nsMgr As XmlNamespaceManager)
            If paragraph Is Nothing OrElse paragraph.XmlNode Is Nothing Then Exit Sub

            Dim normalized As String = If(newText, "")
            Dim currentText As String = If(paragraph.Text, "")
            If String.Equals(currentText, normalized, StringComparison.Ordinal) Then Exit Sub

            Dim pNode As XmlNode = paragraph.XmlNode

            ClearAllAncestorSdtPlaceholderStates(pNode, nsMgr)
            SanitizeCellRunFormatting(pNode, nsMgr)
            SanitizeParagraphProperties(pNode, nsMgr)

            If paragraph.TextNodes IsNot Nothing AndAlso paragraph.TextNodes.Count > 0 Then
                SetTextNodeWithSpacePreserve(paragraph.TextNodes(0), normalized)
                For i As Integer = 1 To paragraph.TextNodes.Count - 1
                    SetTextNodeWithSpacePreserve(paragraph.TextNodes(i), "")
                Next
            Else
                Dim toRemove As New List(Of XmlNode)()
                For Each ch As XmlNode In pNode.ChildNodes
                    If ch.NodeType = XmlNodeType.Element AndAlso ch.LocalName = "pPr" Then Continue For
                    toRemove.Add(ch)
                Next
                For Each ch As XmlNode In toRemove
                    pNode.RemoveChild(ch)
                Next

                Dim newRun As XmlNode = CreateCleanTextRun(pNode, normalized, nsMgr)
                pNode.AppendChild(newRun)

                Dim newTextNode As XmlNode = newRun.SelectSingleNode("w:t", nsMgr)
                paragraph.TextNodes = New List(Of XmlNode)()
                If newTextNode IsNot Nothing Then
                    paragraph.TextNodes.Add(newTextNode)
                End If
            End If

            paragraph.Text = normalized
            paragraph.IsEmpty = String.IsNullOrWhiteSpace(NormalizeTextForPlaceholderDetection(normalized))
            paragraph.IsPlaceholder = IsPlaceholderText(normalized)
        End Sub

        Private Sub ApplyBodySectionUpdates(bodySections As List(Of BodySectionInfo),
                                            updates As List(Of BodySectionUpdateResponse),
                                            nsMgr As XmlNamespaceManager)
            If bodySections Is Nothing OrElse updates Is Nothing Then Exit Sub

            For Each update In updates
                If IsCancelled() Then Exit Sub

                Dim section As BodySectionInfo = bodySections.FirstOrDefault(Function(s) s.SectionIndex = update.SectionIndex)
                If section Is Nothing Then Continue For

                If update.ParagraphUpdates IsNot Nothing Then
                    For Each paragraphUpdate In update.ParagraphUpdates
                        If IsCancelled() Then Exit Sub

                        Dim action As String = NormalizeRequestedTextAction(paragraphUpdate.Action)
                        If action = "keep" Then Continue For

                        Dim paragraph As BodyParagraphInfo =
                            section.Paragraphs.FirstOrDefault(Function(p) p.Index = paragraphUpdate.Index)

                        If paragraph Is Nothing Then Continue For

                        Dim targetText As String = If(action = "clear", "", If(paragraphUpdate.Text, ""))
                        WriteTextIntoParagraph(paragraph, targetText, nsMgr)
                    Next
                End If

                If update.FieldUpdates IsNot Nothing AndAlso section.FormFields IsNot Nothing Then
                    For Each bodyFieldUpdate In update.FieldUpdates
                        If IsCancelled() Then Exit Sub

                        Dim ff As FormFieldInfo =
                            section.FormFields.FirstOrDefault(Function(f) f.FieldId = bodyFieldUpdate.FieldId)

                        If ff Is Nothing OrElse Not ff.IsWritable Then Continue For

                        Dim mappedUpdate As New FormFieldUpdate() With {
                            .FieldId = bodyFieldUpdate.FieldId,
                            .Value = bodyFieldUpdate.Value,
                            .OptionText = bodyFieldUpdate.OptionText,
                            .OptionValue = bodyFieldUpdate.OptionValue,
                            .OptionIndex = bodyFieldUpdate.OptionIndex
                        }

                        ApplyFormFieldUpdate(ff, mappedUpdate, nsMgr)
                    Next
                End If
            Next
        End Sub

        Private Sub ApplyTableUpdates(tables As List(Of TableInfo),
                                      updates As List(Of TableUpdateResponse),
                                      nsMgr As XmlNamespaceManager)
            If tables Is Nothing OrElse updates Is Nothing Then Exit Sub

            For Each update In updates
                If IsCancelled() Then Exit Sub

                Dim tbl As TableInfo = tables.FirstOrDefault(Function(t) t.TableIndex = update.TableIndex)
                If tbl Is Nothing Then Continue For

                Dim templateRowClone As XmlNode = Nothing
                If update.NewRows IsNot Nothing AndAlso update.NewRows.Count > 0 Then
                    templateRowClone = CloneTemplateRowBeforeApply(tbl, update.InsertAfterRow)
                End If

                If update.Updates IsNot Nothing Then
                    For Each cellUpdate In update.Updates
                        If IsCancelled() Then Exit Sub

                        Dim action As String = NormalizeRequestedTextAction(cellUpdate.Action)
                        If action = "keep" Then Continue For

                        Dim row As TableRowInfo = tbl.Rows.FirstOrDefault(Function(r) r.RowIndex = cellUpdate.Row)
                        If row Is Nothing Then Continue For

                        Dim cell As TableCellInfo = row.Cells.FirstOrDefault(Function(c) c.Col = cellUpdate.Col)
                        If cell Is Nothing Then Continue For
                        If Not cell.IsWritable Then Continue For

                        Dim targetText As String = If(action = "clear", "", If(cellUpdate.Text, ""))
                        Dim currentText As String = If(cell.Text, "")

                        If String.Equals(currentText, targetText, StringComparison.Ordinal) Then Continue For

                        WriteTextIntoCell(cell, targetText, nsMgr)
                    Next
                End If

                If update.FieldUpdates IsNot Nothing Then
                    For Each fu In update.FieldUpdates
                        If IsCancelled() Then Exit Sub

                        Dim row As TableRowInfo = tbl.Rows.FirstOrDefault(Function(r) r.RowIndex = fu.Row)
                        If row Is Nothing Then Continue For

                        Dim cell As TableCellInfo = row.Cells.FirstOrDefault(Function(c) c.Col = fu.Col)
                        If cell Is Nothing OrElse cell.FormFields Is Nothing Then Continue For

                        Dim ff As FormFieldInfo = cell.FormFields.FirstOrDefault(Function(f) f.FieldId = fu.FieldId)
                        If ff Is Nothing Then Continue For
                        If Not ff.IsWritable Then Continue For

                        ApplyFormFieldUpdate(ff, fu, nsMgr)
                    Next
                End If

                If update.NewRows IsNot Nothing AndAlso update.NewRows.Count > 0 Then
                    If tbl.Rows.Any(Function(r) r.Cells.Any(Function(c) Not c.IsWritable)) Then
                        Continue For
                    End If

                    InsertNewTableRows(tbl, update.NewRows, update.InsertAfterRow, nsMgr, templateRowClone)
                End If
            Next
        End Sub

        Private Sub ClearSdtPlaceholderState(sdtNode As XmlNode, nsMgr As XmlNamespaceManager)
            Dim showingNode As XmlNode = sdtNode.SelectSingleNode("w:sdtPr/w:showingPlcHdr", nsMgr)
            If showingNode IsNot Nothing Then
                showingNode.ParentNode.RemoveChild(showingNode)
            End If

            Dim sdtContent As XmlNode = sdtNode.SelectSingleNode("w:sdtContent", nsMgr)
            If sdtContent Is Nothing Then Exit Sub

            Dim rStyleNodes As XmlNodeList = sdtContent.SelectNodes(
                ".//w:rPr/w:rStyle[@w:val='PlaceholderText' or @w:val='Platzhaltertext']", nsMgr)
            For Each rStyleNode As XmlNode In rStyleNodes
                Dim rPr As XmlNode = rStyleNode.ParentNode
                rPr.RemoveChild(rStyleNode)
                If rPr.ChildNodes.Count = 0 Then
                    rPr.ParentNode.RemoveChild(rPr)
                End If
            Next
        End Sub

        Private Sub ApplyFormFieldUpdate(ff As FormFieldInfo,
                                         fieldUpdate As FormFieldUpdate,
                                         nsMgr As XmlNamespaceManager)
            If ff Is Nothing OrElse fieldUpdate Is Nothing Then Exit Sub
            If Not ff.IsWritable Then Exit Sub

            Dim newValue As String = fieldUpdate.Value

            Select Case ff.FieldType

                Case FormFieldType.Dropdown
                    Dim selectedOption As FormFieldOptionInfo = ResolveUpdatedDropdownOption(ff, fieldUpdate)
                    If selectedOption Is Nothing Then Exit Sub

                    WriteTextIntoSdtContent(ff.XmlNode, selectedOption.DisplayText, nsMgr)
                    ApplySdtDefaultRunProperties(ff.XmlNode, nsMgr)

                    Dim sdtPr As XmlNode = ff.XmlNode.SelectSingleNode("w:sdtPr", nsMgr)
                    If sdtPr IsNot Nothing Then
                        Dim listNode As XmlNode = sdtPr.SelectSingleNode("w:dropDownList", nsMgr)
                        If listNode Is Nothing Then listNode = sdtPr.SelectSingleNode("w:comboBox", nsMgr)

                        If listNode IsNot Nothing Then
                            Dim lastValAttr As XmlAttribute = listNode.OwnerDocument.CreateAttribute("w", "lastValue", WNs)
                            lastValAttr.Value = selectedOption.StoredValue
                            listNode.Attributes.SetNamedItem(lastValAttr)
                        End If
                    End If

                Case FormFieldType.Checkbox
                    If String.IsNullOrEmpty(newValue) Then Exit Sub

                    Dim sdtNode As XmlNode = ff.XmlNode
                    Dim checkboxNode As XmlNode = sdtNode.SelectSingleNode("w:sdtPr/w14:checkbox", nsMgr)
                    If checkboxNode Is Nothing Then Exit Sub

                    Dim isChecked As Boolean = (newValue.ToLowerInvariant() = "true" OrElse newValue = "1")
                    Dim xmlVal As String = If(isChecked, "1", "0")

                    Dim checkedEl As XmlNode = checkboxNode.SelectSingleNode("w14:checked", nsMgr)
                    If checkedEl Is Nothing Then
                        checkedEl = checkboxNode.OwnerDocument.CreateElement("w14", "checked", W14Ns)
                        checkboxNode.AppendChild(checkedEl)
                    End If

                    Dim valAttr As XmlAttribute = checkedEl.OwnerDocument.CreateAttribute("w14", "val", W14Ns)
                    valAttr.Value = xmlVal
                    checkedEl.Attributes.SetNamedItem(valAttr)

                    Dim charCode As String
                    If isChecked Then
                        Dim checkedState As XmlNode = checkboxNode.SelectSingleNode("w14:checkedState/@w14:val", nsMgr)
                        charCode = If(checkedState?.Value, "2612")
                    Else
                        Dim uncheckedState As XmlNode = checkboxNode.SelectSingleNode("w14:uncheckedState/@w14:val", nsMgr)
                        charCode = If(uncheckedState?.Value, "2610")
                    End If

                    WriteTextIntoSdtContent(sdtNode, ChrW(System.Convert.ToInt32(charCode, 16)).ToString(), nsMgr)
                    ApplySdtDefaultRunProperties(sdtNode, nsMgr)

                Case FormFieldType.TextInput
                    If newValue Is Nothing Then Exit Sub
                    WriteTextIntoSdtContent(ff.XmlNode, newValue, nsMgr)
                    ApplySdtDefaultRunProperties(ff.XmlNode, nsMgr)

                Case FormFieldType.LegacyDropdown
                    Dim selectedOption As FormFieldOptionInfo = ResolveUpdatedDropdownOption(ff, fieldUpdate)
                    If selectedOption Is Nothing Then Exit Sub

                    Dim fldCharNode As XmlNode = ff.XmlNode
                    Dim ffData As XmlNode = fldCharNode.SelectSingleNode("w:ffData", nsMgr)
                    If ffData Is Nothing Then Exit Sub

                    Dim ddList As XmlNode = ffData.SelectSingleNode("w:ddList", nsMgr)
                    If ddList Is Nothing Then Exit Sub

                    Dim targetIdx As Integer = -1
                    For i As Integer = 0 To ff.Options.Count - 1
                        If ff.Options(i).DisplayText.Equals(selectedOption.DisplayText, StringComparison.OrdinalIgnoreCase) AndAlso
                           ff.Options(i).StoredValue.Equals(selectedOption.StoredValue, StringComparison.OrdinalIgnoreCase) Then
                            targetIdx = i
                            Exit For
                        End If
                    Next

                    If targetIdx < 0 Then Exit Sub

                    Dim resultNode As XmlNode = ddList.SelectSingleNode("w:result", nsMgr)
                    If resultNode Is Nothing Then
                        resultNode = ddList.OwnerDocument.CreateElement("w", "result", WNs)
                        ddList.AppendChild(resultNode)
                    End If

                    Dim resultValAttr As XmlAttribute = resultNode.OwnerDocument.CreateAttribute("w", "val", WNs)
                    resultValAttr.Value = targetIdx.ToString()
                    resultNode.Attributes.SetNamedItem(resultValAttr)

                    UpdateLegacyFieldDisplayText(fldCharNode, selectedOption.DisplayText, nsMgr)

                Case FormFieldType.LegacyCheckbox
                    If String.IsNullOrEmpty(newValue) Then Exit Sub

                    Dim fldCharNode As XmlNode = ff.XmlNode
                    Dim ffData As XmlNode = fldCharNode.SelectSingleNode("w:ffData", nsMgr)
                    If ffData Is Nothing Then Exit Sub

                    Dim checkBox As XmlNode = ffData.SelectSingleNode("w:checkBox", nsMgr)
                    If checkBox Is Nothing Then Exit Sub

                    Dim isChecked As Boolean = (newValue.ToLowerInvariant() = "true" OrElse newValue = "1")

                    Dim checkedNode As XmlNode = checkBox.SelectSingleNode("w:checked", nsMgr)
                    If checkedNode Is Nothing Then
                        checkedNode = checkBox.OwnerDocument.CreateElement("w", "checked", WNs)
                        checkBox.AppendChild(checkedNode)
                    End If

                    Dim checkedValAttr As XmlAttribute = checkedNode.OwnerDocument.CreateAttribute("w", "val", WNs)
                    checkedValAttr.Value = If(isChecked, "1", "0")
                    checkedNode.Attributes.SetNamedItem(checkedValAttr)

            End Select
        End Sub

        Private Function ResolveUpdatedDropdownOption(ff As FormFieldInfo, fieldUpdate As FormFieldUpdate) As FormFieldOptionInfo
            If ff Is Nothing OrElse ff.Options Is Nothing OrElse ff.Options.Count = 0 OrElse fieldUpdate Is Nothing Then Return Nothing

            If Not String.IsNullOrWhiteSpace(fieldUpdate.OptionValue) Then
                Dim optionByStoredValue As FormFieldOptionInfo =
                    ff.Options.FirstOrDefault(Function(o) o.StoredValue.Equals(fieldUpdate.OptionValue, StringComparison.OrdinalIgnoreCase))
                If optionByStoredValue IsNot Nothing Then Return optionByStoredValue
            End If

            If Not String.IsNullOrWhiteSpace(fieldUpdate.OptionText) Then
                Dim optionByDisplayText As FormFieldOptionInfo =
                    ff.Options.FirstOrDefault(Function(o) o.DisplayText.Equals(fieldUpdate.OptionText, StringComparison.OrdinalIgnoreCase))
                If optionByDisplayText IsNot Nothing Then Return optionByDisplayText
            End If

            If Not String.IsNullOrWhiteSpace(fieldUpdate.Value) Then
                Dim optionByValueAsStored As FormFieldOptionInfo =
                    ff.Options.FirstOrDefault(Function(o) o.StoredValue.Equals(fieldUpdate.Value, StringComparison.OrdinalIgnoreCase))
                If optionByValueAsStored IsNot Nothing Then Return optionByValueAsStored

                Dim optionByValueAsDisplay As FormFieldOptionInfo =
                    ff.Options.FirstOrDefault(Function(o) o.DisplayText.Equals(fieldUpdate.Value, StringComparison.OrdinalIgnoreCase))
                If optionByValueAsDisplay IsNot Nothing Then Return optionByValueAsDisplay
            End If

            If fieldUpdate.OptionIndex >= 0 AndAlso fieldUpdate.OptionIndex < ff.Options.Count Then
                Return ff.Options(fieldUpdate.OptionIndex)
            End If

            Return Nothing
        End Function

        Private Sub UpdateLegacyFieldDisplayText(fldCharBeginRun As XmlNode, newText As String, nsMgr As XmlNamespaceManager)
            Dim beginRun As XmlNode = fldCharBeginRun.ParentNode
            If beginRun Is Nothing Then Exit Sub

            Dim currentNode As XmlNode = beginRun.NextSibling
            Dim foundSeparate As Boolean = False
            Dim textUpdated As Boolean = False

            While currentNode IsNot Nothing
                If currentNode.LocalName = "r" Then
                    Dim fc As XmlNode = currentNode.SelectSingleNode("w:fldChar", nsMgr)
                    If fc IsNot Nothing Then
                        Dim fcType As XmlNode = fc.Attributes("w:fldCharType")
                        If fcType IsNot Nothing Then
                            If fcType.Value = "separate" Then
                                foundSeparate = True
                            ElseIf fcType.Value = "end" Then
                                Exit While
                            End If
                        End If
                    ElseIf foundSeparate Then
                        Dim tNodes As XmlNodeList = currentNode.SelectNodes("w:t", nsMgr)
                        If tNodes.Count > 0 AndAlso Not textUpdated Then
                            SetTextNodeWithSpacePreserve(tNodes(0), newText)
                            textUpdated = True
                            For i As Integer = 1 To tNodes.Count - 1
                                SetTextNodeWithSpacePreserve(tNodes(i), "")
                            Next
                        ElseIf tNodes.Count > 0 Then
                            For Each tn As XmlNode In tNodes
                                SetTextNodeWithSpacePreserve(tn, "")
                            Next
                        End If
                    End If
                End If
                currentNode = currentNode.NextSibling
            End While
        End Sub

        Private Sub InsertNewTableRows(tbl As TableInfo,
                                       newRows As List(Of List(Of String)),
                                       insertAfterRow As Integer,
                                       nsMgr As XmlNamespaceManager,
                                       Optional templateRowClone As XmlNode = Nothing)
            If tbl Is Nothing OrElse newRows Is Nothing OrElse newRows.Count = 0 Then Exit Sub
            If IsCancelled() Then Exit Sub

            Dim templateRowNode As XmlNode = templateRowClone
            If templateRowNode Is Nothing Then
                templateRowNode = CloneTemplateRowBeforeApply(tbl, insertAfterRow)
            End If
            If templateRowNode Is Nothing Then Exit Sub

            If RowContainsVerticalMerge(templateRowNode, nsMgr) Then
                Exit Sub
            End If

            Dim refNode As XmlNode
            If insertAfterRow >= 0 AndAlso insertAfterRow < tbl.Rows.Count Then
                refNode = tbl.Rows(insertAfterRow).XmlNode
            Else
                refNode = tbl.Rows(tbl.Rows.Count - 1).XmlNode
            End If
            If refNode Is Nothing Then Exit Sub

            Dim insertParent As XmlNode = refNode.ParentNode
            If insertParent Is Nothing Then Exit Sub

            For Each rowTexts In newRows
                If IsCancelled() Then Exit Sub

                Dim clonedRow As XmlNode = templateRowNode.CloneNode(deep:=True)

                Dim tblHeaderNode As XmlNode = clonedRow.SelectSingleNode("w:trPr/w:tblHeader", nsMgr)
                If tblHeaderNode IsNot Nothing Then
                    tblHeaderNode.ParentNode.RemoveChild(tblHeaderNode)
                End If

                Dim vMergeNodes As XmlNodeList = clonedRow.SelectNodes(".//w:tcPr/w:vMerge", nsMgr)
                For Each vmNode As XmlNode In vMergeNodes
                    vmNode.ParentNode.RemoveChild(vmNode)
                Next

                For Each innerSdt As XmlNode In clonedRow.SelectNodes(".//w:sdt", nsMgr)
                    ClearSdtPlaceholderState(innerSdt, nsMgr)
                Next

                Dim tcNodes As XmlNodeList = clonedRow.SelectNodes(".//w:tc", nsMgr)
                For ci As Integer = 0 To tcNodes.Count - 1
                    Dim tcNode As XmlNode = tcNodes(ci)
                    Dim cellText As String = If(ci < rowTexts.Count, rowTexts(ci), "")

                    Dim pNode As XmlNode = tcNode.SelectSingleNode("w:p", nsMgr)
                    If pNode Is Nothing Then
                        pNode = tcNode.OwnerDocument.CreateElement("w", "p", WNs)
                        tcNode.AppendChild(pNode)
                    End If

                    Dim toRemove As New List(Of XmlNode)()
                    For Each child As XmlNode In pNode.ChildNodes
                        If child.NodeType = XmlNodeType.Element AndAlso child.LocalName = "pPr" Then Continue For
                        toRemove.Add(child)
                    Next
                    For Each child As XmlNode In toRemove
                        pNode.RemoveChild(child)
                    Next

                    pNode.AppendChild(CreateCleanTextRun(pNode, cellText, nsMgr))
                Next

                insertParent.InsertAfter(clonedRow, refNode)
                refNode = clonedRow
            Next
        End Sub

        Private Sub WriteTextIntoCell(cell As TableCellInfo,
                                      newText As String,
                                      nsMgr As XmlNamespaceManager)
            If cell Is Nothing OrElse cell.XmlNode Is Nothing Then Exit Sub

            Dim normalized As String = If(newText, "")
            Dim currentText As String = If(cell.Text, "")
            If String.Equals(currentText, normalized, StringComparison.Ordinal) Then Exit Sub

            Dim tcNode As XmlNode = cell.XmlNode

            ClearAllAncestorSdtPlaceholderStates(tcNode, nsMgr)
            SanitizeCellRunFormatting(tcNode, nsMgr)

            If cell.TextNodes IsNot Nothing AndAlso cell.TextNodes.Count > 0 Then
                SetTextNodeWithSpacePreserve(cell.TextNodes(0), normalized)
                For i As Integer = 1 To cell.TextNodes.Count - 1
                    SetTextNodeWithSpacePreserve(cell.TextNodes(i), "")
                Next
                cell.Text = normalized
                Exit Sub
            End If

            Dim pNode As XmlNode = tcNode.SelectSingleNode("w:p", nsMgr)
            If pNode Is Nothing Then
                pNode = tcNode.OwnerDocument.CreateElement("w", "p", WNs)
                tcNode.AppendChild(pNode)
            End If

            Dim childrenToRemove As New List(Of XmlNode)()
            For Each ch As XmlNode In pNode.ChildNodes
                If ch.NodeType = XmlNodeType.Element AndAlso ch.LocalName = "pPr" Then Continue For
                childrenToRemove.Add(ch)
            Next
            For Each ch As XmlNode In childrenToRemove
                pNode.RemoveChild(ch)
            Next

            SanitizeParagraphProperties(pNode, nsMgr)

            Dim newRun As XmlNode = CreateCleanTextRun(pNode, normalized, nsMgr)
            pNode.AppendChild(newRun)

            Dim newTextNode As XmlNode = newRun.SelectSingleNode("w:t", nsMgr)
            If newTextNode IsNot Nothing Then
                If cell.TextNodes Is Nothing Then cell.TextNodes = New List(Of XmlNode)()
                cell.TextNodes.Add(newTextNode)
            End If

            cell.Text = normalized
        End Sub

        Private Sub WriteTextIntoSdtContent(sdtNode As XmlNode,
                                            newText As String,
                                            nsMgr As XmlNamespaceManager)
            If sdtNode Is Nothing Then Exit Sub

            ClearAllAncestorSdtPlaceholderStates(sdtNode, nsMgr)

            Dim sdtContent As XmlNode = sdtNode.SelectSingleNode("w:sdtContent", nsMgr)
            If sdtContent Is Nothing Then Exit Sub

            SanitizeCellRunFormatting(sdtContent, nsMgr)

            For Each pInside As XmlNode In sdtContent.SelectNodes(".//w:p", nsMgr)
                SanitizeParagraphProperties(pInside, nsMgr)
            Next

            Dim contentTextNodes As XmlNodeList = sdtContent.SelectNodes(".//w:t", nsMgr)
            If contentTextNodes.Count > 0 Then
                SetTextNodeWithSpacePreserve(contentTextNodes(0), If(newText, ""))
                For i As Integer = 1 To contentTextNodes.Count - 1
                    SetTextNodeWithSpacePreserve(contentTextNodes(i), "")
                Next
                Exit Sub
            End If

            Dim ownerDoc As XmlDocument = sdtContent.OwnerDocument
            Dim sdtParentName As String = If(sdtNode.ParentNode?.LocalName, "")

            sdtContent.InnerXml = ""

            If sdtParentName = "tc" OrElse sdtParentName = "body" Then
                Dim pNode As XmlElement = ownerDoc.CreateElement("w", "p", WNs)
                sdtContent.AppendChild(pNode)
                pNode.AppendChild(CreateCleanTextRun(pNode, If(newText, ""), nsMgr))
            Else
                Dim hostParagraph As XmlNode = sdtNode.ParentNode
                While hostParagraph IsNot Nothing AndAlso hostParagraph.LocalName <> "p"
                    hostParagraph = hostParagraph.ParentNode
                End While
                If hostParagraph Is Nothing Then hostParagraph = sdtContent
                sdtContent.AppendChild(CreateCleanTextRun(hostParagraph, If(newText, ""), nsMgr))
            End If
        End Sub

        Private Sub ClearAllAncestorSdtPlaceholderStates(startNode As XmlNode,
                                                         nsMgr As XmlNamespaceManager)
            If startNode Is Nothing Then Exit Sub

            Dim current As XmlNode = startNode
            While current IsNot Nothing
                If current.NodeType = XmlNodeType.Element AndAlso current.LocalName = "sdt" Then
                    ClearSdtPlaceholderState(current, nsMgr)
                End If
                current = current.ParentNode
            End While
        End Sub

        Private Sub SanitizeCellRunFormatting(scope As XmlNode,
                                              nsMgr As XmlNamespaceManager)
            If scope Is Nothing Then Exit Sub

            Dim rPrNodes As XmlNodeList = scope.SelectNodes(".//w:rPr", nsMgr)
            For Each rPr As XmlNode In rPrNodes
                Dim toRemove As New List(Of XmlNode)()

                For Each child As XmlNode In rPr.ChildNodes
                    Select Case child.LocalName
                        Case "vanish", "webHidden"
                            toRemove.Add(child)

                        Case "rStyle"
                            Dim valAttr As XmlNode = child.Attributes("w:val")
                            If valAttr IsNot Nothing AndAlso
                               (valAttr.Value.Equals("Platzhaltertext", StringComparison.OrdinalIgnoreCase) OrElse
                                valAttr.Value.Equals("PlaceholderText", StringComparison.OrdinalIgnoreCase)) Then
                                toRemove.Add(child)
                            End If
                    End Select
                Next

                For Each n As XmlNode In toRemove
                    rPr.RemoveChild(n)
                Next

                If rPr.ChildNodes.Count = 0 AndAlso rPr.ParentNode IsNot Nothing Then
                    rPr.ParentNode.RemoveChild(rPr)
                End If
            Next
        End Sub

        Private Sub SanitizeParagraphProperties(pNode As XmlNode,
                                                nsMgr As XmlNamespaceManager)
            If pNode Is Nothing Then Exit Sub
            Dim pPr As XmlNode = pNode.SelectSingleNode("w:pPr", nsMgr)
            If pPr Is Nothing Then Exit Sub

            Dim toRemove As New List(Of XmlNode)()
            For Each pStyle As XmlNode In pPr.SelectNodes("w:pStyle", nsMgr)
                Dim valAttr As XmlNode = pStyle.Attributes("w:val")
                If valAttr Is Nothing Then Continue For
                If valAttr.Value.Equals("Platzhaltertext", StringComparison.OrdinalIgnoreCase) OrElse
                   valAttr.Value.Equals("PlaceholderText", StringComparison.OrdinalIgnoreCase) Then
                    toRemove.Add(pStyle)
                End If
            Next
            For Each n As XmlNode In toRemove
                pPr.RemoveChild(n)
            Next
        End Sub

#End Region

#Region "Logging / Debug / Host Shims"

        Private Function IsCancelled() As Boolean
            Return _ct.IsCancellationRequested
        End Function

        Private Sub LogInfo(message As String)
            If String.IsNullOrWhiteSpace(message) Then Exit Sub
            Try : _context.Log(message) : Catch : End Try
            Try : _owner.ApDashboardLog(message, "step") : Catch : End Try
        End Sub

        Private Sub LogWarn(message As String)
            If String.IsNullOrWhiteSpace(message) Then Exit Sub
            Try : _context.Log(message, "warn") : Catch : End Try
            Try : _owner.ApDashboardLog(message, "warn") : Catch : End Try
        End Sub

        Private Sub SetProgress(label As String, current As Integer, total As Integer)
            Dim safeTotal As Integer = Math.Max(1, total)
            Dim safeCurrent As Integer = Math.Max(0, Math.Min(current, safeTotal))
            LogInfo($"{label} ({safeCurrent}/{safeTotal})")
        End Sub

        Private Sub StartTableCompletionDebugSession(inputPath As String, outputPath As String, userInstructions As String)
            If Not TableCompleteDebugEnabled Then Exit Sub

            Dim sb As New StringBuilder()
            sb.AppendLine(New String("="c, 80))
            sb.AppendLine($"UTC: {Date.UtcNow:yyyy-MM-dd HH:mm:ss.fff}")
            sb.AppendLine($"Input: {inputPath}")
            sb.AppendLine($"Output: {outputPath}")
            sb.AppendLine("User instructions:")
            sb.AppendLine(userInstructions)
            sb.AppendLine(New String("="c, 80))

            WriteTableCompletionDebug("Session start", sb.ToString())
        End Sub

        Private Sub WriteTableCompletionDebug(sectionName As String, content As String)
            If Not TableCompleteDebugEnabled Then Exit Sub

            Try
                Dim logPath As String = GetTableCompletionDebugLogPath()

                SyncLock TableCompleteDebugSyncRoot
                    Dim sb As New StringBuilder()
                    sb.AppendLine()
                    sb.AppendLine($"[{Date.UtcNow:yyyy-MM-dd HH:mm:ss.fff}] {sectionName}")
                    sb.AppendLine(New String("-"c, 80))
                    sb.AppendLine(If(content, ""))
                    sb.AppendLine(New String("-"c, 80))
                    File.AppendAllText(logPath, sb.ToString(), Encoding.UTF8)
                End SyncLock
            Catch ex As Exception
                Debug.WriteLine($"WriteTableCompletionDebug failed: {ex.Message}")
            End Try
        End Sub

        Private Function GetTableCompletionDebugLogPath() As String
            Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
            Return Path.Combine(desktopPath, TableCompleteDebugLogFileName)
        End Function

        Private Function BuildFormFieldDebugReport(tables As List(Of TableInfo)) As String
            If tables Is Nothing Then Return ""
            Dim sb As New StringBuilder()

            For Each tbl In tables
                sb.AppendLine($"Table {tbl.TableIndex}")

                For Each row In tbl.Rows
                    For Each cell In row.Cells
                        If cell.FormFields Is Nothing OrElse cell.FormFields.Count = 0 Then Continue For

                        sb.AppendLine($"  Cell r={cell.Row}, c={cell.Col}, text='{cell.Text}'")

                        For Each ff In cell.FormFields
                            sb.AppendLine($"    Field {ff.FieldId} | type={ff.FieldType} | name='{ff.Name}' | display='{ff.CurrentValue}' | stored='{ff.CurrentOptionValue}'")
                        Next
                    Next
                Next
            Next

            Return sb.ToString()
        End Function

        Private Sub DumpDocxPackageDiagnostics(docxPath As String)
            If Not TableCompleteDebugEnabled Then Exit Sub
        End Sub

        Private Sub VerifySavedDocxCellUpdates(docxPath As String, updates As List(Of TableUpdateResponse))
            If Not TableCompleteDebugEnabled Then Exit Sub
        End Sub

        Private Sub VerifySavedDocxWithWordInterop(docxPath As String)
            If Not TableCompleteDebugEnabled Then Exit Sub
        End Sub

#End Region

    End Class

End Class