' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.CompleteTables.vb
' Purpose: Completes Word document tables using AI by extracting table structure
'          to JSON, having the LLM fill/expand them, then patching OpenXML.
'
' Architecture / Key Ideas:
'  - OpenXML Processing: Operates directly on DOCX XML, modifying only <w:t>
'    nodes within <w:tc> cells and cloning <w:tr> rows for insertions.
'  - JSON Round-Trip: Extracts table structure (headings, row labels, cell text,
'    merged cell spans, column widths) into compact JSON, sends to LLM, parses
'    response JSON, and applies changes back to XML preserving all formatting.
'  - Form Field Support: Detects and exposes structured document tags (SDTs)
'    including dropdown/comboBox selections and checkbox states. The LLM can
'    select dropdown items by exact option value and toggle checkboxes.
'    Content control plain-text fields are also extracted and modifiable.
'    Legacy form fields (w:ffData) with dropdown lists and checkboxes are
'    also extracted and can be set by the LLM.
'  - Full Document Context: The LLM receives the entire document body text
'    (paragraphs between and around tables) so it can understand document
'    purpose, fill tables consistently, and detect placeholders anywhere.
'  - Placeholder Detection: Cells containing patterns like [...], ___, TBD,
'    <enter ...>, N/A, or similar are flagged with "ph":true so the LLM
'    knows they require completion.
'  - Column Width Hints: Relative column widths (as percentages) are included
'    so the LLM can gauge how much text fits in each cell.
'  - Row Cloning: New rows are created by deep-cloning the last data row of the
'    table, preserving cell properties (w:tcPr), paragraph properties (w:pPr),
'    run properties (w:rPr), borders, widths, and shading.
'  - Batch Processing: Large documents with many tables are processed in batches
'    to stay within LLM token limits.
'  - Formatting Preservation: Only text content changes; all styles, borders,
'    widths, merged cell definitions, and document structure are untouched.
'  - Prompt Persistence: The user's last instructions are saved/restored via
'    My.Settings.TableComplete_LastInstructions.
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
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Maximum characters of document context to include per batch.
    ''' </summary>
    Private Const TableCompleteMaxContextChars As Integer = 12000

    ''' <summary>
    ''' Maximum characters of the generated form-consistency summary included in the final prompt.
    ''' </summary>
    Private Const TableCompleteMaxConsistencySummaryChars As Integer = 4000

    ''' <summary>
    ''' XML namespace URI for w14 extensions (checkbox content controls).
    ''' </summary>
    Private Const W14Ns As String = "http://schemas.microsoft.com/office/word/2010/wordml"

    ''' <summary>
    ''' XML namespace URI for the main wordprocessingml namespace.
    ''' </summary>
    Private Const WNs As String = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    ''' <summary>
    ''' Enables verbose table-completion debugging.
    ''' </summary>
    Private Const TableCompleteDebugEnabled As Boolean = False

    ''' <summary>
    ''' Includes raw XML payloads in debug logs.
    ''' </summary>
    Private Const TableCompleteDebugIncludeXml As Boolean = False

    ''' <summary>
    ''' Includes full DOCX package diagnostics in debug logs.
    ''' </summary>
    Private Const TableCompleteDebugIncludePackageDump As Boolean = False

    ''' <summary>
    ''' Logs detailed diagnostics for plain table-cell apply operations.
    ''' </summary>
    Private Const TableCompleteDebugLogCellApply As Boolean = False

    ''' <summary>
    ''' Reopens the saved DOCX after repacking and verifies that requested cell updates are present.
    ''' </summary>
    Private Const TableCompleteDebugVerifySavedDocx As Boolean = False

    ''' <summary>
    ''' Reopens the saved DOCX in hidden Word and logs what Word sees in the target table.
    ''' </summary>
    Private Const TableCompleteDebugVerifySavedDocxWithWordInterop As Boolean = False

    ''' <summary>
    ''' Desktop log filename for table-completion debugging.
    ''' </summary>
    Private Const TableCompleteDebugLogFileName As String = "RI_TableCompletion_Debug.log"

    Private Shared ReadOnly TableCompleteDebugSyncRoot As New Object()

    ''' <summary>
    ''' Regex pattern matching common placeholder content in table cells.
    ''' Matches: [...], [TBD], ___, …, TBD, N/A, n/a, &lt;enter ...&gt;, {placeholder}, etc.
    ''' </summary>
    Private Shared ReadOnly PlaceholderPattern As New Regex(
        "^\s*(\[(?!\s*\d+\s*\]$)[^\]\r\n]+\]|\{[^}\r\n]*\}|<[^>\r\n]*>|_{3,}|\.{3,}|…+|TBD|TBC|N/?A|TODO|FIXME|XXX|INSERT|ENTER)\s*$",
        RegexOptions.IgnoreCase Or RegexOptions.Compiled Or RegexOptions.CultureInvariant)

#Region "Data Classes"

    ''' <summary>
    ''' Represents a single cell extracted from a table for JSON serialization.
    ''' </summary>
    Private Class TableCellInfo
        Public Property Row As Integer
        Public Property Col As Integer
        Public Property Text As String
        Public Property GridSpan As Integer
        Public Property VMerge As String
        Public Property TextNodes As List(Of System.Xml.XmlNode)
        Public Property IsHeader As Boolean
        Public Property IsPlaceholder As Boolean
        Public Property WidthPct As Double
        ''' <summary>The effective w:tc node for this cell, whether direct or wrapper-owned.</summary>
        Public Property XmlNode As System.Xml.XmlNode
        ''' <summary>Whether the cell content may be modified.</summary>
        Public Property IsWritable As Boolean
        ''' <summary>Optional reason why the cell is not writable.</summary>
        Public Property ProtectionReason As String
        ''' <summary>Form fields (SDT and legacy) found inside or owning this cell.</summary>
        Public Property FormFields As List(Of FormFieldInfo)
    End Class

    ''' <summary>
    ''' Represents a table row extracted from the document.
    ''' </summary>
    Private Class TableRowInfo
        Public Property RowIndex As Integer
        Public Property Cells As List(Of TableCellInfo)
        Public Property XmlNode As System.Xml.XmlNode
    End Class

    Private Class CompletionBatchesResult
        Public Property TableUpdates As List(Of TableUpdateResponse)
        Public Property BodySectionUpdates As List(Of BodySectionUpdateResponse)
    End Class

    ''' <summary>
    ''' Represents a complete table extracted from the document.
    ''' </summary>
    Private Class TableInfo
        Public Property TableIndex As Integer
        Public Property Rows As List(Of TableRowInfo)
        Public Property XmlNode As System.Xml.XmlNode
        Public Property ColumnCount As Integer
        Public Property ContextBefore As String
        Public Property ColumnWidths As List(Of Double)
    End Class

    ''' <summary>
    ''' Represents a cell update or new row from the LLM response.
    ''' </summary>
    Private Class TableCellUpdate
        Public Property Row As Integer
        Public Property Col As Integer
        Public Property Text As String
        Public Property Action As String
    End Class


    ''' <summary>
    ''' Represents the LLM's response for a single table.
    ''' </summary>
    Private Class TableUpdateResponse
        Public Property TableIndex As Integer
        Public Property Updates As List(Of TableCellUpdate)
        Public Property NewRows As List(Of List(Of String))
        Public Property InsertAfterRow As Integer
        Public Property FieldUpdates As List(Of FormFieldUpdate)
    End Class

    ''' <summary>
    ''' The type of a form field found in the document.
    ''' </summary>
    Private Enum FormFieldType
        ''' <summary>SDT dropdown or comboBox content control.</summary>
        Dropdown
        ''' <summary>SDT checkbox content control (w14:checkbox).</summary>
        Checkbox
        ''' <summary>SDT plain-text or rich-text content control.</summary>
        TextInput
        ''' <summary>Legacy form field dropdown (w:ffData/w:ddList).</summary>
        LegacyDropdown
        ''' <summary>Legacy form field checkbox (w:ffData/w:checkBox).</summary>
        LegacyCheckbox
    End Enum

    ''' <summary>
    ''' Represents one selectable option for a dropdown or combo box.
    ''' </summary>
    Private Class FormFieldOptionInfo
        Public Property DisplayText As String
        Public Property StoredValue As String
    End Class

    ''' <summary>
    ''' Extracted information about a form field (SDT or legacy).
    ''' </summary>
    Private Class FormFieldInfo
        ''' <summary>Unique ID within the cell for LLM round-trip ("f0", "f1", ...).</summary>
        Public Property FieldId As String
        Public Property FieldType As FormFieldType
        ''' <summary>Tag or alias of the content control (w:sdtPr/w:tag or w:sdtPr/w:alias), or legacy field name.</summary>
        Public Property Name As String
        ''' <summary>Current display value shown in Word.</summary>
        Public Property CurrentValue As String
        ''' <summary>Current stored/internal value used by dropdown-like controls.</summary>
        Public Property CurrentOptionValue As String
        ''' <summary>Available selectable options for the field.</summary>
        Public Property Options As List(Of FormFieldOptionInfo)
        ''' <summary>The w:sdt node (for SDTs) or the w:fldChar node (for legacy fields).</summary>
        Public Property XmlNode As System.Xml.XmlNode
        ''' <summary>Whether the field may be modified.</summary>
        Public Property IsWritable As Boolean
        ''' <summary>Optional reason why the field is not writable.</summary>
        Public Property ProtectionReason As String
    End Class

    ''' <summary>
    ''' An LLM-requested change to a form field.
    ''' </summary>
    Private Class FormFieldUpdate
        Public Property Row As Integer
        Public Property Col As Integer
        Public Property FieldId As String
        Public Property Value As String
        Public Property OptionText As String
        Public Property OptionValue As String
        Public Property OptionIndex As Integer
    End Class

    ''' <summary>
    ''' Represents a section of body-level content (paragraphs between tables) that
    ''' may contain placeholders, empty text, or form fields that need completion.
    ''' </summary>
    Private Class BodySectionInfo
        ''' <summary>Sequential index among all body sections.</summary>
        Public Property SectionIndex As Integer
        ''' <summary>Heading or context text preceding the section paragraphs.</summary>
        Public Property HeadingText As String
        ''' <summary>Paragraphs in this section that are candidates for completion.</summary>
        Public Property Paragraphs As List(Of BodyParagraphInfo)
        ''' <summary>Body-level SDT form fields in this section.</summary>
        Public Property FormFields As List(Of FormFieldInfo)
    End Class

    ''' <summary>
    ''' A single body-level paragraph that may be empty or contain placeholder text.
    ''' </summary>
    Private Class BodyParagraphInfo
        Public Property Index As Integer
        Public Property Text As String
        Public Property IsPlaceholder As Boolean
        Public Property IsEmpty As Boolean
        Public Property TextNodes As List(Of System.Xml.XmlNode)
        Public Property XmlNode As System.Xml.XmlNode
    End Class

    ''' <summary>
    ''' LLM response for body section updates.
    ''' </summary>
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

#End Region

#Region "Entry Point and Pipeline"

    ''' <summary>
    ''' Entry point: prompts for file and user instructions, then completes tables.
    ''' </summary>
    Public Async Sub directCompleteWordDocumentTables()
        Dim selectedPath As String = ""

        If INI_AllowLegacyDocFiles Then
            Globals.ThisAddIn.DragDropFormLabel = "Select a Word document with tables to complete"
            Globals.ThisAddIn.DragDropFormFilter = "Word Documents|*.doc;*.docx|Word Document (*.docx)|*.docx|Word 97-2003 (*.doc)|*.doc"
        Else
            Globals.ThisAddIn.DragDropFormLabel = "Select a Word document with tables to complete"
            Globals.ThisAddIn.DragDropFormFilter = "Word Documents (*.docx)|*.docx"
        End If

        Try
            Using frm As New DragDropForm(DragDropMode.FileOrDirectory)
                If frm.ShowDialog() = DialogResult.OK Then
                    selectedPath = frm.SelectedFilePath
                End If
            End Using
        Finally
            Globals.ThisAddIn.DragDropFormLabel = ""
            Globals.ThisAddIn.DragDropFormFilter = ""
        End Try

        If String.IsNullOrWhiteSpace(selectedPath) Then Exit Sub
        If Not File.Exists(selectedPath) Then
            ShowCustomMessageBox("The selected file does not exist.")
            Exit Sub
        End If

        Dim ext As String = Path.GetExtension(selectedPath).ToLowerInvariant()
        If ext = ".doc" AndAlso Not INI_AllowLegacyDocFiles Then
            ShowCustomMessageBox("The .doc format is disabled for security. Please convert to .docx first.")
            Exit Sub
        End If
        If ext <> ".doc" AndAlso ext <> ".docx" Then
            ShowCustomMessageBox($"File type '{ext}' is not supported. Please select a .docx file.")
            Exit Sub
        End If

        ' Restore last instructions from My.Settings
        Dim lastInstructions As String = ""
        Try
            lastInstructions = If(My.Settings.TableComplete_LastInstructions, "")
        Catch
        End Try

        ' Get user instructions for table completion
        Dim userInstructions As String = ShowCustomInputBox(
            "Describe how the tables should be completed:" & vbCrLf &
            "(e.g., 'Fill in the empty cells based on headings and row labels. " &
            "Add rows where the table appears incomplete.')",
            AN & " Complete Tables", False, lastInstructions, Context:=_context)

        If String.IsNullOrWhiteSpace(userInstructions) Then Exit Sub
        userInstructions = userInstructions.Trim()

        ' Save instructions to My.Settings
        Try
            My.Settings.TableComplete_LastInstructions = userInstructions
            My.Settings.Save()
        Catch
        End Try

        ' Determine output path
        Dim dir As String = Path.GetDirectoryName(selectedPath)
        Dim nameWithoutExt As String = Path.GetFileNameWithoutExtension(selectedPath)
        Dim outputPath As String = Path.Combine(dir, $"{nameWithoutExt}_completed.docx")

        If File.Exists(outputPath) Then
            Dim overwrite As Integer = ShowCustomYesNoBox(
                $"Output file already exists:{vbCrLf}{Path.GetFileName(outputPath)}{vbCrLf}{vbCrLf}Overwrite?",
                "Yes, overwrite", "No, cancel")
            If overwrite <> 1 Then Exit Sub
        End If

        ProgressBarModule.GlobalProgressValue = 0
        ProgressBarModule.GlobalProgressMax = 1
        ProgressBarModule.GlobalProgressLabel = "Analyzing tables..."
        ProgressBarModule.CancelOperation = False
        ProgressBarModule.ShowProgressBarInSeparateThread(AN & " Complete Tables", "Starting...")

        Try
            StartTableCompletionDebugSession(selectedPath, outputPath, userInstructions)

            Dim success As Boolean = Await ProcessTableCompletion(selectedPath, outputPath, userInstructions)

            If success Then
                ShowCustomMessageBox($"Tables completed successfully.{vbCrLf}Output: {Path.GetFileName(outputPath)}", AN & " Complete Tables")
            Else
                ShowCustomMessageBox("Table completion failed or was cancelled.", AN & " Complete Tables")
            End If
        Catch ex As Exception
            WriteTableCompletionDebug("Unhandled exception", ex.ToString())
            ShowCustomMessageBox($"Error: {ex.Message}", AN & " Complete Tables")
        Finally
            ProgressBarModule.CancelOperation = True
        End Try
    End Sub

    Public Async Function CompleteWordDocumentTables(Optional promptOverride As String = Nothing,
                                                 Optional useSecondAPI As Boolean = False) As System.Threading.Tasks.Task
        Dim selectedPath As String = ""

        If INI_AllowLegacyDocFiles Then
            Globals.ThisAddIn.DragDropFormLabel = "Select a Word document with tables to complete"
            Globals.ThisAddIn.DragDropFormFilter = "Word Documents|*.doc;*.docx|Word Document (*.docx)|*.docx|Word 97-2003 (*.doc)|*.doc"
        Else
            Globals.ThisAddIn.DragDropFormLabel = "Select a Word document with tables to complete"
            Globals.ThisAddIn.DragDropFormFilter = "Word Documents (*.docx)|*.docx"
        End If

        Try
            Using frm As New DragDropForm(DragDropMode.FileOrDirectory)
                If frm.ShowDialog() = DialogResult.OK Then
                    selectedPath = frm.SelectedFilePath
                End If
            End Using
        Finally
            Globals.ThisAddIn.DragDropFormLabel = ""
            Globals.ThisAddIn.DragDropFormFilter = ""
        End Try

        If String.IsNullOrWhiteSpace(selectedPath) Then Exit Function
        If Not File.Exists(selectedPath) Then
            ShowCustomMessageBox("The selected file does not exist.")
            Exit Function
        End If

        Dim ext As String = Path.GetExtension(selectedPath).ToLowerInvariant()
        If ext = ".doc" AndAlso Not INI_AllowLegacyDocFiles Then
            ShowCustomMessageBox("The .doc format is disabled for security. Please convert to .docx first.")
            Exit Function
        End If
        If ext <> ".doc" AndAlso ext <> ".docx" Then
            ShowCustomMessageBox($"File type '{ext}' is not supported. Please select a .docx file.")
            Exit Function
        End If

        Dim lastInstructions As String = ""
        Try
            lastInstructions = If(My.Settings.TableComplete_LastInstructions, "")
        Catch
        End Try

        Dim userInstructions As String
        If String.IsNullOrWhiteSpace(promptOverride) Then
            userInstructions = ShowCustomInputBox(
            "Describe how the tables should be completed:" & vbCrLf &
            "(e.g., 'Fill in the empty cells based on headings and row labels. " &
            "Add rows where the table appears incomplete.')",
            AN & " Complete Tables", False, lastInstructions, Context:=_context)

            If String.IsNullOrWhiteSpace(userInstructions) Then Exit Function
            userInstructions = userInstructions.Trim()
        Else
            userInstructions = InterpolateAtRuntime(promptOverride).Trim()
            If String.IsNullOrWhiteSpace(userInstructions) Then
                ShowCustomMessageBox("No form instructions were provided.", AN & " Complete Tables")
                Exit Function
            End If
        End If

        Try
            My.Settings.TableComplete_LastInstructions = userInstructions
            My.Settings.Save()
        Catch
        End Try

        Dim dir As String = Path.GetDirectoryName(selectedPath)
        Dim nameWithoutExt As String = Path.GetFileNameWithoutExtension(selectedPath)
        Dim outputPath As String = Path.Combine(dir, $"{nameWithoutExt}_completed.docx")

        If File.Exists(outputPath) Then
            Dim overwrite As Integer = ShowCustomYesNoBox(
            $"Output file already exists:{vbCrLf}{Path.GetFileName(outputPath)}{vbCrLf}{vbCrLf}Overwrite?",
            "Yes, overwrite", "No, cancel")
            If overwrite <> 1 Then Exit Function
        End If

        ProgressBarModule.GlobalProgressValue = 0
        ProgressBarModule.GlobalProgressMax = 1
        ProgressBarModule.GlobalProgressLabel = "Analyzing tables..."
        ProgressBarModule.CancelOperation = False
        ProgressBarModule.ShowProgressBarInSeparateThread(AN & " Complete Tables", "Starting...")

        Try
            StartTableCompletionDebugSession(selectedPath, outputPath, userInstructions)

            Dim success As Boolean = Await ProcessTableCompletion(selectedPath, outputPath, userInstructions, useSecondAPI)

            If success Then
                ShowCustomMessageBox($"Completed document saved as:{vbCrLf}{outputPath}", AN & " Complete Tables")
            End If
        Catch ex As Exception
            WriteTableCompletionDebug("Unhandled exception", ex.ToString())
            ShowCustomMessageBox($"Error: {ex.Message}", AN & " Complete Tables")
        Finally
            ProgressBarModule.CancelOperation = True
        End Try
    End Function

    Private Async Function ProcessTableCompletion(inputPath As String,
                                                  outputPath As String,
                                                  userInstructions As String,
                                                  Optional useSecondAPI As Boolean = False) As Task(Of Boolean)
        Dim tempDocxPath As String = Nothing
        Dim wordApp As Word.Application = Nothing
        Dim doc As Word.Document = Nothing

        Try
            If Path.GetExtension(inputPath).ToLowerInvariant() = ".doc" Then
                tempDocxPath = Path.Combine(Path.GetTempPath(), $"{AN2}_tblconv_{Guid.NewGuid():N}.docx")
                wordApp = Globals.ThisAddIn.Application
                wordApp.ScreenUpdating = False
                doc = wordApp.Documents.Open(inputPath, ReadOnly:=True, Visible:=False, AddToRecentFiles:=False)
                doc.SaveAs2(tempDocxPath, WdSaveFormat.wdFormatXMLDocument)
                doc.Close(WdSaveOptions.wdDoNotSaveChanges)
                doc = Nothing
                wordApp.ScreenUpdating = True
            Else
                tempDocxPath = inputPath
            End If

            File.Copy(tempDocxPath, outputPath, overwrite:=True)

            Dim success As Boolean = Await ProcessDocxTables(outputPath, userInstructions, useSecondAPI)
            If Not success Then Return False

            Dim compareSourcePath As String = tempDocxPath
            If String.IsNullOrWhiteSpace(compareSourcePath) OrElse Not File.Exists(compareSourcePath) Then
                compareSourcePath = inputPath
            End If

            Dim compareCreated As Boolean = TableCreateWordCompareDocument(compareSourcePath, outputPath)
            If Not compareCreated Then
                ShowCustomMessageBox("The completed document was created, but the compare document could not be generated.", AN & " Complete Tables")
            End If

            Return True

        Finally
            If doc IsNot Nothing Then
                Try : doc.Close(WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
            End If
            If tempDocxPath IsNot Nothing AndAlso tempDocxPath <> inputPath AndAlso File.Exists(tempDocxPath) Then
                Try : File.Delete(tempDocxPath) : Catch : End Try
            End If
        End Try
    End Function


    Private Async Function ProcessDocxTables(docxPath As String,
                                         userInstructions As String,
                                         Optional useSecondAPI As Boolean = False) As Task(Of Boolean)

        Dim tempDir As String = Path.Combine(Path.GetTempPath(), $"{AN2}_tbl_{Guid.NewGuid():N}")

        Try
            ZipFile.ExtractToDirectory(docxPath, tempDir)

            If TableCompleteDebugIncludePackageDump Then DumpDocxPackageDiagnostics(docxPath)

            Dim documentXmlPath As String = Path.Combine(tempDir, "word", "document.xml")
            If Not File.Exists(documentXmlPath) Then
                ShowCustomMessageBox("Invalid DOCX structure - document.xml not found.")
                Return False
            End If

            Dim xmlDoc As New System.Xml.XmlDocument()
            xmlDoc.PreserveWhitespace = True
            xmlDoc.Load(documentXmlPath)

            Dim nsMgr As New System.Xml.XmlNamespaceManager(xmlDoc.NameTable)
            nsMgr.AddNamespace("w", WNs)
            nsMgr.AddNamespace("w14", W14Ns)

            Dim documentContext As String = ExtractDocumentBodyContext(xmlDoc, nsMgr)
            Dim tables As List(Of TableInfo) = ExtractTablesFromXml(xmlDoc, nsMgr)
            Dim bodySections As List(Of BodySectionInfo) = ExtractBodySectionsFromXml(xmlDoc, nsMgr)

            WriteTableCompletionDebug("Extracted fields", BuildFormFieldDebugReport(tables))

            If tables.Count = 0 AndAlso bodySections.Count = 0 Then
                ShowCustomMessageBox("No tables or body sections found in the document.")
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
            confirmMsg &= ". Continue with AI completion?"

            Dim confirm As Integer = ShowCustomYesNoBox(confirmMsg, "Yes, continue", "No, cancel")
            If confirm <> 1 Then Return False

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

    Private Function TableCreateWordCompareDocument(originalPath As String, completedPath As String) As Boolean
        Dim wordApp As Word.Application = Nothing
        Dim originalDoc As Word.Document = Nothing
        Dim completedDoc As Word.Document = Nothing
        Dim compareDoc As Word.Document = Nothing
        Dim wasScreenUpdating As Boolean = True

        Try
            If String.IsNullOrWhiteSpace(originalPath) OrElse String.IsNullOrWhiteSpace(completedPath) Then Return False
            If Not File.Exists(originalPath) OrElse Not File.Exists(completedPath) Then Return False

            wordApp = Globals.ThisAddIn.Application
            wasScreenUpdating = wordApp.ScreenUpdating
            wordApp.ScreenUpdating = False

            Dim comparePath As String = GetTableCompletionCompareFilePath(completedPath)

            originalDoc = wordApp.Documents.Open(originalPath, ReadOnly:=True, Visible:=False, AddToRecentFiles:=False)
            completedDoc = wordApp.Documents.Open(completedPath, ReadOnly:=True, Visible:=False, AddToRecentFiles:=False)

            compareDoc = wordApp.CompareDocuments(
                OriginalDocument:=originalDoc,
                RevisedDocument:=completedDoc,
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

            compareDoc.SaveAs2(GetTableCompletionCompareFilePath(completedPath), WdSaveFormat.wdFormatXMLDocument)

            compareDoc.Close(WdSaveOptions.wdDoNotSaveChanges)
            compareDoc = Nothing

            completedDoc.Close(WdSaveOptions.wdDoNotSaveChanges)
            completedDoc = Nothing

            originalDoc.Close(WdSaveOptions.wdDoNotSaveChanges)
            originalDoc = Nothing

            wordApp.ScreenUpdating = wasScreenUpdating
            Return True

        Catch
            Return False
        Finally
            If compareDoc IsNot Nothing Then
                Try : compareDoc.Close(WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
            End If
            If completedDoc IsNot Nothing Then
                Try : completedDoc.Close(WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
            End If
            If originalDoc IsNot Nothing Then
                Try : originalDoc.Close(WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
            End If
            If wordApp IsNot Nothing Then
                Try : wordApp.ScreenUpdating = wasScreenUpdating : Catch : End Try
            End If
        End Try
    End Function


#End Region

#Region "Extraction"

    Private Function ExtractVisibleParagraphTextExcludingSdts(paragraphNode As System.Xml.XmlNode,
                                                             nsMgr As System.Xml.XmlNamespaceManager,
                                                             ByRef textNodes As List(Of System.Xml.XmlNode)) As String
        textNodes = New List(Of System.Xml.XmlNode)()
        If paragraphNode Is Nothing Then Return ""

        Dim sb As New StringBuilder()
        Dim allTextNodes As System.Xml.XmlNodeList = paragraphNode.SelectNodes(".//w:t", nsMgr)

        For Each tNode As System.Xml.XmlNode In allTextNodes
            Dim ancestor As System.Xml.XmlNode = tNode.ParentNode
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

    Private Function ExtractBestContextParagraphTextFromNode(node As System.Xml.XmlNode,
                                                             nsMgr As System.Xml.XmlNamespaceManager) As String
        If node Is Nothing Then Return ""

        Dim paragraphNodes As New List(Of System.Xml.XmlNode)()

        If node.LocalName = "p" Then
            paragraphNodes.Add(node)
        Else
            For Each pNode As System.Xml.XmlNode In node.SelectNodes(".//w:p[not(ancestor::w:tbl)]", nsMgr)
                paragraphNodes.Add(pNode)
            Next
        End If

        For Each pNode As System.Xml.XmlNode In paragraphNodes
            Dim textNodes As List(Of System.Xml.XmlNode) = Nothing
            Dim text As String = NormalizeTextForPlaceholderDetection(ExtractVisibleParagraphTextExcludingSdts(pNode, nsMgr, textNodes))
            If Not String.IsNullOrWhiteSpace(text) AndAlso Not IsPlaceholderText(text) Then
                Return text
            End If
        Next

        Return ""
    End Function

    Private Function ExtractCandidateBodyParagraphsFromNode(node As System.Xml.XmlNode,
                                                            nsMgr As System.Xml.XmlNamespaceManager) As List(Of BodyParagraphInfo)
        Dim results As New List(Of BodyParagraphInfo)()
        If node Is Nothing Then Return results

        Dim paragraphNodes As New List(Of System.Xml.XmlNode)()

        If node.LocalName = "p" Then
            paragraphNodes.Add(node)
        Else
            For Each pNode As System.Xml.XmlNode In node.SelectNodes(".//w:p[not(ancestor::w:tbl)]", nsMgr)
                paragraphNodes.Add(pNode)
            Next
        End If

        Dim paragraphIndex As Integer = 0

        For Each pNode As System.Xml.XmlNode In paragraphNodes
            Dim textNodes As List(Of System.Xml.XmlNode) = Nothing
            Dim rawText As String = ExtractVisibleParagraphTextExcludingSdts(pNode, nsMgr, textNodes)
            Dim normalizedText As String = NormalizeTextForPlaceholderDetection(rawText)

            results.Add(New BodyParagraphInfo() With {
                .Index = paragraphIndex,
                .Text = rawText,
                .IsPlaceholder = False,
                .IsEmpty = String.IsNullOrWhiteSpace(normalizedText),
                .TextNodes = If(textNodes, New List(Of System.Xml.XmlNode)()),
                .XmlNode = pNode
            })

            paragraphIndex += 1
        Next

        Return results
    End Function

    Private Function ExtractFormFieldsFromNode(containerNode As System.Xml.XmlNode,
                                               nsMgr As System.Xml.XmlNamespaceManager,
                                               isWritable As Boolean,
                                               protectionReason As String) As List(Of FormFieldInfo)
        Dim results As New List(Of FormFieldInfo)()
        If containerNode Is Nothing Then Return results

        Dim fieldIndex As Integer = 0
        Dim sdtNodes As New List(Of System.Xml.XmlNode)()

        Dim owningSdtNode As System.Xml.XmlNode = GetOwningSdtNodeForCell(containerNode)
        If owningSdtNode IsNot Nothing Then
            sdtNodes.Add(owningSdtNode)
        End If

        Dim descendantSdtNodes As System.Xml.XmlNodeList = containerNode.SelectNodes(".//w:sdt", nsMgr)
        For Each sdtNode As System.Xml.XmlNode In descendantSdtNodes
            If Not sdtNodes.Any(Function(existingNode) existingNode Is sdtNode) Then
                sdtNodes.Add(sdtNode)
            End If
        Next

        For Each sdtNode As System.Xml.XmlNode In sdtNodes
            Dim sdtPr As System.Xml.XmlNode = sdtNode.SelectSingleNode("w:sdtPr", nsMgr)
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

            Dim tagNode As System.Xml.XmlNode = sdtPr.SelectSingleNode("w:tag/@w:val", nsMgr)
            Dim aliasNode As System.Xml.XmlNode = sdtPr.SelectSingleNode("w:alias/@w:val", nsMgr)
            ff.Name = If(tagNode?.Value, If(aliasNode?.Value, ""))

            Dim showingPlcHdr As Boolean = (sdtPr.SelectSingleNode("w:showingPlcHdr", nsMgr) IsNot Nothing)

            Dim dropDownNode As System.Xml.XmlNode = sdtPr.SelectSingleNode("w:dropDownList", nsMgr)
            Dim comboBoxNode As System.Xml.XmlNode = sdtPr.SelectSingleNode("w:comboBox", nsMgr)
            Dim checkboxNode As System.Xml.XmlNode = sdtPr.SelectSingleNode("w14:checkbox", nsMgr)
            Dim textNode As System.Xml.XmlNode = sdtPr.SelectSingleNode("w:text", nsMgr)
            Dim dateNode As System.Xml.XmlNode = sdtPr.SelectSingleNode("w:date", nsMgr)

            If dropDownNode IsNot Nothing OrElse comboBoxNode IsNot Nothing Then
                ff.FieldType = FormFieldType.Dropdown

                Dim listParent As System.Xml.XmlNode = If(dropDownNode, comboBoxNode)
                Dim listItems As System.Xml.XmlNodeList = listParent.SelectNodes("w:listItem", nsMgr)
                For Each li As System.Xml.XmlNode In listItems
                    Dim displayAttr As System.Xml.XmlNode = li.Attributes("w:displayText")
                    Dim valueAttr As System.Xml.XmlNode = li.Attributes("w:value")
                    AddFormFieldOption(ff.Options, If(displayAttr?.Value, ""), If(valueAttr?.Value, ""))
                Next

                Dim rawDisplayValue As String = GetSdtContentText(sdtNode, nsMgr)
                Dim lastValueAttr As System.Xml.XmlNode = listParent.Attributes("w:lastValue")
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

                Dim checkedNode As System.Xml.XmlNode = checkboxNode.SelectSingleNode("w14:checked/@w14:val", nsMgr)
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

        Dim fldCharNodes As System.Xml.XmlNodeList = containerNode.SelectNodes(".//w:r[w:fldChar[@w:fldCharType='begin']]", nsMgr)
        For Each fldRunNode As System.Xml.XmlNode In fldCharNodes
            Dim fldChar As System.Xml.XmlNode = fldRunNode.SelectSingleNode("w:fldChar", nsMgr)
            If fldChar Is Nothing Then Continue For

            Dim ffData As System.Xml.XmlNode = fldChar.SelectSingleNode("w:ffData", nsMgr)
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

            Dim nameNode As System.Xml.XmlNode = ffData.SelectSingleNode("w:name/@w:val", nsMgr)
            ff.Name = If(nameNode?.Value, "")

            Dim ddList As System.Xml.XmlNode = ffData.SelectSingleNode("w:ddList", nsMgr)
            If ddList IsNot Nothing Then
                ff.FieldType = FormFieldType.LegacyDropdown

                Dim listEntries As System.Xml.XmlNodeList = ddList.SelectNodes("w:listEntry/@w:val", nsMgr)
                For Each entry As System.Xml.XmlNode In listEntries
                    AddFormFieldOption(ff.Options, entry.Value, entry.Value)
                Next

                Dim resultNode As System.Xml.XmlNode = ddList.SelectSingleNode("w:result/@w:val", nsMgr)
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

            Dim checkBox As System.Xml.XmlNode = ffData.SelectSingleNode("w:checkBox", nsMgr)
            If checkBox IsNot Nothing Then
                ff.FieldType = FormFieldType.LegacyCheckbox
                AddFormFieldOption(ff.Options, "true", "true")
                AddFormFieldOption(ff.Options, "false", "false")

                Dim checkedNode As System.Xml.XmlNode = checkBox.SelectSingleNode("w:checked/@w:val", nsMgr)
                If checkedNode IsNot Nothing Then
                    ff.CurrentValue = If(checkedNode.Value = "1" OrElse checkedNode.Value.ToLowerInvariant() = "true", "true", "false")
                Else
                    Dim defaultNode As System.Xml.XmlNode = checkBox.SelectSingleNode("w:default/@w:val", nsMgr)
                    ff.CurrentValue = If(defaultNode IsNot Nothing AndAlso (defaultNode.Value = "1" OrElse defaultNode.Value.ToLowerInvariant() = "true"), "true", "false")
                End If

                ff.CurrentOptionValue = ff.CurrentValue
                results.Add(ff)
                fieldIndex += 1
            End If
        Next

        Return results
    End Function

    Private Function TryBuildBodySection(node As System.Xml.XmlNode,
                                         sectionIndex As Integer,
                                         headingText As String,
                                         nsMgr As System.Xml.XmlNamespaceManager) As BodySectionInfo
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


    Private Function ExtractBodySectionsFromXml(xmlDoc As System.Xml.XmlDocument,
                                                nsMgr As System.Xml.XmlNamespaceManager) As List(Of BodySectionInfo)
        Dim results As New List(Of BodySectionInfo)()
        Dim bodyNode As System.Xml.XmlNode = xmlDoc.SelectSingleNode("//w:body", nsMgr)
        If bodyNode Is Nothing Then Return results

        Dim sectionIndex As Integer = 0
        Dim lastContextText As String = ""

        For Each child As System.Xml.XmlNode In bodyNode.ChildNodes
            If child.NodeType <> System.Xml.XmlNodeType.Element Then Continue For

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

    Private Function IsTextNodeInsideSdt(textNode As System.Xml.XmlNode, tcNode As System.Xml.XmlNode) As Boolean
        Dim current As System.Xml.XmlNode = textNode.ParentNode

        While current IsNot Nothing AndAlso current IsNot tcNode
            If current.LocalName = "sdt" Then Return True
            current = current.ParentNode
        End While

        Return False
    End Function

    Private Function TryGetCellOverrideWidthPct(tcNode As System.Xml.XmlNode,
                                                totalGridWidth As Double,
                                                nsMgr As System.Xml.XmlNamespaceManager,
                                                ByRef widthPct As Double) As Boolean
        widthPct = 0

        If tcNode Is Nothing OrElse nsMgr Is Nothing Then Return False

        Dim tcWidthValueNode As System.Xml.XmlNode = tcNode.SelectSingleNode("w:tcPr/w:tcW/@w:w", nsMgr)
        If tcWidthValueNode Is Nothing Then Return False

        Dim tcWidthTypeNode As System.Xml.XmlNode = tcNode.SelectSingleNode("w:tcPr/w:tcW/@w:type", nsMgr)
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

    Private Function GetDocumentOrderTableIndexMap(xmlDoc As System.Xml.XmlDocument,
                                                   nsMgr As System.Xml.XmlNamespaceManager) As Dictionary(Of System.Xml.XmlNode, Integer)
        Dim result As New Dictionary(Of System.Xml.XmlNode, Integer)()
        Dim tblNodes As System.Xml.XmlNodeList = xmlDoc.SelectNodes("//w:tbl", nsMgr)
        Dim index As Integer = 0

        For Each tblNode As System.Xml.XmlNode In tblNodes
            result(tblNode) = index
            index += 1
        Next

        Return result
    End Function

    Private Sub AppendBodyContextNode(node As System.Xml.XmlNode,
                                      nsMgr As System.Xml.XmlNamespaceManager,
                                      sb As StringBuilder,
                                      tableIndexMap As Dictionary(Of System.Xml.XmlNode, Integer))
        If node Is Nothing Then Exit Sub
        If node.NodeType <> System.Xml.XmlNodeType.Element Then Exit Sub

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
                For Each child As System.Xml.XmlNode In node.ChildNodes
                    If child.NodeType <> System.Xml.XmlNodeType.Element Then Continue For
                    AppendBodyContextNode(child, nsMgr, sb, tableIndexMap)
                Next

            Case "AlternateContent"
                Dim branch As System.Xml.XmlNode = Nothing

                For Each child As System.Xml.XmlNode In node.ChildNodes
                    If child.NodeType <> System.Xml.XmlNodeType.Element Then Continue For
                    If child.LocalName = "Choice" Then
                        branch = child
                        Exit For
                    End If
                Next

                If branch Is Nothing Then
                    For Each child As System.Xml.XmlNode In node.ChildNodes
                        If child.NodeType <> System.Xml.XmlNodeType.Element Then Continue For
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

    Private Function TruncateContextAtLineBoundaries(text As String, maxChars As Integer) As String
        If String.IsNullOrEmpty(text) OrElse text.Length <= maxChars Then Return text

        Dim normalized As String = text.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
        Dim lines As List(Of String) = normalized.Split({vbLf}, StringSplitOptions.None).ToList()

        Dim separator As String = vbCrLf
        Dim marker As String = "[... document content truncated ...]"
        Dim halfBudget As Integer = Math.Max(1, (maxChars - marker.Length - (separator.Length * 2)) \ 2)

        Dim head As New List(Of String)()
        Dim tail As New List(Of String)()

        Dim currentHeadLength As Integer = 0
        For Each line In lines
            Dim addLen As Integer = If(head.Count = 0, 0, separator.Length) + line.Length
            If currentHeadLength + addLen > halfBudget Then Exit For
            head.Add(line)
            currentHeadLength += addLen
        Next

        Dim currentTailLength As Integer = 0
        For i As Integer = lines.Count - 1 To 0 Step -1
            Dim line As String = lines(i)
            Dim addLen As Integer = If(tail.Count = 0, 0, separator.Length) + line.Length
            If currentTailLength + addLen > halfBudget Then Exit For
            tail.Insert(0, line)
            currentTailLength += addLen
        Next

        Return String.Join(separator, head) &
               separator & marker & separator &
               String.Join(separator, tail)
    End Function

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

    Private Function CloneTemplateRowBeforeApply(tbl As TableInfo, insertAfterRow As Integer) As System.Xml.XmlNode
        Dim templateRow As TableRowInfo = FindTemplateRowForInsertion(tbl, insertAfterRow)
        If templateRow Is Nothing OrElse templateRow.XmlNode Is Nothing Then Return Nothing
        Return templateRow.XmlNode.CloneNode(deep:=True)
    End Function

    Private Function RowContainsVerticalMerge(rowNode As System.Xml.XmlNode, nsMgr As System.Xml.XmlNamespaceManager) As Boolean
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

    Private Sub ApplySdtDefaultRunProperties(sdtNode As System.Xml.XmlNode, nsMgr As System.Xml.XmlNamespaceManager)
        If sdtNode Is Nothing Then Exit Sub

        Dim sdtPrRunProps As System.Xml.XmlNode = sdtNode.SelectSingleNode("w:sdtPr/w:rPr", nsMgr)
        Dim firstRun As System.Xml.XmlNode = sdtNode.SelectSingleNode("w:sdtContent//w:r", nsMgr)

        If firstRun Is Nothing Then Exit Sub

        Dim existingRpr As System.Xml.XmlNode = firstRun.SelectSingleNode("w:rPr", nsMgr)
        If existingRpr IsNot Nothing Then
            firstRun.RemoveChild(existingRpr)
        End If

        If sdtPrRunProps Is Nothing Then Exit Sub

        Dim clonedRpr As System.Xml.XmlNode = sdtPrRunProps.CloneNode(deep:=True)

        Dim placeholderStyles As System.Xml.XmlNodeList = clonedRpr.SelectNodes("w:rStyle[@w:val='PlaceholderText' or @w:val='Platzhaltertext']", nsMgr)
        For Each styleNode As System.Xml.XmlNode In placeholderStyles
            clonedRpr.RemoveChild(styleNode)
        Next

        firstRun.PrependChild(clonedRpr)
    End Sub


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

    Private Function ExtractDocumentBodyContext(xmlDoc As System.Xml.XmlDocument,
                                                nsMgr As System.Xml.XmlNamespaceManager) As String
        Dim bodyNode As System.Xml.XmlNode = xmlDoc.SelectSingleNode("//w:body", nsMgr)
        If bodyNode Is Nothing Then Return ""

        Dim sb As New StringBuilder()
        Dim tableIndexMap As Dictionary(Of System.Xml.XmlNode, Integer) = GetDocumentOrderTableIndexMap(xmlDoc, nsMgr)

        For Each child As System.Xml.XmlNode In bodyNode.ChildNodes
            If child.NodeType <> System.Xml.XmlNodeType.Element Then Continue For
            AppendBodyContextNode(child, nsMgr, sb, tableIndexMap)

            If sb.Length > TableCompleteMaxContextChars * 2 Then Exit For
        Next

        Return TruncateContextAtLineBoundaries(sb.ToString(), TableCompleteMaxContextChars)
    End Function

    Private Function ExtractTablesFromXml(xmlDoc As System.Xml.XmlDocument,
                                          nsMgr As System.Xml.XmlNamespaceManager) As List(Of TableInfo)
        Dim tables As New List(Of TableInfo)()
        Dim tblNodes As System.Xml.XmlNodeList = xmlDoc.SelectNodes("//w:tbl", nsMgr)
        Dim tableIndex As Integer = 0

        For Each tblNode As System.Xml.XmlNode In tblNodes
            Dim tblInfo As New TableInfo() With {
                .TableIndex = tableIndex,
                .Rows = New List(Of TableRowInfo)(),
                .XmlNode = tblNode,
                .ColumnCount = 0,
                .ContextBefore = "",
                .ColumnWidths = New List(Of Double)()
            }

            Dim gridCols As System.Xml.XmlNodeList = tblNode.SelectNodes("w:tblGrid/w:gridCol", nsMgr)
            Dim rawWidths As New List(Of Double)()
            Dim totalGridWidth As Double = 0

            For Each gridCol As System.Xml.XmlNode In gridCols
                Dim wAttr As System.Xml.XmlNode = gridCol.Attributes("w:w")
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

            Dim prevSibling As System.Xml.XmlNode = tblNode.PreviousSibling
            Dim contextParts As New List(Of String)()
            Dim contextCount As Integer = 0

            While prevSibling IsNot Nothing AndAlso contextCount < 3
                If prevSibling.NodeType = System.Xml.XmlNodeType.Element AndAlso prevSibling.LocalName = "p" Then
                    Dim paraText As String = GetParagraphPlainText(prevSibling, nsMgr)
                    If Not String.IsNullOrWhiteSpace(paraText) Then
                        contextParts.Insert(0, paraText)
                        contextCount += 1
                    End If
                End If
                prevSibling = prevSibling.PreviousSibling
            End While

            tblInfo.ContextBefore = String.Join(" | ", contextParts)

            Dim trNodes As System.Xml.XmlNodeList = tblNode.SelectNodes("w:tr", nsMgr)
            Dim rowIndex As Integer = 0

            For Each trNode As System.Xml.XmlNode In trNodes
                Dim rowInfo As New TableRowInfo() With {
                    .RowIndex = rowIndex,
                    .Cells = New List(Of TableCellInfo)(),
                    .XmlNode = trNode
                }

                Dim isHeaderRow As Boolean = (trNode.SelectSingleNode("w:trPr/w:tblHeader", nsMgr) IsNot Nothing)
                Dim effectiveTcNodes As List(Of System.Xml.XmlNode) = GetEffectiveRowCellNodes(trNode, nsMgr)
                Dim colIndex As Integer = 0

                For Each tcNode As System.Xml.XmlNode In effectiveTcNodes
                    Dim cellInfo As New TableCellInfo() With {
                        .Row = rowIndex,
                        .Col = colIndex,
                        .TextNodes = New List(Of System.Xml.XmlNode)(),
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

                    Dim gridSpanNode As System.Xml.XmlNode = tcNode.SelectSingleNode("w:tcPr/w:gridSpan/@w:val", nsMgr)
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

                    Dim vMergeNode As System.Xml.XmlNode = tcNode.SelectSingleNode("w:tcPr/w:vMerge", nsMgr)
                    If vMergeNode IsNot Nothing Then
                        Dim vMergeVal As System.Xml.XmlNode = vMergeNode.Attributes("w:val")
                        cellInfo.VMerge = If(vMergeVal IsNot Nothing AndAlso vMergeVal.Value = "restart", "restart", "continue")
                    End If

                    Dim cellProtectionReason As String = ""
                    cellInfo.IsWritable = IsCellWritable(tcNode, nsMgr, cellProtectionReason)
                    cellInfo.ProtectionReason = cellProtectionReason

                    ExtractFormFieldsFromCell(tcNode, cellInfo, nsMgr)

                    Dim textBuilder As New StringBuilder()
                    Dim allTextNodes As System.Xml.XmlNodeList = tcNode.SelectNodes(".//w:t", nsMgr)

                    For Each tNode As System.Xml.XmlNode In allTextNodes
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

    Private Function GetEffectiveRowCellNodes(trNode As System.Xml.XmlNode, nsMgr As System.Xml.XmlNamespaceManager) As List(Of System.Xml.XmlNode)
        Dim results As New List(Of System.Xml.XmlNode)()

        If trNode Is Nothing Then Return results

        Dim expectedGridColumns As Integer = GetExpectedRowGridColumnCount(trNode, nsMgr)

        For Each child As System.Xml.XmlNode In trNode.ChildNodes
            If child.NodeType <> System.Xml.XmlNodeType.Element Then Continue For

            CollectEffectiveRowCellNodesCore(child, nsMgr, results, expectedGridColumns)

            If expectedGridColumns > 0 AndAlso GetAccumulatedGridSpan(results, nsMgr) >= expectedGridColumns Then
                Exit For
            End If
        Next

        Return results
    End Function

    Private Sub CollectEffectiveRowCellNodes(node As System.Xml.XmlNode, nsMgr As System.Xml.XmlNamespaceManager, results As List(Of System.Xml.XmlNode))
        Dim expectedGridColumns As Integer = 0

        If node IsNot Nothing Then
            Dim rowNode As System.Xml.XmlNode = GetAncestorRowNode(node)
            If rowNode IsNot Nothing Then
                expectedGridColumns = GetExpectedRowGridColumnCount(rowNode, nsMgr)
            End If
        End If

        CollectEffectiveRowCellNodesCore(node, nsMgr, results, expectedGridColumns)
    End Sub

    Private Sub CollectEffectiveRowCellNodes(node As System.Xml.XmlNode, results As List(Of System.Xml.XmlNode))
        If node Is Nothing Then Exit Sub
        If results Is Nothing Then Exit Sub
        If node.NodeType <> System.Xml.XmlNodeType.Element Then Exit Sub

        Select Case node.LocalName
            Case "tc"
                results.Add(node)
                Exit Sub

            Case "sdt", "customXml", "smartTag", "ins", "moveTo", "hyperlink", "Choice", "Fallback"
                For Each child As System.Xml.XmlNode In node.ChildNodes
                    If child.NodeType <> System.Xml.XmlNodeType.Element Then Continue For
                    CollectEffectiveRowCellNodes(child, results)
                Next

            Case "AlternateContent"
                For Each child As System.Xml.XmlNode In node.ChildNodes
                    If child.NodeType <> System.Xml.XmlNodeType.Element Then Continue For
                    If child.LocalName <> "Choice" AndAlso child.LocalName <> "Fallback" Then Continue For
                    CollectEffectiveRowCellNodes(child, results)
                    If results.Count > 0 Then Exit For
                Next

            Case "del", "moveFrom"
                Exit Sub
        End Select
    End Sub

    Private Sub CollectEffectiveRowCellNodesCore(node As System.Xml.XmlNode, nsMgr As System.Xml.XmlNamespaceManager, results As List(Of System.Xml.XmlNode), expectedGridColumns As Integer)
        If node Is Nothing Then Exit Sub
        If nsMgr Is Nothing Then Exit Sub
        If results Is Nothing Then Exit Sub
        If node.NodeType <> System.Xml.XmlNodeType.Element Then Exit Sub

        If expectedGridColumns > 0 AndAlso GetAccumulatedGridSpan(results, nsMgr) >= expectedGridColumns Then
            Exit Sub
        End If

        Select Case node.LocalName
            Case "tc"
                results.Add(node)
                Exit Sub

            Case "sdt"
                Dim sdtContent As System.Xml.XmlNode = Nothing

                For Each child As System.Xml.XmlNode In node.ChildNodes
                    If child.NodeType <> System.Xml.XmlNodeType.Element Then Continue For
                    If child.LocalName = "sdtContent" Then
                        sdtContent = child
                        Exit For
                    End If
                Next

                If sdtContent Is Nothing Then Exit Sub

                For Each child As System.Xml.XmlNode In sdtContent.ChildNodes
                    If child.NodeType <> System.Xml.XmlNodeType.Element Then Continue For

                    CollectEffectiveRowCellNodesCore(child, nsMgr, results, expectedGridColumns)

                    If expectedGridColumns > 0 AndAlso GetAccumulatedGridSpan(results, nsMgr) >= expectedGridColumns Then
                        Exit For
                    End If
                Next

            Case "customXml", "smartTag", "ins", "moveTo", "hyperlink", "Choice", "Fallback"
                For Each child As System.Xml.XmlNode In node.ChildNodes
                    If child.NodeType <> System.Xml.XmlNodeType.Element Then Continue For

                    CollectEffectiveRowCellNodesCore(child, nsMgr, results, expectedGridColumns)

                    If expectedGridColumns > 0 AndAlso GetAccumulatedGridSpan(results, nsMgr) >= expectedGridColumns Then
                        Exit For
                    End If
                Next

            Case "AlternateContent"
                Dim branch As System.Xml.XmlNode = Nothing

                For Each child As System.Xml.XmlNode In node.ChildNodes
                    If child.NodeType <> System.Xml.XmlNodeType.Element Then Continue For
                    If child.LocalName = "Choice" Then
                        branch = child
                        Exit For
                    End If
                Next

                If branch Is Nothing Then
                    For Each child As System.Xml.XmlNode In node.ChildNodes
                        If child.NodeType <> System.Xml.XmlNodeType.Element Then Continue For
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
                ' Non-visible tracked-change content. Ignore.

            Case Else
                ' Ignore unrelated row-level markup.
        End Select
    End Sub

    Private Function GetAncestorRowNode(node As System.Xml.XmlNode) As System.Xml.XmlNode
        Dim current As System.Xml.XmlNode = node

        While current IsNot Nothing
            If current.LocalName = "tr" Then Return current
            If current.LocalName = "tbl" OrElse current.LocalName = "body" Then Exit While
            current = current.ParentNode
        End While

        Return Nothing
    End Function

    Private Function GetExpectedRowGridColumnCount(trNode As System.Xml.XmlNode, nsMgr As System.Xml.XmlNamespaceManager) As Integer
        If trNode Is Nothing Then Return 0
        If nsMgr Is Nothing Then Return 0

        Dim current As System.Xml.XmlNode = trNode

        While current IsNot Nothing AndAlso current.LocalName <> "tbl"
            current = current.ParentNode
        End While

        If current Is Nothing Then Return 0

        Dim gridCols As System.Xml.XmlNodeList = current.SelectNodes("w:tblGrid/w:gridCol", nsMgr)
        If gridCols Is Nothing Then Return 0

        Return gridCols.Count
    End Function

    Private Function GetAccumulatedGridSpan(tcNodes As List(Of System.Xml.XmlNode), nsMgr As System.Xml.XmlNamespaceManager) As Integer
        If tcNodes Is Nothing OrElse tcNodes.Count = 0 Then Return 0

        Dim total As Integer = 0

        For Each tcNode As System.Xml.XmlNode In tcNodes
            total += GetGridSpanValue(tcNode, nsMgr)
        Next

        Return total
    End Function

    Private Function GetGridSpanValue(tcNode As System.Xml.XmlNode, nsMgr As System.Xml.XmlNamespaceManager) As Integer
        If tcNode Is Nothing Then Return 1
        If nsMgr Is Nothing Then Return 1

        Dim spanNode As System.Xml.XmlNode = tcNode.SelectSingleNode("w:tcPr/w:gridSpan/@w:val", nsMgr)
        If spanNode Is Nothing Then Return 1

        Dim span As Integer
        If Integer.TryParse(spanNode.Value, span) AndAlso span > 0 Then
            Return span
        End If

        Return 1
    End Function



    Private Function GetOwningSdtNodeForCell(tcNode As System.Xml.XmlNode) As System.Xml.XmlNode
        Dim ancestor As System.Xml.XmlNode = tcNode.ParentNode

        While ancestor IsNot Nothing
            If ancestor.LocalName = "sdt" Then Return ancestor
            If ancestor.LocalName = "body" Then Exit While
            ancestor = ancestor.ParentNode
        End While

        Return Nothing
    End Function

    Private Function TryGetLockedSdtReason(node As System.Xml.XmlNode,
                                           nsMgr As System.Xml.XmlNamespaceManager,
                                           ByRef reason As String) As Boolean
        Dim current As System.Xml.XmlNode = node

        While current IsNot Nothing
            If current.NodeType = System.Xml.XmlNodeType.Element AndAlso current.LocalName = "sdt" Then
                Dim lockAttr As System.Xml.XmlNode = current.SelectSingleNode("w:sdtPr/w:lock/@w:val", nsMgr)
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

    Private Function IsCellWritable(tcNode As System.Xml.XmlNode, nsMgr As System.Xml.XmlNamespaceManager, ByRef reason As String) As Boolean
        Return Not TryGetLockedSdtReason(tcNode, nsMgr, reason)
    End Function

    Private Function IsFormFieldWritable(fieldNode As System.Xml.XmlNode, nsMgr As System.Xml.XmlNamespaceManager, ByRef reason As String) As Boolean
        Return Not TryGetLockedSdtReason(fieldNode, nsMgr, reason)
    End Function

    Private Sub ExtractFormFieldsFromCell(tcNode As System.Xml.XmlNode, cellInfo As TableCellInfo, nsMgr As System.Xml.XmlNamespaceManager)
        If cellInfo Is Nothing Then Exit Sub

        Dim protectionReason As String = If(cellInfo.ProtectionReason, "")
        Dim writable As Boolean = cellInfo.IsWritable

        cellInfo.FormFields = ExtractFormFieldsFromNode(tcNode, nsMgr, writable, protectionReason)
    End Sub

    ''' <summary>
    ''' Gets the plain text content from inside a w:sdt/w:sdtContent node.
    ''' </summary>
    Private Function GetSdtContentText(sdtNode As System.Xml.XmlNode, nsMgr As System.Xml.XmlNamespaceManager) As String
        Dim sdtContent As System.Xml.XmlNode = sdtNode.SelectSingleNode("w:sdtContent", nsMgr)
        If sdtContent Is Nothing Then Return ""
        Dim sb As New StringBuilder()
        For Each ct As System.Xml.XmlNode In sdtContent.SelectNodes(".//w:t", nsMgr)
            sb.Append(ct.InnerText)
        Next
        Return sb.ToString()
    End Function

    Private Function GetParagraphPlainText(paraNode As System.Xml.XmlNode, nsMgr As System.Xml.XmlNamespaceManager) As String
        Dim sb As New StringBuilder()
        Dim textNodes As System.Xml.XmlNodeList = paraNode.SelectNodes(".//w:t", nsMgr)
        For Each tNode As System.Xml.XmlNode In textNodes
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
            _context, InterpolateAtRuntime(BuildConsistencySummarySystemPrompt()), summaryPromptBuilder.ToString(),
            "", "", 0, useSecondAPI, True)

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
        ProgressBarModule.GlobalProgressMax = Math.Max(1, totalWorkBlocks + 1)
        ProgressBarModule.GlobalProgressValue = 0
        ProgressBarModule.GlobalProgressLabel = "Building form consistency summary"

        If ProgressBarModule.CancelOperation Then Return Nothing

        Dim consistencySummaryJson As String =
            Await BuildConsistencySummaryJson(tables, bodySections, userInstructions, documentContext, useSecondAPI)

        ProgressBarModule.GlobalProgressValue = 1

        Dim result As New CompletionBatchesResult() With {
            .TableUpdates = New List(Of TableUpdateResponse)(),
            .BodySectionUpdates = New List(Of BodySectionUpdateResponse)()
        }

        Dim processedBlocks As Integer = 0

        For i As Integer = 0 To tableBlocks.Count - 1
            If ProgressBarModule.CancelOperation Then Return Nothing

            ProgressBarModule.GlobalProgressLabel = $"Processing block {processedBlocks + 1} of {totalWorkBlocks} (tables)"
            Dim blockUpdates As List(Of TableUpdateResponse) =
                Await ProcessSingleTableBlock(tableBlocks(i), userInstructions, documentContext, consistencySummaryJson, useSecondAPI)

            If blockUpdates Is Nothing Then Return Nothing
            result.TableUpdates.AddRange(blockUpdates)

            processedBlocks += 1
            ProgressBarModule.GlobalProgressValue = 1 + processedBlocks
        Next

        For i As Integer = 0 To bodyBlocks.Count - 1
            If ProgressBarModule.CancelOperation Then Return Nothing

            ProgressBarModule.GlobalProgressLabel = $"Processing block {processedBlocks + 1} of {totalWorkBlocks} (body)"
            Dim blockUpdates As List(Of BodySectionUpdateResponse) =
                Await ProcessSingleBodyBlock(bodyBlocks(i), userInstructions, documentContext, consistencySummaryJson, useSecondAPI)

            If blockUpdates Is Nothing Then Return Nothing
            result.BodySectionUpdates.AddRange(blockUpdates)

            processedBlocks += 1
            ProgressBarModule.GlobalProgressValue = 1 + processedBlocks
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
            _context, InterpolateAtRuntime(BuildTableCompletionSystemPrompt()), promptBuilder.ToString(),
            "", "", 0, useSecondAPI, True)

        If String.IsNullOrWhiteSpace(response) Then
            ShowCustomMessageBox("LLM returned empty response for a table block.")
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
            _context, InterpolateAtRuntime(BuildTableCoverageSystemPrompt()), coveragePromptBuilder.ToString(),
            "", "", 0, useSecondAPI, True)

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
            _context, InterpolateAtRuntime(BuildBodyCompletionSystemPrompt()), promptBuilder.ToString(),
            "", "", 0, useSecondAPI, True)

        If String.IsNullOrWhiteSpace(response) Then
            ShowCustomMessageBox("LLM returned empty response for a body block.")
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
            ShowCustomMessageBox($"Failed to parse body completion response as JSON: {ex.Message}")
            Return Nothing
        End Try

        Return results
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
        If String.IsNullOrWhiteSpace(response) Then Return "{}"

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

            If compactJson.Length > TableCompleteMaxConsistencySummaryChars Then
                Dim trimmed As New JObject()

                If parsed("documentType") IsNot Nothing Then trimmed("documentType") = parsed("documentType")
                If parsed("theme") IsNot Nothing Then trimmed("theme") = parsed("theme")
                If parsed("terminology") IsNot Nothing Then trimmed("terminology") = parsed("terminology")
                If parsed("entities") IsNot Nothing Then trimmed("entities") = parsed("entities")
                If parsed("styleRules") IsNot Nothing Then trimmed("styleRules") = parsed("styleRules")
                If parsed("crossTableRules") IsNot Nothing Then trimmed("crossTableRules") = parsed("crossTableRules")
                If parsed("dropdownConventions") IsNot Nothing Then trimmed("dropdownConventions") = parsed("dropdownConventions")
                If parsed("dateConventions") IsNot Nothing Then trimmed("dateConventions") = parsed("dateConventions")
                If parsed("numericConventions") IsNot Nothing Then trimmed("numericConventions") = parsed("numericConventions")
                If parsed("warnings") IsNot Nothing Then trimmed("warnings") = parsed("warnings")

                compactJson = trimmed.ToString(Newtonsoft.Json.Formatting.None)

                If compactJson.Length > TableCompleteMaxConsistencySummaryChars Then
                    compactJson = compactJson.Substring(0, TableCompleteMaxConsistencySummaryChars)
                End If
            End If

            Return compactJson
        Catch ex As Exception
            WriteTableCompletionDebug("NormalizeConsistencySummaryJson error", ex.ToString() & vbCrLf & vbCrLf & response)
            Return "{}"
        End Try
    End Function

#End Region

#Region "Response Parsing"

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
            ShowCustomMessageBox($"Failed to parse LLM response as JSON: {ex.Message}")
            Return Nothing
        End Try

        Return results
    End Function

#End Region

#Region "Apply Updates"

    Private Sub WriteTextIntoParagraph(paragraph As BodyParagraphInfo,
                                       newText As String,
                                       nsMgr As System.Xml.XmlNamespaceManager)
        If paragraph Is Nothing OrElse paragraph.XmlNode Is Nothing Then Exit Sub

        Dim normalized As String = If(newText, "")
        Dim currentText As String = If(paragraph.Text, "")
        If String.Equals(currentText, normalized, StringComparison.Ordinal) Then Exit Sub

        Dim pNode As System.Xml.XmlNode = paragraph.XmlNode

        ClearAllAncestorSdtPlaceholderStates(pNode, nsMgr)
        SanitizeCellRunFormatting(pNode, nsMgr)
        SanitizeParagraphProperties(pNode, nsMgr)

        If paragraph.TextNodes IsNot Nothing AndAlso paragraph.TextNodes.Count > 0 Then
            SetTextNodeWithSpacePreserve(paragraph.TextNodes(0), normalized)
            For i As Integer = 1 To paragraph.TextNodes.Count - 1
                SetTextNodeWithSpacePreserve(paragraph.TextNodes(i), "")
            Next
        Else
            Dim toRemove As New List(Of System.Xml.XmlNode)()
            For Each ch As System.Xml.XmlNode In pNode.ChildNodes
                If ch.NodeType = System.Xml.XmlNodeType.Element AndAlso ch.LocalName = "pPr" Then Continue For
                toRemove.Add(ch)
            Next
            For Each ch As System.Xml.XmlNode In toRemove
                pNode.RemoveChild(ch)
            Next

            Dim newRun As System.Xml.XmlNode = CreateCleanTextRun(pNode, normalized, nsMgr)
            pNode.AppendChild(newRun)

            Dim newTextNode As System.Xml.XmlNode = newRun.SelectSingleNode("w:t", nsMgr)
            paragraph.TextNodes = New List(Of System.Xml.XmlNode)()
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
                                        nsMgr As System.Xml.XmlNamespaceManager)
        If bodySections Is Nothing OrElse updates Is Nothing Then Exit Sub

        For Each update In updates
            If ProgressBarModule.CancelOperation Then Exit Sub

            Dim section As BodySectionInfo = bodySections.FirstOrDefault(Function(s) s.SectionIndex = update.SectionIndex)
            If section Is Nothing Then Continue For

            If update.ParagraphUpdates IsNot Nothing Then
                For Each paragraphUpdate In update.ParagraphUpdates
                    If ProgressBarModule.CancelOperation Then Exit Sub

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
                    If ProgressBarModule.CancelOperation Then Exit Sub

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


    Private Function GetXmlNodeVisibleText(node As System.Xml.XmlNode, nsMgr As System.Xml.XmlNamespaceManager) As String
        If node Is Nothing Then Return ""

        Dim sb As New StringBuilder()
        Dim textNodes As System.Xml.XmlNodeList = node.SelectNodes(".//w:t", nsMgr)

        For Each tNode As System.Xml.XmlNode In textNodes
            sb.Append(tNode.InnerText)
        Next

        Return sb.ToString()
    End Function

    Private Function BuildCellApplyDebugSummary(cell As TableCellInfo, nsMgr As System.Xml.XmlNamespaceManager) As String
        If cell Is Nothing Then Return "[cell is Nothing]"

        Dim sb As New StringBuilder()
        Dim textNodeCount As Integer = If(cell.TextNodes IsNot Nothing, cell.TextNodes.Count, 0)

        sb.AppendLine($"Row={cell.Row}, Col={cell.Col}, Writable={cell.IsWritable}, GridSpan={cell.GridSpan}, VMerge='{cell.VMerge}'")
        sb.AppendLine($"StoredText='{If(cell.Text, "")}'")
        sb.AppendLine($"XmlVisibleText='{GetXmlNodeVisibleText(cell.XmlNode, nsMgr)}'")
        sb.AppendLine($"TextNodes.Count={textNodeCount}")

        If cell.TextNodes IsNot Nothing AndAlso cell.TextNodes.Count > 0 Then
            For i As Integer = 0 To cell.TextNodes.Count - 1
                Dim textNode As System.Xml.XmlNode = cell.TextNodes(i)
                sb.AppendLine($"  TextNode[{i}]='{If(textNode?.InnerText, "")}'")
            Next
        End If

        If cell.XmlNode IsNot Nothing Then
            sb.AppendLine(GetDebugXmlPayload(cell.XmlNode.OuterXml, "Cell XML"))
        End If

        Return sb.ToString()
    End Function

    Private Sub ApplyTableUpdates(tables As List(Of TableInfo),
                                  updates As List(Of TableUpdateResponse),
                                  nsMgr As System.Xml.XmlNamespaceManager)
        If tables Is Nothing OrElse updates Is Nothing Then Exit Sub

        For Each update In updates
            If ProgressBarModule.CancelOperation Then Exit Sub

            Dim tbl As TableInfo = tables.FirstOrDefault(Function(t) t.TableIndex = update.TableIndex)
            If tbl Is Nothing Then Continue For

            Dim templateRowClone As System.Xml.XmlNode = Nothing
            If update.NewRows IsNot Nothing AndAlso update.NewRows.Count > 0 Then
                templateRowClone = CloneTemplateRowBeforeApply(tbl, update.InsertAfterRow)
            End If

            If update.Updates IsNot Nothing Then
                For Each cellUpdate In update.Updates
                    If ProgressBarModule.CancelOperation Then Exit Sub

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
                    If ProgressBarModule.CancelOperation Then Exit Sub

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

    ''' <summary>
    ''' Removes the w:showingPlcHdr element from w:sdtPr and strips the PlaceholderText
    ''' run style from all runs inside w:sdtContent. This is required when writing a real
    ''' value into an SDT that was previously showing placeholder text.
    ''' </summary>
    Private Sub ClearSdtPlaceholderState(sdtNode As System.Xml.XmlNode, nsMgr As System.Xml.XmlNamespaceManager)
        Dim showingNode As System.Xml.XmlNode = sdtNode.SelectSingleNode("w:sdtPr/w:showingPlcHdr", nsMgr)
        If showingNode IsNot Nothing Then
            showingNode.ParentNode.RemoveChild(showingNode)
        End If

        Dim sdtContent As System.Xml.XmlNode = sdtNode.SelectSingleNode("w:sdtContent", nsMgr)
        If sdtContent Is Nothing Then Exit Sub

        Dim rStyleNodes As System.Xml.XmlNodeList = sdtContent.SelectNodes(
            ".//w:rPr/w:rStyle[@w:val='PlaceholderText' or @w:val='Platzhaltertext']", nsMgr)
        For Each rStyleNode As System.Xml.XmlNode In rStyleNodes
            Dim rPr As System.Xml.XmlNode = rStyleNode.ParentNode
            rPr.RemoveChild(rStyleNode)
            If rPr.ChildNodes.Count = 0 Then
                rPr.ParentNode.RemoveChild(rPr)
            End If
        Next
    End Sub

    ''' <summary>
    ''' Applies a value change to a single form field (SDT or legacy).
    ''' </summary>
    Private Sub ApplyFormFieldUpdate(ff As FormFieldInfo,
                                     fieldUpdate As FormFieldUpdate,
                                     nsMgr As System.Xml.XmlNamespaceManager)
        If ff Is Nothing OrElse fieldUpdate Is Nothing Then Exit Sub
        If Not ff.IsWritable Then Exit Sub

        Dim newValue As String = fieldUpdate.Value

        Select Case ff.FieldType

            Case FormFieldType.Dropdown
                Dim selectedOption As FormFieldOptionInfo = ResolveUpdatedDropdownOption(ff, fieldUpdate)
                If selectedOption Is Nothing Then Exit Sub

                WriteTextIntoSdtContent(ff.XmlNode, selectedOption.DisplayText, nsMgr)
                ApplySdtDefaultRunProperties(ff.XmlNode, nsMgr)

                Dim sdtPr As System.Xml.XmlNode = ff.XmlNode.SelectSingleNode("w:sdtPr", nsMgr)
                If sdtPr IsNot Nothing Then
                    Dim listNode As System.Xml.XmlNode = sdtPr.SelectSingleNode("w:dropDownList", nsMgr)
                    If listNode Is Nothing Then listNode = sdtPr.SelectSingleNode("w:comboBox", nsMgr)

                    If listNode IsNot Nothing Then
                        Dim lastValAttr As System.Xml.XmlAttribute = listNode.OwnerDocument.CreateAttribute("w", "lastValue", WNs)
                        lastValAttr.Value = selectedOption.StoredValue
                        listNode.Attributes.SetNamedItem(lastValAttr)
                    End If
                End If

            Case FormFieldType.Checkbox
                If String.IsNullOrEmpty(newValue) Then Exit Sub

                Dim sdtNode As System.Xml.XmlNode = ff.XmlNode
                Dim checkboxNode As System.Xml.XmlNode = sdtNode.SelectSingleNode("w:sdtPr/w14:checkbox", nsMgr)
                If checkboxNode Is Nothing Then Exit Sub

                Dim isChecked As Boolean = (newValue.ToLowerInvariant() = "true" OrElse newValue = "1")
                Dim xmlVal As String = If(isChecked, "1", "0")

                Dim checkedEl As System.Xml.XmlNode = checkboxNode.SelectSingleNode("w14:checked", nsMgr)
                If checkedEl Is Nothing Then
                    checkedEl = checkboxNode.OwnerDocument.CreateElement("w14", "checked", W14Ns)
                    checkboxNode.AppendChild(checkedEl)
                End If

                Dim valAttr As System.Xml.XmlAttribute = checkedEl.OwnerDocument.CreateAttribute("w14", "val", W14Ns)
                valAttr.Value = xmlVal
                checkedEl.Attributes.SetNamedItem(valAttr)

                Dim charCode As String
                If isChecked Then
                    Dim checkedState As System.Xml.XmlNode = checkboxNode.SelectSingleNode("w14:checkedState/@w14:val", nsMgr)
                    charCode = If(checkedState?.Value, "2612")
                Else
                    Dim uncheckedState As System.Xml.XmlNode = checkboxNode.SelectSingleNode("w14:uncheckedState/@w14:val", nsMgr)
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

                Dim fldCharNode As System.Xml.XmlNode = ff.XmlNode
                Dim ffData As System.Xml.XmlNode = fldCharNode.SelectSingleNode("w:ffData", nsMgr)
                If ffData Is Nothing Then Exit Sub

                Dim ddList As System.Xml.XmlNode = ffData.SelectSingleNode("w:ddList", nsMgr)
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

                Dim resultNode As System.Xml.XmlNode = ddList.SelectSingleNode("w:result", nsMgr)
                If resultNode Is Nothing Then
                    resultNode = ddList.OwnerDocument.CreateElement("w", "result", WNs)
                    ddList.AppendChild(resultNode)
                End If

                Dim resultValAttr As System.Xml.XmlAttribute = resultNode.OwnerDocument.CreateAttribute("w", "val", WNs)
                resultValAttr.Value = targetIdx.ToString()
                resultNode.Attributes.SetNamedItem(resultValAttr)

                UpdateLegacyFieldDisplayText(fldCharNode, selectedOption.DisplayText, nsMgr)

            Case FormFieldType.LegacyCheckbox
                If String.IsNullOrEmpty(newValue) Then Exit Sub

                Dim fldCharNode As System.Xml.XmlNode = ff.XmlNode
                Dim ffData As System.Xml.XmlNode = fldCharNode.SelectSingleNode("w:ffData", nsMgr)
                If ffData Is Nothing Then Exit Sub

                Dim checkBox As System.Xml.XmlNode = ffData.SelectSingleNode("w:checkBox", nsMgr)
                If checkBox Is Nothing Then Exit Sub

                Dim isChecked As Boolean = (newValue.ToLowerInvariant() = "true" OrElse newValue = "1")

                Dim checkedNode As System.Xml.XmlNode = checkBox.SelectSingleNode("w:checked", nsMgr)
                If checkedNode Is Nothing Then
                    checkedNode = checkBox.OwnerDocument.CreateElement("w", "checked", WNs)
                    checkBox.AppendChild(checkedNode)
                End If

                Dim checkedValAttr As System.Xml.XmlAttribute = checkedNode.OwnerDocument.CreateAttribute("w", "val", WNs)
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

    ''' <summary>
    ''' Updates the display text for a legacy form field dropdown.
    ''' Walks sibling runs after the fldChar begin until fldChar separate/end is found.
    ''' </summary>
    Private Sub UpdateLegacyFieldDisplayText(fldCharBeginRun As System.Xml.XmlNode, newText As String, nsMgr As System.Xml.XmlNamespaceManager)
        Dim beginRun As System.Xml.XmlNode = fldCharBeginRun.ParentNode
        If beginRun Is Nothing Then Exit Sub

        Dim currentNode As System.Xml.XmlNode = beginRun.NextSibling
        Dim foundSeparate As Boolean = False
        Dim textUpdated As Boolean = False

        While currentNode IsNot Nothing
            If currentNode.LocalName = "r" Then
                Dim fc As System.Xml.XmlNode = currentNode.SelectSingleNode("w:fldChar", nsMgr)
                If fc IsNot Nothing Then
                    Dim fcType As System.Xml.XmlNode = fc.Attributes("w:fldCharType")
                    If fcType IsNot Nothing Then
                        If fcType.Value = "separate" Then
                            foundSeparate = True
                        ElseIf fcType.Value = "end" Then
                            Exit While
                        End If
                    End If
                ElseIf foundSeparate Then
                    Dim tNodes As System.Xml.XmlNodeList = currentNode.SelectNodes("w:t", nsMgr)
                    If tNodes.Count > 0 AndAlso Not textUpdated Then
                        SetTextNodeWithSpacePreserve(tNodes(0), newText)
                        textUpdated = True
                        For i As Integer = 1 To tNodes.Count - 1
                            SetTextNodeWithSpacePreserve(tNodes(i), "")
                        Next
                    ElseIf tNodes.Count > 0 Then
                        For Each tn As System.Xml.XmlNode In tNodes
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
                                   nsMgr As System.Xml.XmlNamespaceManager,
                                   Optional templateRowClone As System.Xml.XmlNode = Nothing)
        If tbl Is Nothing OrElse newRows Is Nothing OrElse newRows.Count = 0 Then Exit Sub
        If ProgressBarModule.CancelOperation Then Exit Sub

        Dim templateRowNode As System.Xml.XmlNode = templateRowClone
        If templateRowNode Is Nothing Then
            templateRowNode = CloneTemplateRowBeforeApply(tbl, insertAfterRow)
        End If
        If templateRowNode Is Nothing Then Exit Sub

        If RowContainsVerticalMerge(templateRowNode, nsMgr) Then
            Exit Sub
        End If

        Dim refNode As System.Xml.XmlNode
        If insertAfterRow >= 0 AndAlso insertAfterRow < tbl.Rows.Count Then
            refNode = tbl.Rows(insertAfterRow).XmlNode
        Else
            refNode = tbl.Rows(tbl.Rows.Count - 1).XmlNode
        End If
        If refNode Is Nothing Then Exit Sub

        Dim insertParent As System.Xml.XmlNode = refNode.ParentNode
        If insertParent Is Nothing Then Exit Sub

        For Each rowTexts In newRows
            If ProgressBarModule.CancelOperation Then Exit Sub

            Dim clonedRow As System.Xml.XmlNode = templateRowNode.CloneNode(deep:=True)

            Dim tblHeaderNode As System.Xml.XmlNode = clonedRow.SelectSingleNode("w:trPr/w:tblHeader", nsMgr)
            If tblHeaderNode IsNot Nothing Then
                tblHeaderNode.ParentNode.RemoveChild(tblHeaderNode)
            End If

            Dim vMergeNodes As System.Xml.XmlNodeList = clonedRow.SelectNodes(".//w:tcPr/w:vMerge", nsMgr)
            For Each vmNode As System.Xml.XmlNode In vMergeNodes
                vmNode.ParentNode.RemoveChild(vmNode)
            Next

            For Each innerSdt As System.Xml.XmlNode In clonedRow.SelectNodes(".//w:sdt", nsMgr)
                ClearSdtPlaceholderState(innerSdt, nsMgr)
            Next

            Dim tcNodes As System.Xml.XmlNodeList = clonedRow.SelectNodes(".//w:tc", nsMgr)
            For ci As Integer = 0 To tcNodes.Count - 1
                Dim tcNode As System.Xml.XmlNode = tcNodes(ci)
                Dim cellText As String = If(ci < rowTexts.Count, rowTexts(ci), "")

                Dim pNode As System.Xml.XmlNode = tcNode.SelectSingleNode("w:p", nsMgr)
                If pNode Is Nothing Then
                    pNode = tcNode.OwnerDocument.CreateElement("w", "p", WNs)
                    tcNode.AppendChild(pNode)
                End If

                Dim toRemove As New List(Of System.Xml.XmlNode)()
                For Each child As System.Xml.XmlNode In pNode.ChildNodes
                    If child.NodeType = System.Xml.XmlNodeType.Element AndAlso child.LocalName = "pPr" Then Continue For
                    toRemove.Add(child)
                Next
                For Each child As System.Xml.XmlNode In toRemove
                    pNode.RemoveChild(child)
                Next

                pNode.AppendChild(CreateCleanTextRun(pNode, cellText, nsMgr))
            Next

            insertParent.InsertAfter(clonedRow, refNode)
            refNode = clonedRow
        Next
    End Sub


#End Region

#Region "Debug Logging"

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
        Dim sb As New StringBuilder()

        For Each tbl In tables
            sb.AppendLine($"Table {tbl.TableIndex}")

            For Each row In tbl.Rows
                For Each cell In row.Cells
                    If cell.FormFields Is Nothing OrElse cell.FormFields.Count = 0 Then Continue For

                    sb.AppendLine($"  Cell r={cell.Row}, c={cell.Col}, text='{cell.Text}'")

                    For Each ff In cell.FormFields
                        sb.AppendLine($"    Field {ff.FieldId} | type={ff.FieldType} | name='{ff.Name}' | display='{ff.CurrentValue}' | stored='{ff.CurrentOptionValue}'")

                        If ff.Options IsNot Nothing AndAlso ff.Options.Count > 0 Then
                            For i As Integer = 0 To ff.Options.Count - 1
                                Dim opt As FormFieldOptionInfo = ff.Options(i)
                                sb.AppendLine($"      [{i}] text='{opt.DisplayText}' | value='{opt.StoredValue}'")
                            Next
                        End If
                    Next
                Next
            Next
        Next

        Return sb.ToString()
    End Function

    Private Sub DumpDocxPackageDiagnostics(docxPath As String)
        If Not TableCompleteDebugEnabled Then Exit Sub
        If String.IsNullOrWhiteSpace(docxPath) Then Exit Sub
        If Not File.Exists(docxPath) Then Exit Sub

        Try
            Dim inventory As New StringBuilder()

            Using archive As ZipArchive = ZipFile.OpenRead(docxPath)
                inventory.AppendLine($"Package: {docxPath}")
                inventory.AppendLine($"Entries: {archive.Entries.Count}")
                inventory.AppendLine()

                For Each entry As ZipArchiveEntry In archive.Entries
                    inventory.AppendLine($"{entry.FullName} | Length={entry.Length} | CompressedLength={entry.CompressedLength}")
                Next

                WriteTableCompletionDebug("DOCX package inventory", inventory.ToString())

                For Each entry As ZipArchiveEntry In archive.Entries
                    Dim entryName As String = entry.FullName.Replace("\"c, "/"c)
                    Dim lowerName As String = entryName.ToLowerInvariant()

                    If lowerName.EndsWith(".xml") OrElse lowerName.EndsWith(".rels") Then
                        WriteTableCompletionDebug(
                            $"DOCX XML part: {entryName}",
                            GetDebugXmlPayload(ReadZipEntryAsText(entry), $"DOCX XML part: {entryName}"))
                    Else
                        Dim binarySummary As String = BuildZipEntryBinarySummary(entry)
                        If Not String.IsNullOrWhiteSpace(binarySummary) Then
                            WriteTableCompletionDebug(
                                $"DOCX binary part: {entryName}",
                                binarySummary)
                        End If
                    End If
                Next
            End Using
        Catch ex As Exception
            WriteTableCompletionDebug("DumpDocxPackageDiagnostics error", ex.ToString())
        End Try
    End Sub

    Private Function ReadZipEntryAsText(entry As ZipArchiveEntry) As String
        Try
            Using entryStream As Stream = entry.Open()
                Using reader As New StreamReader(entryStream, Encoding.UTF8, detectEncodingFromByteOrderMarks:=True)
                    Return reader.ReadToEnd()
                End Using
            End Using
        Catch ex As Exception
            Return $"[Failed to read entry as text: {ex.Message}]"
        End Try
    End Function

    Private Function BuildZipEntryBinarySummary(entry As ZipArchiveEntry) As String
        If entry Is Nothing Then Return ""
        If entry.Length = 0 Then Return $"Length=0"

        Dim lowerName As String = entry.FullName.ToLowerInvariant()
        Dim shouldLogBinary As Boolean =
            lowerName.Contains("activex") OrElse
            lowerName.Contains("embedding") OrElse
            lowerName.Contains("embeddings") OrElse
            lowerName.Contains("ole") OrElse
            lowerName.Contains("control") OrElse
            lowerName.Contains("object") OrElse
            lowerName.Contains("customxml") OrElse
            lowerName.EndsWith(".bin")

        If Not shouldLogBinary Then Return ""

        Try
            Dim previewLength As Integer = CInt(Math.Min(128, entry.Length))
            Dim buffer(previewLength - 1) As Byte

            Using entryStream As Stream = entry.Open()
                Dim bytesRead As Integer = entryStream.Read(buffer, 0, previewLength)

                Dim sb As New StringBuilder()
                sb.AppendLine($"Length={entry.Length}")
                sb.AppendLine($"PreviewBytes={bytesRead}")
                sb.AppendLine("HexPreview:")
                sb.AppendLine(ByteArrayToHex(buffer, bytesRead))
                Return sb.ToString()
            End Using
        Catch ex As Exception
            Return $"Length={entry.Length}{vbCrLf}[Failed to read binary preview: {ex.Message}]"
        End Try
    End Function

    Private Function ByteArrayToHex(buffer As Byte(), count As Integer) As String
        If buffer Is Nothing OrElse count <= 0 Then Return ""

        Dim sb As New StringBuilder()

        For i As Integer = 0 To count - 1
            If i > 0 Then
                If i Mod 16 = 0 Then
                    sb.AppendLine()
                Else
                    sb.Append(" ")
                End If
            End If

            sb.Append(buffer(i).ToString("X2"))
        Next

        Return sb.ToString()
    End Function

    Private Function GetDebugXmlPayload(xmlContent As String, xmlLabel As String) As String
        If TableCompleteDebugIncludeXml Then Return If(xmlContent, "")

        Dim sb As New StringBuilder()
        sb.AppendLine($"[{xmlLabel} omitted]")
        sb.AppendLine($"Set TableCompleteDebugIncludeXml = True to include raw XML in this log.")
        Return sb.ToString()
    End Function

#End Region

#Region "Saved DOCX Verification"


    Private Function NormalizeWordInteropCellText(text As String) As String
        If text Is Nothing Then Return ""

        Dim result As String = text

        result = result.Replace(vbCr & ChrW(7), "")
        result = result.Replace(ChrW(7).ToString(), "")
        result = result.Replace(vbCr, "")
        result = result.Replace(vbLf, "")

        Return result
    End Function

    Private Sub VerifySavedDocxWithWordInterop(docxPath As String)
        If Not TableCompleteDebugEnabled Then Exit Sub
        If Not TableCompleteDebugVerifySavedDocxWithWordInterop Then Exit Sub
        If String.IsNullOrWhiteSpace(docxPath) Then Exit Sub
        If Not File.Exists(docxPath) Then Exit Sub

        Dim wordApp As Word.Application = Nothing
        Dim doc As Word.Document = Nothing
        Dim oldScreenUpdating As Boolean = True

        Try
            wordApp = Globals.ThisAddIn.Application
            oldScreenUpdating = wordApp.ScreenUpdating
            wordApp.ScreenUpdating = False

            doc = wordApp.Documents.Open(
                docxPath,
                ReadOnly:=True,
                Visible:=False,
                AddToRecentFiles:=False)

            Dim sb As New StringBuilder()
            sb.AppendLine($"SavedDocx='{docxPath}'")
            sb.AppendLine($"Word.Tables.Count={doc.Tables.Count}")

            Dim targetWordTableIndex As Integer = 8 ' tableIndex 7 => Word table 8

            If doc.Tables.Count < targetWordTableIndex Then
                sb.AppendLine($"Target Word table {targetWordTableIndex} does not exist.")
                WriteTableCompletionDebug("Saved DOCX Word Interop verification", sb.ToString())
                Exit Sub
            End If

            Dim tbl As Word.Table = doc.Tables(targetWordTableIndex)

            sb.AppendLine($"Inspecting Word table {targetWordTableIndex}")
            sb.AppendLine($"Rows.Count={tbl.Rows.Count}")
            sb.AppendLine($"Columns.Count={tbl.Columns.Count}")
            sb.AppendLine()

            For r As Integer = 1 To Math.Min(tbl.Rows.Count, 8)
                For c As Integer = 1 To Math.Min(tbl.Columns.Count, 3)
                    Try
                        Dim rawText As String = tbl.Cell(r, c).Range.Text
                        Dim normalizedText As String = NormalizeWordInteropCellText(rawText)

                        sb.AppendLine($"Cell({r},{c})='{normalizedText}'")
                    Catch ex As Exception
                        sb.AppendLine($"Cell({r},{c})=<error: {ex.Message}>")
                    End Try
                Next

                sb.AppendLine()
            Next

            Try
                Dim targetText As String = NormalizeWordInteropCellText(tbl.Cell(5, 2).Range.Text)
                sb.AppendLine("Target check:")
                sb.AppendLine($"tableIndex=7,row=4,col=1 => Word Cell(5,2)='{targetText}'")
            Catch ex As Exception
                sb.AppendLine("Target check:")
                sb.AppendLine($"tableIndex=7,row=4,col=1 => Word Cell(5,2)=<error: {ex.Message}>")
            End Try

            WriteTableCompletionDebug("Saved DOCX Word Interop verification", sb.ToString())

        Catch ex As Exception
            WriteTableCompletionDebug("VerifySavedDocxWithWordInterop error", ex.ToString())

        Finally
            If doc IsNot Nothing Then
                Try : doc.Close(WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
            End If

            If wordApp IsNot Nothing Then
                Try : wordApp.ScreenUpdating = oldScreenUpdating : Catch : End Try
            End If
        End Try
    End Sub

    Private Function LoadTablesFromSavedDocx(docxPath As String) As List(Of TableInfo)
        If String.IsNullOrWhiteSpace(docxPath) Then Return Nothing
        If Not File.Exists(docxPath) Then Return Nothing

        Dim tempDir As String = Path.Combine(Path.GetTempPath(), $"{AN2}_tblverify_{Guid.NewGuid():N}")

        Try
            ZipFile.ExtractToDirectory(docxPath, tempDir)

            Dim documentXmlPath As String = Path.Combine(tempDir, "word", "document.xml")
            If Not File.Exists(documentXmlPath) Then Return Nothing

            Dim xmlDoc As New System.Xml.XmlDocument()
            xmlDoc.PreserveWhitespace = True
            xmlDoc.Load(documentXmlPath)

            Dim nsMgr As New System.Xml.XmlNamespaceManager(xmlDoc.NameTable)
            nsMgr.AddNamespace("w", WNs)
            nsMgr.AddNamespace("w14", W14Ns)

            Return ExtractTablesFromXml(xmlDoc, nsMgr)

        Catch ex As Exception
            WriteTableCompletionDebug("LoadTablesFromSavedDocx error", ex.ToString())
            Return Nothing

        Finally
            If Directory.Exists(tempDir) Then
                Try : Directory.Delete(tempDir, recursive:=True) : Catch : End Try
            End If
        End Try
    End Function

    Private Sub VerifySavedDocxCellUpdates(docxPath As String,
                                           updates As List(Of TableUpdateResponse))
        If Not TableCompleteDebugEnabled Then Exit Sub
        If Not TableCompleteDebugVerifySavedDocx Then Exit Sub
        If String.IsNullOrWhiteSpace(docxPath) Then Exit Sub
        If Not File.Exists(docxPath) Then Exit Sub
        If updates Is Nothing Then Exit Sub

        Dim savedTables As List(Of TableInfo) = LoadTablesFromSavedDocx(docxPath)
        If savedTables Is Nothing Then
            WriteTableCompletionDebug("Saved DOCX verification",
                                      "Could not reload saved DOCX for verification.")
            Exit Sub
        End If

        Dim checkedCount As Integer = 0
        Dim mismatchCount As Integer = 0
        Dim missingCount As Integer = 0

        For Each update In updates
            If update Is Nothing OrElse update.Updates Is Nothing Then Continue For

            Dim savedTable As TableInfo = savedTables.FirstOrDefault(Function(t) t.TableIndex = update.TableIndex)
            If savedTable Is Nothing Then
                WriteTableCompletionDebug("Saved DOCX verification missing table",
                                          $"TableIndex={update.TableIndex}")
                missingCount += 1
                Continue For
            End If

            ' Updates reference pre-insertion row indices. The saved DOCX contains the
            ' inserted rows, so any original row whose index is strictly greater than
            ' InsertAfterRow has been shifted down by NewRows.Count.
            Dim insertedCount As Integer = If(update.NewRows IsNot Nothing, update.NewRows.Count, 0)
            Dim insertAfter As Integer = update.InsertAfterRow

            For Each cellUpdate In update.Updates
                checkedCount += 1

                Dim expectedRowIndex As Integer = cellUpdate.Row
                If insertedCount > 0 AndAlso insertAfter >= 0 AndAlso cellUpdate.Row > insertAfter Then
                    expectedRowIndex = cellUpdate.Row + insertedCount
                End If

                Dim savedRow As TableRowInfo = savedTable.Rows.FirstOrDefault(Function(r) r.RowIndex = expectedRowIndex)
                If savedRow Is Nothing Then
                    WriteTableCompletionDebug("Saved DOCX verification missing row",
                                              $"Table={update.TableIndex}, OriginalRow={cellUpdate.Row}, ShiftedRow={expectedRowIndex}")
                    missingCount += 1
                    Continue For
                End If

                Dim savedCell As TableCellInfo = savedRow.Cells.FirstOrDefault(Function(c) c.Col = cellUpdate.Col)
                If savedCell Is Nothing Then
                    WriteTableCompletionDebug("Saved DOCX verification missing cell",
                                              $"Table={update.TableIndex}, Row={expectedRowIndex}, Col={cellUpdate.Col}")
                    missingCount += 1
                    Continue For
                End If

                Dim expectedText As String = If(cellUpdate.Text, "")
                Dim actualText As String = If(savedCell.Text, "")

                If Not String.Equals(expectedText, actualText, StringComparison.Ordinal) Then
                    mismatchCount += 1
                    WriteTableCompletionDebug("Saved DOCX verification mismatch",
                        $"Table={update.TableIndex}, OriginalRow={cellUpdate.Row}, ShiftedRow={expectedRowIndex}, Col={cellUpdate.Col}" & vbCrLf &
                        $"Expected='{expectedText}'" & vbCrLf &
                        $"Actual='{actualText}'")
                End If
            Next
        Next

        WriteTableCompletionDebug("Saved DOCX verification summary",
                                  $"CheckedUpdates={checkedCount}, MissingTargets={missingCount}, Mismatches={mismatchCount}, SavedDocx='{docxPath}'")
    End Sub

#End Region

#Region "Cell Write Helpers"

    Private Function NormalizeRequestedTextAction(action As String) As String
        Dim normalized As String = If(action, "").Trim().ToLowerInvariant()

        Select Case normalized
            Case "keep", "replace", "complete", "clear"
                Return normalized
            Case Else
                Return "replace"
        End Select
    End Function


    ''' <summary>
    ''' Returns the w:rPr of a run inside the given scope that is NOT carrying a
    ''' placeholder character style. Used as a template for newly created runs so we
    ''' do not inherit greyed-out placeholder formatting.
    ''' </summary>
    Private Function FindNonPlaceholderRunProperties(scope As System.Xml.XmlNode,
                                                     nsMgr As System.Xml.XmlNamespaceManager) As System.Xml.XmlNode
        If scope Is Nothing Then Return Nothing

        Dim searchScope As System.Xml.XmlNode = scope
        While searchScope IsNot Nothing AndAlso
              searchScope.LocalName <> "tc" AndAlso
              searchScope.LocalName <> "body"
            searchScope = searchScope.ParentNode
        End While
        If searchScope Is Nothing Then searchScope = scope

        Dim rPrNodes As System.Xml.XmlNodeList = searchScope.SelectNodes(".//w:r/w:rPr", nsMgr)
        For Each rPr As System.Xml.XmlNode In rPrNodes
            Dim rStyle As System.Xml.XmlNode = rPr.SelectSingleNode("w:rStyle/@w:val", nsMgr)
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

    ''' <summary>
    ''' Creates a new clean w:r containing a single w:t with the given text, inheriting
    ''' run properties from a non-placeholder run in the same cell when available.
    ''' </summary>
    Private Function CreateCleanTextRun(parentParagraph As System.Xml.XmlNode,
                                        text As String,
                                        nsMgr As System.Xml.XmlNamespaceManager) As System.Xml.XmlNode
        Dim ownerDoc As System.Xml.XmlDocument = parentParagraph.OwnerDocument
        Dim rNode As System.Xml.XmlElement = ownerDoc.CreateElement("w", "r", WNs)

        Dim templateRpr As System.Xml.XmlNode = FindNonPlaceholderRunProperties(parentParagraph, nsMgr)
        If templateRpr IsNot Nothing Then
            rNode.AppendChild(templateRpr.CloneNode(deep:=True))
        End If

        Dim tNode As System.Xml.XmlElement = ownerDoc.CreateElement("w", "t", WNs)
        rNode.AppendChild(tNode)
        SetTextNodeWithSpacePreserve(tNode, If(text, ""))

        Return rNode
    End Function

    ''' <summary>
    ''' Single entry point for writing plain text into a table cell. Performs three
    ''' deterministic invisibility-fix passes before writing:
    ''' 1) Clears w:showingPlcHdr on every ancestor w:sdt up to the document root.
    ''' 2) Strips w:vanish, w:webHidden and placeholder w:rStyle from every w:rPr
    '''    inside the cell, including paragraph-mark rPr inside w:pPr.
    ''' 3) Drops placeholder w:pStyle from the target paragraph.
    ''' For previously-empty cells, all non-pPr children of the first paragraph are
    ''' removed before inserting a fresh clean run, so leftover placeholder runs cannot
    ''' keep the paragraph in placeholder presentation mode.
    ''' </summary>
    Private Sub WriteTextIntoCell(cell As TableCellInfo,
                                  newText As String,
                                  nsMgr As System.Xml.XmlNamespaceManager)
        If cell Is Nothing OrElse cell.XmlNode Is Nothing Then Exit Sub

        Dim normalized As String = If(newText, "")
        Dim currentText As String = If(cell.Text, "")
        If String.Equals(currentText, normalized, StringComparison.Ordinal) Then Exit Sub

        Dim tcNode As System.Xml.XmlNode = cell.XmlNode

        ' (1) Clear placeholder state on every enclosing SDT, all the way to the body.
        ClearAllAncestorSdtPlaceholderStates(tcNode, nsMgr)

        ' (2) Strip invisibility / placeholder run-properties from the whole cell subtree.
        SanitizeCellRunFormatting(tcNode, nsMgr)

        If cell.TextNodes IsNot Nothing AndAlso cell.TextNodes.Count > 0 Then
            ' Update existing text nodes in place (preserves real run formatting).
            SetTextNodeWithSpacePreserve(cell.TextNodes(0), normalized)
            For i As Integer = 1 To cell.TextNodes.Count - 1
                SetTextNodeWithSpacePreserve(cell.TextNodes(i), "")
            Next
            cell.Text = normalized
            Exit Sub
        End If

        ' Previously-empty cell: rebuild the first paragraph deterministically.
        Dim pNode As System.Xml.XmlNode = tcNode.SelectSingleNode("w:p", nsMgr)
        If pNode Is Nothing Then
            pNode = tcNode.OwnerDocument.CreateElement("w", "p", WNs)
            tcNode.AppendChild(pNode)
        End If

        ' (3) Remove every child of the paragraph except w:pPr, so leftover empty runs
        ' carrying placeholder w:rStyle (or w:vanish) cannot suppress our new text.
        Dim childrenToRemove As New List(Of System.Xml.XmlNode)()
        For Each ch As System.Xml.XmlNode In pNode.ChildNodes
            If ch.NodeType = System.Xml.XmlNodeType.Element AndAlso ch.LocalName = "pPr" Then Continue For
            childrenToRemove.Add(ch)
        Next
        For Each ch As System.Xml.XmlNode In childrenToRemove
            pNode.RemoveChild(ch)
        Next

        ' (4) Drop placeholder paragraph style on the target paragraph.
        SanitizeParagraphProperties(pNode, nsMgr)

        ' (5) Insert a single fresh clean run. CreateCleanTextRun does not attach an
        ' explicit w:rPr, which is correct: the run then inherits only from the
        ' paragraph style (which we just sanitized), not from placeholder character
        ' styling carried by leftover empty runs.
        Dim newRun As System.Xml.XmlNode = CreateCleanTextRun(pNode, normalized, nsMgr)
        pNode.AppendChild(newRun)

        Dim newTextNode As System.Xml.XmlNode = newRun.SelectSingleNode("w:t", nsMgr)
        If newTextNode IsNot Nothing Then
            If cell.TextNodes Is Nothing Then cell.TextNodes = New List(Of System.Xml.XmlNode)()
            cell.TextNodes.Add(newTextNode)
        End If

        cell.Text = normalized
    End Sub

    ''' <summary>
    ''' Writes plain text into the w:sdtContent of an SDT. In addition to clearing
    ''' w:showingPlcHdr, this sanitizes every w:rPr inside w:sdtContent so the dropdown /
    ''' text-input value renders normally instead of inheriting the SDT's placeholder
    ''' character style.
    ''' </summary>
    Private Sub WriteTextIntoSdtContent(sdtNode As System.Xml.XmlNode,
                                        newText As String,
                                        nsMgr As System.Xml.XmlNamespaceManager)
        If sdtNode Is Nothing Then Exit Sub

        ' Clear placeholder state on this SDT and on every SDT above it.
        ClearAllAncestorSdtPlaceholderStates(sdtNode, nsMgr)

        Dim sdtContent As System.Xml.XmlNode = sdtNode.SelectSingleNode("w:sdtContent", nsMgr)
        If sdtContent Is Nothing Then Exit Sub

        ' Strip invisibility / placeholder formatting from every run inside the SDT content.
        SanitizeCellRunFormatting(sdtContent, nsMgr)

        ' Also clear the paragraph-level placeholder style on any paragraphs inside.
        For Each pInside As System.Xml.XmlNode In sdtContent.SelectNodes(".//w:p", nsMgr)
            SanitizeParagraphProperties(pInside, nsMgr)
        Next

        Dim contentTextNodes As System.Xml.XmlNodeList = sdtContent.SelectNodes(".//w:t", nsMgr)
        If contentTextNodes.Count > 0 Then
            SetTextNodeWithSpacePreserve(contentTextNodes(0), If(newText, ""))
            For i As Integer = 1 To contentTextNodes.Count - 1
                SetTextNodeWithSpacePreserve(contentTextNodes(i), "")
            Next
            Exit Sub
        End If

        ' No existing text node: build one cleanly.
        Dim ownerDoc As System.Xml.XmlDocument = sdtContent.OwnerDocument
        Dim sdtParentName As String = If(sdtNode.ParentNode?.LocalName, "")

        sdtContent.InnerXml = ""

        If sdtParentName = "tc" OrElse sdtParentName = "body" Then
            Dim pNode As System.Xml.XmlElement = ownerDoc.CreateElement("w", "p", WNs)
            sdtContent.AppendChild(pNode)
            pNode.AppendChild(CreateCleanTextRun(pNode, If(newText, ""), nsMgr))
        Else
            Dim hostParagraph As System.Xml.XmlNode = sdtNode.ParentNode
            While hostParagraph IsNot Nothing AndAlso hostParagraph.LocalName <> "p"
                hostParagraph = hostParagraph.ParentNode
            End While
            If hostParagraph Is Nothing Then hostParagraph = sdtContent
            sdtContent.AppendChild(CreateCleanTextRun(hostParagraph, If(newText, ""), nsMgr))
        End If
    End Sub

    ''' <summary>
    ''' Walks from <paramref name="startNode"/> up to the document root and clears
    ''' w:showingPlcHdr plus placeholder run styling on every enclosing w:sdt — including
    ''' SDTs that wrap the entire row, the entire table, or a whole section. Stopping at
    ''' w:tr / w:tbl (as the previous helper did) misses repeating-section / repeating-row
    ''' content controls that keep the cell's text invisible.
    ''' </summary>
    Private Sub ClearAllAncestorSdtPlaceholderStates(startNode As System.Xml.XmlNode,
                                                     nsMgr As System.Xml.XmlNamespaceManager)
        If startNode Is Nothing Then Exit Sub

        Dim current As System.Xml.XmlNode = startNode
        While current IsNot Nothing
            If current.NodeType = System.Xml.XmlNodeType.Element AndAlso current.LocalName = "sdt" Then
                ClearSdtPlaceholderState(current, nsMgr)
            End If
            current = current.ParentNode
        End While
    End Sub

    ''' <summary>
    ''' Strips every invisibility / placeholder-styling source from every w:rPr inside the
    ''' given cell (or any subtree): w:vanish, w:webHidden, and w:rStyle pointing to
    ''' "Platzhaltertext" / "PlaceholderText". Also drops the rPr if it ends up empty so it
    ''' cannot re-introduce inherited placeholder formatting later.
    ''' </summary>
    Private Sub SanitizeCellRunFormatting(scope As System.Xml.XmlNode,
                                          nsMgr As System.Xml.XmlNamespaceManager)
        If scope Is Nothing Then Exit Sub

        Dim rPrNodes As System.Xml.XmlNodeList = scope.SelectNodes(".//w:rPr", nsMgr)
        For Each rPr As System.Xml.XmlNode In rPrNodes
            Dim toRemove As New List(Of System.Xml.XmlNode)()

            For Each child As System.Xml.XmlNode In rPr.ChildNodes
                Select Case child.LocalName
                    Case "vanish", "webHidden"
                        toRemove.Add(child)

                    Case "rStyle"
                        Dim valAttr As System.Xml.XmlNode = child.Attributes("w:val")
                        If valAttr IsNot Nothing AndAlso
                           (valAttr.Value.Equals("Platzhaltertext", StringComparison.OrdinalIgnoreCase) OrElse
                            valAttr.Value.Equals("PlaceholderText", StringComparison.OrdinalIgnoreCase)) Then
                            toRemove.Add(child)
                        End If
                End Select
            Next

            For Each n As System.Xml.XmlNode In toRemove
                rPr.RemoveChild(n)
            Next

            If rPr.ChildNodes.Count = 0 AndAlso rPr.ParentNode IsNot Nothing Then
                rPr.ParentNode.RemoveChild(rPr)
            End If
        Next
    End Sub

    ''' <summary>
    ''' Removes a placeholder paragraph style (w:pStyle val="Platzhaltertext"/"PlaceholderText")
    ''' from a paragraph's w:pPr. Run-property sanitation inside w:pPr/w:rPr is handled by
    ''' SanitizeCellRunFormatting.
    ''' </summary>
    Private Sub SanitizeParagraphProperties(pNode As System.Xml.XmlNode,
                                            nsMgr As System.Xml.XmlNamespaceManager)
        If pNode Is Nothing Then Exit Sub
        Dim pPr As System.Xml.XmlNode = pNode.SelectSingleNode("w:pPr", nsMgr)
        If pPr Is Nothing Then Exit Sub

        Dim toRemove As New List(Of System.Xml.XmlNode)()
        For Each pStyle As System.Xml.XmlNode In pPr.SelectNodes("w:pStyle", nsMgr)
            Dim valAttr As System.Xml.XmlNode = pStyle.Attributes("w:val")
            If valAttr Is Nothing Then Continue For
            If valAttr.Value.Equals("Platzhaltertext", StringComparison.OrdinalIgnoreCase) OrElse
               valAttr.Value.Equals("PlaceholderText", StringComparison.OrdinalIgnoreCase) Then
                toRemove.Add(pStyle)
            End If
        Next
        For Each n As System.Xml.XmlNode In toRemove
            pPr.RemoveChild(n)
        Next
    End Sub

#End Region

End Class
