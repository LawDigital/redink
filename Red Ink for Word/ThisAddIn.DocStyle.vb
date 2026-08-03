' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.DocStyle.vb
' Purpose: Implements DocStyle template extraction (paragraph/style formatting -> JSON) and template
'          application (LLM mapping -> Word style/format changes) for the active Word document.
'
' Architecture:
'  - Settings Persistence: `DocStyleSettings` loads/saves DocStyle options via `My.Settings`.
'  - Template Creation: `ExtractParagraphStylesToJson` walks selected paragraphs (or document content),
'    captures paragraph/font/list/tab/border/shading info, and writes a JSON template to disk.
'  - Word Style Definitions: `ExtractFullStyleDefinition` serializes Word `Style` definitions including
'    paragraph/font/tab/list template data; `ApplyWdStyleDefinitionsToDocument` can create/update styles.
'  - Template Application: `ApplyStyleTemplate` loads a JSON template file, optionally applies style
'    definitions, builds an LLM prompt for the selected paragraphs, and applies the returned mapping.
'  - LLM Response Handling: `ApplyStylesFromTemplate` parses a JSON array mapping paragraph indices to
'    user style names, confidence and optional delete-prefix instructions.
'  - List Numbering: `ApplyNumberingRestarts` supports rule-based and LLM-assisted numbering restarts,
'    preserving paragraph properties via `CaptureParaProps`/`RestoreParaProps`.
'
' External Dependencies:
'  - Microsoft.Office.Interop.Word for document/style/list operations.
'  - Newtonsoft.Json / JObject / JArray for JSON templates and LLM response parsing.
'  - SharedLibrary.SharedMethods for UI, file editing, clipboard, interpolation and LLM invocation.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports Microsoft.Office.Interop.Word
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Cache for shared list templates during style definition application.
    ''' Key is a template fingerprint cache key, value is the created <see cref="Word.ListTemplate"/> instance.
    ''' </summary>
    Private _sharedListTemplates As New Dictionary(Of String, Word.ListTemplate)()

    ''' <summary>
    ''' When True, bullet list paragraphs break rule-based numbering runs.
    ''' </summary>
    Private Const RuleBreakOnBullets As Boolean = False

    ''' <summary>
    ''' Indentation step (points) used when inferring list levels from paragraph indentation.
    ''' </summary>
    Private Const IndentStepPoints As Single = 10.0F

    ''' <summary>
    ''' Maximum preview length (characters) used in report/debug output for paragraph text.
    ''' </summary>
    Private Const ReportParaPreviewLen As Integer = 50

    ''' <summary>
    ''' Default maximum outline level treated as "heading-like" when rule-based numbering reset is enabled.
    ''' </summary>
    Private Const DefaultRuleHeadingOutlineLevelMax As Integer = 5

    ''' <summary>
    ''' Prompt snippet describing the JSON response format expected from the LLM during style application.
    ''' </summary>
    Public ApplyStylesResponseFormat As String = ""

    ''' <summary>
    ''' Minimal response format used when only style application is required.
    ''' </summary>
    Private Const ApplyStylesResponseFormat_Minimal As String =
                "[" & vbCrLf &
                "  {""paragraphIndex"": 1, ""userStyleName"": ""<user style name>"", ""confidence"": <0-100>, ""preserveList"": <true|false>, ""restartNumbering"": <true|false>}," & vbCrLf &
                "  ..." & vbCrLf &
                "]"

    ''' <summary>
    ''' Extended response format used when optional text amendment instructions (deletePrefix) are enabled.
    ''' </summary>
    Private Const ApplyStylesResponseFormat_Extended As String =
                "[" & vbCrLf &
                "  {""paragraphIndex"": 1, ""userStyleName"": ""<user style name>"", ""confidence"": <0-100>, ""preserveList"": <true|false>, ""restartNumbering"": <true|false>," & vbCrLf &
                "   ""deletePrefix"": { ""kind"": ""HeadingNumber|ListNumber|Bullet|None"", ""text"": ""<exact leading text to remove (incl. trailing space) or empty>"", ""charCount"": <0 if none> }" & vbCrLf &
                "  }," & vbCrLf &
                "  ..." & vbCrLf &
                "]"



#Region "DocStyle Settings Class"

    ''' <summary>
    ''' Holds persistent settings for DocStyle operations.
    ''' Persisted via <c>My.Settings</c> only (no registry usage).
    ''' </summary>
    Private Class DocStyleSettings
        Public Property TrackChanges As Boolean = False
        Public Property ApplyStyleDefinitions As Boolean = True
        Public Property PreviewMode As Boolean = False
        Public Property ConfidenceThreshold As Integer = 70
        Public Property UseConfidenceThreshold As Boolean = False
        Public Property ProcessTables As Boolean = True
        Public Property FastModeStylesOnly As Boolean = True
        Public Property ShowReport As Boolean = True
        Public Property DocumentContext As String = ""
        Public Property UseSecondaryModel As Boolean = False
        Public Property ListNumberingReset As Integer = 0 ' 0=Off, 1=Rule-based, 2=LLM-assisted
        Public Property RuleHeadingOutlineLevelMax As Integer = DefaultRuleHeadingOutlineLevelMax
        Public Property LastStyleTemplateDisplay As String = ""
        Public Property RestoreInlineFormatting As Boolean = False
        Public Property RemoveEmptyLines As Boolean = False

        ''' <summary>
        ''' Loads DocStyle settings from <c>My.Settings</c> and clamps numeric values to valid ranges.
        ''' </summary>
        Public Sub Load()
            ' Strict strongly-typed settings only (compile error if missing)
            TrackChanges = My.Settings.DocStyle_TrackChanges
            ApplyStyleDefinitions = My.Settings.DocStyle_ApplyStyleDefinitions
            PreviewMode = My.Settings.DocStyle_PreviewMode
            UseConfidenceThreshold = My.Settings.DocStyle_UseConfidenceThreshold
            ProcessTables = My.Settings.DocStyle_ProcessTables
            FastModeStylesOnly = My.Settings.DocStyle_FastModeStylesOnly
            ShowReport = My.Settings.DocStyle_ShowReport
            UseSecondaryModel = My.Settings.DocStyle_UseSecondaryModel

            ConfidenceThreshold = My.Settings.DocStyle_ConfidenceThreshold
            ListNumberingReset = My.Settings.DocStyle_ListNumberingReset
            DocumentContext = If(My.Settings.DocStyle_DocumentContext, "")
            RuleHeadingOutlineLevelMax = My.Settings.DocStyle_RuleHeadingOutlineLevelMax

            LastStyleTemplateDisplay = If(My.Settings.DocStyle_LastStyleTemplateDisplay, "")
            RestoreInlineFormatting = My.Settings.DocStyle_RestoreInlineFormatting
            RemoveEmptyLines = My.Settings.DocStyle_RemoveEmptyLines

            ' Clamps
            If ConfidenceThreshold < 0 Then ConfidenceThreshold = 0
            If ConfidenceThreshold > 100 Then ConfidenceThreshold = 100

            If ListNumberingReset < 0 Then ListNumberingReset = 0
            If ListNumberingReset > 2 Then ListNumberingReset = 2

            If RuleHeadingOutlineLevelMax < 0 Then RuleHeadingOutlineLevelMax = 0
            If RuleHeadingOutlineLevelMax > 9 Then RuleHeadingOutlineLevelMax = 9
        End Sub

        ''' <summary>
        ''' Saves DocStyle settings to <c>My.Settings</c>.
        ''' </summary>
        Public Sub Save()
            ' Strict strongly-typed settings only (compile error if missing)
            My.Settings.DocStyle_TrackChanges = TrackChanges
            My.Settings.DocStyle_ApplyStyleDefinitions = ApplyStyleDefinitions
            My.Settings.DocStyle_PreviewMode = PreviewMode
            My.Settings.DocStyle_ConfidenceThreshold = ConfidenceThreshold
            My.Settings.DocStyle_UseConfidenceThreshold = UseConfidenceThreshold
            My.Settings.DocStyle_ProcessTables = ProcessTables
            My.Settings.DocStyle_FastModeStylesOnly = FastModeStylesOnly
            My.Settings.DocStyle_ShowReport = ShowReport
            My.Settings.DocStyle_DocumentContext = If(DocumentContext, "")
            My.Settings.DocStyle_UseSecondaryModel = UseSecondaryModel
            My.Settings.DocStyle_ListNumberingReset = ListNumberingReset
            My.Settings.DocStyle_RuleHeadingOutlineLevelMax = RuleHeadingOutlineLevelMax

            My.Settings.DocStyle_LastStyleTemplateDisplay = If(LastStyleTemplateDisplay, "")
            My.Settings.DocStyle_RestoreInlineFormatting = RestoreInlineFormatting
            My.Settings.DocStyle_RemoveEmptyLines = RemoveEmptyLines

            My.Settings.Save()
        End Sub
    End Class


#End Region


#Region "Step 1: Extract Paragraph Styles to JSON (Style Template Creation) and wdStyles"

    ''' <summary>
    ''' Entry point for style template creation. Prompts the user to choose between
    ''' trainer-based extraction (manual trainer document) or automatic extraction
    ''' (AI-assisted discovery from a properly formatted document), then delegates accordingly.
    ''' </summary>
    Public Sub ExtractParagraphStylesToJson()
        Try
            Dim app As Word.Application = Globals.ThisAddIn.Application
            If app Is Nothing OrElse app.ActiveDocument Is Nothing Then
                ShowCustomMessageBox("No active document found.")
                Return
            End If

            Dim choice As Integer = ShowCustomYesNoBox(
                "How would you like to create the style template?" & vbCrLf & vbCrLf &
                "Trainer Document (manual):" & vbCrLf &
                "  • You prepare a document with paragraphs formatted as 'STYLE NAME: description'" & vbCrLf &
                "  • Full control over style names and descriptions" & vbCrLf &
                "  • Best when you want precise, hand-crafted style definitions" & vbCrLf & vbCrLf &
                "Automatic (AI-assisted):" & vbCrLf &
                "  • Scans any properly formatted document and discovers all used styles" & vbCrLf &
                "  • AI generates style descriptions automatically — no trainer needed" & vbCrLf &
                "  • Best for quickly creating templates from existing documents",
                "Trainer Document", "Automatic (AI)", $"{AN} - Learn Document Style")

            If choice = 0 Then
                ' User closed the dialog
                Return
            ElseIf choice = 1 Then
                ' Trainer-based extraction
                ExtractParagraphStylesToJsonFromTrainer()
            ElseIf choice = 2 Then
                ' Automatic AI-assisted extraction
                AutoExtractStyleTemplate()
            End If

        Catch ex As System.Exception
            ShowCustomMessageBox($"Error: {ex.Message}")
        End Try
    End Sub


    ''' <summary>
    ''' Extracts paragraph formatting from a trainer document (selection or entire document) and generates
    ''' a JSON structure describing each paragraph's text, style, and formatting.
    ''' Expects paragraphs formatted as "STYLE NAME: description..." for automatic parsing.
    ''' </summary>
    Private Sub ExtractParagraphStylesToJsonFromTrainer()
        Try
            Dim app As Word.Application = Globals.ThisAddIn.Application
            Dim doc As Word.Document = app.ActiveDocument

            If doc Is Nothing Then
                ShowCustomMessageBox("No active document found.")
                Return
            End If

            ' Expand paths
            Dim docStylePath As String = ExpandEnvironmentVariables(INI_DocStylePath)
            If Not String.IsNullOrEmpty(docStylePath) AndAlso Not docStylePath.EndsWith("\") Then
                docStylePath &= "\"
            End If

            Dim docStylePathLocal As String = ExpandEnvironmentVariables(INI_DocStylePathLocal)
            If Not String.IsNullOrEmpty(docStylePathLocal) AndAlso Not docStylePathLocal.EndsWith("\") Then
                docStylePathLocal &= "\"
            End If

            Dim hasGlobal As Boolean = Not String.IsNullOrWhiteSpace(docStylePath) AndAlso Directory.Exists(docStylePath)
            Dim hasLocal As Boolean = Not String.IsNullOrWhiteSpace(docStylePathLocal) AndAlso Directory.Exists(docStylePathLocal)

            ' Determine save path
            Dim savePath As String
            Dim isLocal As Boolean = False

            If Not hasGlobal AndAlso Not hasLocal Then
                ' Use desktop and warn
                savePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                ShowCustomMessageBox("Warning: Neither 'DocStylePath' nor 'DocStylePathLocal' is configured. The file will be saved to your Desktop.")
            ElseIf hasGlobal AndAlso hasLocal Then
                ' Ask user
                Dim choice As Integer = ShowCustomYesNoBox("Where do you want to save the style template?",
                    "Central (shared)", "Local (personal)", $"{AN} - Save Location")
                If choice = 0 Then Return
                If choice = 1 Then
                    savePath = docStylePath
                    isLocal = False
                Else
                    savePath = docStylePathLocal
                    isLocal = True
                End If
            ElseIf hasGlobal Then
                savePath = docStylePath
            Else
                savePath = docStylePathLocal
                isLocal = True
            End If

            ' Ask for template display name using ShowCustomInputBox
            Dim safeDocName As String = Regex.Replace(doc.Name.Replace(".docx", "").Replace(".doc", ""), "[^a-zA-Z0-9_-]", "_")
            Dim defaultDisplayName As String = $"{safeDocName}_{DateTime.Now:yyyyMMdd_HHmm}"

            Dim templateDisplayName As String = ShowCustomInputBox("Enter a name for this style template:", $"{AN} - Style Template Name", True, defaultDisplayName)
            If String.IsNullOrWhiteSpace(templateDisplayName) OrElse templateDisplayName.Equals("ESC", StringComparison.OrdinalIgnoreCase) Then
                Return
            End If

            Dim targetRange As Word.Range

            ' Use selection if available, otherwise entire document
            If app.Selection.Type = WdSelectionType.wdSelectionIP Then
                targetRange = doc.Content
            Else
                targetRange = app.Selection.Range.Duplicate
            End If

            Dim userStyles As New JArray()
            Dim wdStyleDefinitions As New JObject()
            Dim collectedStyles As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim openXmlContext As JObject = BuildDocStyleOpenXmlContext(doc)
            Dim processedCount As Integer = 0
            Dim errorCount As Integer = 0
            Dim unparsedStyleNameCount As Integer = 0

            ' Process all paragraphs including those in tables
            For Each para As Word.Paragraph In targetRange.Paragraphs
                Try
                    Dim parseResult As (userStyleJson As JObject, parsedStyleName As Boolean) = ExtractUserStyleFromParagraphEx(para, processedCount + 1, openXmlContext)
                    If parseResult.userStyleJson IsNot Nothing Then
                        userStyles.Add(parseResult.userStyleJson)
                        processedCount += 1

                        If Not parseResult.parsedStyleName Then
                            unparsedStyleNameCount += 1
                        End If

                        ' Collect wdStyle name for style definitions
                        Dim wdStyleName As String = If(parseResult.userStyleJson("wdStyleName") IsNot Nothing, parseResult.userStyleJson("wdStyleName").ToString(), "")
                        If Not String.IsNullOrWhiteSpace(wdStyleName) AndAlso Not collectedStyles.Contains(wdStyleName) Then
                            collectedStyles.Add(wdStyleName)
                        End If
                    End If
                Catch ex As System.Exception
                    errorCount += 1
                    Debug.WriteLine($"Error processing paragraph {processedCount + 1}: {ex.Message}")
                End Try
            Next

            ' Extract full wdStyle definitions for collected styles
            For Each styleName In collectedStyles
                Try
                    Dim styleObj As JObject = ExtractFullStyleDefinition(doc, styleName, openXmlContext)
                    If styleObj IsNot Nothing Then
                        wdStyleDefinitions(styleName) = styleObj
                    End If
                Catch ex As System.Exception
                    Debug.WriteLine($"Error extracting wdStyle definition for '{styleName}': {ex.Message}")
                End Try
            Next

            ' Build the complete JSON structure
            Dim result As New JObject()
            result("schemaVersion") = DocStyleSchemaVersion
            result("templateName") = templateDisplayName
            result("description") = "Style template for intelligent document formatting. Each user style in 'userStyles' includes a 'whenToApply' field describing the situations where that style should be applied. The 'wdStyleDefinitions' section contains Word style definitions that can optionally be created/updated in target documents."
            result("documentInfo") = New JObject From {
                {"extractedFrom", doc.Name},
                {"extractionDate", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")},
                {"totalUserStyles", processedCount},
                {"totalWdStyles", collectedStyles.Count}
            }
            result("userStyles") = userStyles
            result("wdStyleDefinitions") = wdStyleDefinitions
            result("numberingDefinitions") = BuildReferencedNumberingDefinitions(openXmlContext, collectedStyles)
            If openXmlContext("documentDefaults") IsNot Nothing Then
                result("documentDefaults") = openXmlContext("documentDefaults").DeepClone()
            End If

            ' Serialize to formatted JSON
            Dim jsonString As String = JsonConvert.SerializeObject(result, Formatting.Indented)

            ' Generate filename using safe template name
            Dim safeTemplateName As String = Regex.Replace(templateDisplayName, "[^a-zA-Z0-9_-]", "_")
            Dim fileName As String = $"{AN2}-ds-{safeTemplateName}.json"
            Dim filePath As String = Path.Combine(savePath, fileName)

            ' Check if file exists and handle overwrite
            If File.Exists(filePath) Then
                Dim overwrite As Integer = ShowCustomYesNoBox($"A style template with this name already exists.{vbCrLf}{vbCrLf}Do you want to overwrite it?",
                    "Overwrite", "Cancel", $"{AN} - File Exists")
                If overwrite <> 1 Then Return
            End If

            Try
                File.WriteAllText(filePath, jsonString, Encoding.UTF8)
            Catch ioEx As System.Exception
                ShowCustomMessageBox($"Could not save file: {ioEx.Message}")
                Return
            End Try

            ' Build success message - only show tip if some styles were not parsed
            Dim successMessage As String = $"Style template '{templateDisplayName}' has been created successfully.{vbCrLf}{vbCrLf}Location: {filePath}{vbCrLf}{vbCrLf}Would you like to edit the template now?"

            If unparsedStyleNameCount > 0 Then
                successMessage &= $"{vbCrLf}{vbCrLf}Tip: {unparsedStyleNameCount} style(s) could not be auto-named. Edit the 'userStyleName' and 'whenToApply' fields for these styles. Use format 'STYLE NAME: description...' in your source paragraphs for automatic parsing."
            End If

            Dim editChoice As Integer = ShowCustomYesNoBox(successMessage,
                "Edit Template", "Close", $"{AN} - Style Template Created")

            If editChoice = 1 Then
                SLib.ShowTextFileEditor(filePath, $"{AN} - Style Template '{templateDisplayName}'", True, _context)
            End If

        Catch ex As System.Exception
            ShowCustomMessageBox($"Error extracting paragraph styles: {ex.Message}")
        End Try
    End Sub


    ''' <summary>
    ''' Extracts user style information from a single paragraph for template creation.
    ''' Parses paragraph text to extract <c>userStyleName</c> and <c>whenToApply</c> and returns the extracted JSON.
    ''' </summary>
    ''' <param name="para">Word paragraph to read.</param>
    ''' <param name="index">1-based index assigned to the extracted user style.</param>
    ''' <returns>
    ''' A tuple containing the extracted JSON object and a flag indicating whether the style name was parsed from text.
    ''' </returns>
    Private Function ExtractUserStyleFromParagraphEx(para As Word.Paragraph,
                                                     index As Integer,
                                                     Optional openXmlContext As JObject = Nothing) As (userStyleJson As JObject, parsedStyleName As Boolean)
        Dim paraRange As Word.Range = para.Range.Duplicate

        ' Skip empty paragraphs (only paragraph mark)
        Dim text As String = paraRange.Text
        If String.IsNullOrWhiteSpace(text) OrElse text = vbCr OrElse text = vbCrLf Then
            Return (Nothing, False)
        End If

        ' Remove trailing paragraph mark for cleaner text
        text = text.TrimEnd(vbCr, vbLf, ChrW(13), ChrW(10)).Trim()

        Dim result As New JObject()
        result("userStyleIndex") = index

        ' Try to parse "STYLE NAME: description..." format
        Dim parsedStyleName As Boolean = False
        Dim userStyleName As String = $"UserStyle_{index}"
        Dim whenToApply As String = text

        ' Look for colon separator - style name should be short (max ~50 chars) and before colon
        Dim colonIndex As Integer = text.IndexOf(":"c)
        If colonIndex > 0 AndAlso colonIndex <= 50 Then
            Dim potentialName As String = text.Substring(0, colonIndex).Trim()
            Dim potentialDescription As String = text.Substring(colonIndex + 1).Trim()

            ' Validate: name should not contain line breaks and should have some description after
            If Not String.IsNullOrWhiteSpace(potentialName) AndAlso
               Not potentialName.Contains(vbCr) AndAlso
               Not potentialName.Contains(vbLf) AndAlso
               Not String.IsNullOrWhiteSpace(potentialDescription) Then
                userStyleName = potentialName
                whenToApply = potentialDescription
                parsedStyleName = True
            End If
        End If

        result("userStyleName") = userStyleName
        result("whenToApply") = whenToApply

        ' Get the Word style name
        Try
            Dim style As Word.Style = para.Style
            result("wdStyleName") = style.NameLocal
            result("wdStyleBuiltIn") = style.BuiltIn
        Catch
            result("wdStyleName") = "Normal"
            result("wdStyleBuiltIn") = True
        End Try

        Dim styleMeta As JObject = GetStyleOpenXmlMetadata(openXmlContext, If(result("wdStyleName") IsNot Nothing, CStr(result("wdStyleName")), ""))
        If styleMeta IsNot Nothing Then
            If styleMeta("styleId") IsNot Nothing Then result("wdStyleId") = styleMeta("styleId").DeepClone()
            If styleMeta("numberingBinding") IsNot Nothing Then result("numberingBinding") = styleMeta("numberingBinding").DeepClone()
        End If

        ' Check if in table
        Dim isInTable As Boolean = False
        Try
            If paraRange.Tables.Count > 0 OrElse paraRange.Cells.Count > 0 Then
                isInTable = True
            End If
        Catch
        End Try
        result("isInTableCell") = isInTable

        ' Paragraph formatting
        result("paragraphFormatting") = ExtractParagraphFormat(para, paraRange)

        ' Font/Character formatting
        result("fontFormatting") = ExtractFontFormat(paraRange)

        ' List formatting
        result("listFormatting") = ExtractListFormat(paraRange, openXmlContext)

        ' Tab stops (compact format)
        result("tabStops") = ExtractTabStopsCompact(para)

        ' Borders (only if present)
        Dim borders As JObject = ExtractBorders(para)
        If borders.Properties().Any(Function(p) p.Name <> "distanceFromText" AndAlso p.Name <> "error") Then
            result("borders") = borders
        End If

        ' Shading (only if non-default)
        Dim shading As JObject = ExtractShading(para)
        If shading("backgroundColor") IsNot Nothing OrElse shading("foregroundColor") IsNot Nothing Then
            result("shading") = shading
        End If

        Return (result, parsedStyleName)
    End Function

    ''' <summary>
    ''' Extracts tab stops from a paragraph into a compact JSON array.
    ''' </summary>
    ''' <param name="para">Paragraph containing tab stops.</param>
    ''' <returns>Array of tab stop definitions (position, alignment, optional leader).</returns>
    ''' <summary>
    ''' Extracts only explicitly defined paragraph tab stops. Generated default tab stops are represented
    ''' once in documentDefaults.defaultTabStop and are not repeated here.
    ''' </summary>
    Private Function ExtractTabStopsCompact(para As Word.Paragraph) As JArray
        Dim tabs As New JArray()
        If para Is Nothing Then Return tabs

        Try
            For Each tabStop As Word.TabStop In para.TabStops
                Dim isCustom As Boolean = True
                Try
                    isCustom = tabStop.CustomTab
                Catch
                End Try
                If Not isCustom Then Continue For

                Dim tab As New JObject From {
                    {"pos", Math.Round(tabStop.Position, 2)},
                    {"align", tabStop.Alignment.ToString().Replace("wdAlignTab", "")}
                }
                If tabStop.Leader <> WdTabLeader.wdTabLeaderSpaces Then
                    tab("leader") = tabStop.Leader.ToString().Replace("wdTabLeader", "")
                End If
                tabs.Add(tab)
            Next
        Catch
        End Try

        Return tabs
    End Function

#Region "DocStyle Open XML metadata"

    Private Const DocStyleSchemaVersion As Integer = 2

    ''' <summary>
    ''' Builds an extraction-only metadata index from Word's Flat OPC representation.
    ''' No Open XML SDK dependency is required; only System.Xml.Linq is used.
    ''' </summary>
    Private Function BuildDocStyleOpenXmlContext(doc As Word.Document) As JObject
        Dim context As New JObject From {
            {"available", False},
            {"stylesById", New JObject()},
            {"styleIdsByName", New JObject()},
            {"numberingDefinitions", New JObject()},
            {"documentDefaults", New JObject()}
        }

        If doc Is Nothing Then Return context

        Try
            Dim flatOpcXml As String = doc.WordOpenXML
            If String.IsNullOrWhiteSpace(flatOpcXml) Then Return context

            Dim flatOpc As System.Xml.Linq.XDocument =
                System.Xml.Linq.XDocument.Parse(flatOpcXml, System.Xml.Linq.LoadOptions.PreserveWhitespace)

            Dim stylesRoot As System.Xml.Linq.XElement = GetFlatOpcPartRoot(flatOpc, "/word/styles.xml")
            Dim numberingRoot As System.Xml.Linq.XElement = GetFlatOpcPartRoot(flatOpc, "/word/numbering.xml")
            Dim settingsRoot As System.Xml.Linq.XElement = GetFlatOpcPartRoot(flatOpc, "/word/settings.xml")

            If stylesRoot Is Nothing Then Return context

            Dim w As System.Xml.Linq.XNamespace =
                System.Xml.Linq.XNamespace.Get("http://schemas.openxmlformats.org/wordprocessingml/2006/main")
            Dim w15 As System.Xml.Linq.XNamespace =
                System.Xml.Linq.XNamespace.Get("http://schemas.microsoft.com/office/word/2012/wordml")

            Dim stylesById As New JObject()
            Dim styleIdsByName As New JObject()

            For Each styleElement As System.Xml.Linq.XElement In stylesRoot.Elements(w + "style")
                Dim styleId As String = GetOpenXmlAttributeValue(styleElement, w + "styleId")
                If String.IsNullOrWhiteSpace(styleId) Then Continue For

                Dim styleName As String = GetOpenXmlVal(styleElement.Element(w + "name"))
                If String.IsNullOrWhiteSpace(styleName) Then styleName = styleId

                Dim styleMeta As New JObject From {
                    {"styleId", styleId},
                    {"styleName", styleName},
                    {"styleTypeOpenXml", GetOpenXmlAttributeValue(styleElement, w + "type")},
                    {"customStyle", OpenXmlOnOffAttribute(styleElement, w + "customStyle", False)}
                }

                Dim basedOnStyleId As String = GetOpenXmlVal(styleElement.Element(w + "basedOn"))
                If Not String.IsNullOrWhiteSpace(basedOnStyleId) Then styleMeta("basedOnStyleId") = basedOnStyleId

                Dim nextStyleId As String = GetOpenXmlVal(styleElement.Element(w + "next"))
                If Not String.IsNullOrWhiteSpace(nextStyleId) Then styleMeta("nextStyleId") = nextStyleId

                Dim uiPriorityValue As String = GetOpenXmlVal(styleElement.Element(w + "uiPriority"))
                Dim uiPriority As Integer
                If Integer.TryParse(uiPriorityValue, uiPriority) Then styleMeta("uiPriority") = uiPriority

                styleMeta("quickFormat") = (styleElement.Element(w + "qFormat") IsNot Nothing)

                Dim directNumbering As New JObject()
                Dim pPr As System.Xml.Linq.XElement = styleElement.Element(w + "pPr")
                Dim numPr As System.Xml.Linq.XElement = If(pPr IsNot Nothing, pPr.Element(w + "numPr"), Nothing)
                If numPr IsNot Nothing Then
                    Dim ilvlValue As String = GetOpenXmlVal(numPr.Element(w + "ilvl"))
                    Dim ilvlZeroBased As Integer
                    If Integer.TryParse(ilvlValue, ilvlZeroBased) Then
                        directNumbering("ilvlZeroBased") = ilvlZeroBased
                        directNumbering("listLevelNumber") = ilvlZeroBased + 1
                    End If

                    Dim numIdValue As String = GetOpenXmlVal(numPr.Element(w + "numId"))
                    Dim numId As Integer
                    If Integer.TryParse(numIdValue, numId) Then directNumbering("numId") = numId
                End If
                If directNumbering.HasValues Then styleMeta("directNumbering") = directNumbering

                stylesById(styleId) = styleMeta
                styleIdsByName(styleName) = styleId
            Next

            ' Resolve display names after every style ID has been indexed.
            For Each styleProperty As JProperty In stylesById.Properties()
                Dim styleMeta As JObject = CType(styleProperty.Value, JObject)

                If styleMeta("basedOnStyleId") IsNot Nothing Then
                    Dim baseName As String = GetStyleNameFromContext(stylesById, CStr(styleMeta("basedOnStyleId")))
                    If Not String.IsNullOrWhiteSpace(baseName) Then styleMeta("basedOnStyleName") = baseName
                End If

                If styleMeta("nextStyleId") IsNot Nothing Then
                    Dim nextName As String = GetStyleNameFromContext(stylesById, CStr(styleMeta("nextStyleId")))
                    If Not String.IsNullOrWhiteSpace(nextName) Then styleMeta("nextStyleName") = nextName
                End If
            Next

            Dim abstractNumberingById As New JObject()
            If numberingRoot IsNot Nothing Then
                For Each abstractElement As System.Xml.Linq.XElement In numberingRoot.Elements(w + "abstractNum")
                    Dim abstractNumIdText As String = GetOpenXmlAttributeValue(abstractElement, w + "abstractNumId")
                    Dim abstractNumId As Integer
                    If Not Integer.TryParse(abstractNumIdText, abstractNumId) Then Continue For

                    Dim definition As New JObject From {
                        {"abstractNumId", abstractNumId},
                        {"outlineNumbered", True}
                    }

                    Dim nsid As String = GetOpenXmlVal(abstractElement.Element(w + "nsid"))
                    If Not String.IsNullOrWhiteSpace(nsid) Then definition("nsid") = nsid

                    Dim templateCode As String = GetOpenXmlVal(abstractElement.Element(w + "tmpl"))
                    If Not String.IsNullOrWhiteSpace(templateCode) Then definition("templateCode") = templateCode

                    Dim multiLevelType As String = GetOpenXmlVal(abstractElement.Element(w + "multiLevelType"))
                    If Not String.IsNullOrWhiteSpace(multiLevelType) Then
                        definition("multiLevelType") = multiLevelType
                        definition("outlineNumbered") = Not String.Equals(multiLevelType, "singleLevel", StringComparison.OrdinalIgnoreCase)
                    End If

                    Dim styleLink As String = GetOpenXmlVal(abstractElement.Element(w + "styleLink"))
                    If Not String.IsNullOrWhiteSpace(styleLink) Then definition("numberingStyleId") = styleLink

                    Dim numStyleLink As String = GetOpenXmlVal(abstractElement.Element(w + "numStyleLink"))
                    If Not String.IsNullOrWhiteSpace(numStyleLink) Then definition("numStyleLinkId") = numStyleLink

                    Dim restartAfterBreakAttribute As System.Xml.Linq.XAttribute = abstractElement.Attribute(w15 + "restartNumberingAfterBreak")
                    If restartAfterBreakAttribute IsNot Nothing Then
                        definition("restartNumberingAfterBreak") = ParseOpenXmlOnOffValue(restartAfterBreakAttribute.Value, False)
                    End If

                    Dim levels As New JArray()
                    For Each levelElement As System.Xml.Linq.XElement In abstractElement.Elements(w + "lvl")
                        levels.Add(ExtractOpenXmlNumberingLevel(levelElement, stylesById, w))
                    Next
                    definition("levels") = levels
                    definition("templateFingerprintV2") = ComputeJsonListTemplateFingerprintV2(definition)

                    abstractNumberingById(abstractNumId.ToString(System.Globalization.CultureInfo.InvariantCulture)) = definition
                Next
            End If

            Dim numberingDefinitions As New JObject()
            If numberingRoot IsNot Nothing Then
                For Each numElement As System.Xml.Linq.XElement In numberingRoot.Elements(w + "num")
                    Dim numIdText As String = GetOpenXmlAttributeValue(numElement, w + "numId")
                    Dim numId As Integer
                    If Not Integer.TryParse(numIdText, numId) Then Continue For

                    Dim abstractNumIdText As String = GetOpenXmlVal(numElement.Element(w + "abstractNumId"))
                    Dim abstractNumId As Integer
                    If Not Integer.TryParse(abstractNumIdText, abstractNumId) Then Continue For

                    Dim definition As JObject = Nothing
                    Dim abstractToken As JToken = abstractNumberingById(abstractNumId.ToString(System.Globalization.CultureInfo.InvariantCulture))
                    If abstractToken IsNot Nothing AndAlso TypeOf abstractToken Is JObject Then
                        definition = CType(abstractToken.DeepClone(), JObject)
                    Else
                        definition = New JObject From {
                            {"abstractNumId", abstractNumId},
                            {"levels", New JArray()},
                            {"outlineNumbered", True}
                        }
                    End If

                    definition("sourceNumId") = numId
                    definition("sharedListInstanceKey") = $"numId:{numId}"

                    ApplyOpenXmlNumberingOverrides(definition, numElement, stylesById, w)

                    If definition("numberingStyleId") IsNot Nothing Then
                        Dim numberingStyleName As String = GetStyleNameFromContext(stylesById, CStr(definition("numberingStyleId")))
                        If Not String.IsNullOrWhiteSpace(numberingStyleName) Then definition("numberingStyleName") = numberingStyleName
                    End If

                    definition("templateFingerprintV2") = ComputeJsonListTemplateFingerprintV2(definition)
                    numberingDefinitions(CStr(definition("sharedListInstanceKey"))) = definition
                Next
            End If

            ' Resolve effective numId and list level independently through basedOn chains.
            For Each styleProperty As JProperty In stylesById.Properties()
                Dim styleId As String = styleProperty.Name
                Dim styleMeta As JObject = CType(styleProperty.Value, JObject)
                Dim levelResolution As JObject = ResolveStyleNumberingProperty(stylesById, styleId, "listLevelNumber")
                Dim numIdResolution As JObject = ResolveStyleNumberingProperty(stylesById, styleId, "numId")

                If levelResolution IsNot Nothing OrElse numIdResolution IsNot Nothing Then
                    Dim binding As New JObject()
                    Dim directNumbering As JObject = TryCast(styleMeta("directNumbering"), JObject)

                    binding("hasDirectListLevel") =
                        (directNumbering IsNot Nothing AndAlso directNumbering("listLevelNumber") IsNot Nothing)
                    binding("hasDirectNumId") =
                        (directNumbering IsNot Nothing AndAlso directNumbering("numId") IsNot Nothing)
                    binding("direct") = CBool(binding("hasDirectListLevel")) AndAlso CBool(binding("hasDirectNumId"))
                    binding("inherited") = Not CBool(binding("direct"))

                    If styleMeta("basedOnStyleId") IsNot Nothing Then
                        binding("inheritedViaStyleId") = styleMeta("basedOnStyleId")
                        If styleMeta("basedOnStyleName") IsNot Nothing Then
                            binding("inheritedViaStyleName") = styleMeta("basedOnStyleName")
                        End If
                    End If

                    If levelResolution IsNot Nothing Then
                        binding("effectiveListLevel") = levelResolution("value")
                        binding("listLevelSourceStyleId") = levelResolution("sourceStyleId")
                        binding("listLevelSourceStyleName") = levelResolution("sourceStyleName")
                    End If

                    If numIdResolution IsNot Nothing Then
                        Dim effectiveNumId As Integer = CInt(numIdResolution("value"))
                        binding("numId") = effectiveNumId
                        binding("numIdSourceStyleId") = numIdResolution("sourceStyleId")
                        binding("numIdSourceStyleName") = numIdResolution("sourceStyleName")
                        binding("sharedListInstanceKey") = $"numId:{effectiveNumId}"

                        Dim numberingDefinition As JObject = TryCast(numberingDefinitions($"numId:{effectiveNumId}"), JObject)
                        If numberingDefinition IsNot Nothing Then
                            If numberingDefinition("abstractNumId") IsNot Nothing Then binding("abstractNumId") = numberingDefinition("abstractNumId")
                            If numberingDefinition("numberingStyleId") IsNot Nothing Then binding("numberingStyleId") = numberingDefinition("numberingStyleId")
                            If numberingDefinition("numberingStyleName") IsNot Nothing Then binding("numberingStyleName") = numberingDefinition("numberingStyleName")
                        End If
                    End If

                    styleMeta("numberingBinding") = binding
                End If
            Next

            context("stylesById") = stylesById
            context("styleIdsByName") = styleIdsByName
            context("numberingDefinitions") = numberingDefinitions
            context("documentDefaults") = ExtractOpenXmlDocumentDefaults(stylesRoot, settingsRoot, w)
            context("available") = True

        Catch ex As System.Exception
            context("error") = ex.Message
            Debug.WriteLine($"[DocStyle] Open XML metadata extraction failed: {ex.Message}")
        End Try

        Return context
    End Function

    Private Function GetFlatOpcPartRoot(flatOpc As System.Xml.Linq.XDocument,
                                        partName As String) As System.Xml.Linq.XElement
        If flatOpc Is Nothing OrElse String.IsNullOrWhiteSpace(partName) Then Return Nothing

        Dim pkg As System.Xml.Linq.XNamespace =
            System.Xml.Linq.XNamespace.Get("http://schemas.microsoft.com/office/2006/xmlPackage")

        For Each part As System.Xml.Linq.XElement In flatOpc.Descendants(pkg + "part")
            Dim nameAttribute As System.Xml.Linq.XAttribute = part.Attribute(pkg + "name")
            If nameAttribute Is Nothing OrElse
               Not String.Equals(nameAttribute.Value, partName, StringComparison.OrdinalIgnoreCase) Then
                Continue For
            End If

            Dim xmlData As System.Xml.Linq.XElement = part.Element(pkg + "xmlData")
            If xmlData IsNot Nothing Then Return xmlData.Elements().FirstOrDefault()
        Next

        Return Nothing
    End Function

    Private Function GetOpenXmlAttributeValue(element As System.Xml.Linq.XElement,
                                              attributeName As System.Xml.Linq.XName) As String
        If element Is Nothing Then Return ""
        Dim attribute As System.Xml.Linq.XAttribute = element.Attribute(attributeName)
        Return If(attribute IsNot Nothing, attribute.Value, "")
    End Function

    Private Function GetOpenXmlVal(element As System.Xml.Linq.XElement) As String
        If element Is Nothing Then Return ""
        Dim w As System.Xml.Linq.XNamespace =
            System.Xml.Linq.XNamespace.Get("http://schemas.openxmlformats.org/wordprocessingml/2006/main")
        Return GetOpenXmlAttributeValue(element, w + "val")
    End Function

    Private Function ParseOpenXmlOnOffValue(value As String, defaultValue As Boolean) As Boolean
        If String.IsNullOrWhiteSpace(value) Then Return defaultValue
        Select Case value.Trim().ToLowerInvariant()
            Case "1", "true", "on", "yes"
                Return True
            Case "0", "false", "off", "no"
                Return False
            Case Else
                Return defaultValue
        End Select
    End Function

    Private Function OpenXmlOnOffAttribute(element As System.Xml.Linq.XElement,
                                           attributeName As System.Xml.Linq.XName,
                                           defaultValue As Boolean) As Boolean
        If element Is Nothing Then Return defaultValue
        Dim attribute As System.Xml.Linq.XAttribute = element.Attribute(attributeName)
        If attribute Is Nothing Then Return defaultValue
        Return ParseOpenXmlOnOffValue(attribute.Value, True)
    End Function

    Private Function GetStyleNameFromContext(stylesById As JObject, styleId As String) As String
        If stylesById Is Nothing OrElse String.IsNullOrWhiteSpace(styleId) Then Return ""
        Dim styleMeta As JObject = TryCast(stylesById(styleId), JObject)
        If styleMeta Is Nothing OrElse styleMeta("styleName") Is Nothing Then Return ""
        Return CStr(styleMeta("styleName"))
    End Function

    Private Function GetStyleOpenXmlMetadata(openXmlContext As JObject, styleName As String) As JObject
        If openXmlContext Is Nothing OrElse String.IsNullOrWhiteSpace(styleName) Then Return Nothing
        Dim styleIdsByName As JObject = TryCast(openXmlContext("styleIdsByName"), JObject)
        Dim stylesById As JObject = TryCast(openXmlContext("stylesById"), JObject)
        If styleIdsByName Is Nothing OrElse stylesById Is Nothing Then Return Nothing

        Dim styleIdToken As JToken = styleIdsByName(styleName)
        If styleIdToken Is Nothing Then
            For Each propertyItem As JProperty In styleIdsByName.Properties()
                If String.Equals(propertyItem.Name, styleName, StringComparison.OrdinalIgnoreCase) Then
                    styleIdToken = propertyItem.Value
                    Exit For
                End If
            Next
        End If

        If styleIdToken Is Nothing Then Return Nothing
        Return TryCast(stylesById(CStr(styleIdToken)), JObject)
    End Function

    Private Function ResolveStyleNumberingProperty(stylesById As JObject,
                                                    startStyleId As String,
                                                    propertyName As String) As JObject
        If stylesById Is Nothing OrElse String.IsNullOrWhiteSpace(startStyleId) Then Return Nothing

        Dim visited As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim currentStyleId As String = startStyleId

        While Not String.IsNullOrWhiteSpace(currentStyleId) AndAlso visited.Add(currentStyleId)
            Dim styleMeta As JObject = TryCast(stylesById(currentStyleId), JObject)
            If styleMeta Is Nothing Then Exit While

            Dim directNumbering As JObject = TryCast(styleMeta("directNumbering"), JObject)
            If directNumbering IsNot Nothing AndAlso directNumbering(propertyName) IsNot Nothing Then
                Return New JObject From {
                    {"value", directNumbering(propertyName).DeepClone()},
                    {"sourceStyleId", currentStyleId},
                    {"sourceStyleName", If(styleMeta("styleName") IsNot Nothing, CStr(styleMeta("styleName")), currentStyleId)}
                }
            End If

            currentStyleId = If(styleMeta("basedOnStyleId") IsNot Nothing,
                                CStr(styleMeta("basedOnStyleId")), "")
        End While

        Return Nothing
    End Function

    Private Function ExtractOpenXmlNumberingLevel(levelElement As System.Xml.Linq.XElement,
                                                  stylesById As JObject,
                                                  w As System.Xml.Linq.XNamespace) As JObject
        Dim levelInfo As New JObject()

        Dim ilvlText As String = GetOpenXmlAttributeValue(levelElement, w + "ilvl")
        Dim ilvl As Integer
        If Not Integer.TryParse(ilvlText, ilvl) Then ilvl = 0

        levelInfo("level") = ilvl + 1
        levelInfo("ilvlZeroBased") = ilvl

        Dim startAt As Integer
        If Integer.TryParse(GetOpenXmlVal(levelElement.Element(w + "start")), startAt) Then
            levelInfo("startAt") = startAt
        Else
            levelInfo("startAt") = 1
        End If

        Dim numFmtOpenXml As String = GetOpenXmlVal(levelElement.Element(w + "numFmt"))
        levelInfo("numberStyleOpenXml") = numFmtOpenXml
        levelInfo("numberStyle") = OpenXmlNumberFormatToWordEnumName(numFmtOpenXml)
        levelInfo("numberFormat") = GetOpenXmlVal(levelElement.Element(w + "lvlText"))

        Dim alignmentOpenXml As String = GetOpenXmlVal(levelElement.Element(w + "lvlJc"))
        levelInfo("alignmentOpenXml") = alignmentOpenXml
        levelInfo("alignment") = OpenXmlListAlignmentToWordEnumName(alignmentOpenXml)

        Dim suffixOpenXml As String = GetOpenXmlVal(levelElement.Element(w + "suff"))
        If String.IsNullOrWhiteSpace(suffixOpenXml) Then suffixOpenXml = "tab"
        levelInfo("trailingCharacterOpenXml") = suffixOpenXml
        levelInfo("trailingCharacter") = OpenXmlSuffixToWordEnumName(suffixOpenXml)

        Dim paragraphStyleId As String = GetOpenXmlVal(levelElement.Element(w + "pStyle"))
        If Not String.IsNullOrWhiteSpace(paragraphStyleId) Then
            levelInfo("paragraphStyleId") = paragraphStyleId
            Dim paragraphStyleName As String = GetStyleNameFromContext(stylesById, paragraphStyleId)
            If Not String.IsNullOrWhiteSpace(paragraphStyleName) Then levelInfo("paragraphStyleName") = paragraphStyleName
        End If

        Dim pPr As System.Xml.Linq.XElement = levelElement.Element(w + "pPr")
        Dim ind As System.Xml.Linq.XElement = If(pPr IsNot Nothing, pPr.Element(w + "ind"), Nothing)
        If ind IsNot Nothing Then
            Dim leftTwips As Integer = ParseOpenXmlIntegerAttribute(ind, w + "left", 0)
            Dim hangingTwips As Integer = ParseOpenXmlIntegerAttribute(ind, w + "hanging", 0)
            Dim firstLineTwips As Integer = ParseOpenXmlIntegerAttribute(ind, w + "firstLine", 0)

            levelInfo("textPosition") = CSng(leftTwips / 20.0R)
            If hangingTwips <> 0 Then
                levelInfo("numberPosition") = CSng((leftTwips - hangingTwips) / 20.0R)
            ElseIf firstLineTwips <> 0 Then
                levelInfo("numberPosition") = CSng((leftTwips + firstLineTwips) / 20.0R)
            Else
                levelInfo("numberPosition") = CSng(leftTwips / 20.0R)
            End If
        End If

        Dim tabPositionFound As Boolean = False
        If pPr IsNot Nothing Then
            Dim tabs As System.Xml.Linq.XElement = pPr.Element(w + "tabs")
            If tabs IsNot Nothing Then
                For Each tabElement As System.Xml.Linq.XElement In tabs.Elements(w + "tab")
                    Dim tabValue As String = GetOpenXmlAttributeValue(tabElement, w + "val")
                    If String.Equals(tabValue, "num", StringComparison.OrdinalIgnoreCase) OrElse
                       String.Equals(tabValue, "left", StringComparison.OrdinalIgnoreCase) Then
                        Dim tabTwips As Integer = ParseOpenXmlIntegerAttribute(tabElement, w + "pos", Integer.MinValue)
                        If tabTwips <> Integer.MinValue Then
                            levelInfo("tabPosition") = CSng(tabTwips / 20.0R)
                            levelInfo("tabPositionState") = "defined"
                            tabPositionFound = True
                            Exit For
                        End If
                    End If
                Next
            End If
        End If
        If Not tabPositionFound Then
            levelInfo("tabPosition") = JValue.CreateNull()
            levelInfo("tabPositionState") = "wdUndefined"
        End If

        Dim rPr As System.Xml.Linq.XElement = levelElement.Element(w + "rPr")
        Dim numberRunFormatting As JObject = ExtractOpenXmlRunFormatting(rPr, w)
        If numberRunFormatting.HasValues Then levelInfo("numberRunFormatting") = numberRunFormatting

        If String.Equals(numFmtOpenXml, "bullet", StringComparison.OrdinalIgnoreCase) Then
            Dim bulletText As String = If(levelInfo("numberFormat") IsNot Nothing, CStr(levelInfo("numberFormat")), "")
            If Not String.IsNullOrEmpty(bulletText) Then levelInfo("bulletCharCode") = AscW(bulletText.Chars(0))

            Dim bulletFont As String = ""
            Dim fontFields As String() = {"fontAscii", "fontHighAnsi", "fontEastAsia", "fontComplexScript"}
            For Each fontField As String In fontFields
                If numberRunFormatting(fontField) IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(CStr(numberRunFormatting(fontField))) Then
                    bulletFont = CStr(numberRunFormatting(fontField))
                    Exit For
                End If
            Next
            If Not String.IsNullOrWhiteSpace(bulletFont) Then levelInfo("bulletFont") = bulletFont
        End If

        Return levelInfo
    End Function

    Private Function ParseOpenXmlIntegerAttribute(element As System.Xml.Linq.XElement,
                                                  attributeName As System.Xml.Linq.XName,
                                                  defaultValue As Integer) As Integer
        If element Is Nothing Then Return defaultValue
        Dim valueText As String = GetOpenXmlAttributeValue(element, attributeName)
        Dim value As Integer
        If Integer.TryParse(valueText, value) Then Return value
        Return defaultValue
    End Function

    Private Function ExtractOpenXmlRunFormatting(rPr As System.Xml.Linq.XElement,
                                                 w As System.Xml.Linq.XNamespace) As JObject
        Dim result As New JObject()
        If rPr Is Nothing Then Return result

        Dim rFonts As System.Xml.Linq.XElement = rPr.Element(w + "rFonts")
        If rFonts IsNot Nothing Then
            AddOpenXmlAttributeIfPresent(result, "fontAscii", rFonts, w + "ascii")
            AddOpenXmlAttributeIfPresent(result, "fontHighAnsi", rFonts, w + "hAnsi")
            AddOpenXmlAttributeIfPresent(result, "fontEastAsia", rFonts, w + "eastAsia")
            AddOpenXmlAttributeIfPresent(result, "fontComplexScript", rFonts, w + "cs")
            AddOpenXmlAttributeIfPresent(result, "asciiTheme", rFonts, w + "asciiTheme")
            AddOpenXmlAttributeIfPresent(result, "highAnsiTheme", rFonts, w + "hAnsiTheme")
            AddOpenXmlAttributeIfPresent(result, "eastAsiaTheme", rFonts, w + "eastAsiaTheme")
            AddOpenXmlAttributeIfPresent(result, "complexScriptTheme", rFonts, w + "cstheme")
            AddOpenXmlAttributeIfPresent(result, "fontHint", rFonts, w + "hint")
        End If

        If rPr.Element(w + "b") IsNot Nothing Then result("bold") = GetOpenXmlOnOffElementValue(rPr.Element(w + "b"), w, True)
        If rPr.Element(w + "i") IsNot Nothing Then result("italic") = GetOpenXmlOnOffElementValue(rPr.Element(w + "i"), w, True)
        If rPr.Element(w + "caps") IsNot Nothing Then result("allCaps") = GetOpenXmlOnOffElementValue(rPr.Element(w + "caps"), w, True)
        If rPr.Element(w + "smallCaps") IsNot Nothing Then result("smallCaps") = GetOpenXmlOnOffElementValue(rPr.Element(w + "smallCaps"), w, True)

        Dim colorElement As System.Xml.Linq.XElement = rPr.Element(w + "color")
        If colorElement IsNot Nothing Then AddOpenXmlAttributeIfPresent(result, "colorOpenXml", colorElement, w + "val")

        Dim sizeHalfPoints As Integer
        If Integer.TryParse(GetOpenXmlVal(rPr.Element(w + "sz")), sizeHalfPoints) Then
            result("size") = CSng(sizeHalfPoints / 2.0R)
        End If

        Return result
    End Function

    Private Function GetOpenXmlOnOffElementValue(element As System.Xml.Linq.XElement,
                                                 w As System.Xml.Linq.XNamespace,
                                                 defaultWhenPresent As Boolean) As Boolean
        If element Is Nothing Then Return False
        Dim value As String = GetOpenXmlAttributeValue(element, w + "val")
        If String.IsNullOrWhiteSpace(value) Then Return defaultWhenPresent
        Return ParseOpenXmlOnOffValue(value, defaultWhenPresent)
    End Function

    Private Sub AddOpenXmlAttributeIfPresent(target As JObject,
                                             propertyName As String,
                                             element As System.Xml.Linq.XElement,
                                             attributeName As System.Xml.Linq.XName)
        Dim value As String = GetOpenXmlAttributeValue(element, attributeName)
        If Not String.IsNullOrWhiteSpace(value) Then target(propertyName) = value
    End Sub

    Private Sub ApplyOpenXmlNumberingOverrides(definition As JObject,
                                               numElement As System.Xml.Linq.XElement,
                                               stylesById As JObject,
                                               w As System.Xml.Linq.XNamespace)
        If definition Is Nothing OrElse numElement Is Nothing Then Return
        Dim levels As JArray = TryCast(definition("levels"), JArray)
        If levels Is Nothing Then
            levels = New JArray()
            definition("levels") = levels
        End If

        Dim levelOverrides As New Newtonsoft.Json.Linq.JArray()
        For Each overrideElement As System.Xml.Linq.XElement In numElement.Elements(w + "lvlOverride")
            Dim ilvlText As String = GetOpenXmlAttributeValue(overrideElement, w + "ilvl")
            Dim ilvl As Integer
            If Not Integer.TryParse(ilvlText, ilvl) Then Continue For

            Dim overrideInfo As New JObject From {
                {"level", ilvl + 1},
                {"ilvlZeroBased", ilvl}
            }

            Dim startOverride As Integer
            If Integer.TryParse(GetOpenXmlVal(overrideElement.Element(w + "startOverride")), startOverride) Then
                overrideInfo("startOverride") = startOverride
                Dim existingLevel As JObject = FindLevelDefinition(levels, ilvl + 1)
                If existingLevel IsNot Nothing Then existingLevel("startAt") = startOverride
            End If

            Dim overrideLevelElement As System.Xml.Linq.XElement = overrideElement.Element(w + "lvl")
            If overrideLevelElement IsNot Nothing Then
                Dim overrideLevel As JObject = ExtractOpenXmlNumberingLevel(overrideLevelElement, stylesById, w)
                overrideInfo("levelDefinition") = overrideLevel.DeepClone()
                ReplaceOrAddLevelDefinition(levels, overrideLevel)
            End If

            levelOverrides.Add(overrideInfo)
        Next

        If levelOverrides.Count > 0 Then definition("levelOverrides") = levelOverrides
    End Sub

    Private Function FindLevelDefinition(levels As JArray, levelNumber As Integer) As JObject
        If levels Is Nothing Then Return Nothing
        For Each token As JToken In levels
            Dim level As JObject = TryCast(token, JObject)
            If level IsNot Nothing AndAlso level("level") IsNot Nothing AndAlso CInt(level("level")) = levelNumber Then
                Return level
            End If
        Next
        Return Nothing
    End Function

    Private Sub ReplaceOrAddLevelDefinition(levels As JArray, replacement As JObject)
        If levels Is Nothing OrElse replacement Is Nothing OrElse replacement("level") Is Nothing Then Return
        Dim levelNumber As Integer = CInt(replacement("level"))

        For index As Integer = 0 To levels.Count - 1
            Dim existing As JObject = TryCast(levels(index), JObject)
            If existing IsNot Nothing AndAlso existing("level") IsNot Nothing AndAlso CInt(existing("level")) = levelNumber Then
                levels(index) = replacement
                Return
            End If
        Next

        levels.Add(replacement)
    End Sub

    Private Function ExtractOpenXmlDocumentDefaults(stylesRoot As System.Xml.Linq.XElement,
                                                    settingsRoot As System.Xml.Linq.XElement,
                                                    w As System.Xml.Linq.XNamespace) As JObject
        Dim defaults As New JObject()

        If settingsRoot IsNot Nothing Then
            Dim defaultTabStopTwips As Integer
            If Integer.TryParse(GetOpenXmlVal(settingsRoot.Element(w + "defaultTabStop")), defaultTabStopTwips) Then
                defaults("defaultTabStop") = CSng(defaultTabStopTwips / 20.0R)
            End If
        End If

        If stylesRoot IsNot Nothing Then
            Dim docDefaults As System.Xml.Linq.XElement = stylesRoot.Element(w + "docDefaults")
            Dim rPrDefault As System.Xml.Linq.XElement = If(docDefaults IsNot Nothing, docDefaults.Element(w + "rPrDefault"), Nothing)
            Dim rPr As System.Xml.Linq.XElement = If(rPrDefault IsNot Nothing, rPrDefault.Element(w + "rPr"), Nothing)

            If rPr IsNot Nothing Then
                Dim fontDefaults As JObject = ExtractOpenXmlRunFormatting(rPr, w)
                If fontDefaults.HasValues Then defaults("fontDefaults") = fontDefaults

                Dim lang As System.Xml.Linq.XElement = rPr.Element(w + "lang")
                If lang IsNot Nothing Then
                    Dim languageDefaults As New JObject()
                    AddOpenXmlAttributeIfPresent(languageDefaults, "latin", lang, w + "val")
                    AddOpenXmlAttributeIfPresent(languageDefaults, "eastAsia", lang, w + "eastAsia")
                    AddOpenXmlAttributeIfPresent(languageDefaults, "bidi", lang, w + "bidi")
                    If languageDefaults.HasValues Then defaults("languageDefaults") = languageDefaults
                End If
            End If
        End If

        ' Safe defaults: the global tab interval is applied automatically; changing Normal's
        ' script fonts or languages remains opt-in because it can affect unrelated document content.
        defaults("applyOnImport") = New JObject From {
            {"defaultTabStop", True},
            {"fontDefaultsToNormalStyle", False},
            {"languageDefaultsToNormalStyle", False}
        }

        Return defaults
    End Function

    Private Function BuildReferencedNumberingDefinitions(openXmlContext As JObject,
                                                         collectedStyles As IEnumerable(Of String)) As JObject
        Dim result As New JObject()
        If openXmlContext Is Nothing OrElse collectedStyles Is Nothing Then Return result

        Dim allDefinitions As JObject = TryCast(openXmlContext("numberingDefinitions"), JObject)
        If allDefinitions Is Nothing Then Return result

        For Each styleName As String In collectedStyles
            Dim styleMeta As JObject = GetStyleOpenXmlMetadata(openXmlContext, styleName)
            Dim binding As JObject = If(styleMeta IsNot Nothing, TryCast(styleMeta("numberingBinding"), JObject), Nothing)
            If binding Is Nothing OrElse binding("sharedListInstanceKey") Is Nothing Then Continue For

            Dim sharedKey As String = CStr(binding("sharedListInstanceKey"))
            If result(sharedKey) Is Nothing AndAlso allDefinitions(sharedKey) IsNot Nothing Then
                result(sharedKey) = allDefinitions(sharedKey).DeepClone()
            End If
        Next

        Return result
    End Function

    Private Function ComputeJsonListTemplateFingerprintV2(definition As JObject) As String
        If definition Is Nothing Then Return ""
        Dim fingerprint As New StringBuilder()

        If definition("outlineNumbered") IsNot Nothing Then fingerprint.Append($"outline={definition("outlineNumbered")};")
        If definition("numberingStyleId") IsNot Nothing Then fingerprint.Append($"style={definition("numberingStyleId")};")

        Dim levels As JArray = TryCast(definition("levels"), JArray)
        If levels IsNot Nothing Then
            For Each token As JToken In levels.OrderBy(Function(item) If(item("level") IsNot Nothing, CInt(item("level")), 0))
                Dim level As JObject = TryCast(token, JObject)
                If level Is Nothing Then Continue For
                fingerprint.Append(If(level("level"), "")).Append(":"c)
                fingerprint.Append(If(level("numberStyle"), "")).Append(":"c)
                fingerprint.Append(If(level("numberFormat"), "")).Append(":"c)
                fingerprint.Append(If(level("numberPosition"), "")).Append(":"c)
                fingerprint.Append(If(level("textPosition"), "")).Append(":"c)
                fingerprint.Append(If(level("tabPosition"), "undefined")).Append(":"c)
                fingerprint.Append(If(level("trailingCharacter"), "")).Append(":"c)
                fingerprint.Append(If(level("paragraphStyleId"), "")).Append("|"c)
            Next
        End If

        Return fingerprint.ToString()
    End Function

    Private Function OpenXmlNumberFormatToWordEnumName(numberFormat As String) As String
        Select Case If(numberFormat, "").Trim().ToLowerInvariant()
            Case "decimal" : Return "wdListNumberStyleArabic"
            Case "upperroman" : Return "wdListNumberStyleUppercaseRoman"
            Case "lowerroman" : Return "wdListNumberStyleLowercaseRoman"
            Case "upperletter" : Return "wdListNumberStyleUppercaseLetter"
            Case "lowerletter" : Return "wdListNumberStyleLowercaseLetter"
            Case "ordinal" : Return "wdListNumberStyleOrdinal"
            Case "cardinaltext" : Return "wdListNumberStyleCardinalText"
            Case "ordinaltext" : Return "wdListNumberStyleOrdinalText"
            Case "bullet" : Return "wdListNumberStyleBullet"
            Case "none" : Return "wdListNumberStyleNone"
            Case "decimalzero" : Return "wdListNumberStyleArabicLZ"
            Case "chicago" : Return "wdListNumberStyleLegal"
            Case "hebrew1" : Return "wdListNumberStyleHebrew1"
            Case "hebrew2" : Return "wdListNumberStyleHebrew2"
            Case "arabicalpha" : Return "wdListNumberStyleArabic1"
            Case "arabicabjad" : Return "wdListNumberStyleArabic2"
            Case "hindivowels" : Return "wdListNumberStyleHindiLetter1"
            Case "hindiconsonants" : Return "wdListNumberStyleHindiLetter2"
            Case "hindinumbers" : Return "wdListNumberStyleHindiArabic"
            Case "hindicounting" : Return "wdListNumberStyleHindiCardinalText"
            Case "thailetters" : Return "wdListNumberStyleThaiLetter"
            Case "thainumbers" : Return "wdListNumberStyleThaiArabic"
            Case "thaicounting" : Return "wdListNumberStyleThaiCardinalText"
            Case "vietnamesecounting" : Return "wdListNumberStyleVietCardinalText"
            Case "russianlower" : Return "wdListNumberStyleLowercaseRussian"
            Case "russianupper" : Return "wdListNumberStyleUppercaseRussian"
            Case Else : Return "wdListNumberStyleArabic"
        End Select
    End Function

    Private Function OpenXmlListAlignmentToWordEnumName(alignment As String) As String
        Select Case If(alignment, "").Trim().ToLowerInvariant()
            Case "center" : Return "wdListLevelAlignCenter"
            Case "right" : Return "wdListLevelAlignRight"
            Case Else : Return "wdListLevelAlignLeft"
        End Select
    End Function

    Private Function OpenXmlSuffixToWordEnumName(suffix As String) As String
        Select Case If(suffix, "").Trim().ToLowerInvariant()
            Case "space" : Return "wdTrailingSpace"
            Case "nothing" : Return "wdTrailingNone"
            Case Else : Return "wdTrailingTab"
        End Select
    End Function

    Private Sub MergeOpenXmlStyleMetadata(styleDef As JObject,
                                          styleMeta As JObject,
                                          openXmlContext As JObject)
        If styleDef Is Nothing OrElse styleMeta Is Nothing Then Return

        Dim metadataFields As String() = {
            "styleId", "styleTypeOpenXml", "customStyle", "basedOnStyleId", "basedOnStyleName",
            "nextStyleId", "nextStyleName", "uiPriority", "quickFormat"
        }

        For Each fieldName As String In metadataFields
            If styleMeta(fieldName) IsNot Nothing Then styleDef(fieldName) = styleMeta(fieldName).DeepClone()
        Next

        If styleMeta("numberingBinding") IsNot Nothing Then
            styleDef("numberingBinding") = styleMeta("numberingBinding").DeepClone()

            Dim listFormat As JObject = TryCast(styleDef("listFormat"), JObject)
            If listFormat Is Nothing Then
                listFormat = New JObject From {
                    {"hasListTemplate", True}
                }
                styleDef("listFormat") = listFormat
            End If

            listFormat("numberingBinding") = styleMeta("numberingBinding").DeepClone()
            If styleMeta("numberingBinding")("sharedListInstanceKey") IsNot Nothing Then
                listFormat("sharedListInstanceKey") = styleMeta("numberingBinding")("sharedListInstanceKey").DeepClone()
            End If
            If styleMeta("numberingBinding")("effectiveListLevel") IsNot Nothing Then
                listFormat("linkedLevel") = styleMeta("numberingBinding")("effectiveListLevel").DeepClone()
            End If

            Dim sharedKey As String = If(styleMeta("numberingBinding")("sharedListInstanceKey") IsNot Nothing,
                                         CStr(styleMeta("numberingBinding")("sharedListInstanceKey")), "")
            Dim numberingDefinitions As JObject = If(openXmlContext IsNot Nothing,
                                                      TryCast(openXmlContext("numberingDefinitions"), JObject), Nothing)
            Dim numberingDefinition As JObject = If(numberingDefinitions IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(sharedKey),
                                                     TryCast(numberingDefinitions(sharedKey), JObject), Nothing)
            If numberingDefinition IsNot Nothing Then
                If numberingDefinition("outlineNumbered") IsNot Nothing Then listFormat("outlineNumbered") = numberingDefinition("outlineNumbered").DeepClone()
                If numberingDefinition("numberingStyleId") IsNot Nothing Then listFormat("numberingStyleId") = numberingDefinition("numberingStyleId").DeepClone()
                If numberingDefinition("numberingStyleName") IsNot Nothing Then listFormat("numberingStyleName") = numberingDefinition("numberingStyleName").DeepClone()
                If numberingDefinition("templateFingerprintV2") IsNot Nothing Then listFormat("templateFingerprintV2") = numberingDefinition("templateFingerprintV2").DeepClone()

                Dim sourceLevels As JArray = TryCast(numberingDefinition("levels"), JArray)
                If sourceLevels IsNot Nothing Then
                    Dim targetLevels As JArray = TryCast(listFormat("levels"), JArray)
                    If targetLevels Is Nothing OrElse targetLevels.Count = 0 Then
                        listFormat("levels") = sourceLevels.DeepClone()
                    Else
                        MergeNumberingLevelMetadata(targetLevels, sourceLevels)
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub MergeNumberingLevelMetadata(targetLevels As JArray, sourceLevels As JArray)
        If targetLevels Is Nothing OrElse sourceLevels Is Nothing Then Return

        For Each sourceToken As JToken In sourceLevels
            Dim sourceLevel As JObject = TryCast(sourceToken, JObject)
            If sourceLevel Is Nothing OrElse sourceLevel("level") Is Nothing Then Continue For

            Dim targetLevel As JObject = FindLevelDefinition(targetLevels, CInt(sourceLevel("level")))
            If targetLevel Is Nothing Then
                targetLevels.Add(sourceLevel.DeepClone())
                Continue For
            End If

            For Each propertyItem As JProperty In sourceLevel.Properties()
                Select Case propertyItem.Name
                    Case "paragraphStyleId", "paragraphStyleName", "numberRunFormatting",
                         "tabPositionState", "numberStyleOpenXml", "alignmentOpenXml",
                         "trailingCharacterOpenXml", "ilvlZeroBased"
                        targetLevel(propertyItem.Name) = propertyItem.Value.DeepClone()
                End Select
            Next
        Next
    End Sub

#End Region
    ''' <summary>
    ''' Extracts a Word style definition (paragraph formatting, font formatting, tab stops, and list template data)
    ''' into a JSON object for inclusion in a DocStyle template.
    ''' </summary>
    ''' <param name="doc">Active document containing the style.</param>
    ''' <param name="styleName">Style name to extract.</param>
    ''' <returns>Extracted style definition JSON, or Nothing if the style cannot be resolved.</returns>
    ''' <summary>
    ''' Extracts a Word style definition and augments it with style IDs, inheritance,
    ''' shared numbering-instance metadata and Open XML list-level links.
    ''' </summary>
    Private Function ExtractFullStyleDefinition(doc As Word.Document,
                                                styleName As String,
                                                Optional openXmlContext As JObject = Nothing) As JObject
        Try
            Dim style As Word.Style = doc.Styles(styleName)
            If style Is Nothing Then Return Nothing

            Dim styleDef As New JObject()
            styleDef("wdStyleName") = style.NameLocal
            styleDef("styleType") = style.Type.ToString()
            styleDef("builtIn") = style.BuiltIn

            Try
                styleDef("quickFormat") = style.QuickStyle
            Catch
            End Try
            Try
                styleDef("uiPriority") = style.Priority
            Catch
            End Try

            ' Base style hierarchy. Empty names and cycles are ignored.
            Dim baseStyles As New JArray()
            Dim seenBaseStyles As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim currentStyle As Word.Style = style
            Try
                While currentStyle IsNot Nothing AndAlso currentStyle.BaseStyle IsNot Nothing
                    Dim baseStyle As Word.Style = TryCast(currentStyle.BaseStyle, Word.Style)
                    If baseStyle Is Nothing Then Exit While
                    Dim baseName As String = If(baseStyle.NameLocal, "").Trim()
                    If String.IsNullOrWhiteSpace(baseName) OrElse Not seenBaseStyles.Add(baseName) Then Exit While
                    baseStyles.Add(baseName)
                    currentStyle = baseStyle
                End While
            Catch
            End Try
            If baseStyles.Count > 0 Then styleDef("baseStyleHierarchy") = baseStyles

            Try
                If style.BaseStyle IsNot Nothing Then
                    Dim directBaseStyle As Word.Style = TryCast(style.BaseStyle, Word.Style)
                    If directBaseStyle IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(directBaseStyle.NameLocal) Then
                        styleDef("basedOnStyleName") = directBaseStyle.NameLocal
                    End If
                End If
            Catch
            End Try

            Try
                If style.NextParagraphStyle IsNot Nothing Then
                    Dim nextStyle As Word.Style = TryCast(style.NextParagraphStyle, Word.Style)
                    If nextStyle IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(nextStyle.NameLocal) Then
                        styleDef("nextParagraphStyle") = nextStyle.NameLocal
                        styleDef("nextStyleName") = nextStyle.NameLocal
                    End If
                End If
            Catch
            End Try

            Try
                Dim pf As Word.ParagraphFormat = style.ParagraphFormat
                styleDef("paragraphFormat") = New JObject From {
                    {"alignment", pf.Alignment.ToString()},
                    {"leftIndent", pf.LeftIndent},
                    {"rightIndent", pf.RightIndent},
                    {"firstLineIndent", pf.FirstLineIndent},
                    {"spaceBefore", pf.SpaceBefore},
                    {"spaceAfter", pf.SpaceAfter},
                    {"lineSpacing", pf.LineSpacing},
                    {"lineSpacingRule", pf.LineSpacingRule.ToString()},
                    {"keepTogether", pf.KeepTogether},
                    {"keepWithNext", pf.KeepWithNext},
                    {"pageBreakBefore", pf.PageBreakBefore},
                    {"widowControl", pf.WidowControl},
                    {"outlineLevel", pf.OutlineLevel.ToString()}
                }
            Catch
            End Try

            ' Only custom tab stops are style data. Word also enumerates generated default tabs.
            Try
                Dim tabs As New JArray()
                For Each tabStop As Word.TabStop In style.ParagraphFormat.TabStops
                    Dim isCustom As Boolean = True
                    Try
                        isCustom = tabStop.CustomTab
                    Catch
                    End Try
                    If Not isCustom Then Continue For

                    tabs.Add(New JObject From {
                        {"position", Math.Round(tabStop.Position, 2)},
                        {"alignment", tabStop.Alignment.ToString()},
                        {"leader", tabStop.Leader.ToString()}
                    })
                Next
                styleDef("tabStops") = tabs
            Catch
            End Try

            Try
                Dim font As Word.Font = style.Font
                Dim fontFormat As New JObject From {
                    {"name", font.Name},
                    {"size", font.Size},
                    {"bold", (font.Bold = -1)},
                    {"italic", (font.Italic = -1)},
                    {"underline", font.Underline.ToString()},
                    {"allCaps", (font.AllCaps = -1)},
                    {"smallCaps", (font.SmallCaps = -1)},
                    {"strikeThrough", (font.StrikeThrough = -1)},
                    {"doubleStrikeThrough", (font.DoubleStrikeThrough = -1)},
                    {"subscript", (font.Subscript = -1)},
                    {"superscript", (font.Superscript = -1)},
                    {"color", font.Color.ToString()},
                    {"colorRGB", ColorToRGB(font.Color)}
                }
                Try
                    fontFormat("scaling") = font.Scaling
                    fontFormat("spacing") = font.Spacing
                    fontFormat("position") = font.Position
                    fontFormat("kerning") = font.Kerning
                Catch
                End Try
                AddScriptFontPropertiesToJson(font, fontFormat)
                styleDef("fontFormat") = fontFormat
            Catch
            End Try

            Try
                Dim listTemplate As Word.ListTemplate = style.ListTemplate
                If listTemplate IsNot Nothing Then
                    Dim listFormat As New JObject From {
                        {"hasListTemplate", True},
                        {"outlineNumbered", listTemplate.OutlineNumbered}
                    }

                    Try
                        listFormat("linkedLevel") = style.ListLevelNumber
                    Catch
                        listFormat("linkedLevel") = 1
                    End Try

                    Try
                        Dim fingerprint As New StringBuilder()
                        For levelNumber As Integer = 1 To Math.Min(9, listTemplate.ListLevels.Count)
                            Try
                                Dim level As Word.ListLevel = listTemplate.ListLevels(levelNumber)
                                fingerprint.Append($"{levelNumber}:{level.NumberStyle}:{level.NumberFormat}|")
                            Catch
                            End Try
                        Next
                        listFormat("templateFingerprint") = fingerprint.ToString()
                    Catch
                    End Try

                    Dim levels As New JArray()
                    For levelNumber As Integer = 1 To listTemplate.ListLevels.Count
                        Try
                            Dim level As Word.ListLevel = listTemplate.ListLevels(levelNumber)
                            Dim levelInfo As New JObject From {
                                {"level", levelNumber},
                                {"numberStyle", level.NumberStyle.ToString()},
                                {"textPosition", level.TextPosition},
                                {"numberPosition", level.NumberPosition},
                                {"alignment", level.Alignment.ToString()},
                                {"startAt", level.StartAt}
                            }

                            If IsWordUndefinedPosition(level.TabPosition) Then
                                levelInfo("tabPosition") = JValue.CreateNull()
                                levelInfo("tabPositionState") = "wdUndefined"
                            Else
                                levelInfo("tabPosition") = level.TabPosition
                                levelInfo("tabPositionState") = "defined"
                            End If

                            If level.NumberStyle = WdListNumberStyle.wdListNumberStyleBullet Then
                                Try
                                    Dim bulletFontName As String = ""
                                    Try
                                        If level.Font IsNot Nothing AndAlso Not String.IsNullOrEmpty(level.Font.Name) Then
                                            bulletFontName = level.Font.Name
                                        End If
                                    Catch
                                    End Try
                                    levelInfo("bulletFont") = If(String.IsNullOrEmpty(bulletFontName), "Symbol", bulletFontName)

                                    If Not String.IsNullOrEmpty(level.NumberFormat) Then
                                        levelInfo("bulletCharCode") = AscW(level.NumberFormat.Chars(0))
                                    Else
                                        levelInfo("bulletCharCode") = &H2022
                                    End If
                                    levelInfo("numberFormat") = ""
                                Catch ex As System.Exception
                                    Debug.WriteLine($"Error extracting bullet info for level {levelNumber}: {ex.Message}")
                                    levelInfo("bulletCharCode") = &H2022
                                    levelInfo("bulletFont") = "Symbol"
                                    levelInfo("numberFormat") = ""
                                End Try
                            Else
                                levelInfo("numberFormat") = level.NumberFormat
                            End If

                            Try
                                levelInfo("trailingCharacter") = level.TrailingCharacter.ToString()
                            Catch
                            End Try

                            Try
                                Dim numberRunFormatting As New JObject()
                                AddScriptFontPropertiesToJson(level.Font, numberRunFormatting)
                                If Not String.IsNullOrWhiteSpace(level.Font.Name) Then numberRunFormatting("name") = level.Font.Name
                                If level.Font.Size > 0 Then numberRunFormatting("size") = level.Font.Size
                                If level.Font.Bold <> CInt(WdConstants.wdUndefined) Then numberRunFormatting("bold") = (level.Font.Bold = -1)
                                If level.Font.Italic <> CInt(WdConstants.wdUndefined) Then numberRunFormatting("italic") = (level.Font.Italic = -1)
                                If numberRunFormatting.HasValues Then levelInfo("numberRunFormatting") = numberRunFormatting
                            Catch
                            End Try

                            levels.Add(levelInfo)
                        Catch
                        End Try
                    Next

                    listFormat("levels") = levels
                    styleDef("listFormat") = listFormat
                End If
            Catch
                ' The Open XML metadata below can still supply inherited numbering information.
            End Try

            Dim styleMeta As JObject = GetStyleOpenXmlMetadata(openXmlContext, style.NameLocal)
            If styleMeta IsNot Nothing Then MergeOpenXmlStyleMetadata(styleDef, styleMeta, openXmlContext)

            Return styleDef

        Catch ex As System.Exception
            Debug.WriteLine($"Error extracting style definition for '{styleName}': {ex.Message}")
            Return Nothing
        End Try
    End Function
#End Region

#Region "Step 2: Apply Style Template"

    ''' <summary>
    ''' Applies a DocStyle template to the current selection or (if no selection) the active document.
    ''' This method collects user parameters, loads the selected template JSON, optionally applies Word style
    ''' definitions, and then applies LLM-provided style mappings to target paragraphs.
    ''' </summary>
    Public Async Sub ApplyStyleTemplate()
        If INILoadFail() OrElse Not IsDocumentEditable() Then Return

        Dim do2ndModel As Boolean = False
        Dim settings As New DocStyleSettings()
        settings.Load()

        Try
            ' Expand paths
            Dim docStylePath As String = ExpandEnvironmentVariables(INI_DocStylePath)
            If Not String.IsNullOrEmpty(docStylePath) AndAlso Not docStylePath.EndsWith("\") Then docStylePath &= "\"

            Dim docStylePathLocal As String = ExpandEnvironmentVariables(INI_DocStylePathLocal)
            If Not String.IsNullOrEmpty(docStylePathLocal) AndAlso Not docStylePathLocal.EndsWith("\") Then docStylePathLocal &= "\"

            Dim hasGlobal As Boolean = Not String.IsNullOrWhiteSpace(docStylePath) AndAlso Directory.Exists(docStylePath)
            Dim hasLocal As Boolean = Not String.IsNullOrWhiteSpace(docStylePathLocal) AndAlso Directory.Exists(docStylePathLocal)

            If Not hasGlobal AndAlso Not hasLocal Then
                ShowCustomMessageBox("No style template paths are configured. Please configure 'DocStylePath' or 'DocStylePathLocal' in your INI file.")
                Return
            End If

            Dim app As Word.Application = Globals.ThisAddIn.Application
            If app Is Nothing OrElse app.Documents Is Nothing OrElse app.Documents.Count = 0 Then
                ShowCustomMessageBox("No open document.")
                Return
            End If

            Dim doc As Word.Document = app.ActiveDocument
            _sharedListTemplates.Clear()
            If doc Is Nothing Then
                ShowCustomMessageBox("Active document was not found.")
                Return
            End If

            Dim currentSelection As Word.Selection = app.Selection
            Dim targetRange As Word.Range

            If currentSelection.Type = WdSelectionType.wdSelectionIP Then
                Dim answer As Integer = ShowCustomYesNoBox("You have not selected any text. Do you want to apply the style template to the entire document?",
                    "Yes, entire document", "No, cancel", $"{AN} - Apply Style Template")
                If answer <> 1 Then Return
                targetRange = doc.Content
            Else
                targetRange = currentSelection.Range.Duplicate
            End If

            Dim templates As List(Of DocStyleTemplate) = LoadStyleTemplates(docStylePath, docStylePathLocal)
            If templates Is Nothing OrElse templates.Count = 0 Then
                ShowCustomMessageBox($"No valid style templates found. Create templates using 'Extract Style Template' and save them as '{AN2}-ds-*.json' files.")
                Return
            End If

            Dim displayToTemplate As New Dictionary(Of String, DocStyleTemplate)(StringComparer.OrdinalIgnoreCase)
            Dim displayOptions As New List(Of String)()
            For Each t In templates
                Dim display As String = t.DisplayName
                If t.IsLocal Then display &= " (local)"
                Dim originalDisplay As String = display
                Dim counter As Integer = 1
                While displayToTemplate.ContainsKey(display)
                    counter += 1
                    display = $"{originalDisplay} ({counter})"
                End While
                displayToTemplate(display) = t
                displayOptions.Add(display)
            Next

            ' Parameter form
            Dim defaultTemplateDisplay As String = If(displayOptions.Count > 0, displayOptions(0), "")
            If Not String.IsNullOrWhiteSpace(settings.LastStyleTemplateDisplay) AndAlso displayOptions.Contains(settings.LastStyleTemplateDisplay) Then
                defaultTemplateDisplay = settings.LastStyleTemplateDisplay
            End If

            Dim p0 As New SLib.InputParameter("Style Template", defaultTemplateDisplay)
            p0.Options = displayOptions

            Dim p2 As New SLib.InputParameter("Create/update Word styles from template", settings.ApplyStyleDefinitions)
            Dim p3 As New SLib.InputParameter("Preview each change", settings.PreviewMode)
            Dim p4 As New SLib.InputParameter("Use confidence threshold", settings.UseConfidenceThreshold)
            Dim p5 As New SLib.InputParameter("Confidence threshold (0-100%)", settings.ConfidenceThreshold)
            Dim p6 As New SLib.InputParameter("Process table cells", settings.ProcessTables)

            Dim p7 As New SLib.InputParameter("Apply only style, no text updating (faster, safer)", settings.FastModeStylesOnly)

            Dim listResetOptions As New List(Of String) From {"Off", "Rule-based (after non-list paragraphs)", "LLM-assisted (semantic analysis)"}
            Dim p8b As New SLib.InputParameter("List numbering reset", listResetOptions(settings.ListNumberingReset))
            p8b.Options = listResetOptions

            Dim headingOutlineLevelOptions As New List(Of String) From {
                "0 (never treat outline levels as headings)",
                "1", "2", "3", "4", "5", "6", "7", "8", "9"
            }

            Dim initialHeadingMax As Integer = settings.RuleHeadingOutlineLevelMax
            If initialHeadingMax < 0 Then initialHeadingMax = 0
            If initialHeadingMax > 9 Then initialHeadingMax = 9

            Dim p8c As New SLib.InputParameter("Maximum heading levels (rule-based reset)", headingOutlineLevelOptions(initialHeadingMax))
            p8c.Options = headingOutlineLevelOptions

            Dim p10 As New SLib.InputParameter("Document context (optional)", settings.DocumentContext)

            Dim p11 As SLib.InputParameter
            If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                p11 = New SLib.InputParameter("Use a secondary model", settings.UseSecondaryModel)
            ElseIf INI_SecondAPI Then
                p11 = New SLib.InputParameter("Use the secondary model", settings.UseSecondaryModel)
            Else
                p11 = New SLib.InputParameter("Use the secondary model", CType(Nothing, Boolean?))
            End If

            Dim p1 As New SLib.InputParameter("Apply in Track Changes", settings.TrackChanges)

            ' Seeds the UI option from the persisted setting value.            
            Dim p1b As New SLib.InputParameter("Restore in-para formatting (requires Track Changes)", settings.RestoreInlineFormatting)

            Dim p1c As New SLib.InputParameter("Remove empty paragraphs after styling", settings.RemoveEmptyLines)

            Dim p9 As New SLib.InputParameter("Show report at end", settings.ShowReport)

            Dim params() As SLib.InputParameter = {p0, p2, p3, p4, p5, p6, p7, p8b, p8c, p10, p11, p1, p1b, p1c, p9}
            If Not ShowCustomVariableInputForm("Configure Style Template Application:", $"{AN} - Apply Style Template", params) Then
                Return
            End If

            ' Read back values
            Dim chosenDisplay As String = System.Convert.ToString(params(0).Value)

            settings.LastStyleTemplateDisplay = chosenDisplay
            settings.ApplyStyleDefinitions = System.Convert.ToBoolean(params(1).Value)
            settings.PreviewMode = System.Convert.ToBoolean(params(2).Value)
            settings.UseConfidenceThreshold = System.Convert.ToBoolean(params(3).Value)
            settings.ConfidenceThreshold = System.Convert.ToInt32(params(4).Value)
            settings.ProcessTables = System.Convert.ToBoolean(params(5).Value)
            settings.FastModeStylesOnly = System.Convert.ToBoolean(params(6).Value)

            settings.ListNumberingReset = listResetOptions.IndexOf(System.Convert.ToString(params(7).Value))
            If settings.ListNumberingReset < 0 Then settings.ListNumberingReset = 0

            Dim headingMaxRaw As String = System.Convert.ToString(params(8).Value)
            Dim m As Match = Regex.Match(headingMaxRaw, "^\s*(\d+)")
            If m.Success Then
                settings.RuleHeadingOutlineLevelMax = Math.Max(0, Math.Min(9, Integer.Parse(m.Groups(1).Value)))
            Else
                settings.RuleHeadingOutlineLevelMax = DefaultRuleHeadingOutlineLevelMax
            End If

            settings.DocumentContext = System.Convert.ToString(params(9).Value)

            Dim secondModel = params(10).Value
            If TypeOf secondModel Is Boolean Then
                do2ndModel = CBool(secondModel)
                settings.UseSecondaryModel = do2ndModel
            End If

            settings.TrackChanges = System.Convert.ToBoolean(params(11).Value)
            settings.RestoreInlineFormatting = System.Convert.ToBoolean(params(12).Value)
            settings.RemoveEmptyLines = System.Convert.ToBoolean(params(13).Value)
            settings.ShowReport = System.Convert.ToBoolean(params(14).Value)

            settings.Save()

            Dim chosenTemplate As DocStyleTemplate = Nothing
            If Not displayToTemplate.TryGetValue(chosenDisplay, chosenTemplate) Then
                ShowCustomMessageBox("Selected style template could not be resolved.")
                Return
            End If

            Dim templateJson As String
            Try
                templateJson = File.ReadAllText(chosenTemplate.FilePath, Encoding.UTF8)
            Catch ex As System.Exception
                ShowCustomMessageBox($"Could not read template file: {ex.Message}")
                Return
            End Try

            Dim templateObj As JObject
            Try
                templateObj = JObject.Parse(templateJson)
            Catch ex As System.Exception
                ShowCustomMessageBox($"Invalid JSON in template file: {ex.Message}")
                Return
            End Try

            If do2ndModel Then
                If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                    If Not ShowModelSelection(_context, INI_AlternateModelPath) Then
                        originalConfigLoaded = False
                        ShowCustomMessageBox("The secondary model could not be loaded - aborting.")
                        Return
                    End If
                End If
            End If

            Dim wdStyleDefinitions As JObject = TryCast(templateObj("wdStyleDefinitions"), JObject)
            Dim numberingDefinitions As JObject = TryCast(templateObj("numberingDefinitions"), JObject)

            If settings.ApplyStyleDefinitions Then
                ApplyDocumentDefaultsToDocument(doc, TryCast(templateObj("documentDefaults"), JObject))
                If wdStyleDefinitions IsNot Nothing Then
                    ApplyWdStyleDefinitionsToDocument(doc, wdStyleDefinitions, numberingDefinitions)
                End If
            End If

            If templateObj("wdStyleDefinitions") IsNot Nothing Then templateObj.Remove("wdStyleDefinitions")

            Dim templateForLLM As String = CreateMinimalTemplateForLLM(templateObj)

            ' Track changes independent.
            Dim originalTrackChanges As Boolean = doc.TrackRevisions
            If settings.TrackChanges Then
                doc.TrackRevisions = True
            End If

            Try
                Using New WordUndoScope(app, $"{AN} - Apply Style Template")
                    Dim report As String = Await ApplyStylesFromTemplate(doc, targetRange, templateForLLM, templateObj, settings, do2ndModel)

                    If settings.ShowReport AndAlso Not String.IsNullOrWhiteSpace(report) Then
                        Dim reportResult As String = ShowCustomWindow(
                            "Style Template Application Complete",
                            report,
                            "The report has been copied to your clipboard.",
                            $"{AN} - Application Report",
                            NoRTF:=True,
                            Getfocus:=True)
                        SLib.PutInClipboard(report)
                    End If
                End Using

                ' Remove empty paragraphs if requested (done after styling, respects Track Changes setting)
                If settings.RemoveEmptyLines Then
                    Dim removedCount As Integer = RemoveEmptyParagraphs(targetRange)
                    Debug.WriteLine($"[DocStyle] Removed {removedCount} empty paragraph(s)")
                End If

                If settings.TrackChanges AndAlso settings.RestoreInlineFormatting Then
                    UndoInlineFormattingRevisions(doc)
                End If

            Finally
                doc.TrackRevisions = originalTrackChanges
            End Try

        Catch ex As System.Exception
            ShowCustomMessageBox($"Error applying style template: {ex.Message}")
        Finally
            _sharedListTemplates.Clear()
            If do2ndModel AndAlso originalConfigLoaded Then
                RestoreDefaults(_context, originalConfig)
                originalConfigLoaded = False
            End If
        End Try
    End Sub

    ''' <summary>
    ''' Applies styles and optional text amendments to paragraphs in <paramref name="targetRange"/> based on an LLM
    ''' mapping response. Produces a text report of applied and skipped operations.
    ''' </summary>
    ''' <param name="doc">Active document being modified.</param>
    ''' <param name="targetRange">Range whose paragraphs are processed.</param>
    ''' <param name="templateJson">Minimal template JSON passed to the LLM.</param>
    ''' <param name="templateObj">Full template object used for lookups of formatting overrides.</param>
    ''' <param name="settings">DocStyle runtime settings.</param>
    ''' <param name="useSecondAPI">When True, calls the secondary model/API for the LLM request.</param>
    ''' <returns>Plain text report.</returns>
    Private Async Function ApplyStylesFromTemplate(doc As Word.Document, targetRange As Word.Range,
                                        templateJson As String, templateObj As JObject,
                                        settings As DocStyleSettings,
                                        useSecondAPI As Boolean) As System.Threading.Tasks.Task(Of String)
        Dim report As New StringBuilder()
        report.AppendLine("=== Style Template Application Report (Fast Mode) ===")
        report.AppendLine($"Date: {DateTime.Now:yyyy-MM-dd HH:mm:ss}")
        report.AppendLine()

        Try
            ' Text amendment (deletePrefix) is controlled by FastModeStylesOnly.
            ' TrackChanges is independent; caller toggles doc.TrackRevisions.
            Dim doTextAmendment As Boolean = (Not settings.FastModeStylesOnly)

            ApplyStylesResponseFormat = If(doTextAmendment, ApplyStylesResponseFormat_Extended, ApplyStylesResponseFormat_Minimal)

            ' Build user style lookup tables
            Dim userStyleToWdStyle As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Dim userStyleNameToDef As New Dictionary(Of String, JObject)(StringComparer.OrdinalIgnoreCase)

            If templateObj("userStyles") IsNot Nothing Then
                For Each userStyle As JObject In CType(templateObj("userStyles"), JArray)
                    Dim rawUserStyleName As String = If(userStyle("userStyleName") IsNot Nothing, userStyle("userStyleName").ToString(), "")
                    Dim userStyleKey As String = NormalizeStyleKey(rawUserStyleName)

                    If String.IsNullOrWhiteSpace(userStyleKey) Then Continue For

                    Dim wdStyleName As String = If(userStyle("wdStyleName") IsNot Nothing, userStyle("wdStyleName").ToString(), "Normal")
                    userStyleToWdStyle(userStyleKey) = wdStyleName
                    userStyleNameToDef(userStyleKey) = userStyle
                Next
            End If

            ' Build numbered paragraph list with formatting hints
            Dim paragraphTexts As New StringBuilder()
            Dim paragraphList As New List(Of Word.Paragraph)()
            Dim paragraphHasList As New List(Of Boolean)()
            Dim idx As Integer = 1

            For Each para As Word.Paragraph In targetRange.Paragraphs
                Dim text As String = para.Range.Text.TrimEnd(vbCr, vbLf, ChrW(13), ChrW(10))
                If String.IsNullOrWhiteSpace(text) Then Continue For

                If Not settings.ProcessTables Then
                    Try
                        If para.Range.Cells.Count > 0 Then Continue For
                    Catch
                    End Try
                End If

                Dim hints As New StringBuilder()

                Dim hasList As Boolean = False
                Try
                    Dim lf As Word.ListFormat = para.Range.ListFormat
                    If lf.ListType <> WdListType.wdListNoNumbering Then
                        hasList = True

                        Dim listString As String = lf.ListString
                        Dim isBullet As Boolean = False

                        If lf.ListType = WdListType.wdListBullet Then
                            isBullet = True
                        ElseIf Not String.IsNullOrEmpty(listString) Then
                            Dim trimmed As String = listString.Trim()
                            If trimmed.Length > 0 Then
                                Dim firstChar As Char = trimmed.Chars(0)
                                If Not Char.IsLetterOrDigit(firstChar) Then
                                    isBullet = True
                                End If
                            End If
                        End If

                        Dim listTypeHint As String = If(isBullet, "Bullet", lf.ListType.ToString().Replace("wdList", ""))
                        hints.Append($"[LIST:{listTypeHint}/L{lf.ListLevelNumber}] ")
                    End If
                Catch
                End Try

                Try
                    If para.LeftIndent > 0 OrElse para.FirstLineIndent <> 0 Then
                        hints.Append($"[INDENT:{Math.Round(para.LeftIndent, 0)}/{Math.Round(para.FirstLineIndent, 0)}] ")
                    End If
                Catch
                End Try

                paragraphTexts.AppendLine($"[{idx}] {hints}{text}")
                paragraphList.Add(para)
                paragraphHasList.Add(hasList)
                idx += 1
            Next

            If paragraphList.Count = 0 Then
                report.AppendLine("No paragraphs to process.")
                Return report.ToString()
            End If

            Dim systemPrompt As String = SP_ApplyDocStyle
            If settings.ListNumberingReset = 2 Then systemPrompt &= SP_ApplyDocStyle_NumberingHint
            systemPrompt = InterpolateAtRuntime(systemPrompt)

            Dim userPrompt As New StringBuilder()
            userPrompt.AppendLine("<STYLETEMPLATE>")
            userPrompt.AppendLine(templateJson)
            userPrompt.AppendLine("</STYLETEMPLATE>")
            userPrompt.AppendLine()
            userPrompt.AppendLine("<DOCUMENT>")
            userPrompt.AppendLine(paragraphTexts.ToString())
            userPrompt.AppendLine("</DOCUMENT>")

            If Not String.IsNullOrWhiteSpace(settings.DocumentContext) Then
                userPrompt.AppendLine()
                userPrompt.AppendLine("<CONTEXT>")
                userPrompt.AppendLine(settings.DocumentContext)
                userPrompt.AppendLine("</CONTEXT>")
            End If

            Dim response As String = Await LLM(systemPrompt, userPrompt.ToString(), "", "", 0, useSecondAPI)

            Dim mappingArray As JArray
            Try
                response = ExtractJsonFromResponse(response)
                mappingArray = JArray.Parse(response)
            Catch ex As System.Exception
                report.AppendLine($"Error parsing LLM response: {ex.Message}")
                Return report.ToString()
            End Try

            Dim appliedCount As Integer = 0
            Dim skippedCount As Integer = 0

            ShowProgressBarInSeparateThread($"{AN} - Applying Styles", "Applying styles...")
            ProgressBarModule.CancelOperation = False
            GlobalProgressMax = mappingArray.Count
            GlobalProgressValue = 0

            Dim numberingRestartParas As New List(Of Word.Paragraph)()
            Dim templateBaselineByIndex As New Dictionary(Of Integer, JObject)()

            Try
                For Each mapping As JObject In mappingArray
                    If ProgressBarModule.CancelOperation Then
                        report.AppendLine("Operation cancelled by user.")
                        Exit For
                    End If

                    Dim paraIdx As Integer = CInt(mapping("paragraphIndex")) - 1

                    Dim userStyleNameRaw As String = If(mapping("userStyleName") IsNot Nothing, CStr(mapping("userStyleName")), "")
                    Dim userStyleName As String = NormalizeStyleKey(userStyleNameRaw)
                    Dim confidence As Integer = If(mapping("confidence") IsNot Nothing, CInt(mapping("confidence")), 100)

                    GlobalProgressValue += 1
                    GlobalProgressLabel = $"Processing paragraph {paraIdx + 1} of {paragraphList.Count}"

                    If paraIdx < 0 OrElse paraIdx >= paragraphList.Count Then Continue For

                    Dim para As Word.Paragraph = paragraphList(paraIdx)
                    Dim paraPreview As String = GetParaPreview(para, ReportParaPreviewLen)

                    Dim wdStyleName As String = "Normal"
                    If userStyleToWdStyle.ContainsKey(userStyleName) Then
                        wdStyleName = userStyleToWdStyle(userStyleName)
                    ElseIf Not String.IsNullOrWhiteSpace(userStyleName) Then
                        wdStyleName = userStyleName
                    End If

                    If settings.UseConfidenceThreshold AndAlso confidence < settings.ConfidenceThreshold Then
                        skippedCount += 1
                        report.AppendLine($"Skipped paragraph {paraIdx + 1}: confidence {confidence}% below threshold | suggested='{userStyleNameRaw}' norm='{userStyleName}' wdStyle='{wdStyleName}' | text='{paraPreview}'")
                        Continue For
                    End If

                    Dim userStyleDef As JObject = Nothing
                    userStyleNameToDef.TryGetValue(userStyleName, userStyleDef)

                    ' Skip if no actual changes would occur
                    Dim wouldChange As Boolean = WouldStyleApplicationChangeAnything(para, doc, wdStyleName, userStyleDef)
                    If Not wouldChange Then
                        skippedCount += 1
                        report.AppendLine($"Skipped paragraph {paraIdx + 1}: no changes needed | style='{wdStyleName}' | text='{paraPreview}'")
                        Continue For
                    End If

                    If settings.PreviewMode Then
                        ' Temporarily highlight the paragraph for visibility
                        Dim originalHighlight As WdColorIndex = WdColorIndex.wdNoHighlight
                        Try
                            originalHighlight = para.Range.HighlightColorIndex
                            para.Range.HighlightColorIndex = WdColorIndex.wdYellow
                        Catch
                        End Try

                        para.Range.Select()

                        ' Ensure the selection is visible in the document window
                        Try
                            Globals.ThisAddIn.Application.ActiveWindow.ScrollIntoView(para.Range)
                            Globals.ThisAddIn.Application.ScreenRefresh()
                        Catch
                        End Try

                        Dim preview As Integer = ShowCustomYesNoBox(
                        $"Apply user style '{userStyleNameRaw}' (normalized: '{userStyleName}') (Word style: '{wdStyleName}') to this paragraph?{vbCrLf}{vbCrLf}Confidence: {confidence}%{vbCrLf}Text: {para.Range.Text.Substring(0, Math.Min(100, para.Range.Text.Length))}...",
                        "Yes", "Skip", $"{AN} - Preview")

                        ' Restore original highlight
                        Try
                            para.Range.HighlightColorIndex = originalHighlight
                        Catch
                        End Try

                        If preview = 0 Then
                            Dim continueChoice As Integer = ShowCustomYesNoBox(
                            "You closed the preview dialog without making a selection." & vbCrLf & vbCrLf &
                            "Do you want to continue applying styles without individual preview, or abort the operation?",
                            "Continue without preview", "Abort", $"{AN} - Continue?")
                            If continueChoice = 1 Then
                                settings.PreviewMode = False
                            Else
                                report.AppendLine("Operation aborted by user.")
                                Exit For
                            End If
                        ElseIf preview <> 1 Then
                            skippedCount += 1
                            report.AppendLine($"Skipped paragraph {paraIdx + 1}: user skipped in preview | suggested='{userStyleNameRaw}' norm='{userStyleName}' wdStyle='{wdStyleName}' | text='{paraPreview}'")
                            Continue For
                        End If
                    End If

                    Try
                        ' Optional text amendment: deletes a manual prefix marker (before applying the style).
                        If doTextAmendment AndAlso mapping("deletePrefix") IsNot Nothing AndAlso TypeOf mapping("deletePrefix") Is JObject Then
                            TryDeleteManualPrefixMarker(para, CType(mapping("deletePrefix"), JObject), report, paraIdx + 1)
                        End If

                        para.Style = doc.Styles(wdStyleName)

                        If userStyleDef IsNot Nothing Then
                            ApplyUserStyleFormattingFromTemplate(para, userStyleDef)
                        Else
                            Debug.WriteLine($"[DocStyle] FastMode: No userStyleDef for para {paraIdx + 1}: raw='{userStyleNameRaw}' norm='{userStyleName}' wdStyle='{wdStyleName}' preview='{paraPreview}'")
                        End If

                        templateBaselineByIndex(paraIdx) = CaptureParaProps(para)
                        appliedCount += 1

                        If settings.ListNumberingReset > 0 Then
                            Dim shouldRestart As Boolean = False
                            If settings.ListNumberingReset = 2 Then
                                shouldRestart = If(mapping("restartNumbering") IsNot Nothing, CBool(mapping("restartNumbering")), False)
                            End If
                            If shouldRestart Then numberingRestartParas.Add(para)
                        End If

                    Catch ex As System.Exception
                        report.AppendLine($"Error applying style '{wdStyleName}' to paragraph {paraIdx + 1}: {ex.Message} | suggested='{userStyleNameRaw}' norm='{userStyleName}' wdStyle='{wdStyleName}' | text='{paraPreview}'")
                    End Try
                Next
            Finally
                ProgressBarModule.CancelOperation = True
            End Try

            Dim restartCount As Integer = 0
            If settings.ListNumberingReset > 0 Then
                restartCount = ApplyNumberingRestarts(doc, paragraphList, paragraphHasList, settings, numberingRestartParas, templateBaselineByIndex)
            End If

            report.AppendLine()
            report.AppendLine($"Total paragraphs: {paragraphList.Count}")
            report.AppendLine($"Styles applied: {appliedCount}")
            report.AppendLine($"Skipped: {skippedCount}")
            If settings.ListNumberingReset > 0 Then report.AppendLine($"Numbering restarts applied: {restartCount}")

        Catch ex As System.Exception
            report.AppendLine($"Error in fast mode: {ex.Message}")
        End Try

        Return report.ToString()
    End Function


    ''' <summary>
    ''' Determines whether applying a style and formatting to a paragraph would result in any actual changes.
    ''' </summary>
    ''' <param name="para">Paragraph to check.</param>
    ''' <param name="doc">Document containing the paragraph.</param>
    ''' <param name="wdStyleName">Word style name to apply.</param>
    ''' <param name="userStyleDef">User style definition containing formatting overrides.</param>
    ''' <returns>True if applying the style would change the paragraph; otherwise False.</returns>
    Private Function WouldStyleApplicationChangeAnything(para As Word.Paragraph, doc As Word.Document, wdStyleName As String, userStyleDef As JObject) As Boolean
        Try
            ' Check if style would change
            Dim currentStyleName As String = ""
            Try
                currentStyleName = para.Style.NameLocal
            Catch
                currentStyleName = ""
            End Try

            If Not String.Equals(currentStyleName, wdStyleName, StringComparison.OrdinalIgnoreCase) Then
                Return True ' Style name differs
            End If

            ' Check paragraph formatting differences if userStyleDef is provided
            If userStyleDef IsNot Nothing AndAlso userStyleDef("paragraphFormatting") IsNot Nothing Then
                Dim pf = userStyleDef("paragraphFormatting")

                If pf("alignment") IsNot Nothing Then
                    Dim desired = ParseAlignment(CStr(pf("alignment")))
                    If para.Alignment <> desired Then Return True
                End If

                If pf("leftIndent") IsNot Nothing Then
                    If Math.Abs(para.LeftIndent - CSng(pf("leftIndent"))) > 0.5 Then Return True
                End If

                If pf("rightIndent") IsNot Nothing Then
                    If Math.Abs(para.RightIndent - CSng(pf("rightIndent"))) > 0.5 Then Return True
                End If

                If pf("firstLineIndent") IsNot Nothing Then
                    If Math.Abs(para.FirstLineIndent - CSng(pf("firstLineIndent"))) > 0.5 Then Return True
                End If

                If pf("spaceBefore") IsNot Nothing Then
                    If Math.Abs(para.SpaceBefore - CSng(pf("spaceBefore"))) > 0.5 Then Return True
                End If

                If pf("spaceAfter") IsNot Nothing Then
                    If Math.Abs(para.SpaceAfter - CSng(pf("spaceAfter"))) > 0.5 Then Return True
                End If
            End If

            ' Check font formatting differences
            If userStyleDef IsNot Nothing AndAlso userStyleDef("fontFormatting") IsNot Nothing Then
                Dim ff = userStyleDef("fontFormatting")
                Dim rng As Word.Range = para.Range

                If ff("fontName") IsNot Nothing AndAlso CStr(ff("fontName")) <> "mixed" Then
                    If rng.Font.Name <> CStr(ff("fontName")) Then Return True
                End If

                If ff("fontSize") IsNot Nothing AndAlso CStr(ff("fontSize")) <> "mixed" Then
                    If Math.Abs(rng.Font.Size - CSng(ff("fontSize"))) > 0.1 Then Return True
                End If

                If ff("bold") IsNot Nothing AndAlso ff("bold").Type = JTokenType.Boolean Then
                    Dim desired As Integer = If(CBool(ff("bold")), -1, 0)
                    If rng.Font.Bold <> desired AndAlso rng.Font.Bold <> CInt(WdConstants.wdUndefined) Then Return True
                End If

                If ff("italic") IsNot Nothing AndAlso ff("italic").Type = JTokenType.Boolean Then
                    Dim desired As Integer = If(CBool(ff("italic")), -1, 0)
                    If rng.Font.Italic <> desired AndAlso rng.Font.Italic <> CInt(WdConstants.wdUndefined) Then Return True
                End If

                If HasScriptFontDifference(rng.Font, ff) Then Return True
            End If

            ' Check list formatting transitions
            If userStyleDef IsNot Nothing AndAlso userStyleDef("listFormatting") IsNot Nothing Then
                Dim lf = userStyleDef("listFormatting")
                Dim templateHasList As Boolean = If(lf("hasList") IsNot Nothing, CBool(lf("hasList")), False)

                Dim currentHasList As Boolean = False
                Try
                    currentHasList = (para.Range.ListFormat.ListType <> WdListType.wdListNoNumbering)
                Catch
                End Try

                If templateHasList <> currentHasList Then Return True
                If templateHasList AndAlso currentHasList AndAlso lf("listLevelNumber") IsNot Nothing Then
                    Try
                        If para.Range.ListFormat.ListLevelNumber <> CInt(lf("listLevelNumber")) Then Return True
                    Catch
                        Return True
                    End Try
                End If
            End If

            ' No differences detected
            Return False

        Catch ex As System.Exception
            Debug.WriteLine($"[DocStyle] WouldStyleApplicationChangeAnything error: {ex.Message}")
            ' On error, assume change is needed to be safe
            Return True
        End Try
    End Function

    ''' <summary>
    ''' Rejects formatting-only revisions that affect only a portion of a paragraph.
    ''' </summary>
    ''' <param name="doc">Document whose revisions are inspected.</param>
    Private Sub UndoInlineFormattingRevisions(doc As Word.Document)
        If doc Is Nothing Then Exit Sub

        Try
            If doc.Revisions Is Nothing OrElse doc.Revisions.Count = 0 Then Exit Sub

            ' Snapshot revisions first (O(1) enumerator access instead of O(n) ordinal
            ' seeks per item), then reject from the end so removals do not shift the
            ' indices of items still to be processed.
            Dim allRevisions As New List(Of Word.Revision)()
            For Each snapRev As Word.Revision In doc.Revisions
                allRevisions.Add(snapRev)
            Next

            For i As Integer = allRevisions.Count - 1 To 0 Step -1
                Dim rev As Word.Revision = allRevisions(i)
                If rev Is Nothing Then Continue For

                ' Check by enum value, not string
                If rev.Type <> WdRevisionType.wdRevisionProperty Then Continue For

                Dim rng As Word.Range = Nothing
                Try
                    rng = rev.Range
                Catch
                    Continue For
                End Try
                If rng Is Nothing Then Continue For

                Dim para As Word.Paragraph = Nothing
                Try
                    If rng.Paragraphs Is Nothing OrElse rng.Paragraphs.Count <> 1 Then Continue For
                    para = rng.Paragraphs(1)
                Catch
                    Continue For
                End Try
                If para Is Nothing Then Continue For

                ' Get paragraph text length (excluding paragraph mark)
                Dim paraTextStart As Integer = para.Range.Start
                Dim paraTextLen As Integer = 0
                Try
                    Dim paraText As String = para.Range.Text
                    ' Remove trailing paragraph mark(s)
                    paraText = paraText.TrimEnd(vbCr, vbLf, ChrW(13), ChrW(10))
                    paraTextLen = paraText.Length
                Catch
                    Continue For
                End Try

                If paraTextLen <= 0 Then Continue For

                Dim paraTextEnd As Integer = paraTextStart + paraTextLen

                ' Calculate effective revision span within paragraph
                Dim revStart As Integer = Math.Max(rng.Start, paraTextStart)
                Dim revEnd As Integer = Math.Min(rng.End, paraTextEnd)

                If revEnd <= revStart Then Continue For

                ' Inline = revision does NOT cover entire paragraph text
                Dim revLen As Integer = revEnd - revStart
                If revLen >= paraTextLen Then
                    ' Full paragraph formatting - leave tracked
                    Continue For
                End If

                ' Reject (revert) inline formatting revision
                Try
                    rev.Reject()
                Catch ex As System.Exception
                    Debug.WriteLine($"[DocStyle] Could not reject inline formatting revision {i}: {ex.Message}")
                End Try
            Next

        Catch ex As System.Exception
            Debug.WriteLine($"[DocStyle] UndoInlineFormattingRevisions error: {ex.Message}")
        End Try
    End Sub



    ''' <summary>
    ''' Normalizes style names for dictionary lookups by removing repeated whitespace and normalizing non-breaking spaces.
    ''' </summary>
    ''' <param name="s">Raw style name.</param>
    ''' <returns>Normalized style key.</returns>
    Private Shared Function NormalizeStyleKey(s As String) As String
        If s Is Nothing Then Return ""

        ' Normalizes common invisible differences from LLM / copy-paste.
        s = s.Replace(ChrW(&HA0), " ") ' NBSP -> space
        s = Regex.Replace(s, "\s+", " ").Trim()

        Return s
    End Function

    ''' <summary>
    ''' Deletes a leading manual prefix marker from a paragraph as instructed by an LLM mapping response.
    ''' Deletion occurs only at the paragraph start and never removes the paragraph mark.
    ''' </summary>
    ''' <param name="para">Paragraph to amend.</param>
    ''' <param name="deletePrefix">Delete-prefix instruction object.</param>
    ''' <param name="report">Report sink for non-fatal messages.</param>
    ''' <param name="logicalIdx1Based">1-based paragraph index used in the report.</param>
    ''' <returns>True if a prefix was deleted; otherwise False.</returns>
    Private Function TryDeleteManualPrefixMarker(para As Word.Paragraph, deletePrefix As JObject, ByRef report As StringBuilder, logicalIdx1Based As Integer) As Boolean
        Try
            If para Is Nothing OrElse deletePrefix Is Nothing Then Return False

            Dim kind As String = ""
            Try : kind = CStr(deletePrefix("kind")) : Catch : kind = "" : End Try

            Dim charCount As Integer = 0
            Try : charCount = CInt(deletePrefix("charCount")) : Catch : charCount = 0 : End Try

            Dim expectedText As String = ""
            Try : expectedText = CStr(deletePrefix("text")) : Catch : expectedText = "" : End Try

            If charCount <= 0 Then Return False
            If String.Equals(kind, "None", StringComparison.OrdinalIgnoreCase) Then Return False

            Dim pr As Word.Range = para.Range.Duplicate

            ' Exclude paragraph mark to avoid deleting it
            If pr.End > pr.Start Then
                Try
                    Dim lastChar As String = pr.Duplicate.Characters.Last.Text
                    If lastChar = vbCr OrElse lastChar = vbLf Then
                        pr.End -= 1
                    End If
                Catch
                End Try
            End If

            If pr.End <= pr.Start Then Return False

            ' Bound check
            Dim maxLen As Integer = pr.End - pr.Start
            If charCount > maxLen Then charCount = maxLen
            If charCount <= 0 Then Return False

            Dim prefixRng As Word.Range = para.Range.Document.Range(pr.Start, pr.Start + charCount)

            ' Optional verification by exact text match
            If Not String.IsNullOrEmpty(expectedText) Then
                Dim actual As String = prefixRng.Text
                actual = actual.Replace(ChrW(&HA0), " ") ' NBSP normalize
                Dim exp As String = expectedText.Replace(ChrW(&HA0), " ")

                If Not String.Equals(actual, exp, StringComparison.Ordinal) Then
                    report.AppendLine($"Did not delete prefix for paragraph {logicalIdx1Based}: expected '{exp}' but found '{actual}'.")
                    Return False
                End If
            End If

            prefixRng.Delete()
            Return True

        Catch ex As System.Exception
            report.AppendLine($"Error deleting prefix for paragraph {logicalIdx1Based}: {ex.Message}")
            Return False
        End Try
    End Function


    ''' <summary>
    ''' Determines whether a paragraph is part of a numbered list (bullets excluded).
    ''' </summary>
    ''' <param name="p">Paragraph to inspect.</param>
    ''' <returns>True if the paragraph is part of a non-bullet list; otherwise False.</returns>
    Private Function IsNumberedListParagraph(p As Word.Paragraph) As Boolean
        Try
            Dim lf As Word.ListFormat = p.Range.ListFormat
            If lf Is Nothing Then Return False

            Dim t As WdListType = lf.ListType
            If t = WdListType.wdListNoNumbering Then Return False

            ' Explicit bullets => not numbered
            If t = WdListType.wdListBullet OrElse t = WdListType.wdListPictureBullet Then
                Return False
            End If

            ' Treat all other list types as numbered.
            Return True
        Catch
            Return False
        End Try
    End Function



    ''' <summary>
    ''' Returns a single-line paragraph text preview (no paragraph marks/newlines), truncated to <paramref name="maxLen"/>.
    ''' </summary>
    ''' <param name="p">Paragraph to preview.</param>
    ''' <param name="maxLen">Maximum number of characters returned.</param>
    ''' <returns>Preview string.</returns>
    Private Function GetParaPreview(p As Word.Paragraph, Optional maxLen As Integer = 25) As String
        Try
            Dim t As String = p.Range.Text
            t = t.Replace(vbCr, " ").Replace(vbLf, " ").Replace(ChrW(13), " ").Replace(ChrW(10), " ")
            t = t.Trim()

            If t.Length <= maxLen Then Return t
            Return t.Substring(0, maxLen) & "..."
        Catch
            Return ""
        End Try
    End Function


    ''' <summary>
    ''' Captures paragraph-level properties that are commonly disturbed by list operations.
    ''' The style is intentionally excluded.
    ''' </summary>
    ''' <param name="p">Paragraph to snapshot.</param>
    ''' <returns>JSON object containing captured properties.</returns>
    Private Function CaptureParaProps(p As Word.Paragraph) As JObject
        Dim o As New JObject()

        Try : o("alignment") = p.Alignment.ToString() : Catch : End Try
        Try : o("leftIndent") = p.LeftIndent : Catch : End Try
        Try : o("rightIndent") = p.RightIndent : Catch : End Try
        Try : o("firstLineIndent") = p.FirstLineIndent : Catch : End Try
        Try : o("spaceBefore") = p.SpaceBefore : Catch : End Try
        Try : o("spaceAfter") = p.SpaceAfter : Catch : End Try
        Try : o("lineSpacing") = p.LineSpacing : Catch : End Try
        Try : o("lineSpacingRule") = p.LineSpacingRule.ToString() : Catch : End Try
        Try : o("keepTogether") = p.KeepTogether : Catch : End Try
        Try : o("keepWithNext") = p.KeepWithNext : Catch : End Try
        Try : o("pageBreakBefore") = p.PageBreakBefore : Catch : End Try
        Try : o("widowControl") = p.WidowControl : Catch : End Try
        Try : o("outlineLevel") = p.OutlineLevel.ToString() : Catch : End Try

        ' Tab stops are often clobbered by list operations; preserve if present.
        Try
            Dim tabs As New JArray()
            For Each ts As Word.TabStop In p.TabStops
                Dim t As New JObject()
                t("pos") = Math.Round(ts.Position, 2)
                t("align") = ts.Alignment.ToString()
                t("leader") = ts.Leader.ToString()
                tabs.Add(t)
            Next
            o("tabStops") = tabs
        Catch
        End Try

        Return o
    End Function

    ''' <summary>
    ''' Restores paragraph properties from a snapshot previously created by <see cref="CaptureParaProps"/>.
    ''' </summary>
    ''' <param name="p">Paragraph to restore.</param>
    ''' <param name="props">Snapshot containing paragraph properties and optional tab stop definitions.</param>
    Private Sub RestoreParaProps(p As Word.Paragraph, props As JObject)
        If props Is Nothing Then Exit Sub
        Try
            If props("alignment") IsNot Nothing Then p.Alignment = ParseAlignment(CStr(props("alignment")))
            If props("leftIndent") IsNot Nothing Then p.LeftIndent = CSng(props("leftIndent"))
            If props("rightIndent") IsNot Nothing Then p.RightIndent = CSng(props("rightIndent"))
            If props("firstLineIndent") IsNot Nothing Then p.FirstLineIndent = CSng(props("firstLineIndent"))
            If props("spaceBefore") IsNot Nothing Then p.SpaceBefore = CSng(props("spaceBefore"))
            If props("spaceAfter") IsNot Nothing Then p.SpaceAfter = CSng(props("spaceAfter"))
            If props("lineSpacing") IsNot Nothing Then p.LineSpacing = CSng(props("lineSpacing"))

            If props("lineSpacingRule") IsNot Nothing Then
                p.LineSpacingRule = ParseLineSpacingRule(CStr(props("lineSpacingRule")))
            End If

            If props("keepTogether") IsNot Nothing Then p.KeepTogether = CInt(props("keepTogether"))
            If props("keepWithNext") IsNot Nothing Then p.KeepWithNext = CInt(props("keepWithNext"))
            If props("pageBreakBefore") IsNot Nothing Then p.PageBreakBefore = CInt(props("pageBreakBefore"))
            If props("widowControl") IsNot Nothing Then p.WidowControl = CInt(props("widowControl"))
            If props("outlineLevel") IsNot Nothing Then p.OutlineLevel = ParseOutlineLevel(CStr(props("outlineLevel")))

            ' Restore tab stops
            If props("tabStops") IsNot Nothing AndAlso TypeOf props("tabStops") Is JArray Then
                Try
                    p.TabStops.ClearAll()
                    For Each tabDef As JObject In CType(props("tabStops"), JArray)
                        Dim pos As Single = CSng(tabDef("pos"))
                        Dim al As WdTabAlignment = ParseTabAlignment(CStr(tabDef("align")))
                        Dim ld As WdTabLeader = ParseTabLeader(CStr(tabDef("leader")))
                        p.TabStops.Add(pos, al, ld)
                    Next
                Catch
                End Try
            End If
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Restarts numbering for a single paragraph by re-applying its current list template with
    ''' <c>ContinuePreviousList:=False</c> and restoring paragraph properties afterwards.
    ''' </summary>
    ''' <param name="p">Paragraph whose numbering is restarted.</param>
    ''' <param name="logicalIndex1Based">1-based paragraph index used in debug output.</param>
    ''' <param name="restoreProps">
    ''' Optional baseline properties to restore after list operations; when omitted, a snapshot is captured from <paramref name="p"/>.
    ''' </param>
    Private Sub RestartNumberingForParagraph(p As Word.Paragraph,
                                        Optional logicalIndex1Based As Integer = -1,
                                        Optional restoreProps As JObject = Nothing)
        Try
            Dim preview As String = GetParaPreview(p, 25)

            Dim beforeStyle As String = ""
            Try : beforeStyle = p.Style.NameLocal : Catch : End Try

            ' If caller provided a baseline (e.g., template baseline), use it; otherwise snapshot current properties.
            Dim propsToRestore As JObject = If(restoreProps, CaptureParaProps(p))

            Dim lf As Word.ListFormat = Nothing
            Try : lf = p.Range.ListFormat : Catch : lf = Nothing : End Try

            If lf Is Nothing OrElse lf.ListType = WdListType.wdListNoNumbering Then
                Debug.WriteLine($"[DocStyle] RestartNumbering SKIP idx={logicalIndex1Based} preview='{preview}' sig='{GetListSignature(p)}'")
                Exit Sub
            End If

            Dim tpl As Word.ListTemplate = Nothing
            Try : tpl = lf.ListTemplate : Catch : tpl = Nothing : End Try
            If tpl Is Nothing Then
                Debug.WriteLine($"[DocStyle] RestartNumbering SKIP idx={logicalIndex1Based} preview='{preview}' (no ListTemplate) sig='{GetListSignature(p)}'")
                Exit Sub
            End If

            Dim lvl As Integer = 1
            Try : lvl = lf.ListLevelNumber : Catch : lvl = 1 : End Try

            Debug.WriteLine($"[DocStyle] RestartNumbering START idx={logicalIndex1Based} preview='{preview}' style='{beforeStyle}' sig='{GetListSignature(p)}'")

            Dim applyTo As WdListApplyTo = WdListApplyTo.wdListApplyToWholeList
            Try
                p.Range.ListFormat.ApplyListTemplateWithLevel(
                ListTemplate:=tpl,
                ContinuePreviousList:=False,
                ApplyTo:=applyTo,
                DefaultListBehavior:=WdDefaultListBehavior.wdWord10ListBehavior,
                ApplyLevel:=lvl
            )
            Catch
                p.Range.ListFormat.ApplyListTemplateWithLevel(
                ListTemplate:=tpl,
                ContinuePreviousList:=False,
                ApplyTo:=WdListApplyTo.wdListApplyToThisPointForward,
                DefaultListBehavior:=WdDefaultListBehavior.wdWord10ListBehavior,
                ApplyLevel:=lvl
            )
            End Try

            ' Restore the intended baseline (template baseline if provided).
            RestoreParaProps(p, propsToRestore)

            Debug.WriteLine($"[DocStyle] RestartNumbering END   idx={logicalIndex1Based} preview='{preview}' styleBefore='{beforeStyle}' styleAfter='{TryGetStyleName(p)}' sigAfter='{GetListSignature(p)}'")
        Catch ex As System.Exception
            Debug.WriteLine($"[DocStyle] RestartNumberingForParagraph error idx={logicalIndex1Based}: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Returns the paragraph style name, or an empty string if unavailable.
    ''' </summary>
    ''' <param name="p">Paragraph to inspect.</param>
    ''' <returns>Style name (<c>NameLocal</c>) or empty string.</returns>
    Private Function TryGetStyleName(p As Word.Paragraph) As String
        Try
            Return p.Style.NameLocal
        Catch
            Return ""
        End Try
    End Function


    ''' <summary>
    ''' Determines whether a paragraph is considered "heading-like" based on its outline level and the configured maximum.
    ''' </summary>
    ''' <param name="p">Paragraph to inspect.</param>
    ''' <param name="headingOutlineLevelMax">Maximum outline level treated as heading-like (0 disables).</param>
    ''' <returns>True if the paragraph outline level is within the configured range; otherwise False.</returns>
    Private Function IsHeadingLikeParagraph(p As Word.Paragraph, headingOutlineLevelMax As Integer) As Boolean
        If headingOutlineLevelMax <= 0 Then Return False
        If headingOutlineLevelMax > 9 Then headingOutlineLevelMax = 9

        Try
            Dim ol As WdOutlineLevel = p.OutlineLevel

            Select Case ol
                Case WdOutlineLevel.wdOutlineLevel1 : Return (headingOutlineLevelMax >= 1)
                Case WdOutlineLevel.wdOutlineLevel2 : Return (headingOutlineLevelMax >= 2)
                Case WdOutlineLevel.wdOutlineLevel3 : Return (headingOutlineLevelMax >= 3)
                Case WdOutlineLevel.wdOutlineLevel4 : Return (headingOutlineLevelMax >= 4)
                Case WdOutlineLevel.wdOutlineLevel5 : Return (headingOutlineLevelMax >= 5)
                Case WdOutlineLevel.wdOutlineLevel6 : Return (headingOutlineLevelMax >= 6)
                Case WdOutlineLevel.wdOutlineLevel7 : Return (headingOutlineLevelMax >= 7)
                Case WdOutlineLevel.wdOutlineLevel8 : Return (headingOutlineLevelMax >= 8)
                Case WdOutlineLevel.wdOutlineLevel9 : Return (headingOutlineLevelMax >= 9)
                Case Else
                    Return False ' BodyText
            End Select
        Catch
            Return False
        End Try
    End Function


    ''' <summary>
    ''' Determines an effective list nesting level for a paragraph, preferring the Word list level when available
    ''' and otherwise inferring from paragraph indentation relative to <paramref name="baseIndent"/>.
    ''' </summary>
    ''' <param name="p">Paragraph to inspect.</param>
    ''' <param name="baseIndent">Base indent (points) used as the reference for level 1.</param>
    ''' <returns>Effective nesting level in the range 1..9.</returns>
    Private Function GuessNestingLevel(p As Word.Paragraph, baseIndent As Single) As Integer
        Try
            Dim lf As Word.ListFormat = p.Range.ListFormat
            Dim lvl As Integer = 1
            Try : lvl = lf.ListLevelNumber : Catch : lvl = 1 : End Try
            If lvl > 1 Then Return lvl

            ' Manual indent nesting: compare paragraph left indent to the run's base indent
            Dim li As Single = 0
            Try : li = CSng(p.LeftIndent) : Catch : li = 0 : End Try

            Dim delta As Single = li - baseIndent
            If delta <= (IndentStepPoints * 0.5F) Then Return 1

            Dim inferred As Integer = 1 + CInt(Math.Floor(delta / IndentStepPoints))
            If inferred < 1 Then inferred = 1
            If inferred > 9 Then inferred = 9
            Return inferred
        Catch
            Return 1
        End Try
    End Function


    ''' <summary>
    ''' Applies numbering restarts according to the configured mode and returns the number of restarts applied.
    ''' </summary>
    ''' <param name="doc">Active document being modified.</param>
    ''' <param name="paragraphList">Paragraphs processed in display order.</param>
    ''' <param name="paragraphHasList">Per-paragraph indicator whether the paragraph has list formatting.</param>
    ''' <param name="settings">DocStyle runtime settings (controls restart mode and heading detection).</param>
    ''' <param name="llmRestartParas">Paragraphs marked for restart by the LLM (LLM-assisted mode only).</param>
    ''' <param name="templateBaselineByIndex">Baseline paragraph property snapshots captured during style application.</param>
    ''' <returns>Number of paragraphs for which numbering restart was applied.</returns>
    Private Function ApplyNumberingRestarts(doc As Word.Document,
                                       paragraphList As List(Of Word.Paragraph),
                                       paragraphHasList As List(Of Boolean),
                                       settings As DocStyleSettings,
                                       llmRestartParas As List(Of Word.Paragraph),
                                       templateBaselineByIndex As Dictionary(Of Integer, JObject)) As Integer
        Dim restartCount As Integer = 0

        Try
            Dim restartIndices As New HashSet(Of Integer)()

            If settings.ListNumberingReset = 1 Then
                Dim inBodyNumberedRun As Boolean = False
                Dim restartedMainThisRun As Boolean = False

                Dim lastSeenIndexByLevel As New Dictionary(Of Integer, Integer)()
                Dim restartedForParentKey As New HashSet(Of String)(StringComparer.Ordinal)

                Dim runBaseIndent As Single = 0
                Dim haveRunBaseIndent As Boolean = False

                For i As Integer = 0 To paragraphList.Count - 1
                    Dim p = paragraphList(i)

                    Dim isNumbered As Boolean = IsNumberedListParagraph(p)
                    Dim isBullet As Boolean = If(RuleBreakOnBullets, IsBulletListParagraph(p), False)
                    Dim isHeadingLike As Boolean = IsHeadingLikeParagraph(p, settings.RuleHeadingOutlineLevelMax)

                    If (Not isNumbered) OrElse isBullet OrElse isHeadingLike Then
                        inBodyNumberedRun = False
                        restartedMainThisRun = False
                        lastSeenIndexByLevel.Clear()
                        restartedForParentKey.Clear()
                        haveRunBaseIndent = False
                        runBaseIndent = 0
                        Continue For
                    End If

                    If Not inBodyNumberedRun Then
                        inBodyNumberedRun = True
                        restartedMainThisRun = False
                        lastSeenIndexByLevel.Clear()
                        restartedForParentKey.Clear()
                        haveRunBaseIndent = False
                        runBaseIndent = 0
                    End If

                    If Not haveRunBaseIndent Then
                        Try : runBaseIndent = CSng(p.LeftIndent) : Catch : runBaseIndent = 0 : End Try
                        haveRunBaseIndent = True
                    End If

                    Dim effLevel As Integer = GuessNestingLevel(p, runBaseIndent)

                    Dim parentIdx As Integer = -1
                    If effLevel > 1 Then
                        If lastSeenIndexByLevel.ContainsKey(effLevel - 1) Then
                            parentIdx = lastSeenIndexByLevel(effLevel - 1)
                        Else
                            For pl As Integer = effLevel - 2 To 1 Step -1
                                If lastSeenIndexByLevel.ContainsKey(pl) Then
                                    parentIdx = lastSeenIndexByLevel(pl)
                                    Exit For
                                End If
                            Next
                        End If
                    End If

                    If effLevel = 1 Then
                        If Not restartedMainThisRun Then
                            restartIndices.Add(i)
                            restartedMainThisRun = True
                            Debug.WriteLine($"[DocStyle] RuleRestart(main) idx={i + 1} effLevel=1 leftIndent={Math.Round(p.LeftIndent, 1)} preview='{GetParaPreview(p, 25)}' sig='{GetListSignature(p)}'")
                        End If
                    Else
                        Dim key As String = $"{effLevel}:{parentIdx}"
                        If parentIdx >= 0 AndAlso Not restartedForParentKey.Contains(key) Then
                            restartIndices.Add(i)
                            restartedForParentKey.Add(key)
                            Debug.WriteLine($"[DocStyle] RuleRestart(sub) idx={i + 1} effLevel={effLevel} parentIdx={parentIdx + 1} leftIndent={Math.Round(p.LeftIndent, 1)} preview='{GetParaPreview(p, 25)}' sig='{GetListSignature(p)}'")
                        ElseIf parentIdx < 0 Then
                            Dim fallbackKey As String = $"{effLevel}:-1"
                            If Not restartedForParentKey.Contains(fallbackKey) Then
                                restartIndices.Add(i)
                                restartedForParentKey.Add(fallbackKey)
                                Debug.WriteLine($"[DocStyle] RuleRestart(sub-fallback) idx={i + 1} effLevel={effLevel} leftIndent={Math.Round(p.LeftIndent, 1)} preview='{GetParaPreview(p, 25)}' sig='{GetListSignature(p)}'")
                            End If
                        End If
                    End If

                    lastSeenIndexByLevel(effLevel) = i

                    Dim deeper = lastSeenIndexByLevel.Keys.Where(Function(k) k > effLevel).ToList()
                    For Each k In deeper
                        lastSeenIndexByLevel.Remove(k)
                    Next
                Next

            ElseIf settings.ListNumberingReset = 2 Then
                For i As Integer = 0 To paragraphList.Count - 1
                    If llmRestartParas.Contains(paragraphList(i)) AndAlso IsNumberedListParagraph(paragraphList(i)) Then
                        restartIndices.Add(i)
                    End If
                Next
            End If

            For Each idx In restartIndices.OrderBy(Function(x) x)
                Dim p As Word.Paragraph = paragraphList(idx)

                Dim baseline As JObject = Nothing
                If templateBaselineByIndex IsNot Nothing Then
                    templateBaselineByIndex.TryGetValue(idx, baseline)
                End If

                Debug.WriteLine($"[DocStyle] ApplyNumberingRestarts restarting paragraph {idx + 1} preview='{GetParaPreview(p, 25)}' sig='{GetListSignature(p)}'")

                ' Restart using template baseline restore
                RestartNumberingForParagraph(p, idx + 1, baseline)
                restartCount += 1

                ' Word can perturb sibling paragraphs in the same list/template after a restart.
                ' Apply baseline to paragraphs that share the same ListTemplate instance as the restarted paragraph.
                Try
                    Dim restartedTplId As Integer = GetListTemplateId(p)
                    If restartedTplId <> 0 AndAlso templateBaselineByIndex IsNot Nothing Then
                        For j As Integer = 0 To paragraphList.Count - 1
                            If j = idx Then Continue For

                            Dim pj As Word.Paragraph = paragraphList(j)

                            ' Same list template instance? Then reapply the template baseline props we captured earlier.
                            If GetListTemplateId(pj) = restartedTplId Then
                                Dim bj As JObject = Nothing
                                If templateBaselineByIndex.TryGetValue(j, bj) AndAlso bj IsNot Nothing Then
                                    RestoreParaProps(pj, bj)
                                End If
                            End If
                        Next
                    End If
                Catch ex As System.Exception
                    Debug.WriteLine($"[DocStyle] Repair after restart failed: {ex.Message}")
                End Try
            Next

        Catch ex As System.Exception
            Debug.WriteLine($"Error in ApplyNumberingRestarts: {ex.Message}")
        End Try

        Return restartCount
    End Function

    ''' <summary>
    ''' Determines whether a paragraph is part of a bullet list.
    ''' </summary>
    ''' <param name="p">Paragraph to inspect.</param>
    ''' <returns>True if the paragraph list type is a bullet list; otherwise False.</returns>
    Private Function IsBulletListParagraph(p As Word.Paragraph) As Boolean
        Try
            Dim lf = p.Range.ListFormat
            If lf Is Nothing Then Return False
            Return lf.ListType = WdListType.wdListBullet OrElse lf.ListType = WdListType.wdListPictureBullet
        Catch
            Return False
        End Try
    End Function


    ''' <summary>
    ''' Builds a debug signature describing a paragraph's list type, list-template identity, list level, and list string.
    ''' </summary>
    ''' <param name="p">Paragraph to inspect.</param>
    ''' <returns>Debug signature string.</returns>
    Private Function GetListSignature(p As Word.Paragraph) As String
        Try
            Dim lf = p.Range.ListFormat
            If lf Is Nothing OrElse lf.ListType = WdListType.wdListNoNumbering Then Return "NoList"

            Dim tplHash As String = "tpl=?"
            Try
                ' Not stable across sessions, but good enough for in-run debug
                tplHash = $"tpl#{RuntimeHelpers.GetHashCode(lf.ListTemplate)}"
            Catch
            End Try

            Dim lvl As Integer = 1
            Try : lvl = lf.ListLevelNumber : Catch : End Try

            Dim ls As String = ""
            Try : ls = lf.ListString : Catch : End Try

            Return $"{lf.ListType}/{tplHash}/L{lvl}/'{ls}'"
        Catch
            Return "ListSigError"
        End Try
    End Function

    ''' <summary>
    ''' Returns an in-process identifier for a paragraph's list template.
    ''' </summary>
    ''' <param name="p">Paragraph to inspect.</param>
    ''' <returns>Runtime hash code of the list template instance, or 0 if unavailable.</returns>
    Private Function GetListTemplateId(p As Word.Paragraph) As Integer
        Try
            Dim tpl = p.Range.ListFormat.ListTemplate
            If tpl Is Nothing Then Return 0
            Return RuntimeHelpers.GetHashCode(tpl)
        Catch
            Return 0
        End Try
    End Function




    ''' <summary>
    ''' Creates a minimal style template JSON string for LLM consumption.
    ''' Only includes <c>userStyleName</c>, <c>whenToApply</c>, and an optional <c>hasList</c> hint.
    ''' </summary>
    ''' <param name="templateObj">Full template JSON object containing <c>templateName</c> and <c>userStyles</c>.</param>
    ''' <returns>Serialized minimal JSON string (no indentation).</returns>
    Private Function CreateMinimalTemplateForLLM(templateObj As JObject) As String
        Dim minimal As New JObject()
        minimal("templateName") = templateObj("templateName")

        Dim minimalStyles As New JArray()
        If templateObj("userStyles") IsNot Nothing Then
            For Each userStyle As JObject In CType(templateObj("userStyles"), JArray)
                Dim minStyle As New JObject()
                minStyle("userStyleName") = userStyle("userStyleName")
                minStyle("whenToApply") = userStyle("whenToApply")

                ' Includes a list presence hint when available in the template.
                If userStyle("listFormatting") IsNot Nothing AndAlso
                   userStyle("listFormatting")("hasList") IsNot Nothing Then
                    minStyle("hasList") = userStyle("listFormatting")("hasList")
                End If

                minimalStyles.Add(minStyle)
            Next
        End If
        minimal("userStyles") = minimalStyles

        Return minimal.ToString(Formatting.None)
    End Function


#Region "DocStyle metadata application helpers"

    Private Function IsWordUndefinedPosition(value As Single) As Boolean
        Return value = CSng(WdConstants.wdUndefined) OrElse Math.Abs(CDbl(value)) > 1000000.0R
    End Function

    Private Sub AddScriptFontPropertiesToJson(font As Word.Font, target As JObject)
        If font Is Nothing OrElse target Is Nothing Then Return

        Try
            Dim fontObject As Object = font
            AddLateBoundStringProperty(target, "nameAscii", fontObject, "NameAscii")
            AddLateBoundStringProperty(target, "nameFarEast", fontObject, "NameFarEast")
            AddLateBoundStringProperty(target, "nameBi", fontObject, "NameBi")
            AddLateBoundStringProperty(target, "nameOther", fontObject, "NameOther")
            AddLateBoundStringProperty(target, "languageId", fontObject, "LanguageID")
            AddLateBoundStringProperty(target, "languageIdFarEast", fontObject, "LanguageIDFarEast")
            AddLateBoundStringProperty(target, "languageIdOther", fontObject, "LanguageIDOther")
        Catch
        End Try
    End Sub

    Private Sub AddLateBoundStringProperty(target As JObject,
                                           jsonPropertyName As String,
                                           source As Object,
                                           comPropertyName As String)
        Try
            Dim value As Object = source.GetType().InvokeMember(
                comPropertyName,
                System.Reflection.BindingFlags.GetProperty,
                Nothing,
                source,
                Nothing,
                System.Globalization.CultureInfo.InvariantCulture)
            If value IsNot Nothing Then target(jsonPropertyName) = value.ToString()
        Catch
        End Try
    End Sub

    Private Sub SetLateBoundProperty(target As Object, propertyName As String, value As Object)
        If target Is Nothing OrElse value Is Nothing Then Return
        Try
            target.GetType().InvokeMember(
                propertyName,
                System.Reflection.BindingFlags.SetProperty,
                Nothing,
                target,
                New Object() {value},
                System.Globalization.CultureInfo.InvariantCulture)
        Catch
        End Try
    End Sub

    Private Sub ApplyScriptFontPropertiesToFont(font As Word.Font, definition As JToken)
        If font Is Nothing OrElse definition Is Nothing Then Return
        Dim fontObject As Object = font

        If definition("nameAscii") IsNot Nothing Then SetLateBoundProperty(fontObject, "NameAscii", CStr(definition("nameAscii")))
        If definition("nameFarEast") IsNot Nothing Then SetLateBoundProperty(fontObject, "NameFarEast", CStr(definition("nameFarEast")))
        If definition("nameBi") IsNot Nothing Then SetLateBoundProperty(fontObject, "NameBi", CStr(definition("nameBi")))
        If definition("nameOther") IsNot Nothing Then SetLateBoundProperty(fontObject, "NameOther", CStr(definition("nameOther")))

        If definition("languageId") IsNot Nothing Then
            SetLateBoundProperty(fontObject, "LanguageID", ParseWdLanguageId(definition("languageId")))
        End If
        If definition("languageIdFarEast") IsNot Nothing Then
            SetLateBoundProperty(fontObject, "LanguageIDFarEast", ParseWdLanguageId(definition("languageIdFarEast")))
        End If
        If definition("languageIdOther") IsNot Nothing Then
            SetLateBoundProperty(fontObject, "LanguageIDOther", ParseWdLanguageId(definition("languageIdOther")))
        End If
    End Sub

    Private Function ParseWdLanguageId(token As JToken) As Object
        If token Is Nothing Then Return Nothing
        Dim text As String = CStr(token)
        Dim numericValue As Integer
        If Integer.TryParse(text, numericValue) Then Return CType(numericValue, WdLanguageID)

        If text.StartsWith("wd", StringComparison.OrdinalIgnoreCase) Then
            Try
                Return CType([Enum].Parse(GetType(WdLanguageID), text, True), WdLanguageID)
            Catch
            End Try
        End If

        Try
            Dim culture As System.Globalization.CultureInfo = System.Globalization.CultureInfo.GetCultureInfo(text)
            Return CType(culture.LCID, WdLanguageID)
        Catch
            Return Nothing
        End Try
    End Function

    Private Function GetLateBoundPropertyString(source As Object, propertyName As String) As String
        If source Is Nothing Then Return ""
        Try
            Dim value As Object = source.GetType().InvokeMember(
                propertyName,
                System.Reflection.BindingFlags.GetProperty,
                Nothing,
                source,
                Nothing,
                System.Globalization.CultureInfo.InvariantCulture)
            Return If(value IsNot Nothing, value.ToString(), "")
        Catch
            Return ""
        End Try
    End Function

    Private Function ApplyScriptFontPropertiesIfDifferent(font As Word.Font, definition As JToken) As Boolean
        If font Is Nothing OrElse definition Is Nothing Then Return False
        Dim changed As Boolean = False
        Dim fontObject As Object = font

        Dim propertyMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
            {"nameAscii", "NameAscii"},
            {"nameFarEast", "NameFarEast"},
            {"nameBi", "NameBi"},
            {"nameOther", "NameOther"},
            {"languageId", "LanguageID"},
            {"languageIdFarEast", "LanguageIDFarEast"},
            {"languageIdOther", "LanguageIDOther"}
        }

        For Each mapping As KeyValuePair(Of String, String) In propertyMap
            If definition(mapping.Key) Is Nothing Then Continue For
            Dim desired As Object = If(mapping.Key.StartsWith("languageId", StringComparison.OrdinalIgnoreCase),
                                       ParseWdLanguageId(definition(mapping.Key)),
                                       CObj(CStr(definition(mapping.Key))))
            If desired Is Nothing Then Continue For

            Dim current As String = GetLateBoundPropertyString(fontObject, mapping.Value)
            If Not String.Equals(current, desired.ToString(), StringComparison.OrdinalIgnoreCase) Then
                SetLateBoundProperty(fontObject, mapping.Value, desired)
                changed = True
            End If
        Next

        Return changed
    End Function

    Private Function HasScriptFontDifference(font As Word.Font, definition As JToken) As Boolean
        If font Is Nothing OrElse definition Is Nothing Then Return False
        Dim fontObject As Object = font
        Dim propertyMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
            {"nameAscii", "NameAscii"},
            {"nameFarEast", "NameFarEast"},
            {"nameBi", "NameBi"},
            {"nameOther", "NameOther"},
            {"languageId", "LanguageID"},
            {"languageIdFarEast", "LanguageIDFarEast"},
            {"languageIdOther", "LanguageIDOther"}
        }

        For Each mapping As KeyValuePair(Of String, String) In propertyMap
            If definition(mapping.Key) Is Nothing Then Continue For
            Dim desired As Object = If(mapping.Key.StartsWith("languageId", StringComparison.OrdinalIgnoreCase),
                                       ParseWdLanguageId(definition(mapping.Key)),
                                       CObj(CStr(definition(mapping.Key))))
            If desired Is Nothing Then Continue For
            If Not String.Equals(GetLateBoundPropertyString(fontObject, mapping.Value),
                                 desired.ToString(),
                                 StringComparison.OrdinalIgnoreCase) Then Return True
        Next

        Return False
    End Function

    Private Sub ApplyDocumentDefaultsToDocument(doc As Word.Document, documentDefaults As JObject)
        If doc Is Nothing OrElse documentDefaults Is Nothing Then Return

        Try
            Dim policy As JObject = TryCast(documentDefaults("applyOnImport"), JObject)
            Dim applyDefaultTabStop As Boolean = True
            Dim applyFontsToNormal As Boolean = False
            Dim applyLanguagesToNormal As Boolean = False

            If policy IsNot Nothing Then
                If policy("defaultTabStop") IsNot Nothing Then applyDefaultTabStop = CBool(policy("defaultTabStop"))
                If policy("fontDefaultsToNormalStyle") IsNot Nothing Then applyFontsToNormal = CBool(policy("fontDefaultsToNormalStyle"))
                If policy("languageDefaultsToNormalStyle") IsNot Nothing Then applyLanguagesToNormal = CBool(policy("languageDefaultsToNormalStyle"))
            End If

            If applyDefaultTabStop AndAlso documentDefaults("defaultTabStop") IsNot Nothing Then
                Try
                    doc.DefaultTabStop = CSng(documentDefaults("defaultTabStop"))
                Catch ex As System.Exception
                    Debug.WriteLine($"[DocStyle] Could not apply default tab stop: {ex.Message}")
                End Try
            End If

            If Not applyFontsToNormal AndAlso Not applyLanguagesToNormal Then Return

            Dim normalStyle As Word.Style = Nothing
            Try
                normalStyle = doc.Styles(WdBuiltinStyle.wdStyleNormal)
            Catch
            End Try
            If normalStyle Is Nothing Then Return

            If applyFontsToNormal Then
                Dim fontDefaults As JObject = TryCast(documentDefaults("fontDefaults"), JObject)
                If fontDefaults IsNot Nothing Then
                    Dim translated As New JObject()
                    If fontDefaults("fontAscii") IsNot Nothing Then translated("nameAscii") = fontDefaults("fontAscii").DeepClone()
                    If fontDefaults("fontHighAnsi") IsNot Nothing Then translated("nameOther") = fontDefaults("fontHighAnsi").DeepClone()
                    If fontDefaults("fontEastAsia") IsNot Nothing Then translated("nameFarEast") = fontDefaults("fontEastAsia").DeepClone()
                    If fontDefaults("fontComplexScript") IsNot Nothing Then translated("nameBi") = fontDefaults("fontComplexScript").DeepClone()
                    ApplyScriptFontPropertiesToFont(normalStyle.Font, translated)
                End If
            End If

            If applyLanguagesToNormal Then
                Dim languageDefaults As JObject = TryCast(documentDefaults("languageDefaults"), JObject)
                If languageDefaults IsNot Nothing Then
                    Dim translated As New JObject()
                    If languageDefaults("latin") IsNot Nothing Then translated("languageId") = languageDefaults("latin").DeepClone()
                    If languageDefaults("eastAsia") IsNot Nothing Then translated("languageIdFarEast") = languageDefaults("eastAsia").DeepClone()
                    If languageDefaults("bidi") IsNot Nothing Then translated("languageIdOther") = languageDefaults("bidi").DeepClone()
                    ApplyScriptFontPropertiesToFont(normalStyle.Font, translated)
                End If
            End If

        Catch ex As System.Exception
            Debug.WriteLine($"[DocStyle] Could not apply document defaults: {ex.Message}")
        End Try
    End Sub

    Private Function ParseStyleType(styleDef As JObject) As WdStyleType
        Dim styleType As String = ""
        If styleDef IsNot Nothing Then
            If styleDef("styleType") IsNot Nothing Then
                styleType = CStr(styleDef("styleType"))
            ElseIf styleDef("styleTypeOpenXml") IsNot Nothing Then
                styleType = CStr(styleDef("styleTypeOpenXml"))
            End If
        End If

        Select Case styleType.Trim().ToLowerInvariant()
            Case "wdstyletypecharacter", "character"
                Return WdStyleType.wdStyleTypeCharacter
            Case "wdstyletypetable", "table"
                Return WdStyleType.wdStyleTypeTable
            Case "wdstyletypelist", "numbering", "list"
                Return WdStyleType.wdStyleTypeList
            Case "wdstyletypelinked", "linked"
                Return WdStyleType.wdStyleTypeParagraph
            Case Else
                Return WdStyleType.wdStyleTypeParagraph
        End Select
    End Function

    Private Function ApplyStyleRelationshipMetadata(doc As Word.Document,
                                                    style As Word.Style,
                                                    styleDef As JObject) As Boolean
        If doc Is Nothing OrElse style Is Nothing OrElse styleDef Is Nothing Then Return False
        Dim changed As Boolean = False

        Dim baseStyleName As String = ""
        If styleDef("basedOnStyleName") IsNot Nothing Then
            baseStyleName = CStr(styleDef("basedOnStyleName"))
        ElseIf styleDef("baseStyleHierarchy") IsNot Nothing AndAlso TypeOf styleDef("baseStyleHierarchy") Is JArray Then
            Dim hierarchy As JArray = CType(styleDef("baseStyleHierarchy"), JArray)
            If hierarchy.Count > 0 Then baseStyleName = CStr(hierarchy(0))
        End If

        If Not String.IsNullOrWhiteSpace(baseStyleName) Then
            Try
                Dim desiredBase As Word.Style = doc.Styles(baseStyleName)
                Dim currentBaseName As String = ""
                Try
                    Dim currentBase As Word.Style = TryCast(style.BaseStyle, Word.Style)
                    If currentBase IsNot Nothing Then currentBaseName = currentBase.NameLocal
                Catch
                End Try

                If Not String.Equals(currentBaseName, desiredBase.NameLocal, StringComparison.OrdinalIgnoreCase) Then
                    style.BaseStyle = desiredBase
                    changed = True
                End If
            Catch ex As System.Exception
                Debug.WriteLine($"[DocStyle] Base style '{baseStyleName}' for '{style.NameLocal}' is unavailable: {ex.Message}")
            End Try
        End If

        Dim nextStyleName As String = ""
        If styleDef("nextStyleName") IsNot Nothing Then
            nextStyleName = CStr(styleDef("nextStyleName"))
        ElseIf styleDef("nextParagraphStyle") IsNot Nothing Then
            nextStyleName = CStr(styleDef("nextParagraphStyle"))
        End If

        If Not String.IsNullOrWhiteSpace(nextStyleName) AndAlso style.Type = WdStyleType.wdStyleTypeParagraph Then
            Try
                Dim desiredNext As Word.Style = doc.Styles(nextStyleName)
                Dim currentNextName As String = ""
                Try
                    Dim currentNext As Word.Style = TryCast(style.NextParagraphStyle, Word.Style)
                    If currentNext IsNot Nothing Then currentNextName = currentNext.NameLocal
                Catch
                End Try

                If Not String.Equals(currentNextName, desiredNext.NameLocal, StringComparison.OrdinalIgnoreCase) Then
                    style.NextParagraphStyle = desiredNext
                    changed = True
                End If
            Catch ex As System.Exception
                Debug.WriteLine($"[DocStyle] Next style '{nextStyleName}' for '{style.NameLocal}' is unavailable: {ex.Message}")
            End Try
        End If

        Try
            Dim desiredQuickStyle As Boolean = If(styleDef("quickFormat") IsNot Nothing, CBool(styleDef("quickFormat")), True)
            If style.QuickStyle <> desiredQuickStyle Then
                style.QuickStyle = desiredQuickStyle
                changed = True
            End If
        Catch
        End Try

        If styleDef("uiPriority") IsNot Nothing Then
            Try
                Dim desiredPriority As Integer = CInt(styleDef("uiPriority"))
                If style.Priority <> desiredPriority Then
                    style.Priority = desiredPriority
                    changed = True
                End If
            Catch
            End Try
        End If

        Return changed
    End Function


    Private Function OrderStyleDefinitionsBaseFirst(wdStyleDefinitions As JObject) As List(Of JProperty)
        Dim ordered As New List(Of JProperty)()
        If wdStyleDefinitions Is Nothing Then Return ordered

        Dim propertiesByName As New Dictionary(Of String, JProperty)(StringComparer.OrdinalIgnoreCase)
        For Each propertyItem As JProperty In wdStyleDefinitions.Properties()
            propertiesByName(propertyItem.Name) = propertyItem
        Next

        Dim visiting As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim visited As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each propertyItem As JProperty In wdStyleDefinitions.Properties()
            AddStyleDefinitionBaseFirst(propertyItem.Name, propertiesByName, visiting, visited, ordered)
        Next

        Return ordered
    End Function

    Private Sub AddStyleDefinitionBaseFirst(styleName As String,
                                            propertiesByName As Dictionary(Of String, JProperty),
                                            visiting As HashSet(Of String),
                                            visited As HashSet(Of String),
                                            ordered As List(Of JProperty))
        If String.IsNullOrWhiteSpace(styleName) OrElse visited.Contains(styleName) Then Return
        If Not visiting.Add(styleName) Then
            Debug.WriteLine($"[DocStyle] Cyclic basedOn relationship detected at '{styleName}'.")
            Return
        End If

        Dim propertyItem As JProperty = Nothing
        If propertiesByName.TryGetValue(styleName, propertyItem) Then
            Dim styleDef As JObject = TryCast(propertyItem.Value, JObject)
            Dim baseStyleName As String = ""
            If styleDef IsNot Nothing Then
                If styleDef("basedOnStyleName") IsNot Nothing Then
                    baseStyleName = CStr(styleDef("basedOnStyleName"))
                ElseIf styleDef("baseStyleHierarchy") IsNot Nothing AndAlso TypeOf styleDef("baseStyleHierarchy") Is JArray Then
                    Dim hierarchy As JArray = CType(styleDef("baseStyleHierarchy"), JArray)
                    If hierarchy.Count > 0 Then baseStyleName = CStr(hierarchy(0))
                End If
            End If

            If Not String.IsNullOrWhiteSpace(baseStyleName) AndAlso propertiesByName.ContainsKey(baseStyleName) Then
                AddStyleDefinitionBaseFirst(baseStyleName, propertiesByName, visiting, visited, ordered)
            End If

            If visited.Add(styleName) Then ordered.Add(propertyItem)
        End If

        visiting.Remove(styleName)
    End Sub

    Private Sub HydrateListFormatFromNumberingDefinitions(styleDef As JObject,
                                                          numberingDefinitions As JObject)
        If styleDef Is Nothing OrElse numberingDefinitions Is Nothing Then Return
        Dim listFormat As JObject = TryCast(styleDef("listFormat"), JObject)
        If listFormat Is Nothing Then Return

        Dim sharedKey As String = GetSharedListInstanceKey(listFormat)
        If String.IsNullOrWhiteSpace(sharedKey) Then Return
        Dim definition As JObject = TryCast(numberingDefinitions(sharedKey), JObject)
        If definition Is Nothing Then Return

        Dim copiedFields As String() = {
            "outlineNumbered", "numberingStyleId", "numberingStyleName",
            "templateFingerprintV2", "sourceNumId", "abstractNumId"
        }
        For Each fieldName As String In copiedFields
            If listFormat(fieldName) Is Nothing AndAlso definition(fieldName) IsNot Nothing Then
                listFormat(fieldName) = definition(fieldName).DeepClone()
            End If
        Next

        If (listFormat("levels") Is Nothing OrElse Not TypeOf listFormat("levels") Is JArray OrElse CType(listFormat("levels"), JArray).Count = 0) AndAlso
           definition("levels") IsNot Nothing Then
            listFormat("levels") = definition("levels").DeepClone()
        ElseIf TypeOf listFormat("levels") Is JArray AndAlso TypeOf definition("levels") Is JArray Then
            MergeNumberingLevelMetadata(CType(listFormat("levels"), JArray), CType(definition("levels"), JArray))
        End If
    End Sub

    Private Function GetListTemplateCacheKey(listFormatDef As JToken) As String
        If listFormatDef Is Nothing Then Return ""

        Dim binding As JObject = TryCast(listFormatDef("numberingBinding"), JObject)
        Dim sharedKey As String = ""
        If binding IsNot Nothing AndAlso binding("sharedListInstanceKey") IsNot Nothing Then
            sharedKey = CStr(binding("sharedListInstanceKey"))
        ElseIf listFormatDef("sharedListInstanceKey") IsNot Nothing Then
            sharedKey = CStr(listFormatDef("sharedListInstanceKey"))
        End If
        If Not String.IsNullOrWhiteSpace(sharedKey) Then Return "S:" & sharedKey

        Dim isOutline As Boolean = If(listFormatDef("outlineNumbered") IsNot Nothing, CBool(listFormatDef("outlineNumbered")), False)
        Dim fingerprint As String = If(listFormatDef("templateFingerprintV2") IsNot Nothing,
                                       CStr(listFormatDef("templateFingerprintV2")),
                                       If(listFormatDef("templateFingerprint") IsNot Nothing,
                                          CStr(listFormatDef("templateFingerprint")), ""))
        Return $"{If(isOutline, "O", "N")}:{fingerprint}"
    End Function

    Private Function GetSharedListInstanceKey(listFormatDef As JToken) As String
        If listFormatDef Is Nothing Then Return ""
        Dim binding As JObject = TryCast(listFormatDef("numberingBinding"), JObject)
        If binding IsNot Nothing AndAlso binding("sharedListInstanceKey") IsNot Nothing Then
            Return CStr(binding("sharedListInstanceKey"))
        End If
        If listFormatDef("sharedListInstanceKey") IsNot Nothing Then Return CStr(listFormatDef("sharedListInstanceKey"))
        Return ""
    End Function

    Private Function TryGetCachedListTemplate(listFormatDef As JToken) As Word.ListTemplate
        Dim cacheKey As String = GetListTemplateCacheKey(listFormatDef)
        If String.IsNullOrWhiteSpace(cacheKey) Then Return Nothing
        Dim result As Word.ListTemplate = Nothing
        If _sharedListTemplates.TryGetValue(cacheKey, result) Then Return result
        Return Nothing
    End Function

    Private Sub ConfigureListTemplateFromDefinition(listTemplate As Word.ListTemplate,
                                                    listFormatDef As JToken)
        If listTemplate Is Nothing OrElse listFormatDef Is Nothing OrElse listFormatDef("levels") Is Nothing Then Return

        For Each levelDef As JObject In CType(listFormatDef("levels"), JArray)
            Try
                Dim levelNumber As Integer = CInt(levelDef("level"))
                If levelNumber < 1 OrElse levelNumber > listTemplate.ListLevels.Count Then Continue For

                Dim level As Word.ListLevel = listTemplate.ListLevels(levelNumber)
                Dim numberStyleText As String = If(levelDef("numberStyle") IsNot Nothing,
                                                   CStr(levelDef("numberStyle")),
                                                   "wdListNumberStyleArabic")
                Dim numberStyle As WdListNumberStyle = ParseNumberStyle(numberStyleText)
                Dim isBullet As Boolean = (numberStyle = WdListNumberStyle.wdListNumberStyleBullet) OrElse
                                          numberStyleText.IndexOf("bullet", StringComparison.OrdinalIgnoreCase) >= 0

                If isBullet Then
                    Dim bulletFontName As String = If(levelDef("bulletFont") IsNot Nothing, CStr(levelDef("bulletFont")), "Symbol")
                    Dim bulletCharCode As Integer = If(levelDef("bulletCharCode") IsNot Nothing, CInt(levelDef("bulletCharCode")), &H2022)
                    Try : level.Font.Reset() : Catch : End Try
                    Try : level.Font.Name = bulletFontName : Catch : End Try
                    Try : level.NumberStyle = WdListNumberStyle.wdListNumberStyleBullet : Catch : End Try
                    Try : level.NumberFormat = ChrW(bulletCharCode) : Catch : End Try
                Else
                    Try : level.NumberStyle = numberStyle : Catch : End Try
                    If levelDef("numberFormat") IsNot Nothing Then
                        Try : level.NumberFormat = CStr(levelDef("numberFormat")) : Catch : End Try
                    End If
                End If

                If levelDef("numberPosition") IsNot Nothing AndAlso levelDef("numberPosition").Type <> JTokenType.Null Then
                    level.NumberPosition = CSng(levelDef("numberPosition"))
                End If
                If levelDef("textPosition") IsNot Nothing AndAlso levelDef("textPosition").Type <> JTokenType.Null Then
                    level.TextPosition = CSng(levelDef("textPosition"))
                End If

                Dim tabPositionIsDefined As Boolean =
                    levelDef("tabPosition") IsNot Nothing AndAlso
                    levelDef("tabPosition").Type <> JTokenType.Null AndAlso
                    (levelDef("tabPositionState") Is Nothing OrElse
                     Not String.Equals(CStr(levelDef("tabPositionState")), "wdUndefined", StringComparison.OrdinalIgnoreCase))
                If tabPositionIsDefined Then
                    Dim tabPosition As Single = CSng(levelDef("tabPosition"))
                    If Not IsWordUndefinedPosition(tabPosition) Then level.TabPosition = tabPosition
                End If

                If levelDef("startAt") IsNot Nothing Then level.StartAt = CInt(levelDef("startAt"))
                If levelDef("alignment") IsNot Nothing Then level.Alignment = ParseListLevelAlignment(CStr(levelDef("alignment")))
                If levelDef("trailingCharacter") IsNot Nothing Then
                    level.TrailingCharacter = ParseTrailingCharacter(CStr(levelDef("trailingCharacter")))
                End If

                If levelDef("numberRunFormatting") IsNot Nothing Then
                    ApplyNumberRunFormatting(level.Font, levelDef("numberRunFormatting"))
                End If

            Catch ex As System.Exception
                Debug.WriteLine($"[DocStyle] Error configuring list level: {ex.Message}")
            End Try
        Next
    End Sub

    Private Sub ApplyNumberRunFormatting(font As Word.Font, definition As JToken)
        If font Is Nothing OrElse definition Is Nothing Then Return
        Try
            If definition("name") IsNot Nothing Then font.Name = CStr(definition("name"))
            If definition("fontAscii") IsNot Nothing Then SetLateBoundProperty(font, "NameAscii", CStr(definition("fontAscii")))
            If definition("fontHighAnsi") IsNot Nothing Then SetLateBoundProperty(font, "NameOther", CStr(definition("fontHighAnsi")))
            If definition("fontEastAsia") IsNot Nothing Then SetLateBoundProperty(font, "NameFarEast", CStr(definition("fontEastAsia")))
            If definition("fontComplexScript") IsNot Nothing Then SetLateBoundProperty(font, "NameBi", CStr(definition("fontComplexScript")))
            If definition("size") IsNot Nothing Then font.Size = CSng(definition("size"))
            If definition("bold") IsNot Nothing Then font.Bold = If(CBool(definition("bold")), -1, 0)
            If definition("italic") IsNot Nothing Then font.Italic = If(CBool(definition("italic")), -1, 0)
            If definition("allCaps") IsNot Nothing Then font.AllCaps = If(CBool(definition("allCaps")), -1, 0)
            If definition("smallCaps") IsNot Nothing Then font.SmallCaps = If(CBool(definition("smallCaps")), -1, 0)
        Catch ex As System.Exception
            Debug.WriteLine($"[DocStyle] Error applying number run formatting: {ex.Message}")
        End Try
    End Sub

    Private Sub ApplyParagraphStyleLinksForListTemplate(doc As Word.Document,
                                                        listTemplate As Word.ListTemplate,
                                                        listFormatDef As JToken,
                                                        currentStyle As Word.Style)
        If doc Is Nothing OrElse listTemplate Is Nothing OrElse listFormatDef Is Nothing OrElse
           listFormatDef("levels") Is Nothing OrElse currentStyle Is Nothing Then Return

        ' Only touch the style currently being imported. The JSON still retains all pStyle links,
        ' but unrelated target-document styles are not modified unless they are also imported.
        For Each levelDef As JObject In CType(listFormatDef("levels"), JArray)
            Dim paragraphStyleName As String = If(levelDef("paragraphStyleName") IsNot Nothing,
                                                  CStr(levelDef("paragraphStyleName")), "")
            If String.IsNullOrWhiteSpace(paragraphStyleName) OrElse
               Not String.Equals(paragraphStyleName, currentStyle.NameLocal, StringComparison.OrdinalIgnoreCase) Then
                Continue For
            End If

            Try
                Dim levelNumber As Integer = CInt(levelDef("level"))
                currentStyle.LinkToListTemplate(listTemplate, levelNumber)
            Catch ex As System.Exception
                Debug.WriteLine($"[DocStyle] Could not link list level to paragraph style '{paragraphStyleName}': {ex.Message}")
            End Try
        Next
    End Sub

    Private Sub EnsureNumberingStyleLinkedToTemplate(doc As Word.Document,
                                                     listTemplate As Word.ListTemplate,
                                                     listFormatDef As JToken)
        If doc Is Nothing OrElse listTemplate Is Nothing OrElse listFormatDef Is Nothing Then Return
        Dim numberingStyleName As String = If(listFormatDef("numberingStyleName") IsNot Nothing,
                                              CStr(listFormatDef("numberingStyleName")), "")
        If String.IsNullOrWhiteSpace(numberingStyleName) Then Return

        Try
            Dim numberingStyle As Word.Style = Nothing
            Try
                numberingStyle = doc.Styles(numberingStyleName)
            Catch
                numberingStyle = doc.Styles.Add(numberingStyleName, WdStyleType.wdStyleTypeList)
            End Try
            If numberingStyle IsNot Nothing Then numberingStyle.LinkToListTemplate(listTemplate)
        Catch ex As System.Exception
            Debug.WriteLine($"[DocStyle] Could not create/link numbering style '{numberingStyleName}': {ex.Message}")
        End Try
    End Sub

    Private Function IsStyleMappedToListLevel(style As Word.Style,
                                              linkedLevel As Integer,
                                              listFormatDef As JToken) As Boolean
        If style Is Nothing OrElse listFormatDef Is Nothing OrElse listFormatDef("levels") Is Nothing Then Return False

        For Each levelDef As JObject In CType(listFormatDef("levels"), JArray)
            If levelDef("level") Is Nothing OrElse CInt(levelDef("level")) <> linkedLevel Then Continue For
            If levelDef("paragraphStyleName") IsNot Nothing AndAlso
               String.Equals(CStr(levelDef("paragraphStyleName")), style.NameLocal, StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If
        Next

        Return False
    End Function

    Private Function IsInheritedOnlyNumbering(listFormatDef As JToken) As Boolean
        If listFormatDef Is Nothing Then Return False
        Dim binding As JObject = TryCast(listFormatDef("numberingBinding"), JObject)
        If binding Is Nothing Then Return False
        Dim hasDirectLevel As Boolean = If(binding("hasDirectListLevel") IsNot Nothing, CBool(binding("hasDirectListLevel")), False)
        Dim hasDirectNumId As Boolean = If(binding("hasDirectNumId") IsNot Nothing, CBool(binding("hasDirectNumId")), False)
        Return Not hasDirectLevel AndAlso Not hasDirectNumId
    End Function

    Private Function TryGetStyleListTemplate(style As Word.Style) As Word.ListTemplate
        If style Is Nothing Then Return Nothing
        Try
            Return style.ListTemplate
        Catch
            Return Nothing
        End Try
    End Function

#End Region

    ''' <summary>
    ''' Applies all formatting sections from a style definition JSON object to a newly created Word style.
    ''' </summary>
    ''' <param name="doc">Active document that owns the created style.</param>
    ''' <param name="style">Style to update.</param>
    ''' <param name="styleDef">Style definition JSON object.</param>
    Private Sub ApplyAllFormattingToNewStyle(doc As Word.Document, style As Word.Style, styleDef As JObject)
        ' Apply paragraph formatting
        If styleDef("paragraphFormat") IsNot Nothing Then
            ApplyParagraphFormatToStyle(style, styleDef("paragraphFormat"))
        End If

        ' Apply tab stops
        If styleDef("tabStops") IsNot Nothing Then
            ApplyTabStopsToStyle(style, CType(styleDef("tabStops"), JArray))
        End If

        ' Apply font formatting
        If styleDef("fontFormat") IsNot Nothing Then
            ApplyFontFormatToStyle(style, styleDef("fontFormat"))
        End If

        ' Apply list formatting
        If styleDef("listFormat") IsNot Nothing Then
            ApplyListFormatToStyle(doc, style, styleDef("listFormat"))
        End If
    End Sub

    ''' <summary>
    ''' Compares a style's paragraph formatting to a desired JSON definition and applies differences.
    ''' </summary>
    ''' <param name="style">Style to update.</param>
    ''' <param name="pf">Desired paragraph-format JSON token.</param>
    ''' <returns>True if any changes were applied; otherwise False.</returns>
    Private Function ApplyParagraphFormatIfDifferent(style As Word.Style, pf As JToken) As Boolean
        Dim changesMade As Boolean = False
        Dim paraFormat As Word.ParagraphFormat = style.ParagraphFormat

        Try
            If pf("alignment") IsNot Nothing Then
                Dim desired As WdParagraphAlignment = ParseAlignment(CStr(pf("alignment")))
                If paraFormat.Alignment <> desired Then
                    paraFormat.Alignment = desired
                    changesMade = True
                End If
            End If

            If pf("leftIndent") IsNot Nothing Then
                Dim desired As Single = CSng(pf("leftIndent"))
                If Math.Abs(paraFormat.LeftIndent - desired) > 0.1 Then
                    paraFormat.LeftIndent = desired
                    changesMade = True
                End If
            End If

            If pf("rightIndent") IsNot Nothing Then
                Dim desired As Single = CSng(pf("rightIndent"))
                If Math.Abs(paraFormat.RightIndent - desired) > 0.1 Then
                    paraFormat.RightIndent = desired
                    changesMade = True
                End If
            End If

            If pf("firstLineIndent") IsNot Nothing Then
                Dim desired As Single = CSng(pf("firstLineIndent"))
                If Math.Abs(paraFormat.FirstLineIndent - desired) > 0.1 Then
                    paraFormat.FirstLineIndent = desired
                    changesMade = True
                End If
            End If

            If pf("spaceBefore") IsNot Nothing Then
                Dim desired As Single = CSng(pf("spaceBefore"))
                If Math.Abs(paraFormat.SpaceBefore - desired) > 0.1 Then
                    paraFormat.SpaceBefore = desired
                    changesMade = True
                End If
            End If

            If pf("spaceAfter") IsNot Nothing Then
                Dim desired As Single = CSng(pf("spaceAfter"))
                If Math.Abs(paraFormat.SpaceAfter - desired) > 0.1 Then
                    paraFormat.SpaceAfter = desired
                    changesMade = True
                End If
            End If

            If pf("lineSpacing") IsNot Nothing Then
                Dim desired As Single = CSng(pf("lineSpacing"))
                If Math.Abs(paraFormat.LineSpacing - desired) > 0.1 Then
                    paraFormat.LineSpacing = desired
                    changesMade = True
                End If
            End If

            If pf("lineSpacingRule") IsNot Nothing Then
                Dim desired As WdLineSpacing = ParseLineSpacingRule(CStr(pf("lineSpacingRule")))
                If paraFormat.LineSpacingRule <> desired Then
                    paraFormat.LineSpacingRule = desired
                    changesMade = True
                End If
            End If

            If pf("keepTogether") IsNot Nothing Then
                Dim desired As Integer = CInt(pf("keepTogether"))
                If paraFormat.KeepTogether <> desired Then
                    paraFormat.KeepTogether = desired
                    changesMade = True
                End If
            End If

            If pf("keepWithNext") IsNot Nothing Then
                Dim desired As Integer = CInt(pf("keepWithNext"))
                If paraFormat.KeepWithNext <> desired Then
                    paraFormat.KeepWithNext = desired
                    changesMade = True
                End If
            End If

            If pf("pageBreakBefore") IsNot Nothing Then
                Dim desired As Integer = CInt(pf("pageBreakBefore"))
                If paraFormat.PageBreakBefore <> desired Then
                    paraFormat.PageBreakBefore = desired
                    changesMade = True
                End If
            End If

            If pf("widowControl") IsNot Nothing Then
                Dim desired As Integer = CInt(pf("widowControl"))
                If paraFormat.WidowControl <> desired Then
                    paraFormat.WidowControl = desired
                    changesMade = True
                End If
            End If

            If pf("outlineLevel") IsNot Nothing Then
                Dim desired As WdOutlineLevel = ParseOutlineLevel(CStr(pf("outlineLevel")))
                If paraFormat.OutlineLevel <> desired Then
                    paraFormat.OutlineLevel = desired
                    changesMade = True
                End If
            End If

        Catch ex As System.Exception
            Debug.WriteLine($"Error comparing/applying paragraph format: {ex.Message}")
        End Try

        Return changesMade
    End Function

    ''' <summary>
    ''' Applies paragraph formatting from a JSON definition to a style (no comparison).
    ''' </summary>
    ''' <param name="style">Style to update.</param>
    ''' <param name="pf">Desired paragraph-format JSON token.</param>
    Private Sub ApplyParagraphFormatToStyle(style As Word.Style, pf As JToken)
        Dim paraFormat As Word.ParagraphFormat = style.ParagraphFormat

        Try
            If pf("alignment") IsNot Nothing Then paraFormat.Alignment = ParseAlignment(CStr(pf("alignment")))
            If pf("leftIndent") IsNot Nothing Then paraFormat.LeftIndent = CSng(pf("leftIndent"))
            If pf("rightIndent") IsNot Nothing Then paraFormat.RightIndent = CSng(pf("rightIndent"))
            If pf("firstLineIndent") IsNot Nothing Then paraFormat.FirstLineIndent = CSng(pf("firstLineIndent"))
            If pf("spaceBefore") IsNot Nothing Then paraFormat.SpaceBefore = CSng(pf("spaceBefore"))
            If pf("spaceAfter") IsNot Nothing Then paraFormat.SpaceAfter = CSng(pf("spaceAfter"))
            If pf("lineSpacing") IsNot Nothing Then paraFormat.LineSpacing = CSng(pf("lineSpacing"))
            If pf("lineSpacingRule") IsNot Nothing Then paraFormat.LineSpacingRule = ParseLineSpacingRule(CStr(pf("lineSpacingRule")))
            If pf("keepTogether") IsNot Nothing Then paraFormat.KeepTogether = CInt(pf("keepTogether"))
            If pf("keepWithNext") IsNot Nothing Then paraFormat.KeepWithNext = CInt(pf("keepWithNext"))
            If pf("pageBreakBefore") IsNot Nothing Then paraFormat.PageBreakBefore = CInt(pf("pageBreakBefore"))
            If pf("widowControl") IsNot Nothing Then paraFormat.WidowControl = CInt(pf("widowControl"))
            If pf("outlineLevel") IsNot Nothing Then paraFormat.OutlineLevel = ParseOutlineLevel(CStr(pf("outlineLevel")))
        Catch ex As System.Exception
            Debug.WriteLine($"Error applying paragraph format: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Compares a style's tab stops to a desired JSON definition and applies differences.
    ''' </summary>
    ''' <param name="style">Style to update.</param>
    ''' <param name="tabsDef">Desired tab-stop JSON array.</param>
    ''' <returns>True if any changes were applied; otherwise False.</returns>
    Private Function ApplyTabStopsIfDifferent(style As Word.Style, tabsDef As JArray) As Boolean
        Try
            ' Build current tab stops signature
            Dim currentTabs As New List(Of String)()
            For Each tabStop As Word.TabStop In style.ParagraphFormat.TabStops
                Dim isCustom As Boolean = True
                Try : isCustom = tabStop.CustomTab : Catch : End Try
                If isCustom Then currentTabs.Add($"{Math.Round(tabStop.Position, 1)}:{tabStop.Alignment}:{tabStop.Leader}")
            Next

            ' Build desired tab stops signature
            Dim desiredTabs As New List(Of String)()
            For Each tabDef As JObject In tabsDef
                Dim pos As Single = CSng(tabDef("position"))
                Dim align As String = If(tabDef("alignment") IsNot Nothing, CStr(tabDef("alignment")), "wdAlignTabLeft")
                Dim leader As String = If(tabDef("leader") IsNot Nothing, CStr(tabDef("leader")), "wdTabLeaderSpaces")
                desiredTabs.Add($"{Math.Round(pos, 1)}:{ParseTabAlignment(align)}:{ParseTabLeader(leader)}")
            Next

            ' Compare
            If currentTabs.Count = desiredTabs.Count AndAlso
               currentTabs.SequenceEqual(desiredTabs) Then
                Return False ' No changes needed
            End If

            ' Apply changes
            ApplyTabStopsToStyle(style, tabsDef)
            Return True

        Catch ex As System.Exception
            Debug.WriteLine($"Error comparing tab stops: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Applies tab stops from a JSON definition to a style (no comparison).
    ''' </summary>
    ''' <param name="style">Style to update.</param>
    ''' <param name="tabsDef">Desired tab-stop JSON array.</param>
    Private Sub ApplyTabStopsToStyle(style As Word.Style, tabsDef As JArray)
        Try
            style.ParagraphFormat.TabStops.ClearAll()
            For Each tabDef As JObject In tabsDef
                Dim position As Single = CSng(tabDef("position"))
                Dim alignment As WdTabAlignment = ParseTabAlignment(If(tabDef("alignment") IsNot Nothing, CStr(tabDef("alignment")), "wdAlignTabLeft"))
                Dim leader As WdTabLeader = ParseTabLeader(If(tabDef("leader") IsNot Nothing, CStr(tabDef("leader")), "wdTabLeaderSpaces"))
                style.ParagraphFormat.TabStops.Add(position, alignment, leader)
            Next
        Catch ex As System.Exception
            Debug.WriteLine($"Error applying tab stops: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Compares a style's font formatting to a desired JSON definition and applies differences.
    ''' </summary>
    ''' <param name="style">Style to update.</param>
    ''' <param name="ff">Desired font-format JSON token.</param>
    ''' <returns>True if any changes were applied; otherwise False.</returns>
    Private Function ApplyFontFormatIfDifferent(style As Word.Style, ff As JToken) As Boolean
        Dim changesMade As Boolean = False
        Dim font As Word.Font = style.Font

        Try
            If ff("name") IsNot Nothing Then
                Dim desired As String = CStr(ff("name"))
                If font.Name <> desired Then
                    font.Name = desired
                    changesMade = True
                End If
            End If

            If ff("size") IsNot Nothing Then
                Dim desired As Single = CSng(ff("size"))
                If Math.Abs(font.Size - desired) > 0.1 Then
                    font.Size = desired
                    changesMade = True
                End If
            End If

            If ff("bold") IsNot Nothing Then
                Dim desired As Integer = If(CBool(ff("bold")), -1, 0)
                If font.Bold <> desired Then
                    font.Bold = desired
                    changesMade = True
                End If
            End If

            If ff("italic") IsNot Nothing Then
                Dim desired As Integer = If(CBool(ff("italic")), -1, 0)
                If font.Italic <> desired Then
                    font.Italic = desired
                    changesMade = True
                End If
            End If

            If ff("underline") IsNot Nothing Then
                Dim desired As WdUnderline = ParseUnderline(CStr(ff("underline")))
                If font.Underline <> desired Then
                    font.Underline = desired
                    changesMade = True
                End If
            End If

            If ff("allCaps") IsNot Nothing Then
                Dim desired As Integer = If(CBool(ff("allCaps")), -1, 0)
                If font.AllCaps <> desired Then
                    font.AllCaps = desired
                    changesMade = True
                End If
            End If

            If ff("smallCaps") IsNot Nothing Then
                Dim desired As Integer = If(CBool(ff("smallCaps")), -1, 0)
                If font.SmallCaps <> desired Then
                    font.SmallCaps = desired
                    changesMade = True
                End If
            End If

            If ff("strikeThrough") IsNot Nothing Then
                Dim desired As Integer = If(CBool(ff("strikeThrough")), -1, 0)
                If font.StrikeThrough <> desired Then
                    font.StrikeThrough = desired
                    changesMade = True
                End If
            End If

            If ff("doubleStrikeThrough") IsNot Nothing Then
                Dim desired As Integer = If(CBool(ff("doubleStrikeThrough")), -1, 0)
                If font.DoubleStrikeThrough <> desired Then
                    font.DoubleStrikeThrough = desired
                    changesMade = True
                End If
            End If

            If ff("subscript") IsNot Nothing Then
                Dim desired As Integer = If(CBool(ff("subscript")), -1, 0)
                If font.Subscript <> desired Then
                    font.Subscript = desired
                    changesMade = True
                End If
            End If

            If ff("superscript") IsNot Nothing Then
                Dim desired As Integer = If(CBool(ff("superscript")), -1, 0)
                If font.Superscript <> desired Then
                    font.Superscript = desired
                    changesMade = True
                End If
            End If

            If ff("colorRGB") IsNot Nothing Then
                Dim desired As WdColor = ParseColorFromRGB(CStr(ff("colorRGB")))
                If font.Color <> desired Then
                    font.Color = desired
                    changesMade = True
                End If
            End If

            If ff("scaling") IsNot Nothing Then
                Try
                    Dim desired As Integer = CInt(ff("scaling"))
                    If font.Scaling <> desired Then
                        font.Scaling = desired
                        changesMade = True
                    End If
                Catch
                End Try
            End If

            If ff("spacing") IsNot Nothing Then
                Try
                    Dim desired As Single = CSng(ff("spacing"))
                    If Math.Abs(font.Spacing - desired) > 0.1 Then
                        font.Spacing = desired
                        changesMade = True
                    End If
                Catch
                End Try
            End If

            If ff("position") IsNot Nothing Then
                Try
                    Dim desired As Single = CSng(ff("position"))
                    If Math.Abs(font.Position - desired) > 0.1 Then
                        font.Position = desired
                        changesMade = True
                    End If
                Catch
                End Try
            End If

            If ff("kerning") IsNot Nothing Then
                Try
                    Dim desired As Single = CSng(ff("kerning"))
                    If Math.Abs(font.Kerning - desired) > 0.1 Then
                        font.Kerning = desired
                        changesMade = True
                    End If
                Catch
                End Try
            End If

            changesMade = ApplyScriptFontPropertiesIfDifferent(font, ff) OrElse changesMade

        Catch ex As System.Exception
            Debug.WriteLine($"Error comparing/applying font format: {ex.Message}")
        End Try

        Return changesMade
    End Function

    ''' <summary>
    ''' Applies font formatting from a JSON definition to a style (no comparison).
    ''' </summary>
    ''' <param name="style">Style to update.</param>
    ''' <param name="ff">Desired font-format JSON token.</param>
    Private Sub ApplyFontFormatToStyle(style As Word.Style, ff As JToken)
        Dim font As Word.Font = style.Font

        Try
            If ff("name") IsNot Nothing Then font.Name = CStr(ff("name"))
            If ff("size") IsNot Nothing Then font.Size = CSng(ff("size"))
            If ff("bold") IsNot Nothing Then font.Bold = If(CBool(ff("bold")), -1, 0)
            If ff("italic") IsNot Nothing Then font.Italic = If(CBool(ff("italic")), -1, 0)
            If ff("underline") IsNot Nothing Then font.Underline = ParseUnderline(CStr(ff("underline")))
            If ff("allCaps") IsNot Nothing Then font.AllCaps = If(CBool(ff("allCaps")), -1, 0)
            If ff("smallCaps") IsNot Nothing Then font.SmallCaps = If(CBool(ff("smallCaps")), -1, 0)
            If ff("strikeThrough") IsNot Nothing Then font.StrikeThrough = If(CBool(ff("strikeThrough")), -1, 0)
            If ff("doubleStrikeThrough") IsNot Nothing Then font.DoubleStrikeThrough = If(CBool(ff("doubleStrikeThrough")), -1, 0)
            If ff("subscript") IsNot Nothing Then font.Subscript = If(CBool(ff("subscript")), -1, 0)
            If ff("superscript") IsNot Nothing Then font.Superscript = If(CBool(ff("superscript")), -1, 0)
            If ff("colorRGB") IsNot Nothing Then font.Color = ParseColorFromRGB(CStr(ff("colorRGB")))
            Try
                If ff("scaling") IsNot Nothing Then font.Scaling = CInt(ff("scaling"))
                If ff("spacing") IsNot Nothing Then font.Spacing = CSng(ff("spacing"))
                If ff("position") IsNot Nothing Then font.Position = CSng(ff("position"))
                If ff("kerning") IsNot Nothing Then font.Kerning = CSng(ff("kerning"))
            Catch
            End Try
            ApplyScriptFontPropertiesToFont(font, ff)
        Catch ex As System.Exception
            Debug.WriteLine($"Error applying font format: {ex.Message}")
        End Try
    End Sub




    ''' <summary>
    ''' Applies Word style definitions from a template to the document by creating missing styles and updating existing ones.
    ''' </summary>
    ''' <param name="doc">Active document whose <c>Styles</c> collection is updated.</param>
    ''' <param name="wdStyleDefinitions">JSON object containing style definitions keyed by Word style name.</param>
    ''' <summary>
    ''' Creates/updates Word styles in three passes: create all styles, restore style relationships,
    ''' then apply formatting and shared numbering definitions. Older JSON templates remain supported.
    ''' </summary>
    Private Sub ApplyWdStyleDefinitionsToDocument(doc As Word.Document,
                                                  wdStyleDefinitions As JObject,
                                                  Optional numberingDefinitions As JObject = Nothing)
        If doc Is Nothing OrElse wdStyleDefinitions Is Nothing Then Return

        _sharedListTemplates.Clear()

        Dim allStylesExist As Boolean = True
        Dim existingStyles As New List(Of String)()
        Dim missingStyles As New List(Of String)()

        For Each propertyItem As JProperty In wdStyleDefinitions.Properties()
            Try
                Dim existingStyle As Word.Style = doc.Styles(propertyItem.Name)
                existingStyles.Add(propertyItem.Name)
            Catch
                allStylesExist = False
                missingStyles.Add(propertyItem.Name)
            End Try
        Next

        If allStylesExist AndAlso existingStyles.Count > 0 Then
            Dim confirmResult As Integer = ShowCustomYesNoBox(
                $"All {existingStyles.Count} Word styles from the template already exist in this document." & vbCrLf & vbCrLf &
                "Updating existing styles may cause formatting issues if applied multiple times (also, for the same reason, do not apply DocStyle templates twice)." & vbCrLf & vbCrLf &
                "Do you want to update the existing styles anyway?",
                "Yes, update styles", "No, skip style update", $"{AN} - Style Update")

            If confirmResult <> 1 Then
                Debug.WriteLine($"[DocStyle] User skipped style update - all {existingStyles.Count} styles already exist")
                Return
            End If
        End If

        Dim createdStyles As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim stylesCreatedOrUpdated As New List(Of String)()
        Dim stylesSkipped As New List(Of String)()

        ' Pass 1: create every missing style before assigning basedOn/next relationships.
        For Each propertyItem As JProperty In wdStyleDefinitions.Properties()
            Dim styleName As String = propertyItem.Name
            Dim styleDef As JObject = TryCast(propertyItem.Value, JObject)
            If styleDef Is Nothing Then Continue For

            Try
                Dim existingStyle As Word.Style = Nothing
                Try
                    existingStyle = doc.Styles(styleName)
                Catch
                End Try
                If existingStyle IsNot Nothing Then Continue For

                Dim styleType As WdStyleType = ParseStyleType(styleDef)
                Dim createdStyle As Word.Style = doc.Styles.Add(styleName, styleType)
                createdStyles.Add(styleName)
                Debug.WriteLine($"[DocStyle] Created new wdStyle: {styleName} ({styleType})")
            Catch ex As System.Exception
                Debug.WriteLine($"[DocStyle] Could not create wdStyle '{styleName}': {ex.Message}")
            End Try
        Next

        Dim orderedProperties As List(Of JProperty) = OrderStyleDefinitionsBaseFirst(wdStyleDefinitions)

        ' Pass 2: apply style inheritance, next-style, Quick Style and priority metadata.
        For Each propertyItem As JProperty In orderedProperties
            Dim styleName As String = propertyItem.Name
            Dim styleDef As JObject = TryCast(propertyItem.Value, JObject)
            If styleDef Is Nothing Then Continue For

            Try
                Dim style As Word.Style = doc.Styles(styleName)
                ApplyStyleRelationshipMetadata(doc, style, styleDef)
            Catch ex As System.Exception
                Debug.WriteLine($"[DocStyle] Could not apply style relationships for '{styleName}': {ex.Message}")
            End Try
        Next

        ' Pass 3: apply visible formatting and shared list templates, base styles first.
        For Each propertyItem As JProperty In orderedProperties
            Dim styleName As String = propertyItem.Name
            Dim styleDef As JObject = TryCast(propertyItem.Value, JObject)
            If styleDef Is Nothing Then Continue For

            Try
                HydrateListFormatFromNumberingDefinitions(styleDef, numberingDefinitions)

                Dim style As Word.Style = doc.Styles(styleName)
                If style Is Nothing Then Continue For

                If createdStyles.Contains(styleName) Then
                    ApplyAllFormattingToNewStyle(doc, style, styleDef)
                    stylesCreatedOrUpdated.Add($"{styleName} (created)")
                    Continue For
                End If

                Dim changesMade As Boolean = ApplyStyleRelationshipMetadata(doc, style, styleDef)

                If styleDef("paragraphFormat") IsNot Nothing AndAlso style.Type = WdStyleType.wdStyleTypeParagraph Then
                    changesMade = ApplyParagraphFormatIfDifferent(style, styleDef("paragraphFormat")) OrElse changesMade
                End If

                If styleDef("tabStops") IsNot Nothing AndAlso style.Type = WdStyleType.wdStyleTypeParagraph Then
                    changesMade = ApplyTabStopsIfDifferent(style, CType(styleDef("tabStops"), JArray)) OrElse changesMade
                End If

                If styleDef("fontFormat") IsNot Nothing Then
                    changesMade = ApplyFontFormatIfDifferent(style, styleDef("fontFormat")) OrElse changesMade
                End If

                If styleDef("listFormat") IsNot Nothing Then
                    Dim listFormatDef As JToken = styleDef("listFormat")
                    Dim templateWantsList As Boolean =
                        listFormatDef("hasListTemplate") IsNot Nothing AndAlso CBool(listFormatDef("hasListTemplate"))

                    Dim currentTemplate As Word.ListTemplate = TryGetStyleListTemplate(style)
                    Dim styleHasListTemplate As Boolean = (currentTemplate IsNot Nothing)

                    If templateWantsList Then
                        ApplyListFormatToStyle(doc, style, listFormatDef)
                        changesMade = True
                    ElseIf styleHasListTemplate AndAlso Not style.BuiltIn Then
                        Try
                            style.LinkToListTemplate(Nothing)
                            changesMade = True
                        Catch ex As System.Exception
                            Debug.WriteLine($"[DocStyle] Could not remove list from '{styleName}': {ex.Message}")
                        End Try
                    End If
                End If

                If changesMade Then
                    stylesCreatedOrUpdated.Add($"{styleName} (updated)")
                Else
                    stylesSkipped.Add(styleName)
                End If

            Catch ex As System.Exception
                Debug.WriteLine($"[DocStyle] Error creating/updating wdStyle '{styleName}': {ex.Message}")
            End Try
        Next

        ' Deliberately keep _sharedListTemplates alive until paragraph application completes.
        If stylesCreatedOrUpdated.Count > 0 Then
            Debug.WriteLine($"[DocStyle] Word styles processed: {String.Join(", ", stylesCreatedOrUpdated)}")
        End If
        If stylesSkipped.Count > 0 Then
            Debug.WriteLine($"[DocStyle] Word styles unchanged: {String.Join(", ", stylesSkipped)}")
        End If
    End Sub
    ''' <summary>
    ''' Applies user style formatting (paragraph, font, list, and tab stops) from the template to a paragraph.
    ''' Formatting is applied as overrides after the Word style assignment.
    ''' </summary>
    ''' <param name="para">Paragraph to format.</param>
    ''' <param name="userStyleDef">User style JSON definition containing formatting overrides.</param>
    ''' <summary>
    ''' Applies paragraph/font/list/tab overrides after assigning the Word style.
    ''' Shared numbering templates are preferred; Word defaults remain the compatibility fallback.
    ''' </summary>
    Private Sub ApplyUserStyleFormattingFromTemplate(para As Word.Paragraph, userStyleDef As JObject)
        If para Is Nothing OrElse userStyleDef Is Nothing Then Return

        Try
            Dim paragraphFormatDef As JToken = userStyleDef("paragraphFormatting")

            If paragraphFormatDef IsNot Nothing Then
                If paragraphFormatDef("alignment") IsNot Nothing Then para.Alignment = ParseAlignment(CStr(paragraphFormatDef("alignment")))
                If paragraphFormatDef("leftIndent") IsNot Nothing Then para.LeftIndent = CSng(paragraphFormatDef("leftIndent"))
                If paragraphFormatDef("rightIndent") IsNot Nothing Then para.RightIndent = CSng(paragraphFormatDef("rightIndent"))
                If paragraphFormatDef("firstLineIndent") IsNot Nothing Then para.FirstLineIndent = CSng(paragraphFormatDef("firstLineIndent"))
                If paragraphFormatDef("spaceBefore") IsNot Nothing Then para.SpaceBefore = CSng(paragraphFormatDef("spaceBefore"))
                If paragraphFormatDef("spaceAfter") IsNot Nothing Then para.SpaceAfter = CSng(paragraphFormatDef("spaceAfter"))
                If paragraphFormatDef("lineSpacing") IsNot Nothing Then para.LineSpacing = CSng(paragraphFormatDef("lineSpacing"))
                If paragraphFormatDef("lineSpacingRule") IsNot Nothing Then para.Format.LineSpacingRule = ParseLineSpacingRule(CStr(paragraphFormatDef("lineSpacingRule")))
                If paragraphFormatDef("keepTogether") IsNot Nothing Then para.KeepTogether = CInt(paragraphFormatDef("keepTogether"))
                If paragraphFormatDef("keepWithNext") IsNot Nothing Then para.KeepWithNext = CInt(paragraphFormatDef("keepWithNext"))
            End If

            If userStyleDef("fontFormatting") IsNot Nothing Then
                Dim fontFormatDef As JToken = userStyleDef("fontFormatting")
                Dim range As Word.Range = para.Range
                If fontFormatDef("fontName") IsNot Nothing AndAlso CStr(fontFormatDef("fontName")) <> "mixed" Then range.Font.Name = CStr(fontFormatDef("fontName"))
                If fontFormatDef("fontSize") IsNot Nothing AndAlso CStr(fontFormatDef("fontSize")) <> "mixed" Then range.Font.Size = CSng(fontFormatDef("fontSize"))
                If fontFormatDef("bold") IsNot Nothing AndAlso fontFormatDef("bold").Type = JTokenType.Boolean Then range.Font.Bold = If(CBool(fontFormatDef("bold")), -1, 0)
                If fontFormatDef("italic") IsNot Nothing AndAlso fontFormatDef("italic").Type = JTokenType.Boolean Then range.Font.Italic = If(CBool(fontFormatDef("italic")), -1, 0)
                If fontFormatDef("allCaps") IsNot Nothing AndAlso fontFormatDef("allCaps").Type = JTokenType.Boolean Then range.Font.AllCaps = If(CBool(fontFormatDef("allCaps")), -1, 0)
                If fontFormatDef("smallCaps") IsNot Nothing AndAlso fontFormatDef("smallCaps").Type = JTokenType.Boolean Then range.Font.SmallCaps = If(CBool(fontFormatDef("smallCaps")), -1, 0)
                If fontFormatDef("underline") IsNot Nothing AndAlso CStr(fontFormatDef("underline")) <> "mixed" Then range.Font.Underline = ParseUnderline(CStr(fontFormatDef("underline")))
                ApplyScriptFontPropertiesToFont(range.Font, fontFormatDef)
            End If

            Dim listOperationChangedFormatting As Boolean = False
            If userStyleDef("listFormatting") IsNot Nothing Then
                Dim listFormatDef As JToken = userStyleDef("listFormatting")
                Dim templateHasList As Boolean = If(listFormatDef("hasList") IsNot Nothing,
                                                    CBool(listFormatDef("hasList")), False)

                Dim currentHasList As Boolean = False
                Dim currentLevel As Integer = 0
                Try
                    currentHasList = para.Range.ListFormat.ListType <> WdListType.wdListNoNumbering
                    If currentHasList Then currentLevel = para.Range.ListFormat.ListLevelNumber
                Catch
                End Try

                If Not templateHasList AndAlso currentHasList Then
                    para.Range.ListFormat.RemoveNumbers()
                    listOperationChangedFormatting = True
                ElseIf templateHasList Then
                    Dim desiredLevel As Integer = 1
                    If listFormatDef("listLevelNumber") IsNot Nothing Then
                        desiredLevel = CInt(listFormatDef("listLevelNumber"))
                    Else
                        Dim binding As JObject = TryCast(listFormatDef("numberingBinding"), JObject)
                        If binding IsNot Nothing AndAlso binding("effectiveListLevel") IsNot Nothing Then
                            desiredLevel = CInt(binding("effectiveListLevel"))
                        End If
                    End If
                    desiredLevel = Math.Max(1, Math.Min(9, desiredLevel))

                    If Not currentHasList OrElse currentLevel <> desiredLevel Then
                        Dim listTemplate As Word.ListTemplate = Nothing
                        Try
                            Dim paragraphStyle As Word.Style = TryCast(para.Style, Word.Style)
                            listTemplate = TryGetStyleListTemplate(paragraphStyle)
                        Catch
                        End Try
                        If listTemplate Is Nothing Then listTemplate = TryGetCachedListTemplate(listFormatDef)

                        If listTemplate IsNot Nothing Then
                            para.Range.ListFormat.ApplyListTemplateWithLevel(
                                ListTemplate:=listTemplate,
                                ContinuePreviousList:=True,
                                ApplyTo:=WdListApplyTo.wdListApplyToSelection,
                                DefaultListBehavior:=WdDefaultListBehavior.wdWord10ListBehavior,
                                ApplyLevel:=desiredLevel)
                            listOperationChangedFormatting = True
                        ElseIf Not currentHasList Then
                            ' Backward-compatible fallback for old JSON files or disabled style-definition import.
                            Dim listType As String = If(listFormatDef("listType") IsNot Nothing, CStr(listFormatDef("listType")), "")
                            If listType.IndexOf("Bullet", StringComparison.OrdinalIgnoreCase) >= 0 Then
                                para.Range.ListFormat.ApplyBulletDefault()
                            ElseIf listType.IndexOf("Outline", StringComparison.OrdinalIgnoreCase) >= 0 Then
                                para.Range.ListFormat.ApplyOutlineNumberDefault()
                            ElseIf listType.IndexOf("Number", StringComparison.OrdinalIgnoreCase) >= 0 Then
                                para.Range.ListFormat.ApplyNumberDefault()
                            End If
                            listOperationChangedFormatting = True
                        End If
                    End If
                End If
            End If

            ' List operations may alter paragraph indents. Restore the extracted effective values.
            If listOperationChangedFormatting AndAlso paragraphFormatDef IsNot Nothing Then
                If paragraphFormatDef("leftIndent") IsNot Nothing Then para.LeftIndent = CSng(paragraphFormatDef("leftIndent"))
                If paragraphFormatDef("rightIndent") IsNot Nothing Then para.RightIndent = CSng(paragraphFormatDef("rightIndent"))
                If paragraphFormatDef("firstLineIndent") IsNot Nothing Then para.FirstLineIndent = CSng(paragraphFormatDef("firstLineIndent"))
            End If

            If userStyleDef("tabStops") IsNot Nothing Then
                Try
                    para.TabStops.ClearAll()
                    For Each tabDef As JObject In CType(userStyleDef("tabStops"), JArray)
                        Dim position As Single = CSng(tabDef("pos"))
                        Dim alignment As WdTabAlignment = ParseTabAlignment(
                            "wdAlignTab" & If(tabDef("align") IsNot Nothing, CStr(tabDef("align")), "Left"))
                        Dim leader As WdTabLeader = ParseTabLeader(
                            "wdTabLeader" & If(tabDef("leader") IsNot Nothing, CStr(tabDef("leader")), "Spaces"))
                        para.TabStops.Add(position, alignment, leader)
                    Next
                Catch
                End Try
            End If

        Catch ex As System.Exception
            Debug.WriteLine($"Error applying user style formatting: {ex.Message}")
        End Try
    End Sub
    ''' <summary>
    ''' Applies list formatting to a style, including list template creation and linking of the template at the configured list level.
    ''' Shared templates are cached to reduce redundant template creation within the current run.
    ''' </summary>
    ''' <param name="doc">Active document that owns the created list templates.</param>
    ''' <param name="style">Style to link to a list template.</param>
    ''' <param name="listFormatDef">List-format JSON definition.</param>
    ''' <summary>
    ''' Applies or reuses a list template. New templates use sharedListInstanceKey (normally source numId)
    ''' before falling back to the legacy fingerprint cache key.
    ''' </summary>
    Private Sub ApplyListFormatToStyle(doc As Word.Document,
                                       style As Word.Style,
                                       listFormatDef As JToken)
        If doc Is Nothing OrElse style Is Nothing OrElse listFormatDef Is Nothing Then Return
        If listFormatDef("hasListTemplate") Is Nothing OrElse Not CBool(listFormatDef("hasListTemplate")) Then Return

        Try
            Dim isOutline As Boolean = If(listFormatDef("outlineNumbered") IsNot Nothing,
                                          CBool(listFormatDef("outlineNumbered")), False)
            Dim linkedLevel As Integer = If(listFormatDef("linkedLevel") IsNot Nothing,
                                            CInt(listFormatDef("linkedLevel")), 1)
            linkedLevel = Math.Max(1, Math.Min(9, linkedLevel))

            Dim cacheKey As String = GetListTemplateCacheKey(listFormatDef)
            Dim desiredLegacyFingerprint As String = If(listFormatDef("templateFingerprint") IsNot Nothing,
                                                         CStr(listFormatDef("templateFingerprint")), "")
            Dim listTemplate As Word.ListTemplate = Nothing

            If Not String.IsNullOrWhiteSpace(cacheKey) Then
                _sharedListTemplates.TryGetValue(cacheKey, listTemplate)
            End If

            ' Seed the shared cache from an already-correct target template when possible.
            If listTemplate Is Nothing Then
                Dim currentTemplate As Word.ListTemplate = TryGetStyleListTemplate(style)
                If currentTemplate IsNot Nothing Then
                    Dim currentFingerprint As String = ComputeListTemplateFingerprint(currentTemplate)
                    If String.IsNullOrWhiteSpace(desiredLegacyFingerprint) OrElse
                       String.Equals(currentFingerprint, desiredLegacyFingerprint, StringComparison.Ordinal) Then
                        listTemplate = currentTemplate
                        If Not String.IsNullOrWhiteSpace(cacheKey) Then _sharedListTemplates(cacheKey) = listTemplate
                    End If
                End If
            End If

            If listTemplate Is Nothing Then
                listTemplate = doc.ListTemplates.Add(OutlineNumbered:=isOutline)
                ConfigureListTemplateFromDefinition(listTemplate, listFormatDef)
                If Not String.IsNullOrWhiteSpace(cacheKey) Then _sharedListTemplates(cacheKey) = listTemplate
                Debug.WriteLine($"[DocStyle] Created shared list template for '{style.NameLocal}' (key={cacheKey})")
            Else
                ' Existing target templates are not destructively rewritten when the legacy fingerprint matches.
                ' A cached template created in this run was already configured at creation time.
                Debug.WriteLine($"[DocStyle] Reusing shared list template for '{style.NameLocal}' (key={cacheKey})")
            End If

            ApplyParagraphStyleLinksForListTemplate(doc, listTemplate, listFormatDef, style)
            EnsureNumberingStyleLinkedToTemplate(doc, listTemplate, listFormatDef)

            Dim inheritedOnly As Boolean = IsInheritedOnlyNumbering(listFormatDef)
            Dim mappedByLevel As Boolean = IsStyleMappedToListLevel(style, linkedLevel, listFormatDef)

            If inheritedOnly Then
                ' Preserve inheritance where possible. If the target base chain does not yield a list,
                ' fall back to a direct link so older/incomplete templates still work.
                If TryGetStyleListTemplate(style) Is Nothing Then
                    style.LinkToListTemplate(listTemplate, linkedLevel)
                End If
            ElseIf Not mappedByLevel Then
                style.LinkToListTemplate(listTemplate, linkedLevel)
            End If

        Catch ex As System.Exception
            Debug.WriteLine($"Error applying list template to style '{style.NameLocal}': {ex.Message}")
        End Try
    End Sub
    ''' <summary>
    ''' Parses a line spacing rule name into a <see cref="WdLineSpacing"/> enum value.
    ''' </summary>
    ''' <param name="rule">Line-spacing rule name (possibly including the Word enum prefix).</param>
    ''' <returns>Parsed line spacing rule.</returns>
    Private Function ParseLineSpacingRule(rule As String) As WdLineSpacing
        If String.IsNullOrWhiteSpace(rule) Then Return WdLineSpacing.wdLineSpaceSingle

        Select Case rule.ToLower().Replace("wdlinespace", "").Replace("_", "")
            Case "single" : Return WdLineSpacing.wdLineSpaceSingle
            Case "1pt5" : Return WdLineSpacing.wdLineSpace1pt5
            Case "double" : Return WdLineSpacing.wdLineSpaceDouble
            Case "atleast" : Return WdLineSpacing.wdLineSpaceAtLeast
            Case "exactly" : Return WdLineSpacing.wdLineSpaceExactly
            Case "multiple" : Return WdLineSpacing.wdLineSpaceMultiple
            Case Else : Return WdLineSpacing.wdLineSpaceSingle
        End Select
    End Function

    ''' <summary>
    ''' Parses an outline level name into a <see cref="WdOutlineLevel"/> enum value.
    ''' </summary>
    ''' <param name="level">Outline level name (possibly including the Word enum prefix).</param>
    ''' <returns>Parsed outline level.</returns>
    Private Function ParseOutlineLevel(level As String) As WdOutlineLevel
        If String.IsNullOrWhiteSpace(level) Then Return WdOutlineLevel.wdOutlineLevelBodyText

        Select Case level.ToLower().Replace("wdoutlinelevel", "").Replace("_", "")
            Case "1" : Return WdOutlineLevel.wdOutlineLevel1
            Case "2" : Return WdOutlineLevel.wdOutlineLevel2
            Case "3" : Return WdOutlineLevel.wdOutlineLevel3
            Case "4" : Return WdOutlineLevel.wdOutlineLevel4
            Case "5" : Return WdOutlineLevel.wdOutlineLevel5
            Case "6" : Return WdOutlineLevel.wdOutlineLevel6
            Case "7" : Return WdOutlineLevel.wdOutlineLevel7
            Case "8" : Return WdOutlineLevel.wdOutlineLevel8
            Case "9" : Return WdOutlineLevel.wdOutlineLevel9
            Case "bodytext" : Return WdOutlineLevel.wdOutlineLevelBodyText
            Case Else : Return WdOutlineLevel.wdOutlineLevelBodyText
        End Select
    End Function

    ''' <summary>
    ''' Parses a tab alignment name into a <see cref="WdTabAlignment"/> enum value.
    ''' </summary>
    ''' <param name="alignment">Tab alignment name (possibly including the Word enum prefix).</param>
    ''' <returns>Parsed tab alignment.</returns>
    Private Function ParseTabAlignment(alignment As String) As WdTabAlignment
        If String.IsNullOrWhiteSpace(alignment) Then Return WdTabAlignment.wdAlignTabLeft

        Select Case alignment.ToLower().Replace("wdaligntab", "").Replace("_", "")
            Case "left" : Return WdTabAlignment.wdAlignTabLeft
            Case "center" : Return WdTabAlignment.wdAlignTabCenter
            Case "right" : Return WdTabAlignment.wdAlignTabRight
            Case "decimal" : Return WdTabAlignment.wdAlignTabDecimal
            Case "bar" : Return WdTabAlignment.wdAlignTabBar
            Case "list" : Return WdTabAlignment.wdAlignTabList
            Case Else : Return WdTabAlignment.wdAlignTabLeft
        End Select
    End Function

    ''' <summary>
    ''' Parses a tab leader name into a <see cref="WdTabLeader"/> enum value.
    ''' </summary>
    ''' <param name="leader">Tab leader name (possibly including the Word enum prefix).</param>
    ''' <returns>Parsed tab leader.</returns>
    Private Function ParseTabLeader(leader As String) As WdTabLeader
        If String.IsNullOrWhiteSpace(leader) Then Return WdTabLeader.wdTabLeaderSpaces

        Select Case leader.ToLower().Replace("wdtableader", "").Replace("_", "")
            Case "spaces" : Return WdTabLeader.wdTabLeaderSpaces
            Case "dots" : Return WdTabLeader.wdTabLeaderDots
            Case "dashes" : Return WdTabLeader.wdTabLeaderDashes
            Case "lines" : Return WdTabLeader.wdTabLeaderLines
            Case "heavy" : Return WdTabLeader.wdTabLeaderHeavy
            Case "middledot" : Return WdTabLeader.wdTabLeaderMiddleDot
            Case Else : Return WdTabLeader.wdTabLeaderSpaces
        End Select
    End Function

    ''' <summary>
    ''' Parses an underline name into a <see cref="WdUnderline"/> enum value.
    ''' </summary>
    ''' <param name="underline">Underline name (possibly including the Word enum prefix).</param>
    ''' <returns>Parsed underline value.</returns>
    Private Function ParseUnderline(underline As String) As WdUnderline
        If String.IsNullOrWhiteSpace(underline) Then Return WdUnderline.wdUnderlineNone

        Select Case underline.ToLower().Replace("wdunderline", "").Replace("_", "")
            Case "none" : Return WdUnderline.wdUnderlineNone
            Case "single" : Return WdUnderline.wdUnderlineSingle
            Case "words" : Return WdUnderline.wdUnderlineWords
            Case "double" : Return WdUnderline.wdUnderlineDouble
            Case "dotted" : Return WdUnderline.wdUnderlineDotted
            Case "thick" : Return WdUnderline.wdUnderlineThick
            Case "dash" : Return WdUnderline.wdUnderlineDash
            Case "dotdash" : Return WdUnderline.wdUnderlineDotDash
            Case "dotdotdash" : Return WdUnderline.wdUnderlineDotDotDash
            Case "wavy" : Return WdUnderline.wdUnderlineWavy
            Case "wavyheavy" : Return WdUnderline.wdUnderlineWavyHeavy
            Case "wavydouble" : Return WdUnderline.wdUnderlineWavyDouble
            Case "dashlong" : Return WdUnderline.wdUnderlineDashLong
            Case "dashheavy" : Return WdUnderline.wdUnderlineDashHeavy
            Case "dotdashheavy" : Return WdUnderline.wdUnderlineDotDashHeavy
            Case "dotdotdashheavy" : Return WdUnderline.wdUnderlineDotDotDashHeavy
            Case "dashlongheavy" : Return WdUnderline.wdUnderlineDashLongHeavy
            Case Else : Return WdUnderline.wdUnderlineNone
        End Select
    End Function

    ''' <summary>
    ''' Parses a 6-digit RGB hex string (<c>#RRGGBB</c>) into a <see cref="WdColor"/> value.
    ''' </summary>
    ''' <param name="rgb">RGB hex string or <c>auto</c>/<c>unknown</c>.</param>
    ''' <returns>Parsed Word color value.</returns>
    Private Function ParseColorFromRGB(rgb As String) As WdColor
        If String.IsNullOrWhiteSpace(rgb) OrElse rgb = "auto" OrElse rgb = "unknown" Then
            Return WdColor.wdColorAutomatic
        End If

        Try
            rgb = rgb.TrimStart("#"c)
            If rgb.Length = 6 Then
                Dim r As Integer = System.Convert.ToInt32(rgb.Substring(0, 2), 16)
                Dim g As Integer = System.Convert.ToInt32(rgb.Substring(2, 2), 16)
                Dim b As Integer = System.Convert.ToInt32(rgb.Substring(4, 2), 16)
                Return CType(r + (g * 256) + (b * 65536), WdColor)
            End If
        Catch
        End Try

        Return WdColor.wdColorAutomatic
    End Function




    ''' <summary>
    ''' Parses a list level alignment name into a <see cref="WdListLevelAlignment"/> enum value.
    ''' </summary>
    ''' <param name="alignment">List level alignment name (possibly including the Word enum prefix).</param>
    ''' <returns>Parsed list level alignment value.</returns>
    Private Function ParseListLevelAlignment(alignment As String) As WdListLevelAlignment
        If String.IsNullOrWhiteSpace(alignment) Then Return WdListLevelAlignment.wdListLevelAlignLeft

        Select Case alignment.ToLower().Replace("wdlistlevelalign", "").Replace("_", "")
            Case "left" : Return WdListLevelAlignment.wdListLevelAlignLeft
            Case "center" : Return WdListLevelAlignment.wdListLevelAlignCenter
            Case "right" : Return WdListLevelAlignment.wdListLevelAlignRight
            Case Else : Return WdListLevelAlignment.wdListLevelAlignLeft
        End Select
    End Function

    ''' <summary>
    ''' Parses a trailing character name into a <see cref="WdTrailingCharacter"/> enum value.
    ''' </summary>
    ''' <param name="trailing">Trailing character name (possibly including the Word enum prefix).</param>
    ''' <returns>Parsed trailing character value.</returns>
    Private Function ParseTrailingCharacter(trailing As String) As WdTrailingCharacter
        If String.IsNullOrWhiteSpace(trailing) Then Return WdTrailingCharacter.wdTrailingTab

        Select Case trailing.ToLower().Replace("wdtrailing", "").Replace("_", "")
            Case "tab" : Return WdTrailingCharacter.wdTrailingTab
            Case "space" : Return WdTrailingCharacter.wdTrailingSpace
            Case "none" : Return WdTrailingCharacter.wdTrailingNone
            Case Else : Return WdTrailingCharacter.wdTrailingTab
        End Select
    End Function

    ''' <summary>
    ''' Parses a list number style name into a <see cref="WdListNumberStyle"/> enum value.
    ''' </summary>
    ''' <param name="style">List number style name (possibly including the Word enum prefix).</param>
    ''' <returns>Parsed list number style value.</returns>
    Private Function ParseNumberStyle(style As String) As WdListNumberStyle
        If String.IsNullOrWhiteSpace(style) Then Return WdListNumberStyle.wdListNumberStyleArabic

        Select Case style.ToLower().Replace("wdlistnumberstyle", "").Replace("_", "")
            Case "arabic" : Return WdListNumberStyle.wdListNumberStyleArabic
            Case "uppercaseroman" : Return WdListNumberStyle.wdListNumberStyleUppercaseRoman
            Case "lowercaseroman" : Return WdListNumberStyle.wdListNumberStyleLowercaseRoman
            Case "uppercaseletter" : Return WdListNumberStyle.wdListNumberStyleUppercaseLetter
            Case "lowercaseletter" : Return WdListNumberStyle.wdListNumberStyleLowercaseLetter
            Case "bullet" : Return WdListNumberStyle.wdListNumberStyleBullet
            Case "none" : Return WdListNumberStyle.wdListNumberStyleNone
            Case "ordinal" : Return WdListNumberStyle.wdListNumberStyleOrdinal
            Case "ordinaltext" : Return WdListNumberStyle.wdListNumberStyleOrdinalText
            Case "cardinaltext" : Return WdListNumberStyle.wdListNumberStyleCardinalText
            Case "legal" : Return WdListNumberStyle.wdListNumberStyleLegal
            Case "legalllz" : Return WdListNumberStyle.wdListNumberStyleLegalLZ
            Case "arabicfullwidth" : Return WdListNumberStyle.wdListNumberStyleArabicFullWidth
            Case "arabicllz" : Return WdListNumberStyle.wdListNumberStyleArabicLZ
            Case Else : Return WdListNumberStyle.wdListNumberStyleArabic
        End Select
    End Function
#End Region

#Region "Helper Classes and Functions"

    ''' <summary>
    ''' Removes empty paragraphs (containing only whitespace or paragraph marks) from the specified range.
    ''' Iterates in reverse order to avoid index shifting issues during deletion.
    ''' </summary>
    ''' <param name="targetRange">Range from which to remove empty paragraphs.</param>
    ''' <returns>Number of paragraphs removed.</returns>
    Private Function RemoveEmptyParagraphs(targetRange As Word.Range) As Integer
        Dim removedCount As Integer = 0

        Try
            ' Build a list of paragraph ranges to delete (iterate forward, delete in reverse)
            Dim parasToDelete As New List(Of Word.Range)()

            For Each para As Word.Paragraph In targetRange.Paragraphs
                Try
                    Dim text As String = para.Range.Text
                    ' Remove paragraph mark and check if remaining content is empty/whitespace
                    text = text.TrimEnd(vbCr, vbLf, ChrW(13), ChrW(10))

                    If String.IsNullOrWhiteSpace(text) Then
                        ' Don't delete if it's in a table cell (could break table structure)
                        Dim isInTable As Boolean = False
                        Try
                            isInTable = (para.Range.Tables.Count > 0 OrElse para.Range.Cells.Count > 0)
                        Catch
                        End Try

                        If Not isInTable Then
                            parasToDelete.Add(para.Range.Duplicate)
                        End If
                    End If
                Catch
                    ' Skip paragraphs that can't be processed
                End Try
            Next

            ' Delete in reverse order to preserve range integrity
            For i As Integer = parasToDelete.Count - 1 To 0 Step -1
                Try
                    parasToDelete(i).Delete()
                    removedCount += 1
                Catch ex As System.Exception
                    Debug.WriteLine($"[DocStyle] Could not delete empty paragraph: {ex.Message}")
                End Try
            Next

        Catch ex As System.Exception
            Debug.WriteLine($"[DocStyle] RemoveEmptyParagraphs error: {ex.Message}")
        End Try

        Return removedCount
    End Function


    ''' <summary>
    ''' Represents a style template file.
    ''' </summary>
    Private Class DocStyleTemplate
        Public Property DisplayName As String
        Public Property FilePath As String
        Public Property IsLocal As Boolean
    End Class

    ''' <summary>
    ''' Loads available style templates from configured paths, reading display names from JSON.
    ''' </summary>
    ''' <param name="pathGlobal">Directory containing shared templates.</param>
    ''' <param name="pathLocal">Directory containing personal templates.</param>
    ''' <returns>List of resolved template metadata objects.</returns>
    Private Function LoadStyleTemplates(pathGlobal As String, pathLocal As String) As List(Of DocStyleTemplate)
        Dim templates As New List(Of DocStyleTemplate)()

        ' Load from global path
        If Not String.IsNullOrWhiteSpace(pathGlobal) AndAlso Directory.Exists(pathGlobal) Then
            Try
                For Each f In Directory.GetFiles(pathGlobal, $"{AN2}-ds-*.json", SearchOption.TopDirectoryOnly)
                    Dim template As DocStyleTemplate = TryLoadTemplateMetadata(f, False)
                    If template IsNot Nothing Then
                        templates.Add(template)
                    End If
                Next
            Catch
            End Try
        End If

        ' Load from local path
        If Not String.IsNullOrWhiteSpace(pathLocal) AndAlso Directory.Exists(pathLocal) Then
            Try
                For Each f In Directory.GetFiles(pathLocal, $"{AN2}-ds-*.json", SearchOption.TopDirectoryOnly)
                    Dim template As DocStyleTemplate = TryLoadTemplateMetadata(f, True)
                    If template IsNot Nothing Then
                        templates.Add(template)
                    End If
                Next
            Catch
            End Try
        End If

        Return templates
    End Function

    ''' <summary>
    ''' Attempts to load template metadata from a JSON file.
    ''' Returns Nothing if the file cannot be parsed as valid JSON.
    ''' </summary>
    ''' <param name="filePath">Path to the template JSON file.</param>
    ''' <param name="isLocal">True when the file was discovered in the local template directory.</param>
    ''' <returns>Template metadata, or Nothing if unreadable/invalid.</returns>
    Private Function TryLoadTemplateMetadata(filePath As String, isLocal As Boolean) As DocStyleTemplate
        Try
            Dim jsonContent As String = File.ReadAllText(filePath, Encoding.UTF8)
            Dim jsonObj As JObject = JObject.Parse(jsonContent)

            ' Get display name from JSON, fall back to filename
            Dim displayName As String = ""
            If jsonObj("templateName") IsNot Nothing Then
                displayName = jsonObj("templateName").ToString()
            End If

            If String.IsNullOrWhiteSpace(displayName) Then
                displayName = Path.GetFileNameWithoutExtension(filePath).Replace($"{AN2}-ds-", "")
            End If

            Return New DocStyleTemplate With {
                .DisplayName = displayName,
                .FilePath = filePath,
                .IsLocal = isLocal
            }
        Catch
            Debug.WriteLine($"Skipping invalid template file: {filePath}")
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Extracts the JSON payload from an LLM response.
    ''' Supports markdown code blocks and raw JSON responses.
    ''' </summary>
    ''' <param name="response">Raw LLM response text.</param>
    ''' <returns>Extracted JSON string (array or object).</returns>
    Private Function ExtractJsonFromResponse(response As String) As String
        If String.IsNullOrWhiteSpace(response) Then Return "{}"

        ' Try to extract from markdown code block
        Dim match As Match = Regex.Match(response, "```(?:json)?\s*([\s\S]*?)```", RegexOptions.IgnoreCase)
        If match.Success Then
            Return match.Groups(1).Value.Trim()
        End If

        ' Try to find JSON array or object
        Dim startIdx As Integer = response.IndexOfAny(New Char() {"["c, "{"c})
        If startIdx >= 0 Then
            Return response.Substring(startIdx).Trim()
        End If

        Return response.Trim()
    End Function

    ''' <summary>
    ''' Parses an alignment name into a <see cref="WdParagraphAlignment"/> enum value.
    ''' </summary>
    ''' <param name="alignment">Alignment name (possibly including the Word enum prefix).</param>
    ''' <returns>Parsed paragraph alignment value.</returns>
    Private Function ParseAlignment(alignment As String) As WdParagraphAlignment
        If String.IsNullOrWhiteSpace(alignment) Then Return WdParagraphAlignment.wdAlignParagraphLeft

        Select Case alignment.ToLower().Replace("wdalignparagraph", "").Replace("_", "")
            Case "left" : Return WdParagraphAlignment.wdAlignParagraphLeft
            Case "center" : Return WdParagraphAlignment.wdAlignParagraphCenter
            Case "right" : Return WdParagraphAlignment.wdAlignParagraphRight
            Case "justify" : Return WdParagraphAlignment.wdAlignParagraphJustify
            Case Else : Return WdParagraphAlignment.wdAlignParagraphLeft
        End Select
    End Function

    ''' <summary>
    ''' Extracts paragraph formatting into a JSON object for template creation.
    ''' </summary>
    ''' <param name="para">Paragraph to inspect.</param>
    ''' <param name="paraRange">Duplicated paragraph range (for table/cell checks and other range-based info).</param>
    ''' <returns>JSON object containing paragraph formatting properties and optional error information.</returns>
    Private Function ExtractParagraphFormat(para As Word.Paragraph, paraRange As Word.Range) As JObject
        Dim pf As New JObject()
        Try
            pf("alignment") = para.Alignment.ToString()
            pf("leftIndent") = para.LeftIndent
            pf("rightIndent") = para.RightIndent
            pf("firstLineIndent") = para.FirstLineIndent
            pf("spaceBefore") = para.SpaceBefore
            pf("spaceAfter") = para.SpaceAfter
            pf("spaceBeforeAuto") = para.SpaceBeforeAuto
            pf("spaceAfterAuto") = para.SpaceAfterAuto
            pf("lineSpacing") = para.LineSpacing
            pf("lineSpacingRule") = para.LineSpacingRule.ToString()
            pf("keepTogether") = para.KeepTogether
            pf("keepWithNext") = para.KeepWithNext
            pf("pageBreakBefore") = para.PageBreakBefore
            pf("widowControl") = para.WidowControl
            pf("outlineLevel") = para.OutlineLevel.ToString()

            Try : pf("hyphenation") = para.Hyphenation : Catch : pf("hyphenation") = "default" : End Try

            Try
                Dim paraFormat As Object = para.Format
                pf("noSpaceBetweenParagraphsOfSameStyle") = paraFormat.NoSpaceBetweenParagraphsOfSameStyle
            Catch
            End Try

            Try
                Dim paraFormat As Object = para.Format
                pf("mirrorIndents") = paraFormat.MirrorIndents
            Catch
            End Try

            Try : pf("readingOrder") = para.ReadingOrder.ToString() : Catch : End Try

            Try
                Dim paraFormat As Object = para.Format
                pf("contextualSpacing") = paraFormat.ContextualSpacing
            Catch
            End Try

        Catch ex As System.Exception
            pf("error") = ex.Message
        End Try
        Return pf
    End Function

    ''' <summary>
    ''' Extracts character/font formatting into a JSON object for template creation.
    ''' Values that are mixed across the selection are returned as the string <c>mixed</c>.
    ''' </summary>
    ''' <param name="rng">Range to inspect.</param>
    ''' <returns>JSON object containing font formatting properties and optional error information.</returns>
    Private Function ExtractFontFormat(rng As Word.Range) As JObject
        Dim ff As New JObject()
        Try
            Dim font As Word.Font = rng.Font

            If font.Name <> CStr(WdConstants.wdUndefined) Then ff("fontName") = font.Name Else ff("fontName") = "mixed"
            If font.Size <> CSng(WdConstants.wdUndefined) Then ff("fontSize") = font.Size Else ff("fontSize") = "mixed"
            If font.Bold <> CInt(WdConstants.wdUndefined) Then ff("bold") = (font.Bold = -1) Else ff("bold") = "mixed"
            If font.Italic <> CInt(WdConstants.wdUndefined) Then ff("italic") = (font.Italic = -1) Else ff("italic") = "mixed"
            If font.Underline <> CType(WdConstants.wdUndefined, WdUnderline) Then ff("underline") = font.Underline.ToString() Else ff("underline") = "mixed"

            Try
                If font.UnderlineColor <> CType(WdConstants.wdUndefined, WdColor) Then
                    ff("underlineColor") = font.UnderlineColor.ToString()
                    ff("underlineColorRGB") = ColorToRGB(font.UnderlineColor)
                End If
            Catch
            End Try

            If font.StrikeThrough <> CInt(WdConstants.wdUndefined) Then ff("strikeThrough") = (font.StrikeThrough = -1) Else ff("strikeThrough") = "mixed"
            If font.DoubleStrikeThrough <> CInt(WdConstants.wdUndefined) Then ff("doubleStrikeThrough") = (font.DoubleStrikeThrough = -1) Else ff("doubleStrikeThrough") = "mixed"
            If font.Subscript <> CInt(WdConstants.wdUndefined) Then ff("subscript") = (font.Subscript = -1) Else ff("subscript") = "mixed"
            If font.Superscript <> CInt(WdConstants.wdUndefined) Then ff("superscript") = (font.Superscript = -1) Else ff("superscript") = "mixed"

            If font.Color <> CType(WdConstants.wdUndefined, WdColor) Then
                ff("color") = font.Color.ToString()
                ff("colorRGB") = ColorToRGB(font.Color)
            Else
                ff("color") = "mixed"
            End If

            Try
                If rng.HighlightColorIndex <> CType(WdConstants.wdUndefined, WdColorIndex) Then
                    ff("highlightColor") = rng.HighlightColorIndex.ToString()
                Else
                    ff("highlightColor") = "mixed"
                End If
            Catch
                ff("highlightColor") = "none"
            End Try

            If font.AllCaps <> CInt(WdConstants.wdUndefined) Then ff("allCaps") = (font.AllCaps = -1) Else ff("allCaps") = "mixed"
            If font.SmallCaps <> CInt(WdConstants.wdUndefined) Then ff("smallCaps") = (font.SmallCaps = -1) Else ff("smallCaps") = "mixed"

            Try
                ff("scaling") = font.Scaling
                ff("spacing") = font.Spacing
                ff("position") = font.Position
                ff("kerning") = font.Kerning
            Catch
            End Try

            Try
                Dim fontObj As Object = font
                ff("themeFont") = fontObj.ThemeFont.ToString()
            Catch
            End Try

            Try
                Dim fontObj As Object = font
                ff("themeColor") = fontObj.ThemeColor.ToString()
                ff("themeTint") = fontObj.TintAndShade
            Catch
            End Try

            AddScriptFontPropertiesToJson(font, ff)

        Catch ex As System.Exception
            ff("error") = ex.Message
        End Try
        Return ff
    End Function

    ''' <summary>
    ''' Extracts list formatting into a JSON object for template creation.
    ''' </summary>
    ''' <param name="rng">Range to inspect.</param>
    ''' <returns>JSON object describing list presence and current list level details when available.</returns>
    ''' <summary>
    ''' Extracts effective paragraph list formatting and adds the source style's shared numbering binding.
    ''' Legacy fields are retained for backward compatibility.
    ''' </summary>
    Private Function ExtractListFormat(rng As Word.Range,
                                       Optional openXmlContext As JObject = Nothing) As JObject
        Dim listResult As New JObject()
        If rng Is Nothing Then Return listResult

        Try
            Dim listFormat As Word.ListFormat = rng.ListFormat

            If listFormat.ListType = WdListType.wdListNoNumbering Then
                listResult("hasList") = False
                listResult("listType") = "none"
            Else
                listResult("hasList") = True
                listResult("listType") = listFormat.ListType.ToString()
                listResult("listLevelNumber") = listFormat.ListLevelNumber

                ' Kept for old consumers; appliedState is the semantically correct location.
                listResult("listString") = listFormat.ListString
                listResult("appliedState") = New JObject From {
                    {"listString", listFormat.ListString},
                    {"listLevelNumber", listFormat.ListLevelNumber}
                }

                Try
                    If listFormat.ListTemplate IsNot Nothing Then
                        Dim listTemplate As Word.ListTemplate = listFormat.ListTemplate
                        listResult("listTemplateOutlineNumbered") = listTemplate.OutlineNumbered

                        Dim level As Word.ListLevel = listTemplate.ListLevels(listFormat.ListLevelNumber)
                        Dim levelInfo As New JObject From {
                            {"numberFormat", level.NumberFormat},
                            {"numberStyle", level.NumberStyle.ToString()},
                            {"textPosition", level.TextPosition},
                            {"numberPosition", level.NumberPosition},
                            {"alignment", level.Alignment.ToString()},
                            {"startAt", level.StartAt}
                        }

                        If IsWordUndefinedPosition(level.TabPosition) Then
                            levelInfo("tabPosition") = JValue.CreateNull()
                            levelInfo("tabPositionState") = "wdUndefined"
                        Else
                            levelInfo("tabPosition") = level.TabPosition
                            levelInfo("tabPositionState") = "defined"
                        End If

                        listResult("currentLevelFormat") = levelInfo
                    End If
                Catch
                End Try
            End If

            Dim styleName As String = ""
            Try
                If rng.Paragraphs.Count > 0 Then
                    Dim style As Word.Style = TryCast(rng.Paragraphs(1).Style, Word.Style)
                    If style IsNot Nothing Then styleName = style.NameLocal
                End If
            Catch
            End Try

            Dim styleMeta As JObject = GetStyleOpenXmlMetadata(openXmlContext, styleName)
            If styleMeta IsNot Nothing Then
                If styleMeta("styleId") IsNot Nothing Then listResult("wdStyleId") = styleMeta("styleId").DeepClone()
                If styleMeta("numberingBinding") IsNot Nothing Then
                    listResult("numberingBinding") = styleMeta("numberingBinding").DeepClone()
                    If styleMeta("numberingBinding")("sharedListInstanceKey") IsNot Nothing Then
                        listResult("sharedListInstanceKey") = styleMeta("numberingBinding")("sharedListInstanceKey").DeepClone()
                    End If

                    Dim sharedKey As String = If(styleMeta("numberingBinding")("sharedListInstanceKey") IsNot Nothing,
                                                 CStr(styleMeta("numberingBinding")("sharedListInstanceKey")), "")
                    Dim numberingDefinitions As JObject = If(openXmlContext IsNot Nothing,
                                                              TryCast(openXmlContext("numberingDefinitions"), JObject), Nothing)
                    Dim numberingDefinition As JObject = If(numberingDefinitions IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(sharedKey),
                                                             TryCast(numberingDefinitions(sharedKey), JObject), Nothing)
                    If numberingDefinition IsNot Nothing Then
                        If numberingDefinition("numberingStyleId") IsNot Nothing Then listResult("numberingStyleId") = numberingDefinition("numberingStyleId").DeepClone()
                        If numberingDefinition("numberingStyleName") IsNot Nothing Then listResult("numberingStyleName") = numberingDefinition("numberingStyleName").DeepClone()

                        Dim currentLevelNumber As Integer = If(listResult("listLevelNumber") IsNot Nothing,
                                                              CInt(listResult("listLevelNumber")),
                                                              If(styleMeta("numberingBinding")("effectiveListLevel") IsNot Nothing,
                                                                 CInt(styleMeta("numberingBinding")("effectiveListLevel")), 1))
                        Dim sourceLevels As JArray = TryCast(numberingDefinition("levels"), JArray)
                        Dim sourceLevel As JObject = FindLevelDefinition(sourceLevels, currentLevelNumber)
                        Dim currentLevelFormat As JObject = TryCast(listResult("currentLevelFormat"), JObject)
                        If sourceLevel IsNot Nothing Then
                            If currentLevelFormat Is Nothing Then
                                currentLevelFormat = New JObject()
                                listResult("currentLevelFormat") = currentLevelFormat
                            End If

                            Dim metadataFields As String() = {
                                "paragraphStyleId", "paragraphStyleName", "numberRunFormatting",
                                "tabPositionState", "numberStyleOpenXml", "alignmentOpenXml",
                                "trailingCharacterOpenXml", "ilvlZeroBased"
                            }
                            For Each fieldName As String In metadataFields
                                If sourceLevel(fieldName) IsNot Nothing Then
                                    currentLevelFormat(fieldName) = sourceLevel(fieldName).DeepClone()
                                End If
                            Next
                        End If
                    End If
                End If
            End If

        Catch ex As System.Exception
            listResult("error") = ex.Message
        End Try

        Return listResult
    End Function
    ''' <summary>
    ''' Extracts paragraph border formatting.
    ''' </summary>
    ''' <param name="para">Paragraph to inspect.</param>
    ''' <returns>JSON object describing borders present on the paragraph.</returns>
    Private Function ExtractBorders(para As Word.Paragraph) As JObject
        Dim borders As New JObject()
        Try
            Dim borderTypes As WdBorderType() = {
                WdBorderType.wdBorderTop,
                WdBorderType.wdBorderBottom,
                WdBorderType.wdBorderLeft,
                WdBorderType.wdBorderRight
            }
            Dim borderNames As String() = {"top", "bottom", "left", "right"}

            For i As Integer = 0 To borderTypes.Length - 1
                Try
                    Dim border As Word.Border = para.Borders(borderTypes(i))
                    If border.LineStyle <> WdLineStyle.wdLineStyleNone Then
                        Dim b As New JObject()
                        b("lineStyle") = border.LineStyle.ToString()
                        b("lineWidth") = border.LineWidth.ToString()
                        b("color") = border.Color.ToString()
                        b("colorRGB") = ColorToRGB(border.Color)
                        borders(borderNames(i)) = b
                    End If
                Catch
                End Try
            Next

            Try
                borders("distanceFromText") = New JObject From {
                    {"top", para.Borders.DistanceFromTop},
                    {"bottom", para.Borders.DistanceFromBottom},
                    {"left", para.Borders.DistanceFromLeft},
                    {"right", para.Borders.DistanceFromRight}
                }
            Catch
            End Try

        Catch ex As System.Exception
            borders("error") = ex.Message
        End Try
        Return borders
    End Function

    ''' <summary>
    ''' Extracts paragraph shading/background properties.
    ''' </summary>
    ''' <param name="para">Paragraph to inspect.</param>
    ''' <returns>JSON object describing shading properties.</returns>
    Private Function ExtractShading(para As Word.Paragraph) As JObject
        Dim shading As New JObject()
        Try
            Dim s As Word.Shading = para.Shading
            shading("texture") = s.Texture.ToString()
            shading("textureDescription") = "wdTextureNone, wdTextureSolid, etc."

            If s.BackgroundPatternColor <> WdColor.wdColorAutomatic Then
                shading("backgroundColor") = s.BackgroundPatternColor.ToString()
                shading("backgroundColorRGB") = ColorToRGB(s.BackgroundPatternColor)
            End If

            If s.ForegroundPatternColor <> WdColor.wdColorAutomatic Then
                shading("foregroundColor") = s.ForegroundPatternColor.ToString()
                shading("foregroundColorRGB") = ColorToRGB(s.ForegroundPatternColor)
            End If

        Catch ex As System.Exception
            shading("error") = ex.Message
        End Try
        Return shading
    End Function


    ''' <summary>
    ''' Converts a Word color value to RGB hex string.
    ''' </summary>
    ''' <param name="color">Word color value.</param>
    ''' <returns>RGB hex string (<c>#RRGGBB</c>), or <c>auto</c>/<c>unknown</c>.</returns>
    Private Function ColorToRGB(color As WdColor) As String
        Try
            Dim colorValue As Long = CLng(color)
            If colorValue < 0 Then Return "auto"

            Dim r As Integer = colorValue And &HFF
            Dim g As Integer = (colorValue >> 8) And &HFF
            Dim b As Integer = (colorValue >> 16) And &HFF

            Return $"#{r:X2}{g:X2}{b:X2}"
        Catch
            Return "unknown"
        End Try
    End Function


    ''' <summary>
    ''' Computes a fingerprint string for a list template.
    ''' Used to detect equivalent list templates and avoid reapplying identical formatting.
    ''' </summary>
    ''' <param name="tpl">List template instance.</param>
    ''' <returns>Fingerprint string, or empty string if unavailable.</returns>
    Private Function ComputeListTemplateFingerprint(tpl As Word.ListTemplate) As String
        If tpl Is Nothing Then Return ""

        Try
            Dim sb As New StringBuilder()
            Dim maxLvl As Integer = 0
            Try
                maxLvl = Math.Min(9, tpl.ListLevels.Count)
            Catch
                maxLvl = 0
            End Try

            For lvl As Integer = 1 To maxLvl
                Try
                    Dim ll As Word.ListLevel = tpl.ListLevels(lvl)

                    Dim ns As String = ""
                    Try : ns = ll.NumberStyle.ToString() : Catch : ns = "" : End Try

                    Dim isBullet As Boolean = ns.IndexOf("bullet", StringComparison.OrdinalIgnoreCase) >= 0

                    If isBullet Then
                        Dim bFont As String = ""
                        Try
                            If ll.Font IsNot Nothing AndAlso Not String.IsNullOrEmpty(ll.Font.Name) Then
                                bFont = ll.Font.Name
                            End If
                        Catch
                            bFont = ""
                        End Try
                        If String.IsNullOrEmpty(bFont) Then bFont = "?"

                        Dim charCode As Integer = 0
                        Try
                            Dim nf As String = ll.NumberFormat
                            If Not String.IsNullOrEmpty(nf) Then
                                charCode = AscW(nf.Chars(0))
                            End If
                        Catch
                            charCode = 0
                        End Try

                        Dim numPos As Single = 0
                        Dim txtPos As Single = 0
                        Dim tabPos As Single = 0
                        Try : numPos = CSng(ll.NumberPosition) : Catch : End Try
                        Try : txtPos = CSng(ll.TextPosition) : Catch : End Try
                        Try : tabPos = CSng(ll.TabPosition) : Catch : End Try

                        sb.Append($"{lvl}:{ns}:BULLET:{bFont}:{charCode}:{Math.Round(numPos, 1)}:{Math.Round(txtPos, 1)}:{Math.Round(tabPos, 1)}|")
                    Else
                        Dim nf As String = ""
                        Try : nf = ll.NumberFormat : Catch : nf = "" : End Try

                        sb.Append($"{lvl}:{ns}:{nf}|")
                    End If
                Catch
                End Try
            Next

            Return sb.ToString()
        Catch
            Return ""
        End Try
    End Function

#End Region


    ''' <summary>
    ''' Automatically extracts a DocStyle template from a properly formatted Word document.
    ''' Instead of requiring a manually authored trainer document with "STYLE NAME: description" paragraphs,
    ''' this method discovers all paragraph styles actually used in the document, collects representative
    ''' samples, uses the LLM to generate semantic "whenToApply" descriptions, and produces the same
    ''' JSON template structure that <see cref="ExtractParagraphStylesToJson"/> creates.
    ''' </summary>
    Public Async Sub AutoExtractStyleTemplate()
        Try
            Dim app As Word.Application = Globals.ThisAddIn.Application
            Dim doc As Word.Document = app.ActiveDocument

            If doc Is Nothing Then
                ShowCustomMessageBox("No active document found.")
                Return
            End If

            ' Expand paths
            Dim docStylePath As String = ExpandEnvironmentVariables(INI_DocStylePath)
            If Not String.IsNullOrEmpty(docStylePath) AndAlso Not docStylePath.EndsWith("\") Then docStylePath &= "\"

            Dim docStylePathLocal As String = ExpandEnvironmentVariables(INI_DocStylePathLocal)
            If Not String.IsNullOrEmpty(docStylePathLocal) AndAlso Not docStylePathLocal.EndsWith("\") Then docStylePathLocal &= "\"

            Dim hasGlobal As Boolean = Not String.IsNullOrWhiteSpace(docStylePath) AndAlso Directory.Exists(docStylePath)
            Dim hasLocal As Boolean = Not String.IsNullOrWhiteSpace(docStylePathLocal) AndAlso Directory.Exists(docStylePathLocal)

            ' Determine save path
            Dim savePath As String
            Dim isLocal As Boolean = False

            If Not hasGlobal AndAlso Not hasLocal Then
                savePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                ShowCustomMessageBox("Warning: Neither 'DocStylePath' nor 'DocStylePathLocal' is configured. The file will be saved to your Desktop.")
            ElseIf hasGlobal AndAlso hasLocal Then
                Dim choice As Integer = ShowCustomYesNoBox("Where do you want to save the style template?",
                    "Central (shared)", "Local (personal)", $"{AN} - Save Location")
                If choice = 0 Then Return
                If choice = 1 Then
                    savePath = docStylePath
                    isLocal = False
                Else
                    savePath = docStylePathLocal
                    isLocal = True
                End If
            ElseIf hasGlobal Then
                savePath = docStylePath
            Else
                savePath = docStylePathLocal
                isLocal = True
            End If

            ' Ask for template display name
            Dim safeDocName As String = Regex.Replace(doc.Name.Replace(".docx", "").Replace(".doc", ""), "[^a-zA-Z0-9_-]", "_")
            Dim defaultDisplayName As String = $"{safeDocName}_{DateTime.Now:yyyyMMdd_HHmm}"

            Dim templateDisplayName As String = ShowCustomInputBox(
                "Enter a name for this style template:" & vbCrLf & vbCrLf &
                "This will auto-extract styles from the active document and use AI to generate 'whenToApply' descriptions.",
                $"{AN} - Auto-Extract Style Template", True, defaultDisplayName)
            If String.IsNullOrWhiteSpace(templateDisplayName) OrElse templateDisplayName.Equals("ESC", StringComparison.OrdinalIgnoreCase) Then
                Return
            End If

            ' Optional: let user provide document context for better LLM descriptions
            Dim documentContext As String = ShowCustomInputBox(
                "Optionally describe the document type/purpose for better style descriptions:" & vbCrLf &
                "(e.g., 'Swiss law firm contract template', 'Corporate board minutes', 'NDA agreement')" & vbCrLf & vbCrLf &
                "Leave empty to skip.",
                $"{AN} - Document Context", True, "")
            If documentContext IsNot Nothing AndAlso documentContext.Equals("ESC", StringComparison.OrdinalIgnoreCase) Then
                Return
            End If

            Dim targetRange As Word.Range
            Dim openXmlContext As JObject = BuildDocStyleOpenXmlContext(doc)

            ' Use selection if available, otherwise entire document
            If app.Selection.Type = WdSelectionType.wdSelectionIP Then
                targetRange = doc.Content
            Else
                targetRange = app.Selection.Range.Duplicate
            End If

            ShowProgressBarInSeparateThread($"{AN} - Auto-Extracting Styles", "Scanning document paragraphs...")
            ProgressBarModule.CancelOperation = False

            ' ──────────────────────────────────────────────────────────────────
            ' Phase 1: Discover all used styles and collect representative samples
            '          ALSO capture a representative Word.Paragraph per style so
            '          we don't need to re-walk targetRange.Paragraphs later
            '          (the COM enumerator can become stale after an Await).
            ' ──────────────────────────────────────────────────────────────────
            Dim styleGroups As New Dictionary(Of String, AutoStyleGroup)(StringComparer.OrdinalIgnoreCase)
            Dim representativeParas As New Dictionary(Of String, Word.Paragraph)(StringComparer.OrdinalIgnoreCase)
            Dim paragraphCount As Integer = 0

            Try
                For Each para As Word.Paragraph In targetRange.Paragraphs
                    If ProgressBarModule.CancelOperation Then
                        ShowCustomMessageBox("Operation cancelled.")
                        Return
                    End If

                    paragraphCount += 1
                    Dim text As String = para.Range.Text.TrimEnd(vbCr, vbLf, ChrW(13), ChrW(10)).Trim()
                    If String.IsNullOrWhiteSpace(text) Then Continue For

                    Dim styleName As String = "Normal"
                    Try
                        styleName = para.Style.NameLocal
                    Catch
                    End Try

                    If Not styleGroups.ContainsKey(styleName) Then
                        styleGroups(styleName) = New AutoStyleGroup() With {
                            .StyleName = styleName,
                            .Samples = New List(Of AutoStyleSample)(),
                            .TotalCount = 0
                        }
                    End If

                    Dim group As AutoStyleGroup = styleGroups(styleName)
                    group.TotalCount += 1

                    ' Capture the first non-empty paragraph as the representative for formatting extraction
                    If Not representativeParas.ContainsKey(styleName) Then
                        representativeParas(styleName) = para
                    End If

                    ' Collect up to 5 representative samples per style for LLM context
                    If group.Samples.Count < 5 Then
                        Dim sample As New AutoStyleSample()
                        sample.Text = If(text.Length > 200, text.Substring(0, 200) & "...", text)

                        ' Gather context hints
                        Try
                            Dim lf As Word.ListFormat = para.Range.ListFormat
                            If lf.ListType <> WdListType.wdListNoNumbering Then
                                If lf.ListType = WdListType.wdListBullet OrElse lf.ListType = WdListType.wdListPictureBullet Then
                                    sample.ListHint = "Bullet"
                                Else
                                    sample.ListHint = "Numbered"
                                End If
                                sample.ListLevel = lf.ListLevelNumber
                            End If
                        Catch
                        End Try

                        Try
                            sample.IsInTable = (para.Range.Tables.Count > 0 OrElse para.Range.Cells.Count > 0)
                        Catch
                        End Try

                        Try
                            sample.OutlineLevel = para.OutlineLevel.ToString()
                        Catch
                        End Try

                        Try
                            sample.IsBold = (para.Range.Font.Bold = -1)
                        Catch
                        End Try

                        Try
                            sample.FontSize = para.Range.Font.Size
                        Catch
                        End Try

                        group.Samples.Add(sample)
                    End If
                Next
            Finally
                ProgressBarModule.CancelOperation = True
            End Try

            If styleGroups.Count = 0 Then
                ShowCustomMessageBox("No styles found in the document.")
                Return
            End If

            ' ──────────────────────────────────────────────────────────────────
            ' Phase 2: Use LLM to generate whenToApply descriptions
            ' ──────────────────────────────────────────────────────────────────
            GlobalProgressLabel = "Generating style descriptions with AI..."

            Dim styleDescriptionPrompt As New StringBuilder()
            styleDescriptionPrompt.AppendLine("You are an expert document formatting analyst. You will receive a list of Word paragraph styles discovered in a document, along with representative sample paragraphs for each style.")
            styleDescriptionPrompt.AppendLine("Your task: For each style, write a concise 'whenToApply' description (1-2 sentences) that explains WHEN and WHERE this style should be used — i.e., what kind of paragraph content or structural role it serves.")
            styleDescriptionPrompt.AppendLine("The descriptions must be generic and reusable (not specific to this document's content).")
            styleDescriptionPrompt.AppendLine()
            styleDescriptionPrompt.AppendLine("RESPOND with a JSON array where each element has:")
            styleDescriptionPrompt.AppendLine("  {""styleName"": ""<exact style name>"", ""whenToApply"": ""<description>"", ""userStyleName"": ""<short human-friendly name for this style>""}")
            styleDescriptionPrompt.AppendLine()
            styleDescriptionPrompt.AppendLine("RULES:")
            styleDescriptionPrompt.AppendLine("- userStyleName should be a short, descriptive label (e.g., 'Main Heading', 'Body Text', 'Bullet List Item', 'Sub-clause Numbering')")
            styleDescriptionPrompt.AppendLine("- whenToApply should describe the paragraph's structural role, not its content")
            styleDescriptionPrompt.AppendLine("- Include hints about list type (bullet vs numbered) when applicable")
            styleDescriptionPrompt.AppendLine("- Respond ONLY with the JSON array, no other text")

            Dim userPrompt As New StringBuilder()
            If Not String.IsNullOrWhiteSpace(documentContext) Then
                userPrompt.AppendLine($"<CONTEXT>Document type: {documentContext}</CONTEXT>")
                userPrompt.AppendLine()
            End If

            userPrompt.AppendLine("<STYLES>")
            Dim styleIndex As Integer = 0
            For Each kvp In styleGroups.OrderByDescending(Function(g) g.Value.TotalCount)
                styleIndex += 1
                Dim group As AutoStyleGroup = kvp.Value
                userPrompt.AppendLine($"--- Style #{styleIndex}: ""{group.StyleName}"" (used {group.TotalCount} times) ---")

                For Each sample In group.Samples
                    Dim hints As New StringBuilder()
                    If Not String.IsNullOrEmpty(sample.ListHint) Then hints.Append($" [LIST:{sample.ListHint}/L{sample.ListLevel}]")
                    If sample.IsInTable Then hints.Append(" [TABLE]")
                    If sample.OutlineLevel IsNot Nothing AndAlso Not sample.OutlineLevel.Contains("BodyText") Then hints.Append($" [OUTLINE:{sample.OutlineLevel}]")
                    If sample.IsBold Then hints.Append(" [BOLD]")
                    If sample.FontSize > 0 AndAlso sample.FontSize <> 9999 Then hints.Append($" [SIZE:{sample.FontSize}pt]")
                    userPrompt.AppendLine($"  Sample:{hints} {sample.Text}")
                Next

                userPrompt.AppendLine()
            Next
            userPrompt.AppendLine("</STYLES>")

            Dim llmResponse As String = Await LLM(styleDescriptionPrompt.ToString(), userPrompt.ToString(), "", "", 0, False)

            ' Parse LLM response
            Dim styleDescriptions As New Dictionary(Of String, (UserStyleName As String, WhenToApply As String))(StringComparer.OrdinalIgnoreCase)
            Try
                Dim jsonResponse As String = ExtractJsonFromResponse(llmResponse)
                Dim descArray As JArray = JArray.Parse(jsonResponse)

                For Each item As JObject In descArray
                    Dim sName As String = If(item("styleName") IsNot Nothing, CStr(item("styleName")), "")
                    Dim whenToApply As String = If(item("whenToApply") IsNot Nothing, CStr(item("whenToApply")), "")
                    Dim userStyleName As String = If(item("userStyleName") IsNot Nothing, CStr(item("userStyleName")), sName)

                    If Not String.IsNullOrWhiteSpace(sName) Then
                        styleDescriptions(sName) = (userStyleName, whenToApply)
                    End If
                Next
            Catch ex As System.Exception
                Debug.WriteLine($"[DocStyle] Auto-extract LLM parse error: {ex.Message}")
                ' Fall back to using style names as descriptions
            End Try

            ' ──────────────────────────────────────────────────────────────────
            ' Phase 3: Build the JSON template (same structure as ExtractParagraphStylesToJson)
            '          Uses representativeParas captured in Phase 1 — no re-enumeration needed.
            ' ──────────────────────────────────────────────────────────────────
            GlobalProgressLabel = "Building style template..."

            Dim userStyles As New JArray()
            Dim wdStyleDefinitions As New JObject()
            Dim collectedStyles As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim userStyleIndex As Integer = 0

            ' Create one userStyle entry per discovered style, using the cached representative paragraph
            For Each kvp In styleGroups.OrderByDescending(Function(g) g.Value.TotalCount)
                Dim group As AutoStyleGroup = kvp.Value
                userStyleIndex += 1

                ' Use the representative paragraph captured during Phase 1
                Dim representativePara As Word.Paragraph = Nothing
                If Not representativeParas.TryGetValue(group.StyleName, representativePara) OrElse representativePara Is Nothing Then
                    Debug.WriteLine($"[DocStyle] Auto-extract: No representative paragraph for style '{group.StyleName}' — skipping")
                    Continue For
                End If

                ' Validate the cached paragraph is still accessible
                Try
                    Dim checkText As String = representativePara.Range.Text
                Catch ex As System.Exception
                    Debug.WriteLine($"[DocStyle] Auto-extract: Representative paragraph for style '{group.StyleName}' is no longer accessible ({ex.Message}) — skipping")
                    Continue For
                End Try

                ' Build userStyle JSON using the same extraction helpers as the manual flow
                Dim userStyleName As String = $"UserStyle_{userStyleIndex}"
                Dim whenToApply As String = $"Paragraphs formatted with the '{group.StyleName}' Word style."

                If styleDescriptions.ContainsKey(group.StyleName) Then
                    userStyleName = styleDescriptions(group.StyleName).UserStyleName
                    whenToApply = styleDescriptions(group.StyleName).WhenToApply
                End If

                Dim result As New JObject()
                result("userStyleIndex") = userStyleIndex
                result("userStyleName") = userStyleName
                result("whenToApply") = whenToApply

                ' Word style info
                Try
                    Dim style As Word.Style = representativePara.Style
                    result("wdStyleName") = style.NameLocal
                    result("wdStyleBuiltIn") = style.BuiltIn
                Catch
                    result("wdStyleName") = group.StyleName
                    result("wdStyleBuiltIn") = False
                End Try

                Dim styleMeta As JObject = GetStyleOpenXmlMetadata(openXmlContext, If(result("wdStyleName") IsNot Nothing, CStr(result("wdStyleName")), group.StyleName))
                If styleMeta IsNot Nothing Then
                    If styleMeta("styleId") IsNot Nothing Then result("wdStyleId") = styleMeta("styleId").DeepClone()
                    If styleMeta("numberingBinding") IsNot Nothing Then result("numberingBinding") = styleMeta("numberingBinding").DeepClone()
                End If

                ' Table context
                Dim isInTable As Boolean = False
                Try
                    If representativePara.Range.Tables.Count > 0 OrElse representativePara.Range.Cells.Count > 0 Then
                        isInTable = True
                    End If
                Catch
                End Try
                result("isInTableCell") = isInTable

                ' Reuse the same extraction helpers
                result("paragraphFormatting") = ExtractParagraphFormat(representativePara, representativePara.Range.Duplicate)
                result("fontFormatting") = ExtractFontFormat(representativePara.Range)
                result("listFormatting") = ExtractListFormat(representativePara.Range, openXmlContext)
                result("tabStops") = ExtractTabStopsCompact(representativePara)

                Dim borders As JObject = ExtractBorders(representativePara)
                If borders.Properties().Any(Function(p) p.Name <> "distanceFromText" AndAlso p.Name <> "error") Then
                    result("borders") = borders
                End If

                Dim shading As JObject = ExtractShading(representativePara)
                If shading("backgroundColor") IsNot Nothing OrElse shading("foregroundColor") IsNot Nothing Then
                    result("shading") = shading
                End If

                ' Add usage statistics as metadata
                result("_autoExtractInfo") = New JObject From {
                    {"paragraphCount", group.TotalCount},
                    {"autoGenerated", True}
                }

                userStyles.Add(result)

                ' Collect wdStyle name for style definitions
                If Not String.IsNullOrWhiteSpace(group.StyleName) AndAlso Not collectedStyles.Contains(group.StyleName) Then
                    collectedStyles.Add(group.StyleName)
                End If
            Next

            ' Extract full wdStyle definitions
            For Each styleName In collectedStyles
                Try
                    Dim styleObj As JObject = ExtractFullStyleDefinition(doc, styleName, openXmlContext)
                    If styleObj IsNot Nothing Then
                        wdStyleDefinitions(styleName) = styleObj
                    End If
                Catch ex As System.Exception
                    Debug.WriteLine($"[DocStyle] Auto-extract: Error extracting wdStyle definition for '{styleName}': {ex.Message}")
                End Try
            Next

            ' ──────────────────────────────────────────────────────────────────
            ' Phase 4: Assemble and save the template
            ' ──────────────────────────────────────────────────────────────────
            Dim templateResult As New JObject()
            templateResult("schemaVersion") = DocStyleSchemaVersion
            templateResult("templateName") = templateDisplayName
            templateResult("description") = "Style template auto-extracted from a formatted document. Each user style includes an AI-generated 'whenToApply' field. Review and refine the 'userStyleName' and 'whenToApply' fields for best results. The 'wdStyleDefinitions' section contains Word style definitions that can optionally be created/updated in target documents."
            templateResult("documentInfo") = New JObject From {
                {"extractedFrom", doc.Name},
                {"extractionDate", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")},
                {"extractionMethod", "auto"},
                {"totalParagraphsScanned", paragraphCount},
                {"totalUserStyles", userStyles.Count},
                {"totalWdStyles", collectedStyles.Count}
            }
            templateResult("userStyles") = userStyles
            templateResult("wdStyleDefinitions") = wdStyleDefinitions
            templateResult("numberingDefinitions") = BuildReferencedNumberingDefinitions(openXmlContext, collectedStyles)
            If openXmlContext("documentDefaults") IsNot Nothing Then
                templateResult("documentDefaults") = openXmlContext("documentDefaults").DeepClone()
            End If

            Dim jsonString As String = JsonConvert.SerializeObject(templateResult, Formatting.Indented)

            Dim safeTemplateName As String = Regex.Replace(templateDisplayName, "[^a-zA-Z0-9_-]", "_")
            Dim fileName As String = $"{AN2}-ds-{safeTemplateName}.json"
            Dim filePath As String = Path.Combine(savePath, fileName)

            If File.Exists(filePath) Then
                Dim overwrite As Integer = ShowCustomYesNoBox(
                    $"A style template with this name already exists.{vbCrLf}{vbCrLf}Do you want to overwrite it?",
                    "Overwrite", "Cancel", $"{AN} - File Exists")
                If overwrite <> 1 Then Return
            End If

            Try
                File.WriteAllText(filePath, jsonString, Encoding.UTF8)
            Catch ioEx As System.Exception
                ShowCustomMessageBox($"Could not save file: {ioEx.Message}")
                Return
            End Try

            Dim successMessage As String =
                $"Auto-extracted style template '{templateDisplayName}' created successfully." & vbCrLf & vbCrLf &
                $"Styles discovered: {styleGroups.Count}" & vbCrLf &
                $"Paragraphs scanned: {paragraphCount}" & vbCrLf &
                $"Location: {filePath}" & vbCrLf & vbCrLf &
                "Tip: Review and refine the 'userStyleName' and 'whenToApply' fields in the template for best results when applying to other documents." & vbCrLf & vbCrLf &
                "Would you like to edit the template now?"

            Dim editChoice As Integer = ShowCustomYesNoBox(successMessage,
                "Edit Template", "Close", $"{AN} - Style Template Created")

            If editChoice = 1 Then
                SLib.ShowTextFileEditor(filePath, $"{AN} - Style Template '{templateDisplayName}'", True, _context)
            End If

        Catch ex As System.Exception
            ShowCustomMessageBox($"Error auto-extracting style template: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Holds metadata for a style group discovered during auto-extraction.
    ''' </summary>
    Private Class AutoStyleGroup
        Public Property StyleName As String
        Public Property Samples As List(Of AutoStyleSample)
        Public Property TotalCount As Integer
    End Class

    ''' <summary>
    ''' Holds a representative paragraph sample with formatting context hints.
    ''' </summary>
    Private Class AutoStyleSample
        Public Property Text As String = ""
        Public Property ListHint As String = ""
        Public Property ListLevel As Integer = 0
        Public Property IsInTable As Boolean = False
        Public Property OutlineLevel As String = ""
        Public Property IsBold As Boolean = False
        Public Property FontSize As Single = 0
    End Class

End Class
