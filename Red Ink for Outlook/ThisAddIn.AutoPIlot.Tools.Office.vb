' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.Tools.Office.vb
' Purpose:
'   Defines and executes AutoPilot internal tools for Office document operations
'   within Outlook AutoPilot Chat-Agent runs, including Word, Excel, and
'   PowerPoint document creation and conversion workflows.
'
' Tools Provided:
'   - create_word_document: Creates Word documents (.docx) with text, tables,
'     images, and advanced formatting
'   - comment_word_document: Adds review comment bubbles to Word attachments
'   - create_excel_spreadsheet: Creates Excel workbooks (.xlsx/.xlsm) with
'     sheets, cells, charts, data validation, formulas, and optional VBA
'   - create_powerpoint: Creates PowerPoint presentations (.pptx) with slides,
'     text, images, and template support
'   - word_to_pdf: Converts Word documents to PDF format
'   - pdf_to_word: Converts PDF documents to editable Word format
'
' Tool Interface Architecture:
'   - Registration:
'       * Tools are exposed as `ModelConfig` entries (`Tool=True`, `ToolOnly=True`)
'         so they participate in the same tool-calling pipeline as external tools.
'       * Tool metadata (`ToolDefinition`, `ToolInstructionsPrompt`) is generated
'         inline and consumed by `ExecuteToolCall` / `ExecuteToolingLoop`.
'   - Dispatch:
'       * `TryExecuteAutoPilotTool` routes parsed tool calls to strongly scoped
'         executor methods (`ExecuteCreateWordDocTool`, `ExecuteCreateExcelTool`,
'         etc.) and returns `ToolResponse` payloads.
'   - Session scope:
'       * All tools use AutoPilot session state from `ThisAddIn.Autopilot.vb`:
'           - `_apCurrentAttachments`: attachment registry for input/output lookups
'           - `_apCurrentTempDir`: per-mail temp directory for file creation
'           - `_apCurrentMailInfo`: metadata about the current email session
'       * Supports tool chaining via output registration (`OutputFiles`) and
'         attachment lookup via `FindAttachment` (original + prior tool outputs).
'   - UI interaction:
'       * Switches to UI thread via `SwitchToUi` for COM-based Office operations.
'       * Late binding avoids hard PIA references where feasible (PowerPoint).
'   - Error handling:
'       * Returns structured `ToolResponse` with success flag, message, and
'         error details. File operations include collision prevention and
'         cleanup of temporary resources.
'   - Logging and UX:
'       * Emits execution traces to tooling context (`context.Log`) and
'         AutoPilot dashboard (`ApDashboardLog`) with concise status summaries.
'
' Security & Safety:
'   - Path containment:
'       * All tool outputs are created in `_apCurrentTempDir` and re-used only
'         via resolved attachment/output references.
'   - File validation:
'       * Size checks prevent oversized attachments from processing.
'       * Extension validation ensures correct file type handling.
'       * Filename collision prevention via counter-based renaming.
'
' =============================================================================



Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.IO.Compression
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Xml
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary
Imports SharedLibrary.Agents
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn


    Private Class AutoPilotDesignResolution
        Public Property RequestedName As String
        Public Property ApplicationName As String
        Public Property Descriptor As SharedLibrary.Agents.DocumentDesignDescriptor
        Public Property ApplicationConfig As JObject
        Public Property TemplatePath As String
        Public Property TemplateWarning As String
        Public Property AppliedDefaultCount As Integer

        Public ReadOnly Property Found As Boolean
            Get
                Return Descriptor IsNot Nothing AndAlso ApplicationConfig IsNot Nothing
            End Get
        End Property

        Public ReadOnly Property Applied As Boolean
            Get
                Return Found AndAlso (AppliedDefaultCount > 0 OrElse Not String.IsNullOrWhiteSpace(TemplatePath))
            End Get
        End Property

        Public ReadOnly Property SourceLabel As String
            Get
                If Descriptor Is Nothing Then Return ""
                Return If(Descriptor.IsLocal, "local design repository", "central design repository")
            End Get
        End Property
    End Class

    Private Shared Function BuildDesignExecutionNote(design As AutoPilotDesignResolution) As String
        If design Is Nothing OrElse String.IsNullOrWhiteSpace(design.RequestedName) Then Return ""
        If design.Descriptor Is Nothing Then
            Return $" Requested design '{design.RequestedName}' was not found; neutral professional design was used."
        End If
        If design.ApplicationConfig Is Nothing Then
            Return $" Configured design '{design.Descriptor.Name}' has no {design.ApplicationName} profile; neutral professional design was used."
        End If
        If design.Applied Then
            Return $" Configured design used: '{design.Descriptor.Name}' ({design.SourceLabel})."
        End If
        Return $" Configured design '{design.Descriptor.Name}' contained no applicable {design.ApplicationName} settings or usable template; neutral professional design was used."
    End Function

    ''' <summary>
    ''' Resolves a named design from AgentResourcesPath[/Local]\designs\designs.json,
    ''' applies only missing creator arguments as defaults, and resolves an optional
    ''' repository-relative Office template. Explicit tool-call parameters always win.
    ''' </summary>
    Private Function ResolveAutoPilotDocumentDesign(args As Dictionary(Of String, Object),
                                                     applicationName As String,
                                                     allowedDefaultKeys As IEnumerable(Of String),
                                                     allowedTemplateExtensions As IEnumerable(Of String),
                                                     context As ToolExecutionContext) As AutoPilotDesignResolution
        Dim result As New AutoPilotDesignResolution() With {
            .RequestedName = If(GetArgString(args, "design_name"), "").Trim(),
            .ApplicationName = If(applicationName, "").Trim()
        }
        If result.RequestedName = "" Then Return result

        result.Descriptor = SharedLibrary.Agents.DesignRepository.FindDesign(result.RequestedName)
        If result.Descriptor Is Nothing Then
            result.TemplateWarning = $"Configured design '{result.RequestedName}' was not found; neutral professional design was used."
            If context IsNot Nothing Then context.Log(result.TemplateWarning)
            ApDashboardLog("⚠ " & result.TemplateWarning, "warn")
            Return result
        End If

        result.ApplicationConfig = result.Descriptor.GetApplicationConfig(applicationName)
        If result.ApplicationConfig Is Nothing Then
            result.TemplateWarning = $"Configured design '{result.Descriptor.Name}' has no {applicationName} profile; neutral professional design was used."
            If context IsNot Nothing Then context.Log(result.TemplateWarning)
            ApDashboardLog("⚠ " & result.TemplateWarning, "warn")
            Return result
        End If

        If allowedDefaultKeys IsNot Nothing Then
            For Each key As String In allowedDefaultKeys
                If String.IsNullOrWhiteSpace(key) OrElse HasMeaningfulToolArgument(args, key) Then Continue For
                Dim token As JToken = result.ApplicationConfig(key)
                If token Is Nothing OrElse token.Type = JTokenType.Null Then Continue For
                args(key) = token.DeepClone()
                result.AppliedDefaultCount += 1
            Next
        End If

        Dim themeSourceRelative As String = If(result.ApplicationConfig.Value(Of String)("theme_source_file"), "").Trim()
        If themeSourceRelative <> "" Then
            Dim themeSourcePath As String = result.Descriptor.ResolveRepositoryFile(themeSourceRelative)
            If themeSourcePath = "" OrElse Not System.IO.File.Exists(themeSourcePath) Then
                result.TemplateWarning = $"Design '{result.Descriptor.Name}' theme_source_file was not found or invalid: {themeSourceRelative}."
            ElseIf HasAllowedDesignThemeSourceExtension(themeSourcePath, applicationName) Then
                result.AppliedDefaultCount += ApplyOfficeThemeDefaultsFromTemplate(args, themeSourcePath, applicationName)
            Else
                result.TemplateWarning = $"Design '{result.Descriptor.Name}' theme_source_file type is not supported for {applicationName}: {System.IO.Path.GetExtension(themeSourcePath)}."
            End If
        End If

        Dim templateRelative As String = If(result.ApplicationConfig.Value(Of String)("template_file"), "").Trim()
        If templateRelative <> "" Then
            Dim resolvedPath As String = result.Descriptor.ResolveRepositoryFile(templateRelative)
            If resolvedPath = "" Then
                result.TemplateWarning = $"Design '{result.Descriptor.Name}' has an invalid template_file path; JSON design settings were used without the template."
            ElseIf Not System.IO.File.Exists(resolvedPath) Then
                result.TemplateWarning = $"Design '{result.Descriptor.Name}' template was not found: {templateRelative}; JSON design settings were used without the template."
            ElseIf Not HasAllowedDesignTemplateExtension(resolvedPath, allowedTemplateExtensions) Then
                result.TemplateWarning = $"Design '{result.Descriptor.Name}' template type is not allowed for {applicationName}: {System.IO.Path.GetExtension(resolvedPath)}; JSON design settings were used without the template."
            Else
                result.TemplatePath = resolvedPath
                result.AppliedDefaultCount += ApplyOfficeThemeDefaultsFromTemplate(args, resolvedPath, applicationName)
                If String.Equals(applicationName, "Word", StringComparison.OrdinalIgnoreCase) AndAlso
                   Not HasMeaningfulToolArgument(args, "use_template_styles") Then
                    args("use_template_styles") = True
                End If
            End If
        End If

        If result.TemplateWarning <> "" Then
            If context IsNot Nothing Then context.Log(result.TemplateWarning)
            ApDashboardLog("⚠ " & result.TemplateWarning, "warn")
        Else
            Dim templateNote As String = If(result.TemplatePath <> "", $" + template '{System.IO.Path.GetFileName(result.TemplatePath)}'", "")
            Dim msg As String = $"Using configured design '{result.Descriptor.Name}' from {result.SourceLabel}{templateNote}."
            If context IsNot Nothing Then context.Log(msg)
            ApDashboardLog("🎨 " & msg, "step")
        End If

        Return result
    End Function

    Private Shared Function HasMeaningfulToolArgument(args As Dictionary(Of String, Object), key As String) As Boolean
        If args Is Nothing OrElse String.IsNullOrWhiteSpace(key) OrElse Not args.ContainsKey(key) Then Return False
        Dim value As Object = args(key)
        If value Is Nothing Then Return False
        Dim token As JToken = TryCast(value, JToken)
        If token IsNot Nothing AndAlso (token.Type = JTokenType.Null OrElse token.Type = JTokenType.Undefined) Then Return False
        If TypeOf value Is String Then Return Not String.IsNullOrWhiteSpace(CStr(value))
        Return True
    End Function

    Private Shared Function HasAllowedDesignTemplateExtension(path As String,
                                                              allowedExtensions As IEnumerable(Of String)) As Boolean
        If String.IsNullOrWhiteSpace(path) OrElse allowedExtensions Is Nothing Then Return False
        Dim ext As String = System.IO.Path.GetExtension(path)
        For Each allowed As String In allowedExtensions
            If String.Equals(ext, allowed, StringComparison.OrdinalIgnoreCase) Then Return True
        Next
        Return False
    End Function

    Private Shared Function HasAllowedDesignThemeSourceExtension(path As String, applicationName As String) As Boolean
        Dim ext As String = If(System.IO.Path.GetExtension(path), "").ToLowerInvariant()
        Select Case If(applicationName, "").Trim().ToLowerInvariant()
            Case "word"
                Return ext = ".docx" OrElse ext = ".dotx" OrElse ext = ".dotm"
            Case "powerpoint"
                Return ext = ".pptx" OrElse ext = ".potx"
            Case "excel"
                Return ext = ".xlsx" OrElse ext = ".xltx"
            Case Else
                Return False
        End Select
    End Function

    ''' <summary>
    ''' Extracts stable Office Open XML theme primitives that the existing renderers
    ''' can actually consume. The template remains the richer source of masters/styles;
    ''' this prevents custom renderer shapes from falling back to unrelated neutral colors/fonts.
    ''' </summary>
    Private Shared Function ApplyOfficeThemeDefaultsFromTemplate(args As Dictionary(Of String, Object),
                                                                 templatePath As String,
                                                                 applicationName As String) As Integer
        If args Is Nothing OrElse String.IsNullOrWhiteSpace(templatePath) OrElse Not System.IO.File.Exists(templatePath) Then Return 0
        Dim appliedCount As Integer = 0
        Try
            Using archive As System.IO.Compression.ZipArchive = System.IO.Compression.ZipFile.OpenRead(templatePath)
                Dim themeEntry As System.IO.Compression.ZipArchiveEntry = Nothing
                Dim normalizedApp As String = If(applicationName, "").Trim().ToLowerInvariant()
                Dim preferredSuffix As String
                Select Case normalizedApp
                    Case "powerpoint" : preferredSuffix = "ppt/theme/theme1.xml"
                    Case "excel" : preferredSuffix = "xl/theme/theme1.xml"
                    Case "word" : preferredSuffix = "word/theme/theme1.xml"
                    Case Else : preferredSuffix = "/theme/theme1.xml"
                End Select

                themeEntry = archive.Entries.FirstOrDefault(
                    Function(e) e.FullName.Replace("\\", "/").EndsWith(preferredSuffix, StringComparison.OrdinalIgnoreCase))
                If themeEntry Is Nothing Then Return appliedCount

                Dim xml As New System.Xml.XmlDocument()
                Using stream As System.IO.Stream = themeEntry.Open()
                    xml.Load(stream)
                End Using

                Dim ns As New System.Xml.XmlNamespaceManager(xml.NameTable)
                ns.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")

                Dim accent1 As String = ReadOfficeThemeColor(xml, ns, "accent1")
                Dim accent2 As String = ReadOfficeThemeColor(xml, ns, "accent2")
                Dim dk1 As String = ReadOfficeThemeColor(xml, ns, "dk1")
                Dim lt1 As String = ReadOfficeThemeColor(xml, ns, "lt1")
                Dim fontName As String = ReadOfficeThemeFont(xml, ns)

                If SetToolArgumentDefault(args, "accent_color", accent1) Then appliedCount += 1
                If SetToolArgumentDefault(args, "secondary_color", accent2) Then appliedCount += 1
                If SetToolArgumentDefault(args, "text_color", dk1) Then appliedCount += 1
                If SetToolArgumentDefault(args, "light_color", lt1) Then appliedCount += 1
                If normalizedApp = "word" Then
                    If SetToolArgumentDefault(args, "base_font_name", fontName) Then appliedCount += 1
                Else
                    If SetToolArgumentDefault(args, "font_name", fontName) Then appliedCount += 1
                End If
            End Using
        Catch ex As System.Exception
            Debug.WriteLine($"Design theme extraction failed for '{templatePath}': {ex.Message}")
        End Try
        Return appliedCount
    End Function

    Private Shared Function ReadOfficeThemeColor(xml As System.Xml.XmlDocument,
                                                 ns As System.Xml.XmlNamespaceManager,
                                                 schemeName As String) As String
        If xml Is Nothing OrElse ns Is Nothing OrElse String.IsNullOrWhiteSpace(schemeName) Then Return ""
        Dim node As System.Xml.XmlNode = xml.SelectSingleNode($"//a:themeElements/a:clrScheme/a:{schemeName}/*[1]", ns)
        If node Is Nothing OrElse node.Attributes Is Nothing Then Return ""
        Dim raw As String = ""
        Dim attr As System.Xml.XmlAttribute = node.Attributes("val")
        If attr IsNot Nothing Then raw = If(attr.Value, "").Trim()
        If raw.Length = 6 AndAlso raw.All(Function(ch) Uri.IsHexDigit(ch)) Then Return "#" & raw.ToUpperInvariant()

        attr = node.Attributes("lastClr")
        If attr Is Nothing Then Return ""
        raw = If(attr.Value, "").Trim()
        If raw.Length = 6 AndAlso raw.All(Function(ch) Uri.IsHexDigit(ch)) Then Return "#" & raw.ToUpperInvariant()
        Return ""
    End Function

    Private Shared Function ReadOfficeThemeFont(xml As System.Xml.XmlDocument,
                                                ns As System.Xml.XmlNamespaceManager) As String
        If xml Is Nothing OrElse ns Is Nothing Then Return ""
        For Each xpath As String In New String() {
            "//a:themeElements/a:fontScheme/a:minorFont/a:latin",
            "//a:themeElements/a:fontScheme/a:majorFont/a:latin"}

            Dim node As System.Xml.XmlNode = xml.SelectSingleNode(xpath, ns)
            If node Is Nothing OrElse node.Attributes Is Nothing Then Continue For
            Dim attr As System.Xml.XmlAttribute = node.Attributes("typeface")
            If attr IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(attr.Value) Then Return attr.Value.Trim()
        Next
        Return ""
    End Function

    Private Shared Function SetToolArgumentDefault(args As Dictionary(Of String, Object), key As String, value As String) As Boolean
        If args Is Nothing OrElse String.IsNullOrWhiteSpace(key) OrElse String.IsNullOrWhiteSpace(value) Then Return False
        If HasMeaningfulToolArgument(args, key) Then Return False
        args(key) = value
        Return True
    End Function





    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: create_powerpoint
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Async Function ExecuteCreatePowerPointTool(
            toolCall As ToolCall, context As ToolExecutionContext, ct As CancellationToken) As System.Threading.Tasks.Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim slidesArray As JArray = Nothing
            If toolCall.Arguments IsNot Nothing AndAlso toolCall.Arguments.ContainsKey("slides") Then
                slidesArray = TryCast(toolCall.Arguments("slides"), JArray)
            End If

            If slidesArray Is Nothing OrElse slidesArray.Count = 0 Then
                response.Success = False
                response.Response = "Missing required parameter: slides (must be a non-empty array of slide objects)"
                Return response
            End If

            Dim design As AutoPilotDesignResolution = ResolveAutoPilotDocumentDesign(
                toolCall.Arguments,
                "PowerPoint",
                New String() {"style_preset", "accent_color", "secondary_color", "font_name", "aspect_ratio", "footer_text", "show_slide_numbers", "text_color", "muted_color", "light_color", "line_color", "green_color", "red_color", "amber_color", "preserve_template_slides"},
                New String() {".pptx", ".potx"},
                context)

            Dim fileName As String = GetArgString(toolCall.Arguments, "file_name")
            If String.IsNullOrWhiteSpace(fileName) Then fileName = "Presentation"
            For Each c As Char In Path.GetInvalidFileNameChars()
                fileName = fileName.Replace(c, "_"c)
            Next
            If Not fileName.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase) Then fileName &= ".pptx"

            Dim outputPath As String = Path.Combine(_apCurrentTempDir, fileName)
            Dim counter As Integer = 1
            While File.Exists(outputPath)
                Dim baseName As String = Path.GetFileNameWithoutExtension(fileName)
                fileName = baseName & $"_{counter}.pptx"
                outputPath = Path.Combine(_apCurrentTempDir, fileName)
                counter += 1
            End While

            Dim presTitle As String = GetArgString(toolCall.Arguments, "title")
            If String.IsNullOrWhiteSpace(presTitle) Then
                Dim firstSlide As JObject = TryCast(slidesArray(0), JObject)
                If firstSlide IsNot Nothing Then presTitle = firstSlide.Value(Of String)("title")
            End If

            Dim templateName As String = GetArgString(toolCall.Arguments, "template_attachment_name")
            Dim templatePath As String = Nothing
            Dim templateFromDesign As Boolean = False
            If Not String.IsNullOrWhiteSpace(templateName) Then
                Dim templateAtt = FindAttachment(templateName)
                If templateAtt IsNot Nothing AndAlso templateAtt.TempFilePath IsNot Nothing AndAlso File.Exists(templateAtt.TempFilePath) Then
                    templatePath = templateAtt.TempFilePath
                    Dim ignoredThemeDefaults As Integer = ApplyOfficeThemeDefaultsFromTemplate(toolCall.Arguments, templatePath, "PowerPoint")
                    ApDashboardLog($"📊 Using attached template: {templateName}", "step")
                Else
                    ApDashboardLog($"⚠ Template '{templateName}' not found, creating from scratch", "warn")
                End If
            ElseIf design IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(design.TemplatePath) Then
                templatePath = design.TemplatePath
                templateName = Path.GetFileName(templatePath)
                templateFromDesign = True
            End If

            context.Log($"Creating PowerPoint presentation: {fileName} ({slidesArray.Count} slides)" &
                        If(templatePath IsNot Nothing, $" from template: {templateName}", ""))
            ApDashboardLog($"📊 Creating PowerPoint: {fileName}", "step")

            Const ppLayoutBlank As Integer = 12
            Const ppSaveAsOpenXMLPresentation As Integer = 24

            Dim success As Boolean = Await SwitchToUi(Function()
                                                          Dim app As Object = Nothing
                                                          Dim pres As Object = Nothing
                                                          Dim weOwnApp As Boolean = False
                                                          Try
                                                              Try
                                                                  app = System.Runtime.InteropServices.Marshal.GetActiveObject("PowerPoint.Application")
                                                              Catch ex As System.Runtime.InteropServices.COMException
                                                                  app = Microsoft.VisualBasic.Interaction.CreateObject("PowerPoint.Application")
                                                                  weOwnApp = True
                                                              End Try

                                                              If templatePath IsNot Nothing Then
                                                                  pres = app.Presentations.Open(templatePath, ReadOnly:=0, Untitled:=-1, WithWindow:=0)
                                                                  If templateFromDesign AndAlso Not GetArgBool(toolCall.Arguments, "preserve_template_slides", False) Then
                                                                      For templateSlideIndex As Integer = CInt(pres.Slides.Count) To 1 Step -1
                                                                          Try : pres.Slides(templateSlideIndex).Delete() : Catch ex As System.Exception : End Try
                                                                      Next
                                                                  End If
                                                              Else
                                                                  pres = app.Presentations.Add(0)
                                                                  ApplyAutoPilotPowerPointPageSetup(pres, toolCall.Arguments)
                                                              End If

                                                              If Not String.IsNullOrWhiteSpace(presTitle) Then
                                                                  Try : pres.BuiltInDocumentProperties("Title").Value = presTitle : Catch : End Try
                                                              End If

                                                              Dim existingSlideCount As Integer = CInt(pres.Slides.Count)
                                                              Dim slideIndex As Integer = existingSlideCount
                                                              For Each slideObj As JObject In slidesArray.OfType(Of JObject)()
                                                                  ExpandAutoPilotPowerPointSlideData(slideObj)
                                                                  slideIndex += 1
                                                                  Dim sld As Object = Nothing
                                                                  Try
                                                                      sld = pres.Slides.Add(slideIndex, ppLayoutBlank)
                                                                      RenderAutoPilotPowerPointSlide(pres, sld, slideObj, slideIndex, existingSlideCount, toolCall.Arguments)
                                                                      ApplyPowerPointNotes(sld, slideObj.Value(Of String)("notes"))
                                                                  Finally
                                                                      If sld IsNot Nothing Then
                                                                          Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(sld) : Catch : End Try
                                                                      End If
                                                                  End Try
                                                              Next

                                                              pres.SaveAs(outputPath, ppSaveAsOpenXMLPresentation)
                                                              Return True
                                                          Catch ex As System.Exception
                                                              Debug.WriteLine($"CreatePowerPoint error: {ex.Message}")
                                                              Return False
                                                          Finally
                                                              Try
                                                                  If pres IsNot Nothing Then
                                                                      Try : pres.Close() : Catch : End Try
                                                                      Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pres) : Catch : End Try
                                                                  End If
                                                              Catch
                                                              End Try
                                                              Try
                                                                  If app IsNot Nothing Then
                                                                      If weOwnApp Then Try : app.Quit() : Catch : End Try
                                                                      Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(app) : Catch : End Try
                                                                  End If
                                                              Catch
                                                              End Try
                                                          End Try
                                                      End Function)

            If success AndAlso File.Exists(outputPath) Then
                RegisterAutoPilotGeneratedOutputFile(outputPath)
                Dim templateNote As String = If(templatePath IsNot Nothing, $", based on template '{templateName}'", "")
                Dim designNote As String = BuildDesignExecutionNote(design)
                response.Success = True
                response.Response = $"PowerPoint presentation created: {fileName} ({slidesArray.Count} new slides{templateNote}, {New FileInfo(outputPath).Length / 1024:F0} KB). The file will be attached to the reply.{designNote}"
                ApDashboardLog($"✓ PowerPoint created: {fileName}", "info")
            Else
                response.Success = False
                response.Response = "Failed to create PowerPoint presentation."
            End If

        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation was cancelled."
            response.Response = response.ErrorMessage
        Catch ex As System.Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error creating PowerPoint presentation: {ex.Message}"
        End Try

        Return response
    End Function

    ''' <summary>
    ''' Expands the optional shallow tool-contract field data_json into the richer
    ''' in-process slide object consumed by the renderer. Keeping nested presentation
    ''' data out of the public tool schema improves interoperability across tool-calling
    ''' model providers while preserving the full renderer feature set.
    ''' Explicit slide properties always win over values supplied through data_json.
    ''' </summary>
    Private Shared Sub ExpandAutoPilotPowerPointSlideData(slideObj As JObject)
        If slideObj Is Nothing Then Exit Sub

        Dim rawData As String = slideObj.Value(Of String)("data_json")
        If String.IsNullOrWhiteSpace(rawData) Then Exit Sub

        Try
            Dim payload As JObject = JObject.Parse(rawData)
            For Each prop As JProperty In payload.Properties()
                Dim existing As JToken = slideObj(prop.Name)
                If existing Is Nothing OrElse existing.Type = JTokenType.Null Then
                    slideObj(prop.Name) = prop.Value.DeepClone()
                End If
            Next

            ' Infer the most suitable rich layout only when the caller did not set one explicitly.
            If String.IsNullOrWhiteSpace(slideObj.Value(Of String)("layout")) Then
                If slideObj("kpis") IsNot Nothing Then
                    slideObj("layout") = "kpi"
                ElseIf slideObj("table") IsNot Nothing Then
                    slideObj("layout") = "table"
                ElseIf slideObj("chart") IsNot Nothing Then
                    slideObj("layout") = "chart"
                ElseIf slideObj("cards") IsNot Nothing Then
                    slideObj("layout") = "cards"
                ElseIf slideObj("steps") IsNot Nothing Then
                    slideObj("layout") = "process"
                ElseIf slideObj("structure") IsNot Nothing Then
                    slideObj("layout") = "structure"
                ElseIf slideObj("timeline") IsNot Nothing OrElse slideObj("events") IsNot Nothing Then
                    slideObj("layout") = "timeline"
                ElseIf slideObj("comparison") IsNot Nothing Then
                    slideObj("layout") = "comparison"
                ElseIf slideObj("matrix") IsNot Nothing Then
                    slideObj("layout") = "matrix"
                End If
            End If
        Catch ex As Newtonsoft.Json.JsonException
            Debug.WriteLine($"PowerPoint data_json parse error: {ex.Message}")
        Catch ex As System.Exception
            Debug.WriteLine($"PowerPoint data_json expansion error: {ex.Message}")
        End Try
    End Sub

    Private Shared Sub ApplyAutoPilotPowerPointPageSetup(pres As Object, args As Dictionary(Of String, Object))
        If pres Is Nothing Then Exit Sub
        Dim ratio As String = If(GetArgString(args, "aspect_ratio"), "16:9").Trim().ToLowerInvariant()
        Try
            If ratio = "4:3" OrElse ratio = "standard" Then
                pres.PageSetup.SlideWidth = 720.0F
                pres.PageSetup.SlideHeight = 540.0F
            Else
                pres.PageSetup.SlideWidth = 960.0F
                pres.PageSetup.SlideHeight = 540.0F
            End If
        Catch
        End Try
    End Sub

    Private Shared Function PptHexColor(hexColor As String, fallbackHex As String) As Integer
        Dim parsed As Integer? = ParseHexColor(hexColor)
        If parsed.HasValue Then Return parsed.Value
        parsed = ParseHexColor(fallbackHex)
        Return If(parsed.HasValue, parsed.Value, 0)
    End Function

    Private Shared Function GetPowerPointTheme(args As Dictionary(Of String, Object)) As JObject
        Dim accent As String = GetArgString(args, "accent_color")
        Dim secondary As String = GetArgString(args, "secondary_color")
        Dim fontName As String = GetArgString(args, "font_name")
        Dim textColor As String = GetArgString(args, "text_color")
        Dim mutedColor As String = GetArgString(args, "muted_color")
        Dim lightColor As String = GetArgString(args, "light_color")
        Dim lineColor As String = GetArgString(args, "line_color")
        Dim greenColor As String = GetArgString(args, "green_color")
        Dim redColor As String = GetArgString(args, "red_color")
        Dim amberColor As String = GetArgString(args, "amber_color")
        If String.IsNullOrWhiteSpace(accent) Then accent = "#17365D"
        If String.IsNullOrWhiteSpace(secondary) Then secondary = "#2F75B5"
        If String.IsNullOrWhiteSpace(fontName) Then fontName = "Aptos"
        If String.IsNullOrWhiteSpace(textColor) Then textColor = "#202124"
        If String.IsNullOrWhiteSpace(mutedColor) Then mutedColor = "#667085"
        If String.IsNullOrWhiteSpace(lightColor) Then lightColor = "#F3F6F9"
        If String.IsNullOrWhiteSpace(lineColor) Then lineColor = "#D9E2EC"
        If String.IsNullOrWhiteSpace(greenColor) Then greenColor = "#2E7D32"
        If String.IsNullOrWhiteSpace(redColor) Then redColor = "#C62828"
        If String.IsNullOrWhiteSpace(amberColor) Then amberColor = "#B7791F"
        Return New JObject From {
            {"accent", accent}, {"secondary", secondary}, {"font", fontName},
            {"text", textColor}, {"muted", mutedColor}, {"light", lightColor},
            {"line", lineColor}, {"green", greenColor}, {"red", redColor}, {"amber", amberColor}
        }
    End Function

    Private Shared Sub RenderAutoPilotPowerPointSlide(pres As Object,
                                                        sld As Object,
                                                        slideObj As JObject,
                                                        slideIndex As Integer,
                                                        existingSlideCount As Integer,
                                                        args As Dictionary(Of String, Object))
        If sld Is Nothing OrElse slideObj Is Nothing Then Exit Sub

        Dim theme As JObject = GetPowerPointTheme(args)
        Dim accent As Integer = PptHexColor(theme.Value(Of String)("accent"), "#17365D")
        Dim secondary As Integer = PptHexColor(theme.Value(Of String)("secondary"), "#2F75B5")
        Dim textColor As Integer = PptHexColor(theme.Value(Of String)("text"), "#202124")
        Dim muted As Integer = PptHexColor(theme.Value(Of String)("muted"), "#667085")
        Dim light As Integer = PptHexColor(theme.Value(Of String)("light"), "#F3F6F9")
        Dim lineColor As Integer = PptHexColor(theme.Value(Of String)("line"), "#D9E2EC")
        Dim fontName As String = theme.Value(Of String)("font")

        Dim slideW As Single = 960.0F
        Dim slideH As Single = 540.0F
        Try : slideW = CSng(pres.PageSetup.SlideWidth) : Catch : End Try
        Try : slideH = CSng(pres.PageSetup.SlideHeight) : Catch : End Try

        Dim layout As String = If(slideObj.Value(Of String)("layout"), "").Trim().ToLowerInvariant()
        If String.IsNullOrWhiteSpace(layout) Then
            If existingSlideCount = 0 AndAlso slideIndex = 1 Then layout = "title" Else layout = "bullets"
        End If

        Dim title As String = slideObj.Value(Of String)("title")
        Dim subtitle As String = slideObj.Value(Of String)("subtitle")
        Dim body As String = slideObj.Value(Of String)("body")
        Dim sourceText As String = slideObj.Value(Of String)("source")

        If layout = "title" OrElse layout = "section" OrElse layout = "closing" Then
            SetPptSlideBackground(sld, accent)
        Else
            SetPptSlideBackground(sld, PptHexColor("#FFFFFF", "#FFFFFF"))
            AddPptShape(sld, 1, 0.0F, 0.0F, slideW, 5.0F, secondary, secondary, 0.0F)
        End If

        Select Case layout
            Case "title"
                AddPptTextBox(sld, title, 58.0F, 142.0F, slideW - 116.0F, 132.0F, 35.0F, True, PptHexColor("#FFFFFF", "#FFFFFF"), fontName, 1, 0.0F)
                If Not String.IsNullOrWhiteSpace(subtitle) Then
                    AddPptTextBox(sld, subtitle, 60.0F, 286.0F, slideW - 120.0F, 76.0F, 18.0F, False, PptHexColor("#DDE7F0", "#DDE7F0"), fontName, 1, 0.0F)
                ElseIf Not String.IsNullOrWhiteSpace(body) Then
                    AddPptTextBox(sld, body, 60.0F, 286.0F, slideW - 120.0F, 76.0F, 18.0F, False, PptHexColor("#DDE7F0", "#DDE7F0"), fontName, 1, 0.0F)
                End If
                AddPptShape(sld, 1, 60.0F, 118.0F, 72.0F, 6.0F, secondary, secondary, 0.0F)

            Case "section"
                Dim sectionNumber As String = slideObj.Value(Of String)("section_number")
                If Not String.IsNullOrWhiteSpace(sectionNumber) Then
                    AddPptTextBox(sld, sectionNumber, 60.0F, 102.0F, 120.0F, 45.0F, 15.0F, True, PptHexColor("#BFD7EA", "#BFD7EA"), fontName, 1, 0.0F)
                End If
                AddPptTextBox(sld, title, 60.0F, 166.0F, slideW - 120.0F, 124.0F, 32.0F, True, PptHexColor("#FFFFFF", "#FFFFFF"), fontName, 1, 0.0F)
                If Not String.IsNullOrWhiteSpace(subtitle) Then
                    AddPptTextBox(sld, subtitle, 60.0F, 305.0F, slideW - 120.0F, 76.0F, 17.0F, False, PptHexColor("#DDE7F0", "#DDE7F0"), fontName, 1, 0.0F)
                End If

            Case "two_column"
                AddPptStandardTitle(sld, title, subtitle, slideW, fontName, textColor, muted, accent)
                RenderPptTwoColumn(sld, slideObj, slideW, fontName, textColor, muted, light, lineColor, accent)

            Case "kpi", "kpis"
                AddPptStandardTitle(sld, title, subtitle, slideW, fontName, textColor, muted, accent)
                RenderPptKpiCards(sld, TryCast(slideObj("kpis"), JArray), slideW, fontName, textColor, muted, light, lineColor, accent)

            Case "table"
                AddPptStandardTitle(sld, title, subtitle, slideW, fontName, textColor, muted, accent)
                RenderPptTable(sld, TryCast(slideObj("table"), JObject), slideW, fontName, textColor, muted, light, lineColor, accent)

            Case "chart"
                AddPptStandardTitle(sld, title, subtitle, slideW, fontName, textColor, muted, accent)
                RenderPptChart(sld, TryCast(slideObj("chart"), JObject), slideW, fontName, textColor, muted, lineColor, accent, secondary)
                Dim callout As String = slideObj.Value(Of String)("callout")
                If Not String.IsNullOrWhiteSpace(callout) Then
                    AddPptTextBox(sld, callout, slideW - 330.0F, 448.0F, 280.0F, 48.0F, 12.0F, True, accent, fontName, 3, 0.0F)
                End If

            Case "cards"
                AddPptStandardTitle(sld, title, subtitle, slideW, fontName, textColor, muted, accent)
                RenderPptCards(sld, TryCast(slideObj("cards"), JArray), slideW, fontName, textColor, muted, light, lineColor, accent, secondary)

            Case "process"
                AddPptStandardTitle(sld, title, subtitle, slideW, fontName, textColor, muted, accent)
                RenderPptProcess(sld, TryCast(slideObj("steps"), JArray), slideW, fontName, textColor, muted, light, lineColor, accent, secondary)

            Case "structure", "org", "organization"
                AddPptStandardTitle(sld, title, subtitle, slideW, fontName, textColor, muted, accent)
                RenderPptStructure(sld, TryCast(slideObj("structure"), JObject), slideW, fontName, textColor, muted, light, lineColor, accent, secondary)

            Case "timeline"
                AddPptStandardTitle(sld, title, subtitle, slideW, fontName, textColor, muted, accent)
                Dim events As JArray = TryCast(slideObj("events"), JArray)
                If events Is Nothing Then events = TryCast(slideObj("timeline"), JArray)
                If events Is Nothing Then
                    Dim timelineObj As JObject = TryCast(slideObj("timeline"), JObject)
                    If timelineObj IsNot Nothing Then events = TryCast(timelineObj("events"), JArray)
                End If
                RenderPptTimeline(sld, events, slideW, fontName, textColor, muted, light, lineColor, accent, secondary)

            Case "comparison"
                AddPptStandardTitle(sld, title, subtitle, slideW, fontName, textColor, muted, accent)
                RenderPptComparison(sld, TryCast(slideObj("comparison"), JObject), slideW, fontName, textColor, muted, light, lineColor, accent, secondary)

            Case "matrix"
                AddPptStandardTitle(sld, title, subtitle, slideW, fontName, textColor, muted, accent)
                RenderPptMatrix(sld, TryCast(slideObj("matrix"), JObject), slideW, fontName, textColor, muted, light, lineColor, accent, secondary)

            Case "quote"
                AddPptStandardTitle(sld, title, subtitle, slideW, fontName, textColor, muted, accent)
                Dim quoteText As String = slideObj.Value(Of String)("quote")
                If String.IsNullOrWhiteSpace(quoteText) Then quoteText = body
                AddPptTextBox(sld, "“" & quoteText & "”", 105.0F, 155.0F, slideW - 210.0F, 200.0F, 26.0F, False, accent, fontName, 2, 0.0F)
                Dim attribution As String = slideObj.Value(Of String)("attribution")
                If Not String.IsNullOrWhiteSpace(attribution) Then AddPptTextBox(sld, attribution, 160.0F, 375.0F, slideW - 320.0F, 45.0F, 13.0F, True, muted, fontName, 2, 0.0F)

            Case "closing"
                AddPptTextBox(sld, title, 60.0F, 174.0F, slideW - 120.0F, 110.0F, 34.0F, True, PptHexColor("#FFFFFF", "#FFFFFF"), fontName, 2, 0.0F)
                If Not String.IsNullOrWhiteSpace(subtitle) Then AddPptTextBox(sld, subtitle, 80.0F, 300.0F, slideW - 160.0F, 76.0F, 17.0F, False, PptHexColor("#DDE7F0", "#DDE7F0"), fontName, 2, 0.0F)

            Case Else
                AddPptStandardTitle(sld, title, subtitle, slideW, fontName, textColor, muted, accent)
                RenderPptBullets(sld, body, slideW, fontName, textColor, light, lineColor, accent)
        End Select

        If layout <> "title" AndAlso layout <> "section" AndAlso layout <> "closing" Then
            Dim footerText As String = GetArgString(args, "footer_text")
            If Not String.IsNullOrWhiteSpace(sourceText) Then footerText = sourceText
            AddPptFooter(sld, footerText, slideIndex, slideW, slideH, fontName, muted, GetArgBool(args, "show_slide_numbers", True))
        End If
    End Sub

    Private Shared Sub SetPptSlideBackground(sld As Object, colorValue As Integer)
        Try
            sld.FollowMasterBackground = 0
            sld.Background.Fill.Solid()
            sld.Background.Fill.ForeColor.RGB = colorValue
        Catch
        End Try
    End Sub

    Private Shared Function AddPptShape(sld As Object, shapeType As Integer, left As Single, top As Single, width As Single, height As Single,
                                         fillColor As Integer, lineColor As Integer, lineWeight As Single) As Object
        Dim shp As Object = Nothing
        Try
            shp = sld.Shapes.AddShape(shapeType, left, top, width, height)
            shp.Fill.Solid()
            shp.Fill.ForeColor.RGB = fillColor
            If lineWeight <= 0 Then
                shp.Line.Visible = 0
            Else
                shp.Line.Visible = -1
                shp.Line.ForeColor.RGB = lineColor
                shp.Line.Weight = lineWeight
            End If
        Catch
        End Try
        Return shp
    End Function

    Private Shared Function AddPptTextBox(sld As Object, text As String, left As Single, top As Single, width As Single, height As Single,
                                           fontSize As Single, bold As Boolean, fontColor As Integer, fontName As String,
                                           alignment As Integer, margin As Single) As Object
        Dim tb As Object = Nothing
        Try
            tb = sld.Shapes.AddTextbox(1, left, top, width, height)
            tb.TextFrame.WordWrap = -1
            tb.TextFrame.AutoSize = 0
            tb.TextFrame.MarginLeft = margin
            tb.TextFrame.MarginRight = margin
            tb.TextFrame.MarginTop = margin
            tb.TextFrame.MarginBottom = margin
            tb.TextFrame.TextRange.Text = If(text, "")
            tb.TextFrame.TextRange.Font.Name = fontName
            tb.TextFrame.TextRange.Font.Size = fontSize
            tb.TextFrame.TextRange.Font.Bold = If(bold, -1, 0)
            tb.TextFrame.TextRange.Font.Color.RGB = fontColor
            tb.TextFrame.TextRange.ParagraphFormat.Alignment = alignment
        Catch
        End Try
        Return tb
    End Function

    Private Shared Sub AddPptStandardTitle(sld As Object, title As String, subtitle As String, slideW As Single,
                                            fontName As String, textColor As Integer, muted As Integer, accent As Integer)
        Dim titleSize As Single = If(String.IsNullOrWhiteSpace(title) OrElse title.Length <= 70, 26.0F, 24.0F)
        AddPptTextBox(sld, title, 48.0F, 26.0F, slideW - 96.0F, 60.0F, titleSize, True, textColor, fontName, 1, 0.0F)
        AddPptShape(sld, 1, 48.0F, 96.0F, 56.0F, 4.0F, accent, accent, 0.0F)
        If Not String.IsNullOrWhiteSpace(subtitle) Then
            AddPptTextBox(sld, subtitle, 118.0F, 85.0F, slideW - 168.0F, 34.0F, 12.5F, False, muted, fontName, 1, 0.0F)
        End If
    End Sub

    Private Shared Function CleanPptBulletText(body As String) As String
        If String.IsNullOrWhiteSpace(body) Then Return ""
        Dim lines As New List(Of String)()
        For Each rawLine As String In body.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
            Dim t As String = rawLine.Trim()
            If t.StartsWith("- ") OrElse t.StartsWith("* ") OrElse t.StartsWith("+ ") Then t = t.Substring(2)
            If t.Length > 2 AndAlso Char.IsDigit(t(0)) Then
                Dim dotIdx As Integer = t.IndexOf(". ", StringComparison.Ordinal)
                If dotIdx > 0 AndAlso dotIdx <= 3 Then
                    Dim prefix As String = t.Substring(0, dotIdx)
                    If prefix.All(Function(ch As Char) Char.IsDigit(ch)) Then t = t.Substring(dotIdx + 2)
                End If
            End If
            If t.Length > 0 Then lines.Add(t)
        Next
        Return String.Join(vbCrLf, lines)
    End Function


    ''' <summary>
    ''' Chooses a readable narrative font size from content density. The renderer
    ''' intentionally prefers larger text and expects the agent to split dense content
    ''' across slides rather than shrinking narrative text to document-sized typography.
    ''' </summary>
    Private Shared Function GetPptNarrativeFontSize(text As String,
                                                     preferred As Single,
                                                     minimum As Single,
                                                     denseMinimum As Single) As Single
        If String.IsNullOrWhiteSpace(text) Then Return preferred
        Dim normalized As String = text.Replace(vbCr, "").Trim()
        Dim lines As Integer = Math.Max(1, normalized.Split({vbLf}, StringSplitOptions.None).Length)
        Dim chars As Integer = normalized.Length
        Dim density As Double = chars + Math.Max(0, lines - 1) * 35.0

        If density <= 260.0 Then Return preferred
        If density <= 430.0 Then Return Math.Max(minimum, preferred - 1.0F)
        If density <= 620.0 Then Return minimum
        Return denseMinimum
    End Function

    Private Shared Function GetPptArrayText(item As JObject, key As String) As String
        If item Is Nothing Then Return ""
        Dim token As JToken = item(key)
        If token Is Nothing OrElse token.Type = JTokenType.Null Then Return ""
        If TypeOf token Is JArray Then
            Return String.Join(vbCrLf, DirectCast(token, JArray).Select(Function(x As JToken) x.ToString()))
        End If
        Return token.ToString()
    End Function

    Private Shared Function GetPptToneColor(tone As String,
                                            accent As Integer,
                                            secondary As Integer,
                                            muted As Integer) As Integer
        Dim t As String = If(tone, "").Trim().ToLowerInvariant()
        Select Case t
            Case "positive", "good", "green", "recommended", "preferred"
                Return PptHexColor("#2E7D32", "#2E7D32")
            Case "negative", "bad", "red", "risk", "high_risk"
                Return PptHexColor("#C62828", "#C62828")
            Case "warning", "amber", "medium", "caution"
                Return PptHexColor("#B7791F", "#B7791F")
            Case "secondary", "blue"
                Return secondary
            Case "muted", "grey", "gray"
                Return muted
            Case Else
                Return accent
        End Select
    End Function

    Private Shared Sub RenderPptBullets(sld As Object, body As String, slideW As Single, fontName As String,
                                         textColor As Integer, light As Integer, lineColor As Integer, accent As Integer)
        Dim cleaned As String = CleanPptBulletText(body)
        If String.IsNullOrWhiteSpace(cleaned) Then Exit Sub
        Dim card As Object = AddPptShape(sld, 5, 48.0F, 128.0F, slideW - 96.0F, 334.0F, PptHexColor("#FFFFFF", "#FFFFFF"), lineColor, 0.75F)
        If card IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(card) : Catch : End Try

        Dim fontSize As Single = GetPptNarrativeFontSize(cleaned, 18.0F, 16.5F, 15.0F)
        Dim tb As Object = AddPptTextBox(sld, cleaned, 78.0F, 153.0F, slideW - 156.0F, 286.0F, fontSize, False, textColor, fontName, 1, 0.0F)
        If tb IsNot Nothing Then
            Try
                tb.TextFrame.TextRange.ParagraphFormat.Bullet.Visible = -1
                tb.TextFrame.TextRange.ParagraphFormat.LeftMargin = 22.0F
                tb.TextFrame.TextRange.ParagraphFormat.FirstLineIndent = -12.0F
                tb.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 14.0F
                tb.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 1.05F
            Catch
            Finally
                Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(tb) : Catch : End Try
            End Try
        End If
    End Sub

    Private Shared Sub RenderPptTwoColumn(sld As Object, slideObj As JObject, slideW As Single, fontName As String,
                                           textColor As Integer, muted As Integer, light As Integer, lineColor As Integer, accent As Integer)
        Dim gap As Single = 24.0F
        Dim left As Single = 48.0F
        Dim top As Single = 130.0F
        Dim colW As Single = (slideW - 96.0F - gap) / 2.0F
        Dim h As Single = 334.0F
        For idx As Integer = 0 To 1
            Dim x As Single = left + idx * (colW + gap)
            Dim card As Object = AddPptShape(sld, 5, x, top, colW, h, PptHexColor("#FFFFFF", "#FFFFFF"), lineColor, 0.75F)
            If card IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(card) : Catch : End Try
            Dim prefix As String = If(idx = 0, "left", "right")
            Dim colTitle As String = slideObj.Value(Of String)(prefix & "_title")
            Dim colBody As String = slideObj.Value(Of String)(prefix & "_body")
            AddPptTextBox(sld, colTitle, x + 22.0F, top + 20.0F, colW - 44.0F, 46.0F, 16.0F, True, accent, fontName, 1, 0.0F)

            Dim cleaned As String = CleanPptBulletText(colBody)
            Dim fontSize As Single = GetPptNarrativeFontSize(cleaned, 16.0F, 15.0F, 14.0F)
            Dim tb As Object = AddPptTextBox(sld, cleaned, x + 22.0F, top + 78.0F, colW - 44.0F, h - 100.0F, fontSize, False, textColor, fontName, 1, 0.0F)
            If tb IsNot Nothing Then
                Try
                    tb.TextFrame.TextRange.ParagraphFormat.Bullet.Visible = -1
                    tb.TextFrame.TextRange.ParagraphFormat.LeftMargin = 19.0F
                    tb.TextFrame.TextRange.ParagraphFormat.FirstLineIndent = -10.0F
                    tb.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 11.0F
                    tb.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 1.02F
                Catch
                Finally
                    Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(tb) : Catch : End Try
                End Try
            End If
        Next
    End Sub

    Private Shared Sub RenderPptKpiCards(sld As Object, kpis As JArray, slideW As Single, fontName As String,
                                          textColor As Integer, muted As Integer, light As Integer, lineColor As Integer, accent As Integer)
        If kpis Is Nothing OrElse kpis.Count = 0 Then Exit Sub
        Dim count As Integer = Math.Min(4, kpis.Count)
        Dim gap As Single = 18.0F
        Dim totalW As Single = slideW - 96.0F
        Dim cardW As Single = (totalW - gap * (count - 1)) / count
        Dim top As Single = 154.0F
        For i As Integer = 0 To count - 1
            Dim kpi As JObject = TryCast(kpis(i), JObject)
            If kpi Is Nothing Then Continue For
            Dim x As Single = 48.0F + i * (cardW + gap)
            Dim card As Object = AddPptShape(sld, 5, x, top, cardW, 232.0F, PptHexColor("#FFFFFF", "#FFFFFF"), lineColor, 0.8F)
            If card IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(card) : Catch : End Try
            AddPptShape(sld, 1, x, top, cardW, 6.0F, accent, accent, 0.0F)
            AddPptTextBox(sld, kpi.Value(Of String)("label"), x + 18.0F, top + 27.0F, cardW - 36.0F, 38.0F, 12.5F, True, muted, fontName, 1, 0.0F)
            AddPptTextBox(sld, kpi.Value(Of String)("value"), x + 18.0F, top + 72.0F, cardW - 36.0F, 72.0F, 28.0F, True, accent, fontName, 1, 0.0F)
            AddPptTextBox(sld, kpi.Value(Of String)("detail"), x + 18.0F, top + 154.0F, cardW - 36.0F, 58.0F, 12.0F, False, textColor, fontName, 1, 0.0F)
        Next
    End Sub

    Private Shared Sub RenderPptTable(sld As Object, tableObj As JObject, slideW As Single, fontName As String,
                                       textColor As Integer, muted As Integer, light As Integer, lineColor As Integer, accent As Integer)
        If tableObj Is Nothing Then Exit Sub
        Dim headers As JArray = TryCast(tableObj("headers"), JArray)
        Dim rows As JArray = TryCast(tableObj("rows"), JArray)
        If headers Is Nothing OrElse headers.Count = 0 Then Exit Sub

        Dim rowCount As Integer = 1 + If(rows Is Nothing, 0, rows.Count)
        Dim colCount As Integer = headers.Count

        ' Scale typography by actual table density. Short executive comparison tables should
        ' be visibly readable from a room; only genuinely dense tables use smaller text.
        Dim headerFont As Single = 13.5F
        Dim bodyFont As Single = 12.5F
        If rowCount >= 8 OrElse colCount >= 6 Then
            headerFont = 12.0F : bodyFont = 11.0F
        End If
        If rowCount >= 11 OrElse colCount >= 8 Then
            headerFont = 11.0F : bodyFont = 10.0F
        End If

        Dim tableTop As Single = 128.0F
        Dim tableHeight As Single = 350.0F
        If rowCount <= 6 Then tableHeight = 330.0F
        If rowCount <= 4 Then tableHeight = 280.0F

        Dim tableShape As Object = Nothing
        Dim tbl As Object = Nothing
        Try
            tableShape = sld.Shapes.AddTable(rowCount, colCount, 48.0F, tableTop, slideW - 96.0F, tableHeight)
            tbl = tableShape.Table

            For c As Integer = 1 To colCount
                Dim cellShape As Object = tbl.Cell(1, c).Shape
                cellShape.TextFrame.TextRange.Text = headers(c - 1).ToString()
                cellShape.Fill.Solid()
                cellShape.Fill.ForeColor.RGB = accent
                cellShape.TextFrame.TextRange.Font.Name = fontName
                cellShape.TextFrame.TextRange.Font.Size = headerFont
                cellShape.TextFrame.TextRange.Font.Bold = -1
                cellShape.TextFrame.TextRange.Font.Color.RGB = PptHexColor("#FFFFFF", "#FFFFFF")
                cellShape.TextFrame.MarginLeft = 10.0F : cellShape.TextFrame.MarginRight = 10.0F
                cellShape.TextFrame.MarginTop = 7.0F : cellShape.TextFrame.MarginBottom = 7.0F
                Try : cellShape.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 0.0F : Catch : End Try
            Next

            If rows IsNot Nothing Then
                For r As Integer = 1 To rows.Count
                    Dim rowArr As JArray = TryCast(rows(r - 1), JArray)
                    If rowArr Is Nothing Then Continue For
                    For c As Integer = 1 To colCount
                        Dim cellShape As Object = tbl.Cell(r + 1, c).Shape
                        cellShape.TextFrame.TextRange.Text = If(c - 1 < rowArr.Count, rowArr(c - 1).ToString(), "")
                        cellShape.Fill.Solid()
                        cellShape.Fill.ForeColor.RGB = If(r Mod 2 = 0, light, PptHexColor("#FFFFFF", "#FFFFFF"))
                        cellShape.TextFrame.TextRange.Font.Name = fontName
                        cellShape.TextFrame.TextRange.Font.Size = bodyFont
                        cellShape.TextFrame.TextRange.Font.Color.RGB = textColor
                        If c = 1 Then cellShape.TextFrame.TextRange.Font.Bold = -1
                        cellShape.TextFrame.MarginLeft = 9.0F : cellShape.TextFrame.MarginRight = 9.0F
                        cellShape.TextFrame.MarginTop = 6.0F : cellShape.TextFrame.MarginBottom = 6.0F
                        Try : cellShape.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 0.0F : Catch : End Try
                    Next
                Next
            End If

            For r As Integer = 1 To rowCount
                For c As Integer = 1 To colCount
                    For b As Integer = 1 To 4
                        Try
                            tbl.Cell(r, c).Borders(b).ForeColor.RGB = lineColor
                            tbl.Cell(r, c).Borders(b).Weight = 0.6F
                        Catch
                        End Try
                    Next
                Next
            Next
        Catch ex As System.Exception
            Debug.WriteLine($"PowerPoint table error: {ex.Message}")
        Finally
            If tbl IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(tbl) : Catch : End Try
            If tableShape IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(tableShape) : Catch : End Try
        End Try
    End Sub


    Private Shared Sub RenderPptCards(sld As Object,
                                      cards As JArray,
                                      slideW As Single,
                                      fontName As String,
                                      textColor As Integer,
                                      muted As Integer,
                                      light As Integer,
                                      lineColor As Integer,
                                      accent As Integer,
                                      secondary As Integer)
        If cards Is Nothing OrElse cards.Count = 0 Then Exit Sub

        Dim count As Integer = Math.Min(6, cards.Count)
        Dim columns As Integer = If(count <= 3, count, 3)
        Dim rows As Integer = CInt(Math.Ceiling(count / CDbl(columns)))
        Dim gapX As Single = 18.0F
        Dim gapY As Single = 18.0F
        Dim left As Single = 48.0F
        Dim top As Single = 135.0F
        Dim totalW As Single = slideW - 96.0F
        Dim cardW As Single = (totalW - gapX * (columns - 1)) / columns
        Dim availableH As Single = 324.0F
        Dim cardH As Single = (availableH - gapY * (rows - 1)) / rows
        Dim palette() As Integer = {accent, secondary, PptHexColor("#61758A", "#61758A"), PptHexColor("#2E7D32", "#2E7D32"), PptHexColor("#B7791F", "#B7791F"), PptHexColor("#7A5C99", "#7A5C99")}

        For i As Integer = 0 To count - 1
            Dim cardObj As JObject = TryCast(cards(i), JObject)
            If cardObj Is Nothing Then Continue For
            Dim rowIndex As Integer = i \ columns
            Dim colIndex As Integer = i Mod columns
            Dim x As Single = left + colIndex * (cardW + gapX)
            Dim y As Single = top + rowIndex * (cardH + gapY)
            Dim toneColor As Integer = GetPptToneColor(cardObj.Value(Of String)("tone"), palette(i Mod palette.Length), secondary, muted)
            Dim explicitColor As Integer = PptHexColor(cardObj.Value(Of String)("color"), "")
            If explicitColor <> 0 Then toneColor = explicitColor

            Dim card As Object = AddPptShape(sld, 5, x, y, cardW, cardH, PptHexColor("#FFFFFF", "#FFFFFF"), lineColor, 0.75F)
            If card IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(card) : Catch : End Try
            AddPptShape(sld, 1, x, y, cardW, 6.0F, toneColor, toneColor, 0.0F)

            Dim badge As String = cardObj.Value(Of String)("badge")
            If Not String.IsNullOrWhiteSpace(badge) Then
                AddPptTextBox(sld, badge, x + 18.0F, y + 22.0F, 36.0F, 28.0F, 12.0F, True, toneColor, fontName, 1, 0.0F)
            End If

            Dim titleLeft As Single = x + 18.0F
            If Not String.IsNullOrWhiteSpace(badge) Then titleLeft = x + 56.0F
            AddPptTextBox(sld, cardObj.Value(Of String)("title"), titleLeft, y + 20.0F, x + cardW - 18.0F - titleLeft, 42.0F, 15.5F, True, toneColor, fontName, 1, 0.0F)

            Dim body As String = GetPptArrayText(cardObj, "body")
            If String.IsNullOrWhiteSpace(body) Then body = GetPptArrayText(cardObj, "items")
            Dim bodySize As Single = GetPptNarrativeFontSize(body, If(rows = 1, 15.0F, 14.0F), 13.0F, 12.0F)
            AddPptTextBox(sld, body, x + 18.0F, y + 69.0F, cardW - 36.0F, cardH - 86.0F, bodySize, False, textColor, fontName, 1, 0.0F)
        Next
    End Sub

    Private Shared Sub RenderPptProcess(sld As Object,
                                        steps As JArray,
                                        slideW As Single,
                                        fontName As String,
                                        textColor As Integer,
                                        muted As Integer,
                                        light As Integer,
                                        lineColor As Integer,
                                        accent As Integer,
                                        secondary As Integer)
        If steps Is Nothing OrElse steps.Count = 0 Then Exit Sub

        Dim count As Integer = Math.Min(5, steps.Count)
        Dim left As Single = 48.0F
        Dim top As Single = 165.0F
        Dim gap As Single = 28.0F
        Dim totalW As Single = slideW - 96.0F
        Dim cardW As Single = (totalW - gap * (count - 1)) / count
        Dim cardH As Single = 230.0F

        For i As Integer = 0 To count - 1
            Dim stepObj As JObject = TryCast(steps(i), JObject)
            If stepObj Is Nothing Then Continue For
            Dim x As Single = left + i * (cardW + gap)
            Dim card As Object = AddPptShape(sld, 5, x, top, cardW, cardH, PptHexColor("#FFFFFF", "#FFFFFF"), lineColor, 0.8F)
            If card IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(card) : Catch : End Try

            Dim numColor As Integer = If(i Mod 2 = 0, accent, secondary)
            Dim circle As Object = AddPptShape(sld, 9, x + cardW / 2.0F - 20.0F, top - 24.0F, 40.0F, 40.0F, numColor, numColor, 0.0F)
            If circle IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(circle) : Catch : End Try
            AddPptTextBox(sld, (i + 1).ToString(Globalization.CultureInfo.InvariantCulture), x + cardW / 2.0F - 16.0F, top - 17.0F, 32.0F, 24.0F, 12.0F, True, PptHexColor("#FFFFFF", "#FFFFFF"), fontName, 2, 0.0F)

            AddPptTextBox(sld, stepObj.Value(Of String)("title"), x + 16.0F, top + 35.0F, cardW - 32.0F, 52.0F, 15.0F, True, numColor, fontName, 2, 0.0F)
            Dim body As String = GetPptArrayText(stepObj, "body")
            If String.IsNullOrWhiteSpace(body) Then body = GetPptArrayText(stepObj, "detail")
            AddPptTextBox(sld, body, x + 16.0F, top + 96.0F, cardW - 32.0F, cardH - 116.0F, 13.0F, False, textColor, fontName, 2, 0.0F)

            If i < count - 1 Then
                AddPptTextBox(sld, "→", x + cardW + 3.0F, top + 90.0F, gap - 6.0F, 45.0F, 23.0F, True, muted, fontName, 2, 0.0F)
            End If
        Next
    End Sub

    Private Shared Sub RenderPptStructure(sld As Object,
                                          structureObj As JObject,
                                          slideW As Single,
                                          fontName As String,
                                          textColor As Integer,
                                          muted As Integer,
                                          light As Integer,
                                          lineColor As Integer,
                                          accent As Integer,
                                          secondary As Integer)
        If structureObj Is Nothing Then Exit Sub

        Dim topObj As JObject = TryCast(structureObj("top"), JObject)
        If topObj Is Nothing Then topObj = TryCast(structureObj("parent"), JObject)
        Dim children As JArray = TryCast(structureObj("children"), JArray)
        If topObj Is Nothing Then Exit Sub

        Dim topW As Single = Math.Min(360.0F, slideW - 220.0F)
        Dim topLeft As Single = (slideW - topW) / 2.0F
        Dim topY As Single = 140.0F
        Dim topH As Single = 112.0F

        Dim topCard As Object = AddPptShape(sld, 5, topLeft, topY, topW, topH, accent, accent, 0.0F)
        If topCard IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(topCard) : Catch : End Try
        AddPptTextBox(sld, topObj.Value(Of String)("title"), topLeft + 18.0F, topY + 20.0F, topW - 36.0F, 36.0F, 17.0F, True, PptHexColor("#FFFFFF", "#FFFFFF"), fontName, 2, 0.0F)
        AddPptTextBox(sld, GetPptArrayText(topObj, "body"), topLeft + 18.0F, topY + 59.0F, topW - 36.0F, 38.0F, 12.5F, False, PptHexColor("#E6EEF6", "#E6EEF6"), fontName, 2, 0.0F)

        If children Is Nothing OrElse children.Count = 0 Then Exit Sub
        Dim count As Integer = Math.Min(4, children.Count)
        Dim gap As Single = 20.0F
        Dim totalW As Single = slideW - 120.0F
        Dim childW As Single = (totalW - gap * (count - 1)) / count
        Dim childY As Single = 330.0F
        Dim childH As Single = 116.0F
        Dim startX As Single = (slideW - totalW) / 2.0F

        Dim trunk As Object = Nothing
        Try
            trunk = sld.Shapes.AddLine(slideW / 2.0F, topY + topH, slideW / 2.0F, childY - 38.0F)
            trunk.Line.ForeColor.RGB = lineColor
            trunk.Line.Weight = 1.5F
        Finally
            If trunk IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(trunk) : Catch : End Try
        End Try

        Dim firstCenter As Single = startX + childW / 2.0F
        Dim lastCenter As Single = startX + (count - 1) * (childW + gap) + childW / 2.0F
        Dim branch As Object = Nothing
        Try
            branch = sld.Shapes.AddLine(firstCenter, childY - 38.0F, lastCenter, childY - 38.0F)
            branch.Line.ForeColor.RGB = lineColor
            branch.Line.Weight = 1.5F
        Finally
            If branch IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(branch) : Catch : End Try
        End Try

        For i As Integer = 0 To count - 1
            Dim childObj As JObject = TryCast(children(i), JObject)
            If childObj Is Nothing Then Continue For
            Dim x As Single = startX + i * (childW + gap)
            Dim centerX As Single = x + childW / 2.0F
            Dim connector As Object = Nothing
            Try
                connector = sld.Shapes.AddLine(centerX, childY - 38.0F, centerX, childY)
                connector.Line.ForeColor.RGB = lineColor
                connector.Line.Weight = 1.5F
            Finally
                If connector IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(connector) : Catch : End Try
            End Try

            Dim card As Object = AddPptShape(sld, 5, x, childY, childW, childH, PptHexColor("#FFFFFF", "#FFFFFF"), lineColor, 0.8F)
            If card IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(card) : Catch : End Try
            AddPptShape(sld, 1, x, childY, childW, 5.0F, secondary, secondary, 0.0F)
            AddPptTextBox(sld, childObj.Value(Of String)("title"), x + 14.0F, childY + 18.0F, childW - 28.0F, 36.0F, 14.5F, True, secondary, fontName, 2, 0.0F)
            AddPptTextBox(sld, GetPptArrayText(childObj, "body"), x + 14.0F, childY + 58.0F, childW - 28.0F, 44.0F, 11.5F, False, textColor, fontName, 2, 0.0F)
        Next
    End Sub

    Private Shared Sub RenderPptTimeline(sld As Object,
                                         events As JArray,
                                         slideW As Single,
                                         fontName As String,
                                         textColor As Integer,
                                         muted As Integer,
                                         light As Integer,
                                         lineColor As Integer,
                                         accent As Integer,
                                         secondary As Integer)
        If events Is Nothing OrElse events.Count = 0 Then Exit Sub
        Dim count As Integer = Math.Min(6, events.Count)
        Dim left As Single = 75.0F
        Dim right As Single = slideW - 75.0F
        Dim lineY As Single = 255.0F
        Dim axis As Object = Nothing
        Try
            axis = sld.Shapes.AddLine(left, lineY, right, lineY)
            axis.Line.ForeColor.RGB = lineColor
            axis.Line.Weight = 2.0F
        Finally
            If axis IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(axis) : Catch : End Try
        End Try

        Dim stepW As Single = If(count <= 1, 0.0F, (right - left) / (count - 1))
        For i As Integer = 0 To count - 1
            Dim eventObj As JObject = TryCast(events(i), JObject)
            If eventObj Is Nothing Then Continue For
            Dim x As Single = If(count <= 1, slideW / 2.0F, left + i * stepW)
            Dim colorValue As Integer = If(i Mod 2 = 0, accent, secondary)
            Dim dot As Object = AddPptShape(sld, 9, x - 9.0F, lineY - 9.0F, 18.0F, 18.0F, colorValue, colorValue, 0.0F)
            If dot IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(dot) : Catch : End Try

            Dim above As Boolean = (i Mod 2 = 0)
            Dim labelY As Single = If(above, lineY - 121.0F, lineY + 32.0F)
            Dim titleY As Single = If(above, lineY - 92.0F, lineY + 60.0F)
            Dim bodyY As Single = If(above, lineY - 58.0F, lineY + 94.0F)
            Dim boxW As Single = Math.Min(150.0F, Math.Max(110.0F, stepW + 18.0F))
            AddPptTextBox(sld, eventObj.Value(Of String)("label"), x - boxW / 2.0F, labelY, boxW, 24.0F, 11.0F, True, colorValue, fontName, 2, 0.0F)
            AddPptTextBox(sld, eventObj.Value(Of String)("title"), x - boxW / 2.0F, titleY, boxW, 32.0F, 13.0F, True, textColor, fontName, 2, 0.0F)
            AddPptTextBox(sld, GetPptArrayText(eventObj, "body"), x - boxW / 2.0F, bodyY, boxW, 48.0F, 10.5F, False, muted, fontName, 2, 0.0F)
        Next
    End Sub

    Private Shared Sub RenderPptComparison(sld As Object,
                                           comparisonObj As JObject,
                                           slideW As Single,
                                           fontName As String,
                                           textColor As Integer,
                                           muted As Integer,
                                           light As Integer,
                                           lineColor As Integer,
                                           accent As Integer,
                                           secondary As Integer)
        If comparisonObj Is Nothing Then Exit Sub
        Dim columns As JArray = TryCast(comparisonObj("columns"), JArray)
        If columns Is Nothing OrElse columns.Count = 0 Then Exit Sub
        Dim count As Integer = Math.Min(3, columns.Count)
        Dim gap As Single = 22.0F
        Dim left As Single = 48.0F
        Dim top As Single = 132.0F
        Dim totalW As Single = slideW - 96.0F
        Dim colW As Single = (totalW - gap * (count - 1)) / count
        Dim h As Single = 328.0F

        For i As Integer = 0 To count - 1
            Dim colObj As JObject = TryCast(columns(i), JObject)
            If colObj Is Nothing Then Continue For
            Dim x As Single = left + i * (colW + gap)
            Dim toneColor As Integer = GetPptToneColor(colObj.Value(Of String)("tone"), If(i = 0, accent, secondary), secondary, muted)
            Dim card As Object = AddPptShape(sld, 5, x, top, colW, h, PptHexColor("#FFFFFF", "#FFFFFF"), lineColor, 0.8F)
            If card IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(card) : Catch : End Try
            AddPptShape(sld, 1, x, top, colW, 7.0F, toneColor, toneColor, 0.0F)
            AddPptTextBox(sld, colObj.Value(Of String)("title"), x + 18.0F, top + 23.0F, colW - 36.0F, 50.0F, 16.0F, True, toneColor, fontName, 1, 0.0F)

            Dim items As String = GetPptArrayText(colObj, "items")
            Dim itemBox As Object = AddPptTextBox(sld, CleanPptBulletText(items), x + 18.0F, top + 82.0F, colW - 36.0F, 180.0F, 13.5F, False, textColor, fontName, 1, 0.0F)
            If itemBox IsNot Nothing Then
                Try
                    itemBox.TextFrame.TextRange.ParagraphFormat.Bullet.Visible = -1
                    itemBox.TextFrame.TextRange.ParagraphFormat.LeftMargin = 18.0F
                    itemBox.TextFrame.TextRange.ParagraphFormat.FirstLineIndent = -9.0F
                    itemBox.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 9.0F
                Catch
                Finally
                    Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(itemBox) : Catch : End Try
                End Try
            End If

            Dim verdict As String = colObj.Value(Of String)("verdict")
            If Not String.IsNullOrWhiteSpace(verdict) Then
                Dim badge As Object = AddPptShape(sld, 5, x + 18.0F, top + h - 54.0F, colW - 36.0F, 36.0F, light, light, 0.0F)
                If badge IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(badge) : Catch : End Try
                AddPptTextBox(sld, verdict, x + 25.0F, top + h - 47.0F, colW - 50.0F, 24.0F, 11.5F, True, toneColor, fontName, 2, 0.0F)
            End If
        Next
    End Sub

    Private Shared Sub RenderPptMatrix(sld As Object,
                                       matrixObj As JObject,
                                       slideW As Single,
                                       fontName As String,
                                       textColor As Integer,
                                       muted As Integer,
                                       light As Integer,
                                       lineColor As Integer,
                                       accent As Integer,
                                       secondary As Integer)
        If matrixObj Is Nothing Then Exit Sub
        Dim quadrants As JArray = TryCast(matrixObj("quadrants"), JArray)
        If quadrants Is Nothing OrElse quadrants.Count = 0 Then Exit Sub

        Dim left As Single = 120.0F
        Dim top As Single = 142.0F
        Dim gridW As Single = slideW - 190.0F
        Dim gridH As Single = 300.0F
        Dim gap As Single = 8.0F
        Dim qW As Single = (gridW - gap) / 2.0F
        Dim qH As Single = (gridH - gap) / 2.0F

        For i As Integer = 0 To Math.Min(4, quadrants.Count) - 1
            Dim qObj As JObject = TryCast(quadrants(i), JObject)
            If qObj Is Nothing Then Continue For
            Dim row As Integer = i \ 2
            Dim col As Integer = i Mod 2
            Dim x As Single = left + col * (qW + gap)
            Dim y As Single = top + row * (qH + gap)
            Dim fill As Integer = If(i = 1, PptHexColor("#EEF4FA", "#EEF4FA"), If(i = 2, PptHexColor("#F4F7F9", "#F4F7F9"), PptHexColor("#FFFFFF", "#FFFFFF")))
            Dim card As Object = AddPptShape(sld, 5, x, y, qW, qH, fill, lineColor, 0.75F)
            If card IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(card) : Catch : End Try
            Dim qColor As Integer = If(i Mod 2 = 0, accent, secondary)
            AddPptTextBox(sld, qObj.Value(Of String)("title"), x + 18.0F, y + 17.0F, qW - 36.0F, 35.0F, 14.5F, True, qColor, fontName, 1, 0.0F)
            AddPptTextBox(sld, GetPptArrayText(qObj, "body"), x + 18.0F, y + 56.0F, qW - 36.0F, qH - 70.0F, 12.0F, False, textColor, fontName, 1, 0.0F)
        Next

        AddPptTextBox(sld, matrixObj.Value(Of String)("x_left"), left, top + gridH + 8.0F, gridW / 2.0F, 22.0F, 10.0F, True, muted, fontName, 1, 0.0F)
        AddPptTextBox(sld, matrixObj.Value(Of String)("x_right"), left + gridW / 2.0F, top + gridH + 8.0F, gridW / 2.0F, 22.0F, 10.0F, True, muted, fontName, 3, 0.0F)
        AddPptTextBox(sld, matrixObj.Value(Of String)("y_top"), 48.0F, top + 12.0F, 65.0F, 30.0F, 10.0F, True, muted, fontName, 3, 0.0F)
        AddPptTextBox(sld, matrixObj.Value(Of String)("y_bottom"), 48.0F, top + gridH - 34.0F, 65.0F, 30.0F, 10.0F, True, muted, fontName, 3, 0.0F)
    End Sub

    Private Shared Sub RenderPptChart(sld As Object, chartObj As JObject, slideW As Single, fontName As String,
                                       textColor As Integer, muted As Integer, lineColor As Integer, accent As Integer, secondary As Integer)
        If chartObj Is Nothing Then Exit Sub
        Dim categories As JArray = TryCast(chartObj("categories"), JArray)
        Dim series As JArray = TryCast(chartObj("series"), JArray)
        If categories Is Nothing OrElse categories.Count = 0 OrElse series Is Nothing OrElse series.Count = 0 Then Exit Sub

        Dim chartType As String = If(chartObj.Value(Of String)("type"), "column").Trim().ToLowerInvariant()
        Dim left As Single = 92.0F, top As Single = 158.0F, width As Single = slideW - 180.0F, height As Single = 270.0F
        Dim palette() As Integer = {accent, secondary, PptHexColor("#7F8C8D", "#7F8C8D"), PptHexColor("#D98E04", "#D98E04")}
        Dim maxVal As Double = 0.0
        For Each ser As JObject In series.OfType(Of JObject)()
            Dim vals As JArray = TryCast(ser("values"), JArray)
            If vals Is Nothing Then Continue For
            For Each v As JToken In vals
                Dim d As Double
                If Double.TryParse(v.ToString(), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, d) Then maxVal = Math.Max(maxVal, Math.Abs(d))
            Next
        Next
        If maxVal <= 0 Then maxVal = 1.0

        Dim yAxis As Object = Nothing, xAxis As Object = Nothing
        Try
            yAxis = sld.Shapes.AddLine(left, top, left, top + height)
            yAxis.Line.ForeColor.RGB = lineColor : yAxis.Line.Weight = 0.9F
            xAxis = sld.Shapes.AddLine(left, top + height, left + width, top + height)
            xAxis.Line.ForeColor.RGB = lineColor : xAxis.Line.Weight = 0.9F
        Finally
            If yAxis IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(yAxis) : Catch : End Try
            If xAxis IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xAxis) : Catch : End Try
        End Try

        Dim catCount As Integer = categories.Count
        Dim serCount As Integer = Math.Min(4, series.Count)
        If chartType = "bar" Then
            Dim bandH As Single = height / Math.Max(1, catCount)
            For ci As Integer = 0 To catCount - 1
                AddPptTextBox(sld, categories(ci).ToString(), 16.0F, top + ci * bandH + 3.0F, 70.0F, bandH - 4.0F, 10.5F, False, muted, fontName, 3, 0.0F)
                For si As Integer = 0 To serCount - 1
                    Dim ser As JObject = TryCast(series(si), JObject)
                    Dim vals As JArray = If(ser Is Nothing, Nothing, TryCast(ser("values"), JArray))
                    If vals Is Nothing OrElse ci >= vals.Count Then Continue For
                    Dim d As Double
                    If Not Double.TryParse(vals(ci).ToString(), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, d) Then Continue For
                    Dim barH As Single = Math.Max(7.0F, (bandH - 8.0F) / serCount)
                    Dim barTop As Single = top + ci * bandH + 3.0F + si * barH
                    Dim barW As Single = CSng(Math.Abs(d) / maxVal * width * 0.88)
                    Dim colorValue As Integer = PptHexColor(If(ser Is Nothing, Nothing, ser.Value(Of String)("color")), "")
                    If colorValue = 0 Then colorValue = palette(si Mod palette.Length)
                    Dim bar As Object = AddPptShape(sld, 1, left + 1.0F, barTop, barW, Math.Max(5.0F, barH - 2.0F), colorValue, colorValue, 0.0F)
                    If bar IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(bar) : Catch : End Try
                    If GetJBool(chartObj, "show_values") Then AddPptTextBox(sld, vals(ci).ToString(), left + barW + 8.0F, barTop - 2.0F, 58.0F, 22.0F, 10.0F, True, muted, fontName, 1, 0.0F)
                Next
            Next
        ElseIf chartType = "line" Then
            For si As Integer = 0 To serCount - 1
                Dim ser As JObject = TryCast(series(si), JObject)
                Dim vals As JArray = If(ser Is Nothing, Nothing, TryCast(ser("values"), JArray))
                If vals Is Nothing Then Continue For
                Dim colorValue As Integer = PptHexColor(If(ser Is Nothing, Nothing, ser.Value(Of String)("color")), "")
                If colorValue = 0 Then colorValue = palette(si Mod palette.Length)
                Dim prevX As Single = 0.0F, prevY As Single = 0.0F, hasPrev As Boolean = False
                For ci As Integer = 0 To Math.Min(catCount, vals.Count) - 1
                    Dim d As Double
                    If Not Double.TryParse(vals(ci).ToString(), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, d) Then Continue For
                    Dim x As Single = left + If(catCount <= 1, 0.0F, CSng(ci / CDbl(catCount - 1)) * width)
                    Dim y As Single = top + height - CSng(Math.Max(0.0, d) / maxVal * height * 0.9)
                    If hasPrev Then
                        Dim ln As Object = sld.Shapes.AddLine(prevX, prevY, x, y)
                        Try : ln.Line.ForeColor.RGB = colorValue : ln.Line.Weight = 2.4F : Catch : End Try
                        Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(ln) : Catch : End Try
                    End If
                    Dim pt As Object = AddPptShape(sld, 9, x - 5.0F, y - 5.0F, 10.0F, 10.0F, colorValue, colorValue, 0.0F)
                    If pt IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pt) : Catch : End Try
                    prevX = x : prevY = y : hasPrev = True
                Next
            Next
            For ci As Integer = 0 To catCount - 1
                Dim x As Single = left + If(catCount <= 1, 0.0F, CSng(ci / CDbl(catCount - 1)) * width)
                AddPptTextBox(sld, categories(ci).ToString(), x - 46.0F, top + height + 8.0F, 92.0F, 28.0F, 10.0F, False, muted, fontName, 2, 0.0F)
            Next
        Else
            Dim groupW As Single = width / Math.Max(1, catCount)
            Dim barW As Single = Math.Max(8.0F, Math.Min(38.0F, (groupW * 0.72F) / serCount))
            For ci As Integer = 0 To catCount - 1
                For si As Integer = 0 To serCount - 1
                    Dim ser As JObject = TryCast(series(si), JObject)
                    Dim vals As JArray = If(ser Is Nothing, Nothing, TryCast(ser("values"), JArray))
                    If vals Is Nothing OrElse ci >= vals.Count Then Continue For
                    Dim d As Double
                    If Not Double.TryParse(vals(ci).ToString(), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, d) Then Continue For
                    Dim barH As Single = CSng(Math.Abs(d) / maxVal * height * 0.9)
                    Dim x As Single = left + ci * groupW + (groupW - barW * serCount) / 2.0F + si * barW
                    Dim y As Single = top + height - barH
                    Dim colorValue As Integer = PptHexColor(If(ser Is Nothing, Nothing, ser.Value(Of String)("color")), "")
                    If colorValue = 0 Then colorValue = palette(si Mod palette.Length)
                    Dim bar As Object = AddPptShape(sld, 1, x, y, barW - 2.0F, barH, colorValue, colorValue, 0.0F)
                    If bar IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(bar) : Catch : End Try
                    If GetJBool(chartObj, "show_values") Then AddPptTextBox(sld, vals(ci).ToString(), x - 10.0F, y - 24.0F, barW + 20.0F, 20.0F, 9.5F, True, muted, fontName, 2, 0.0F)
                Next
                AddPptTextBox(sld, categories(ci).ToString(), left + ci * groupW, top + height + 8.0F, groupW, 28.0F, 10.0F, False, muted, fontName, 2, 0.0F)
            Next
        End If

        If GetJBool(chartObj, "show_legend") OrElse serCount > 1 Then
            Dim lx As Single = left
            For si As Integer = 0 To serCount - 1
                Dim ser As JObject = TryCast(series(si), JObject)
                If ser Is Nothing Then Continue For
                Dim colorValue As Integer = PptHexColor(ser.Value(Of String)("color"), "")
                If colorValue = 0 Then colorValue = palette(si Mod palette.Length)
                Dim dot As Object = AddPptShape(sld, 9, lx, 128.0F, 9.0F, 9.0F, colorValue, colorValue, 0.0F)
                If dot IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(dot) : Catch : End Try
                AddPptTextBox(sld, ser.Value(Of String)("name"), lx + 14.0F, 121.0F, 118.0F, 24.0F, 10.0F, False, muted, fontName, 1, 0.0F)
                lx += 135.0F
            Next
        End If
    End Sub

    Private Shared Sub AddPptFooter(sld As Object, footerText As String, slideIndex As Integer, slideW As Single, slideH As Single,
                                     fontName As String, muted As Integer, showSlideNumber As Boolean)
        If Not String.IsNullOrWhiteSpace(footerText) Then
            AddPptTextBox(sld, footerText, 48.0F, slideH - 29.0F, slideW - 130.0F, 17.0F, 8.0F, False, muted, fontName, 1, 0.0F)
        End If
        If showSlideNumber Then AddPptTextBox(sld, slideIndex.ToString(Globalization.CultureInfo.InvariantCulture), slideW - 70.0F, slideH - 30.0F, 30.0F, 17.0F, 8.0F, False, muted, fontName, 3, 0.0F)
    End Sub

    Private Shared Sub ApplyPowerPointNotes(sld As Object, notes As String)
        If sld Is Nothing OrElse String.IsNullOrWhiteSpace(notes) Then Exit Sub
        Dim notesPage As Object = Nothing
        Dim notesShapes As Object = Nothing
        Try
            notesPage = sld.NotesPage
            notesShapes = notesPage.Shapes
            Dim nCount As Integer = System.Convert.ToInt32(notesShapes.Count, Globalization.CultureInfo.InvariantCulture)
            For k As Integer = 1 To nCount
                Dim nShp As Object = notesShapes(k)
                Try
                    Dim phType As Integer = System.Convert.ToInt32(nShp.PlaceholderFormat.Type, Globalization.CultureInfo.InvariantCulture)
                    If phType = 2 Then
                        nShp.TextFrame.TextRange.Text = notes
                        Exit For
                    End If
                Catch
                Finally
                    Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(nShp) : Catch : End Try
                End Try
            Next
        Catch
        Finally
            If notesShapes IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(notesShapes) : Catch : End Try
            If notesPage IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(notesPage) : Catch : End Try
        End Try
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: create_excel_spreadsheet
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Async Function ExecuteCreateExcelTool(
            toolCall As ToolCall, context As ToolExecutionContext, ct As CancellationToken) As System.Threading.Tasks.Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            ' ── Resolve sheet definitions ──
            ' Support both: top-level "cells" (single sheet) and "sheets" array (multi-sheet)
            Dim sheetDefs As New List(Of (SheetName As String, Cells As JArray))()
            ' Parallel list holding each sheet's source JObject (Nothing for single-sheet mode),
            ' used to resolve per-sheet formatting/appearance overrides.
            Dim sheetObjs As New List(Of JObject)()
            Dim hasVba As Boolean = False

            ' Check for VBA modules — determines .xlsm vs .xlsx
            Dim vbaModules As JArray = Nothing
            If toolCall.Arguments IsNot Nothing AndAlso toolCall.Arguments.ContainsKey("vba_modules") Then
                Dim vbaObj = toolCall.Arguments("vba_modules")
                If TypeOf vbaObj Is JArray AndAlso DirectCast(vbaObj, JArray).Count > 0 Then
                    vbaModules = DirectCast(vbaObj, JArray)
                    hasVba = True
                End If
            End If

            ' Multi-sheet mode
            If toolCall.Arguments IsNot Nothing AndAlso toolCall.Arguments.ContainsKey("sheets") Then
                Dim sheetsObj = toolCall.Arguments("sheets")
                If TypeOf sheetsObj Is JArray Then
                    For Each sheetObj As JObject In DirectCast(sheetsObj, JArray)
                        Dim sName = sheetObj.Value(Of String)("name")
                        If String.IsNullOrWhiteSpace(sName) Then sName = $"Sheet{sheetDefs.Count + 1}"
                        Dim sCells As JArray = Nothing
                        Dim sCellsToken = sheetObj("cells")
                        If TypeOf sCellsToken Is JArray Then sCells = DirectCast(sCellsToken, JArray)
                        If sCells IsNot Nothing AndAlso sCells.Count > 0 Then
                            sheetDefs.Add((sName, sCells))
                            sheetObjs.Add(sheetObj)
                        End If
                    Next
                End If
            End If

            ' Single-sheet mode (backward compatible)
            If sheetDefs.Count = 0 Then
                Dim cellsArray As JArray = Nothing
                If toolCall.Arguments IsNot Nothing AndAlso toolCall.Arguments.ContainsKey("cells") Then
                    Dim cellsObj = toolCall.Arguments("cells")
                    If TypeOf cellsObj Is JArray Then cellsArray = DirectCast(cellsObj, JArray)
                End If

                If cellsArray Is Nothing OrElse cellsArray.Count = 0 Then
                    response.Success = False
                    response.Response = "Missing required parameter: cells or sheets (must contain at least one non-empty cell array)"
                    Return response
                End If

                Dim sheetName = GetArgString(toolCall.Arguments, "sheet_name")
                If String.IsNullOrWhiteSpace(sheetName) Then sheetName = "Sheet1"
                sheetDefs.Add((sheetName, cellsArray))
                sheetObjs.Add(Nothing)
            End If

            Dim design As AutoPilotDesignResolution = ResolveAutoPilotDocumentDesign(
                toolCall.Arguments,
                "Excel",
                New String() {"professional_layout", "style_preset", "accent_color", "secondary_color", "font_name", "smart_format", "header_row", "show_gridlines", "zoom", "text_color", "light_color", "line_color", "band_color"},
                New String() {".xltx"},
                context)

            ' ── Determine file name and extension ──
            Dim fileName = GetArgString(toolCall.Arguments, "file_name")
            If String.IsNullOrWhiteSpace(fileName) Then fileName = "Spreadsheet"
            For Each c In Path.GetInvalidFileNameChars()
                fileName = fileName.Replace(c, "_"c)
            Next

            Dim fileExt As String = If(hasVba, ".xlsm", ".xlsx")
            If Not fileName.EndsWith(fileExt, StringComparison.OrdinalIgnoreCase) Then
                ' Strip wrong extension if present
                If fileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) OrElse
                   fileName.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase) Then
                    fileName = Path.GetFileNameWithoutExtension(fileName)
                End If
                fileName &= fileExt
            End If

            Dim outputPath = Path.Combine(_apCurrentTempDir, fileName)
            Dim counter = 1
            While File.Exists(outputPath)
                Dim baseName = Path.GetFileNameWithoutExtension(fileName)
                fileName = baseName & $"_{counter}{fileExt}"
                outputPath = Path.Combine(_apCurrentTempDir, fileName)
                counter += 1
            End While

            ' ── Parse shared parameters ──
            Dim columnWidths As Dictionary(Of String, Double) = ParseColumnWidths(toolCall.Arguments)
            Dim rowHeights As Dictionary(Of Integer, Double) = ParseRowHeights(toolCall.Arguments)
            Dim mergeRanges = GetArgStringArray(toolCall.Arguments, "merge_ranges")
            Dim freezePane = GetArgString(toolCall.Arguments, "freeze_pane")
            Dim autoFilter = GetArgString(toolCall.Arguments, "auto_filter")
            Dim dataValidations = ParseJsonArray(toolCall.Arguments, "data_validations")
            Dim conditionalFormats = ParseJsonArray(toolCall.Arguments, "conditional_formats")
            Dim charts = ParseJsonArray(toolCall.Arguments, "charts")
            Dim tables = ParseJsonArray(toolCall.Arguments, "tables")
            Dim namedRanges = ParseJsonArray(toolCall.Arguments, "named_ranges")
            Dim printSetup As JObject = Nothing
            If toolCall.Arguments IsNot Nothing AndAlso toolCall.Arguments.ContainsKey("print_setup") Then
                Dim psObj = toolCall.Arguments("print_setup")
                If TypeOf psObj Is JObject Then printSetup = DirectCast(psObj, JObject)
            End If

            ' Local copy of arguments for use inside the worksheet-building lambda
            ' (worksheet appearance and auto-fit settings are read directly from here).
            Dim excelArgs As Dictionary(Of String, Object) = toolCall.Arguments

            Dim totalCells = sheetDefs.Sum(Function(sd) sd.Cells.Count)
            context.Log($"Creating Excel spreadsheet: {fileName} ({sheetDefs.Count} sheet(s), {totalCells} cells)")
            ApDashboardLog($"📊 Creating Excel: {fileName} ({sheetDefs.Count} sheet(s))", "step")

            ' xlOpenXMLWorkbook = 51, xlOpenXMLWorkbookMacroEnabled = 52
            Const xlOpenXMLWorkbook As Integer = 51
            Const xlOpenXMLWorkbookMacroEnabled As Integer = 52

            Dim success = Await SwitchToUi(Function()
                                               Dim excelApp As Microsoft.Office.Interop.Excel.Application = Nothing
                                               Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
                                               Dim weOwnApp As Boolean = False
                                               Try
                                                   ' Try to reuse an existing Excel instance
                                                   Try
                                                       excelApp = CType(System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application"),
                                                                        Microsoft.Office.Interop.Excel.Application)
                                                   Catch ex As System.Runtime.InteropServices.COMException
                                                       excelApp = New Microsoft.Office.Interop.Excel.Application()
                                                       weOwnApp = True
                                                   End Try

                                                   excelApp.Visible = False
                                                   excelApp.DisplayAlerts = False
                                                   excelApp.ScreenUpdating = False

                                                   If design IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(design.TemplatePath) Then
                                                       wb = excelApp.Workbooks.Add(design.TemplatePath)

                                                       ' A repository workbook is a design carrier, not a content source.
                                                       ' Add fresh requested sheets first, then remove every original template
                                                       ' sheet so sample cells/charts/shapes can never leak into generated output.
                                                       Dim originalTemplateSheetCount As Integer = CInt(wb.Sheets.Count)
                                                       For addIndex As Integer = 1 To sheetDefs.Count
                                                           Dim lastTemplateOrNewSheet As Object = wb.Sheets(wb.Sheets.Count)
                                                           Try
                                                               wb.Sheets.Add(After:=lastTemplateOrNewSheet)
                                                           Finally
                                                               Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(lastTemplateOrNewSheet) : Catch ex As System.Exception : End Try
                                                           End Try
                                                       Next
                                                       For deleteIndex As Integer = 1 To originalTemplateSheetCount
                                                           Dim templateSheet As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                                                           Try
                                                               templateSheet = CType(wb.Sheets(1), Microsoft.Office.Interop.Excel.Worksheet)
                                                               templateSheet.Delete()
                                                           Finally
                                                               If templateSheet IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(templateSheet) : Catch ex As System.Exception : End Try
                                                           End Try
                                                       Next
                                                   Else
                                                       wb = excelApp.Workbooks.Add()
                                                   End If

                                                   ' ── Create worksheets ──
                                                   ' Ensure exactly the requested number of fresh sheets.
                                                   While wb.Sheets.Count < sheetDefs.Count
                                                       Dim lastSheet As Object = wb.Sheets(wb.Sheets.Count)
                                                       wb.Sheets.Add(After:=lastSheet)
                                                       Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(lastSheet) : Catch : End Try
                                                   End While

                                                   While wb.Sheets.Count > sheetDefs.Count
                                                       Dim delSheet As Microsoft.Office.Interop.Excel.Worksheet =
                                                           CType(wb.Sheets(wb.Sheets.Count), Microsoft.Office.Interop.Excel.Worksheet)
                                                       delSheet.Delete()
                                                       Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(delSheet) : Catch : End Try
                                                   End While

                                                   For sheetIdx = 0 To sheetDefs.Count - 1
                                                       Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                                                       Try
                                                           ws = CType(wb.Sheets(sheetIdx + 1), Microsoft.Office.Interop.Excel.Worksheet)
                                                           Dim sheetDef = sheetDefs(sheetIdx)
                                                           ws.Name = sheetDef.SheetName

                                                           ' ── Apply cells ──
                                                           ApplyExcelCells(ws, sheetDef.Cells)

                                                           ' ── Resolve per-sheet settings ──
                                                           ' Top-level settings apply to the first sheet for backward compatibility;
                                                           ' any key present on the sheet object overrides the top-level value.
                                                           Dim sheetObjLocal As JObject =
                                                               If(sheetIdx < sheetObjs.Count, sheetObjs(sheetIdx), Nothing)
                                                           Dim sArgs As Dictionary(Of String, Object) =
                                                               BuildSheetArgs(excelArgs, sheetIdx = 0, sheetObjLocal)

                                                           Dim sColumnWidths = ParseColumnWidths(sArgs)
                                                           Dim sRowHeights = ParseRowHeights(sArgs)
                                                           Dim sMergeRanges = GetArgStringArray(sArgs, "merge_ranges")
                                                           Dim sFreezePane = GetArgString(sArgs, "freeze_pane")
                                                           Dim sAutoFilter = GetArgString(sArgs, "auto_filter")
                                                           Dim sDataValidations = ParseJsonArray(sArgs, "data_validations")
                                                           Dim sConditionalFormats = ParseJsonArray(sArgs, "conditional_formats")
                                                           Dim sPrintSetup As JObject = Nothing
                                                           If sArgs.ContainsKey("print_setup") Then
                                                               sPrintSetup = TryCast(sArgs("print_setup"), JObject)
                                                           End If

                                                           ' ── Opinionated professional baseline. Explicit cell formatting is re-applied afterwards. ──
                                                           ApplyProfessionalExcelSheetStyle(ws, sArgs, excelArgs)
                                                           ReapplyExplicitExcelFormatting(ws, sheetDef.Cells)

                                                           ' ── Auto-fit columns/rows (before explicit widths so explicit values win) ──
                                                           ApplyAutoFit(ws, sArgs)

                                                           ' ── Column widths ──
                                                           If sColumnWidths IsNot Nothing Then
                                                               ApplyColumnWidths(ws, sColumnWidths)
                                                           End If

                                                           ' ── Row heights ──
                                                           If sRowHeights IsNot Nothing Then
                                                               ApplyRowHeights(ws, sRowHeights)
                                                           End If

                                                           ' ── Merge ranges ──
                                                           If sMergeRanges IsNot Nothing AndAlso sMergeRanges.Count > 0 Then
                                                               For Each mr In sMergeRanges
                                                                   Dim mrRange As Microsoft.Office.Interop.Excel.Range = Nothing
                                                                   Try
                                                                       mrRange = ws.Range(mr)
                                                                       mrRange.Merge()
                                                                   Catch
                                                                   Finally
                                                                       If mrRange IsNot Nothing Then
                                                                           Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(mrRange) : Catch : End Try
                                                                       End If
                                                                   End Try
                                                               Next
                                                           End If

                                                           ' ── Freeze pane ──
                                                           If Not String.IsNullOrWhiteSpace(sFreezePane) Then
                                                               Dim fpRange As Microsoft.Office.Interop.Excel.Range = Nothing
                                                               Dim activeWin As Microsoft.Office.Interop.Excel.Window = Nothing
                                                               Try
                                                                   ws.Activate()
                                                                   fpRange = ws.Range(sFreezePane)
                                                                   fpRange.Select()
                                                                   activeWin = excelApp.ActiveWindow
                                                                   activeWin.FreezePanes = True
                                                               Catch
                                                               Finally
                                                                   If activeWin IsNot Nothing Then
                                                                       Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(activeWin) : Catch : End Try
                                                                   End If
                                                                   If fpRange IsNot Nothing Then
                                                                       Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(fpRange) : Catch : End Try
                                                                   End If
                                                               End Try
                                                           End If

                                                           ' ── Auto-filter ──
                                                           If Not String.IsNullOrWhiteSpace(sAutoFilter) Then
                                                               Dim afRange As Microsoft.Office.Interop.Excel.Range = Nothing
                                                               Try
                                                                   afRange = ws.Range(sAutoFilter)
                                                                   afRange.AutoFilter()
                                                               Catch
                                                               Finally
                                                                   If afRange IsNot Nothing Then
                                                                       Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(afRange) : Catch : End Try
                                                                   End If
                                                               End Try
                                                           End If

                                                           ' ── Data validations ──
                                                           If sDataValidations IsNot Nothing Then
                                                               ApplyDataValidations(ws, sDataValidations)
                                                           End If

                                                           ' ── Conditional formatting ──
                                                           If sConditionalFormats IsNot Nothing Then
                                                               ApplyConditionalFormats(ws, sConditionalFormats)
                                                           End If

                                                           ' ── Print setup ──
                                                           If sPrintSetup IsNot Nothing Then
                                                               ApplyPrintSetup(ws, sPrintSetup)
                                                           End If

                                                           ' ── Worksheet appearance (tab color, gridlines, zoom, right-to-left) ──
                                                           ApplyWorksheetAppearance(ws, sArgs)
                                                       Finally
                                                           If ws IsNot Nothing Then
                                                               Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(ws) : Catch : End Try
                                                           End If
                                                       End Try
                                                   Next

                                                   ' ── Excel tables (structured ranges with native filtering/banding) ──
                                                   If tables IsNot Nothing Then
                                                       ApplyExcelTables(wb, tables)
                                                   End If

                                                   ' ── Charts (can target any sheet) ──
                                                   If charts IsNot Nothing Then
                                                       ApplyCharts(wb, charts, sheetDefs, excelArgs)
                                                   End If

                                                   ' ── Named ranges ──
                                                   If namedRanges IsNot Nothing Then
                                                       For Each nrObj As JObject In namedRanges
                                                           Try
                                                               Dim nrName = nrObj.Value(Of String)("name")
                                                               Dim nrRange = nrObj.Value(Of String)("range")
                                                               If Not String.IsNullOrWhiteSpace(nrName) AndAlso Not String.IsNullOrWhiteSpace(nrRange) Then
                                                                   wb.Names.Add(Name:=nrName, RefersTo:="=" & nrRange)
                                                               End If
                                                           Catch ex As System.Exception
                                                           End Try
                                                       Next
                                                   End If

                                                   ' ── VBA modules ──
                                                   If hasVba AndAlso vbaModules IsNot Nothing Then
                                                       ApplyVbaModules(wb, vbaModules)
                                                   End If

                                                   ' ── Save ──
                                                   Dim fmt = If(hasVba, xlOpenXMLWorkbookMacroEnabled, xlOpenXMLWorkbook)
                                                   wb.SaveAs(outputPath, fmt)
                                                   Return True

                                               Catch ex As System.Exception
                                                   Debug.WriteLine($"CreateExcel error: {ex.Message}")
                                                   Return False
                                               Finally
                                                   SafeCloseExcel(wb, excelApp, weOwnApp)
                                               End Try
                                           End Function)

            If success AndAlso File.Exists(outputPath) Then
                RegisterAutoPilotGeneratedOutputFile(outputPath)

                Dim featureList As New List(Of String)()
                If sheetDefs.Count > 1 Then featureList.Add($"{sheetDefs.Count} sheets")
                featureList.Add($"{totalCells} cells")
                If mergeRanges.Count > 0 Then featureList.Add($"{mergeRanges.Count} merged range(s)")
                If dataValidations IsNot Nothing AndAlso dataValidations.Count > 0 Then featureList.Add($"{dataValidations.Count} validation(s)")
                If conditionalFormats IsNot Nothing AndAlso conditionalFormats.Count > 0 Then featureList.Add($"{conditionalFormats.Count} conditional format(s)")
                If tables IsNot Nothing AndAlso tables.Count > 0 Then featureList.Add($"{tables.Count} table(s)")
                If charts IsNot Nothing AndAlso charts.Count > 0 Then featureList.Add($"{charts.Count} chart(s)")
                If hasVba Then featureList.Add("VBA macros")

                Dim designNote As String = BuildDesignExecutionNote(design)
                response.Success = True
                response.Response = $"Excel spreadsheet created: {fileName} ({String.Join(", ", featureList)}, {New FileInfo(outputPath).Length / 1024:F0} KB). The file will be attached to the reply.{designNote}"
                ApDashboardLog($"✓ Excel created: {fileName}", "info")
            Else
                response.Success = False
                response.Response = "Failed to create Excel spreadsheet."
            End If

        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation was cancelled."
            response.Response = response.ErrorMessage
        Catch ex As System.Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error creating Excel spreadsheet: {ex.Message}"
        End Try

        Return response
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  EXCEL CREATION HELPERS
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Parses a hex color string like "#FF0000" or "FF0000" to an OLE color integer.
    ''' Returns Nothing if parsing fails.
    ''' </summary>
    Private Shared Function ParseHexColor(hexStr As String) As Integer?
        If String.IsNullOrWhiteSpace(hexStr) Then Return Nothing
        hexStr = hexStr.TrimStart("#"c)
        If hexStr.Length <> 6 Then Return Nothing
        Try
            Dim r = System.Convert.ToInt32(hexStr.Substring(0, 2), 16)
            Dim g = System.Convert.ToInt32(hexStr.Substring(2, 2), 16)
            Dim b = System.Convert.ToInt32(hexStr.Substring(4, 2), 16)
            ' Excel uses BGR (OLE color) format
            Return System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(r, g, b))
        Catch
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Parses column_widths from tool arguments.
    ''' </summary>
    Private Shared Function ParseColumnWidths(args As Dictionary(Of String, Object)) As Dictionary(Of String, Double)
        If args Is Nothing OrElse Not args.ContainsKey("column_widths") Then Return Nothing
        Dim cwObj = args("column_widths")
        If Not TypeOf cwObj Is JObject Then Return Nothing
        Dim result As New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase)
        For Each prop In DirectCast(cwObj, JObject).Properties()
            Dim w As Double
            If Double.TryParse(prop.Value.ToString(), Globalization.NumberStyles.Any,
                              Globalization.CultureInfo.InvariantCulture, w) Then
                result(prop.Name.ToUpperInvariant()) = w
            End If
        Next
        Return If(result.Count > 0, result, Nothing)
    End Function

    ''' <summary>
    ''' Parses row_heights from tool arguments.
    ''' </summary>
    Private Shared Function ParseRowHeights(args As Dictionary(Of String, Object)) As Dictionary(Of Integer, Double)
        If args Is Nothing OrElse Not args.ContainsKey("row_heights") Then Return Nothing
        Dim rhObj = args("row_heights")
        If Not TypeOf rhObj Is JObject Then Return Nothing
        Dim result As New Dictionary(Of Integer, Double)()
        For Each prop In DirectCast(rhObj, JObject).Properties()
            Dim rowNum As Integer
            Dim h As Double
            If Integer.TryParse(prop.Name, rowNum) AndAlso
               Double.TryParse(prop.Value.ToString(), Globalization.NumberStyles.Any,
                              Globalization.CultureInfo.InvariantCulture, h) Then
                result(rowNum) = h
            End If
        Next
        Return If(result.Count > 0, result, Nothing)
    End Function

    ''' <summary>
    ''' Parses a JSON array from tool arguments by key name.
    ''' </summary>
    Private Shared Function ParseJsonArray(args As Dictionary(Of String, Object), key As String) As List(Of JObject)
        If args Is Nothing OrElse Not args.ContainsKey(key) Then Return Nothing
        Dim obj = args(key)
        If Not TypeOf obj Is JArray Then Return Nothing
        Dim arr = DirectCast(obj, JArray)
        If arr.Count = 0 Then Return Nothing
        Return arr.OfType(Of JObject)().ToList()
    End Function

    ''' <summary>
    ''' Builds an effective argument dictionary for a single sheet by overlaying the
    ''' sheet object's own properties on top of the top-level arguments. Top-level
    ''' values are only used as a base for the first sheet, preserving backward
    ''' compatibility with single-sheet workbooks.
    ''' </summary>
    Private Shared Function BuildSheetArgs(topLevel As Dictionary(Of String, Object),
                                           isFirstSheet As Boolean,
                                           sheetObj As JObject) As Dictionary(Of String, Object)
        Dim result As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)

        If isFirstSheet AndAlso topLevel IsNot Nothing Then
            For Each kv In topLevel
                result(kv.Key) = kv.Value
            Next
        End If

        If sheetObj IsNot Nothing Then
            For Each prop In sheetObj.Properties()
                ' "name" and "cells" are structural, not formatting settings.
                If prop.Name.Equals("name", StringComparison.OrdinalIgnoreCase) OrElse
                   prop.Name.Equals("cells", StringComparison.OrdinalIgnoreCase) Then
                    Continue For
                End If
                result(prop.Name) = prop.Value
            Next
        End If

        Return result
    End Function

    ''' <summary>
    ''' Applies cell data, values, formulas, and rich formatting to a worksheet.
    ''' </summary>
    Private Shared Sub ApplyExcelCells(ws As Microsoft.Office.Interop.Excel.Worksheet, cellsArray As JArray)
        For Each cellObj As JObject In cellsArray
            Dim addr = cellObj.Value(Of String)("cell")
            If String.IsNullOrWhiteSpace(addr) Then Continue For

            Dim cell As Microsoft.Office.Interop.Excel.Range = Nothing
            Try
                cell = ws.Range(addr)
            Catch
                Continue For
            End Try

            Try
                ' ── Number format (apply before value so formatting takes effect) ──
                Dim numFmt = cellObj.Value(Of String)("number_format")
                If Not String.IsNullOrWhiteSpace(numFmt) Then
                    Try : cell.NumberFormat = numFmt : Catch : End Try
                End If

                ' ── Formula or value ──
                Dim formula = cellObj.Value(Of String)("formula")
                If Not String.IsNullOrWhiteSpace(formula) Then
                    Try
                        cell.Formula2 = formula
                    Catch
                        Try : cell.Formula = formula
                        Catch ex2 As System.Exception
                            Debug.WriteLine($"Formula error at {addr}: {ex2.Message}")
                        End Try
                    End Try
                Else
                    Dim valToken = cellObj("value")
                    If valToken IsNot Nothing Then
                        Dim valStr = valToken.ToString()
                        Dim numVal As Double
                        If Double.TryParse(valStr, Globalization.NumberStyles.Any,
                                          Globalization.CultureInfo.InvariantCulture, numVal) Then
                            cell.Value2 = numVal
                        Else
                            cell.Value2 = valStr
                        End If
                    End If
                End If

                ' ── Font styles ──
                If GetJBool(cellObj, "bold") Then Try : cell.Font.Bold = True : Catch : End Try
                If GetJBool(cellObj, "italic") Then Try : cell.Font.Italic = True : Catch : End Try
                If GetJBool(cellObj, "underline") Then Try : cell.Font.Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleSingle : Catch : End Try
                If GetJBool(cellObj, "strikethrough") Then Try : cell.Font.Strikethrough = True : Catch : End Try

                Dim fontName = cellObj.Value(Of String)("font_name")
                If Not String.IsNullOrWhiteSpace(fontName) Then Try : cell.Font.Name = fontName : Catch : End Try

                Dim fontSizeToken = cellObj("font_size")
                If fontSizeToken IsNot Nothing Then
                    Dim fs As Double
                    If Double.TryParse(fontSizeToken.ToString(), Globalization.NumberStyles.Any,
                                      Globalization.CultureInfo.InvariantCulture, fs) AndAlso fs > 0 Then
                        Try : cell.Font.Size = fs : Catch : End Try
                    End If
                End If

                ' ── Font color and background color ──
                Dim fontColorHex = cellObj.Value(Of String)("font_color")
                Dim bgColorHex = cellObj.Value(Of String)("bg_color")
                Dim fontColor = ParseHexColor(fontColorHex)
                Dim bgColor = ParseHexColor(bgColorHex)

                ' Safety guard: prevent white (or near-white) font on white/no background.
                ' LLMs sometimes copy the header's #FFFFFF font_color to data rows that have
                ' no bg_color or a light bg_color, resulting in invisible white-on-white text.
                If fontColor.HasValue Then
                    Dim isWhiteFont = False
                    If Not String.IsNullOrWhiteSpace(fontColorHex) Then
                        Dim trimHex = fontColorHex.TrimStart("#"c).ToUpperInvariant()
                        isWhiteFont = (trimHex = "FFFFFF")
                    End If

                    If isWhiteFont Then
                        ' Only allow white font when there is a sufficiently dark background
                        Dim hasDarkBg = False
                        If bgColor.HasValue AndAlso Not String.IsNullOrWhiteSpace(bgColorHex) Then
                            Dim bgHex = bgColorHex.TrimStart("#"c).ToUpperInvariant()
                            ' Consider the background "dark enough" if it's not white/near-white
                            ' Simple check: if any channel is below 0xC0, the bg is dark enough
                            If bgHex.Length = 6 Then
                                Try
                                    Dim rr = System.Convert.ToInt32(bgHex.Substring(0, 2), 16)
                                    Dim gg = System.Convert.ToInt32(bgHex.Substring(2, 2), 16)
                                    Dim bb = System.Convert.ToInt32(bgHex.Substring(4, 2), 16)
                                    If rr < &HC0 OrElse gg < &HC0 OrElse bb < &HC0 Then
                                        hasDarkBg = True
                                    End If
                                Catch
                                End Try
                            End If
                        End If

                        If hasDarkBg Then
                            ' White font on dark background is fine
                            Try : cell.Font.Color = fontColor.Value : Catch : End Try
                        Else
                            ' White font on white/no background → override to black
                            Debug.WriteLine($"[Excel] Safety: overriding white font to black at {addr} (no dark background)")
                            Try : cell.Font.Color = ParseHexColor("#000000").Value : Catch : End Try
                        End If
                    Else
                        Try : cell.Font.Color = fontColor.Value : Catch : End Try
                    End If
                End If

                ' ── Background color ──
                If bgColor.HasValue Then
                    Try
                        cell.Interior.Color = bgColor.Value
                        cell.Interior.Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternSolid
                    Catch
                    End Try
                End If

                ' ── Alignment ──
                Dim hAlign = cellObj.Value(Of String)("h_align")
                If Not String.IsNullOrWhiteSpace(hAlign) Then
                    Try
                        Select Case hAlign.ToLowerInvariant()
                            Case "left" : cell.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
                            Case "center" : cell.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                            Case "right" : cell.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                        End Select
                    Catch
                    End Try
                End If

                Dim vAlign = cellObj.Value(Of String)("v_align")
                If Not String.IsNullOrWhiteSpace(vAlign) Then
                    Try
                        Select Case vAlign.ToLowerInvariant()
                            Case "top" : cell.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop
                            Case "center" : cell.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                            Case "bottom" : cell.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignBottom
                        End Select
                    Catch
                    End Try
                End If

                If GetJBool(cellObj, "wrap_text") Then Try : cell.WrapText = True : Catch : End Try

                ' ── Borders ──
                Dim borderStyle = cellObj.Value(Of String)("border")
                If Not String.IsNullOrWhiteSpace(borderStyle) Then
                    Dim borderColor = ParseHexColor(cellObj.Value(Of String)("border_color"))
                    ApplyBorderStyle(cell, borderStyle, borderColor)
                End If

                ' ── Text rotation (degrees: -90..90, or 255 for stacked/vertical) ──
                Dim rotationToken = cellObj("text_rotation")
                If rotationToken IsNot Nothing Then
                    Dim rot As Integer
                    If Integer.TryParse(rotationToken.ToString(), rot) Then
                        Try : cell.Orientation = rot : Catch : End Try
                    End If
                End If

                ' ── Indent level ──
                Dim indentToken = cellObj("indent")
                If indentToken IsNot Nothing Then
                    Dim ind As Integer
                    If Integer.TryParse(indentToken.ToString(), ind) AndAlso ind >= 0 Then
                        Try : cell.IndentLevel = ind : Catch : End Try
                    End If
                End If

                ' ── Cell note/comment ──
                Dim noteText = cellObj.Value(Of String)("comment")
                If String.IsNullOrWhiteSpace(noteText) Then noteText = cellObj.Value(Of String)("note")
                If Not String.IsNullOrWhiteSpace(noteText) Then
                    Try : cell.ClearComments() : Catch : End Try
                    Try : cell.AddComment(noteText) : Catch : End Try
                End If

                ' ── Hyperlink ──
                Dim linkAddr = cellObj.Value(Of String)("hyperlink")
                If Not String.IsNullOrWhiteSpace(linkAddr) Then
                    Dim linkDisplay = cellObj.Value(Of String)("hyperlink_display")
                    Try
                        ws.Hyperlinks.Add(Anchor:=cell, Address:=linkAddr,
                                          TextToDisplay:=If(String.IsNullOrWhiteSpace(linkDisplay), linkAddr, linkDisplay))
                    Catch
                    End Try
                End If
            Finally
                If cell IsNot Nothing Then
                    Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(cell) : Catch : End Try
                End If
            End Try
        Next
    End Sub

    ''' <summary>
    ''' Re-applies only explicit cell formatting after the professional baseline.
    ''' Values, formulas, notes and hyperlinks are deliberately removed from the replay
    ''' so the styling pass cannot duplicate side effects.
    ''' </summary>
    Private Shared Sub ReapplyExplicitExcelFormatting(ws As Microsoft.Office.Interop.Excel.Worksheet, cellsArray As JArray)
        If ws Is Nothing OrElse cellsArray Is Nothing Then Exit Sub
        Dim formattingCells As New JArray()
        For Each cellObj As JObject In cellsArray.OfType(Of JObject)()
            Dim clone As JObject = DirectCast(cellObj.DeepClone(), JObject)
            clone.Remove("value")
            clone.Remove("formula")
            clone.Remove("comment")
            clone.Remove("note")
            clone.Remove("hyperlink")
            clone.Remove("hyperlink_display")
            If clone.Properties().Count() > 1 Then formattingCells.Add(clone)
        Next
        If formattingCells.Count > 0 Then ApplyExcelCells(ws, formattingCells)
    End Sub

    ''' <summary>
    ''' Helper to read a boolean from a JObject token.
    ''' </summary>
    Private Shared Function GetJBool(obj As JObject, key As String) As Boolean
        Dim token = obj(key)
        If token Is Nothing Then Return False
        If token.Type = JTokenType.Boolean Then Return CBool(token)
        Dim s = token.ToString()
        Dim result As Boolean
        If Boolean.TryParse(s, result) Then Return result
        Return False
    End Function

    ''' <summary>
    ''' Applies border styles to a cell range.
    ''' </summary>
    Private Shared Sub ApplyBorderStyle(cell As Microsoft.Office.Interop.Excel.Range,
                                         borderStyle As String, borderColor As Integer?)
        ' Map style names to Excel line style and weight
        Dim lineStyle As Microsoft.Office.Interop.Excel.XlLineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
        Dim weight As Microsoft.Office.Interop.Excel.XlBorderWeight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin

        Dim style = borderStyle.ToLowerInvariant()

        If style.Contains("medium") Then
            weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium
        ElseIf style.Contains("thick") Then
            weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
        End If

        Try
            If style.StartsWith("all") OrElse style = "thin" OrElse style = "medium" OrElse style = "thick" Then
                ' All four sides
                Dim edges() As Microsoft.Office.Interop.Excel.XlBordersIndex = {
                    Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft,
                    Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop,
                    Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom,
                    Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight
                }
                For Each edge In edges
                    cell.Borders(edge).LineStyle = lineStyle
                    cell.Borders(edge).Weight = weight
                    If borderColor.HasValue Then cell.Borders(edge).Color = borderColor.Value
                Next
            ElseIf style.StartsWith("bottom") Then
                cell.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = lineStyle
                cell.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = weight
                If borderColor.HasValue Then cell.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Color = borderColor.Value
            ElseIf style.StartsWith("outline") Then
                Dim edges() As Microsoft.Office.Interop.Excel.XlBordersIndex = {
                    Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft,
                    Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop,
                    Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom,
                    Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight
                }
                For Each edge In edges
                    cell.Borders(edge).LineStyle = lineStyle
                    cell.Borders(edge).Weight = weight
                    If borderColor.HasValue Then cell.Borders(edge).Color = borderColor.Value
                Next
            End If
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Applies column widths to a worksheet.
    ''' </summary>
    Private Shared Sub ApplyColumnWidths(ws As Microsoft.Office.Interop.Excel.Worksheet,
                                          widths As Dictionary(Of String, Double))
        For Each kv In widths
            Dim colRange As Microsoft.Office.Interop.Excel.Range = Nothing
            Try
                colRange = ws.Columns(kv.Key & ":" & kv.Key)
                colRange.ColumnWidth = kv.Value
            Catch
            Finally
                If colRange IsNot Nothing Then
                    Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(colRange) : Catch : End Try
                End If
            End Try
        Next
    End Sub

    ''' <summary>
    ''' Applies row heights to a worksheet.
    ''' </summary>
    Private Shared Sub ApplyRowHeights(ws As Microsoft.Office.Interop.Excel.Worksheet,
                                        heights As Dictionary(Of Integer, Double))
        For Each kv In heights
            Dim rowRange As Microsoft.Office.Interop.Excel.Range = Nothing
            Try
                rowRange = ws.Rows(kv.Key)
                rowRange.RowHeight = kv.Value
            Catch
            Finally
                If rowRange IsNot Nothing Then
                    Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(rowRange) : Catch : End Try
                End If
            End Try
        Next
    End Sub

    ''' <summary>
    ''' Auto-fits column widths and/or row heights based on tool arguments.
    ''' Accepts true (fit all), "all"/"*", a single letter/number, or an array of letters/numbers.
    ''' </summary>
    Private Shared Sub ApplyAutoFit(ws As Microsoft.Office.Interop.Excel.Worksheet,
                                    args As Dictionary(Of String, Object))
        If args Is Nothing Then Return
        Try
            If args.ContainsKey("auto_fit_columns") Then
                AutoFitColumns(ws, TryCast(args("auto_fit_columns"), JToken))
            End If
            If args.ContainsKey("auto_fit_rows") Then
                AutoFitRows(ws, TryCast(args("auto_fit_rows"), JToken))
            End If
        Catch
        End Try
    End Sub

    Private Shared Sub AutoFitColumns(ws As Microsoft.Office.Interop.Excel.Worksheet, tok As JToken)
        If tok Is Nothing Then Return
        Select Case tok.Type
            Case JTokenType.Boolean
                If CBool(tok) Then AutoFitAllColumns(ws)
            Case JTokenType.String
                Dim s = tok.ToString().Trim()
                If s.Equals("all", StringComparison.OrdinalIgnoreCase) OrElse s = "*" Then
                    AutoFitAllColumns(ws)
                Else
                    AutoFitColumnLetter(ws, s)
                End If
            Case JTokenType.Array
                For Each item As JToken In DirectCast(tok, JArray)
                    AutoFitColumnLetter(ws, item.ToString().Trim())
                Next
        End Select
    End Sub

    Private Shared Sub AutoFitAllColumns(ws As Microsoft.Office.Interop.Excel.Worksheet)
        Dim used As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim cols As Microsoft.Office.Interop.Excel.Range = Nothing
        Try
            used = ws.UsedRange
            cols = used.Columns
            cols.AutoFit()
        Catch
        Finally
            If cols IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(cols) : Catch : End Try
            If used IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(used) : Catch : End Try
        End Try
    End Sub

    Private Shared Sub AutoFitColumnLetter(ws As Microsoft.Office.Interop.Excel.Worksheet, letter As String)
        If String.IsNullOrWhiteSpace(letter) Then Return
        Dim colRange As Microsoft.Office.Interop.Excel.Range = Nothing
        Try
            colRange = ws.Columns(letter & ":" & letter)
            colRange.AutoFit()
        Catch
        Finally
            If colRange IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(colRange) : Catch : End Try
        End Try
    End Sub

    Private Shared Sub AutoFitRows(ws As Microsoft.Office.Interop.Excel.Worksheet, tok As JToken)
        If tok Is Nothing Then Return
        Select Case tok.Type
            Case JTokenType.Boolean
                If CBool(tok) Then AutoFitAllRows(ws)
            Case JTokenType.String
                Dim s = tok.ToString().Trim()
                If s.Equals("all", StringComparison.OrdinalIgnoreCase) OrElse s = "*" Then
                    AutoFitAllRows(ws)
                Else
                    Dim rowNum As Integer
                    If Integer.TryParse(s, rowNum) Then AutoFitRowNumber(ws, rowNum)
                End If
            Case JTokenType.Array
                For Each item As JToken In DirectCast(tok, JArray)
                    Dim rowNum As Integer
                    If Integer.TryParse(item.ToString().Trim(), rowNum) Then AutoFitRowNumber(ws, rowNum)
                Next
        End Select
    End Sub

    Private Shared Sub AutoFitAllRows(ws As Microsoft.Office.Interop.Excel.Worksheet)
        Dim used As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim rws As Microsoft.Office.Interop.Excel.Range = Nothing
        Try
            used = ws.UsedRange
            rws = used.Rows
            rws.AutoFit()
        Catch
        Finally
            If rws IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(rws) : Catch : End Try
            If used IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(used) : Catch : End Try
        End Try
    End Sub

    Private Shared Sub AutoFitRowNumber(ws As Microsoft.Office.Interop.Excel.Worksheet, rowNum As Integer)
        If rowNum < 1 Then Return
        Dim rowRange As Microsoft.Office.Interop.Excel.Range = Nothing
        Try
            rowRange = ws.Rows(rowNum)
            rowRange.AutoFit()
        Catch
        Finally
            If rowRange IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(rowRange) : Catch : End Try
        End Try
    End Sub

    ''' <summary>
    ''' Applies worksheet-level appearance settings: tab color, gridline visibility,
    ''' zoom level, and right-to-left layout.
    ''' </summary>
    Private Shared Sub ApplyWorksheetAppearance(ws As Microsoft.Office.Interop.Excel.Worksheet,
                                                args As Dictionary(Of String, Object))
        If args Is Nothing Then Return

        ' ── Tab color ──
        Dim tabColorStr As String = Nothing
        If args.ContainsKey("tab_color") Then
            Dim tt = TryCast(args("tab_color"), JToken)
            If tt IsNot Nothing Then tabColorStr = tt.ToString()
        End If
        Dim tabColor = ParseHexColor(tabColorStr)
        If tabColor.HasValue Then
            Try : ws.Tab.Color = tabColor.Value : Catch : End Try
        End If

        ' ── Right-to-left ──
        If args.ContainsKey("right_to_left") Then
            Dim rtlTok = TryCast(args("right_to_left"), JToken)
            If rtlTok IsNot Nothing AndAlso rtlTok.Type = JTokenType.Boolean Then
                Try : ws.DisplayRightToLeft = CBool(rtlTok) : Catch : End Try
            End If
        End If

        ' ── Gridlines / zoom (require the active window) ──
        Dim hasGridlines = args.ContainsKey("show_gridlines")
        Dim hasZoom = args.ContainsKey("zoom")
        If hasGridlines OrElse hasZoom Then
            Dim win As Microsoft.Office.Interop.Excel.Window = Nothing
            Try
                ws.Activate()
                win = ws.Application.ActiveWindow
                If hasGridlines Then
                    Dim gTok = TryCast(args("show_gridlines"), JToken)
                    If gTok IsNot Nothing AndAlso gTok.Type = JTokenType.Boolean Then
                        Try : win.DisplayGridlines = CBool(gTok) : Catch : End Try
                    End If
                End If
                If hasZoom Then
                    Dim zTok = TryCast(args("zoom"), JToken)
                    Dim z As Integer
                    If zTok IsNot Nothing AndAlso Integer.TryParse(zTok.ToString(), z) AndAlso z >= 10 AndAlso z <= 400 Then
                        Try : win.Zoom = z : Catch : End Try
                    End If
                End If
            Catch
            Finally
                If win IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(win) : Catch : End Try
            End Try
        End If
    End Sub

    ''' <summary>
    ''' Applies an opinionated consulting-style baseline to a generated worksheet.
    ''' The baseline is intentionally restrained: strong header hierarchy, subtle row banding,
    ''' compact typography, hidden gridlines, practical sizing, and automatic status highlighting.
    ''' Explicit cell formatting is re-applied by the caller afterwards and therefore wins.
    ''' </summary>
    Private Shared Sub ApplyProfessionalExcelSheetStyle(
            ws As Microsoft.Office.Interop.Excel.Worksheet,
            sheetArgs As Dictionary(Of String, Object),
            workbookArgs As Dictionary(Of String, Object))

        If ws Is Nothing Then Exit Sub
        If Not GetEffectiveExcelBool(sheetArgs, workbookArgs, "professional_layout", True) Then Exit Sub

        Dim stylePreset As String = GetEffectiveExcelString(sheetArgs, workbookArgs, "style_preset", "consulting").ToLowerInvariant()
        If stylePreset = "plain" OrElse stylePreset = "none" Then Exit Sub

        Dim accentHex As String = GetEffectiveExcelString(sheetArgs, workbookArgs, "accent_color", "#17365D")
        Dim accent As Integer = PptHexColor(accentHex, "#17365D")
        Dim textHex As String = GetEffectiveExcelString(sheetArgs, workbookArgs, "text_color", "#202124")
        Dim lineHex As String = GetEffectiveExcelString(sheetArgs, workbookArgs, "line_color", "#D9E2EC")
        Dim bandHex As String = GetEffectiveExcelString(sheetArgs, workbookArgs, "band_color", "#F6F8FA")
        Dim lightHex As String = GetEffectiveExcelString(sheetArgs, workbookArgs, "light_color", "#FFFFFF")
        Dim textColor As Integer = PptHexColor(textHex, "#202124")
        Dim mutedLine As Integer = PptHexColor(lineHex, "#D9E2EC")
        Dim bandColor As Integer = PptHexColor(bandHex, "#F6F8FA")
        Dim white As Integer = PptHexColor(lightHex, "#FFFFFF")
        Dim fontName As String = GetEffectiveExcelString(sheetArgs, workbookArgs, "font_name", "Aptos")
        Dim smartFormat As Boolean = GetEffectiveExcelBool(sheetArgs, workbookArgs, "smart_format", True)

        Dim used As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim headerRange As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim dataRange As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim win As Microsoft.Office.Interop.Excel.Window = Nothing
        Try
            used = ws.UsedRange
            If used Is Nothing Then Exit Sub

            Dim firstRow As Integer = used.Row
            Dim firstCol As Integer = used.Column
            Dim rowCount As Integer = used.Rows.Count
            Dim colCount As Integer = used.Columns.Count
            If rowCount <= 0 OrElse colCount <= 0 Then Exit Sub
            Dim lastRow As Integer = firstRow + rowCount - 1
            Dim lastCol As Integer = firstCol + colCount - 1

            Try : used.Font.Name = fontName : Catch : End Try
            Try : used.Font.Size = 10.0F : Catch : End Try
            Try : used.Font.Color = textColor : Catch : End Try
            Try : used.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter : Catch : End Try
            Try : used.Interior.Color = white : used.Interior.Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternSolid : Catch : End Try
            Try
                used.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                used.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline
                used.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).Color = mutedLine
            Catch
            End Try

            Dim headerRow As Integer = GetEffectiveExcelInt(sheetArgs, workbookArgs, "header_row", firstRow)
            If headerRow >= firstRow AndAlso headerRow <= lastRow Then
                headerRange = ws.Range(ws.Cells(headerRow, firstCol), ws.Cells(headerRow, lastCol))
                Try : headerRange.Interior.Color = accent : headerRange.Interior.Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternSolid : Catch : End Try
                Try : headerRange.Font.Color = white : headerRange.Font.Bold = True : headerRange.Font.Size = 10.5F : Catch : End Try
                Try : headerRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft : Catch : End Try
                Try : headerRange.RowHeight = 24.0F : Catch : End Try
                Try
                    headerRange.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    headerRange.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium
                    headerRange.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Color = accent
                Catch
                End Try

                If lastRow > headerRow Then
                    dataRange = ws.Range(ws.Cells(headerRow + 1, firstCol), ws.Cells(lastRow, lastCol))
                    Try : dataRange.RowHeight = 19.0F : Catch : End Try

                    ' Subtle alternating bands. Direct formatting is used so explicit cell styles can override it later.
                    If lastRow - headerRow <= 2000 Then
                        For r As Integer = headerRow + 1 To lastRow
                            If (r - headerRow) Mod 2 = 0 Then
                                Dim band As Microsoft.Office.Interop.Excel.Range = Nothing
                                Try
                                    band = ws.Range(ws.Cells(r, firstCol), ws.Cells(r, lastCol))
                                    band.Interior.Color = bandColor
                                    band.Interior.Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternSolid
                                Catch
                                Finally
                                    If band IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(band) : Catch : End Try
                                End Try
                            End If
                        Next
                    End If
                End If

                ' Sensible defaults when the agent did not specify them explicitly.
                If Not HasEffectiveExcelArg(sheetArgs, workbookArgs, "auto_filter") AndAlso lastRow > headerRow Then
                    Dim filterRange As Microsoft.Office.Interop.Excel.Range = Nothing
                    Try
                        filterRange = ws.Range(ws.Cells(headerRow, firstCol), ws.Cells(lastRow, lastCol))
                        filterRange.AutoFilter()
                    Catch
                    Finally
                        If filterRange IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(filterRange) : Catch : End Try
                    End Try
                End If

                If Not HasEffectiveExcelArg(sheetArgs, workbookArgs, "freeze_pane") AndAlso lastRow > headerRow Then
                    Dim fp As Microsoft.Office.Interop.Excel.Range = Nothing
                    Try
                        ws.Activate()
                        fp = ws.Cells(headerRow + 1, firstCol)
                        fp.Select()
                        win = ws.Application.ActiveWindow
                        win.FreezePanes = True
                    Catch
                    Finally
                        If fp IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(fp) : Catch : End Try
                        If win IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(win) : Catch : End Try : win = Nothing
                    End Try
                End If

                If smartFormat AndAlso lastRow > headerRow Then
                    ApplyAutomaticExcelSemanticFormatting(ws, headerRow, firstCol, lastRow, lastCol)
                End If
            End If

            ' Fit generated content automatically unless explicit widths were supplied.
            If Not HasEffectiveExcelArg(sheetArgs, workbookArgs, "column_widths") Then
                Dim cols As Microsoft.Office.Interop.Excel.Range = Nothing
                Try
                    cols = used.Columns
                    cols.AutoFit()
                    For c As Integer = 1 To colCount
                        Dim col As Microsoft.Office.Interop.Excel.Range = Nothing
                        Try
                            col = used.Columns(c)
                            Dim w As Double = CDbl(col.ColumnWidth)
                            If w < 9.0 Then col.ColumnWidth = 9.0
                            If w > 36.0 Then col.ColumnWidth = 36.0
                        Catch
                        Finally
                            If col IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(col) : Catch : End Try
                        End Try
                    Next
                Catch
                Finally
                    If cols IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(cols) : Catch : End Try
                End Try
            End If

            If Not HasEffectiveExcelArg(sheetArgs, workbookArgs, "tab_color") Then Try : ws.Tab.Color = accent : Catch : End Try

            ' Consulting-style worksheets are cleaner without gridlines and at a slightly reduced zoom.
            Try
                ws.Activate()
                win = ws.Application.ActiveWindow
                If Not HasEffectiveExcelArg(sheetArgs, workbookArgs, "show_gridlines") Then win.DisplayGridlines = False
                If Not HasEffectiveExcelArg(sheetArgs, workbookArgs, "zoom") Then win.Zoom = 90
            Catch
            Finally
                If win IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(win) : Catch : End Try : win = Nothing
            End Try

        Catch ex As System.Exception
            Debug.WriteLine($"Professional Excel styling error: {ex.Message}")
        Finally
            If dataRange IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(dataRange) : Catch : End Try
            If headerRange IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(headerRange) : Catch : End Try
            If used IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(used) : Catch : End Try
        End Try
    End Sub

    Private Shared Function GetEffectiveExcelString(sheetArgs As Dictionary(Of String, Object),
                                                     workbookArgs As Dictionary(Of String, Object),
                                                     key As String,
                                                     defaultValue As String) As String
        Dim v As String = GetArgString(sheetArgs, key)
        If String.IsNullOrWhiteSpace(v) Then v = GetArgString(workbookArgs, key)
        If String.IsNullOrWhiteSpace(v) Then v = defaultValue
        Return v
    End Function

    Private Shared Function GetEffectiveExcelBool(sheetArgs As Dictionary(Of String, Object),
                                                   workbookArgs As Dictionary(Of String, Object),
                                                   key As String,
                                                   defaultValue As Boolean) As Boolean
        If sheetArgs IsNot Nothing AndAlso sheetArgs.ContainsKey(key) Then Return GetArgBool(sheetArgs, key, defaultValue)
        Return GetArgBool(workbookArgs, key, defaultValue)
    End Function

    Private Shared Function GetEffectiveExcelInt(sheetArgs As Dictionary(Of String, Object),
                                                  workbookArgs As Dictionary(Of String, Object),
                                                  key As String,
                                                  defaultValue As Integer) As Integer
        Dim raw As String = Nothing
        If sheetArgs IsNot Nothing AndAlso sheetArgs.ContainsKey(key) Then raw = GetArgString(sheetArgs, key)
        If String.IsNullOrWhiteSpace(raw) Then raw = GetArgString(workbookArgs, key)
        Dim result As Integer
        If Integer.TryParse(raw, Globalization.NumberStyles.Integer, Globalization.CultureInfo.InvariantCulture, result) Then Return result
        Return defaultValue
    End Function

    Private Shared Function HasEffectiveExcelArg(sheetArgs As Dictionary(Of String, Object),
                                                  workbookArgs As Dictionary(Of String, Object),
                                                  key As String) As Boolean
        Return (sheetArgs IsNot Nothing AndAlso sheetArgs.ContainsKey(key)) OrElse
               (workbookArgs IsNot Nothing AndAlso workbookArgs.ContainsKey(key))
    End Function

    Private Shared Sub ApplyAutomaticExcelSemanticFormatting(ws As Microsoft.Office.Interop.Excel.Worksheet,
                                                              headerRow As Integer,
                                                              firstCol As Integer,
                                                              lastRow As Integer,
                                                              lastCol As Integer)
        For c As Integer = firstCol To lastCol
            Dim headerCell As Microsoft.Office.Interop.Excel.Range = Nothing
            Dim dataCol As Microsoft.Office.Interop.Excel.Range = Nothing
            Try
                headerCell = ws.Cells(headerRow, c)
                Dim header As String = If(headerCell.Value2, "").ToString().Trim().ToLowerInvariant()
                If String.IsNullOrWhiteSpace(header) Then Continue For
                dataCol = ws.Range(ws.Cells(headerRow + 1, c), ws.Cells(lastRow, c))
                Dim rangeAddress As String = dataCol.Address(False, False)

                If header.Contains("status") OrElse header = "state" Then
                    ApplyConditionalFormats(ws, New List(Of JObject) From {
                        New JObject From {{"range", rangeAddress}, {"type", "text_contains"}, {"formula1", "Done"}, {"format_bg_color", "#E2F0D9"}, {"format_font_color", "#2E7D32"}, {"format_bold", True}},
                        New JObject From {{"range", rangeAddress}, {"type", "text_contains"}, {"formula1", "Complete"}, {"format_bg_color", "#E2F0D9"}, {"format_font_color", "#2E7D32"}},
                        New JObject From {{"range", rangeAddress}, {"type", "text_contains"}, {"formula1", "In Progress"}, {"format_bg_color", "#FFF2CC"}, {"format_font_color", "#8A6D1D"}},
                        New JObject From {{"range", rangeAddress}, {"type", "text_contains"}, {"formula1", "Pending"}, {"format_bg_color", "#FFF2CC"}, {"format_font_color", "#8A6D1D"}},
                        New JObject From {{"range", rangeAddress}, {"type", "text_contains"}, {"formula1", "Blocked"}, {"format_bg_color", "#FCE8E6"}, {"format_font_color", "#C62828"}, {"format_bold", True}},
                        New JObject From {{"range", rangeAddress}, {"type", "text_contains"}, {"formula1", "Overdue"}, {"format_bg_color", "#FCE8E6"}, {"format_font_color", "#C62828"}}
                    })
                ElseIf header.Contains("priority") Then
                    ApplyConditionalFormats(ws, New List(Of JObject) From {
                        New JObject From {{"range", rangeAddress}, {"type", "text_contains"}, {"formula1", "Critical"}, {"format_bg_color", "#FCE8E6"}, {"format_font_color", "#C62828"}, {"format_bold", True}},
                        New JObject From {{"range", rangeAddress}, {"type", "text_contains"}, {"formula1", "High"}, {"format_bg_color", "#FCE8E6"}, {"format_font_color", "#C62828"}},
                        New JObject From {{"range", rangeAddress}, {"type", "text_contains"}, {"formula1", "Medium"}, {"format_bg_color", "#FFF2CC"}, {"format_font_color", "#8A6D1D"}},
                        New JObject From {{"range", rangeAddress}, {"type", "text_contains"}, {"formula1", "Low"}, {"format_bg_color", "#E2F0D9"}, {"format_font_color", "#2E7D32"}}
                    })
                End If
            Catch
            Finally
                If dataCol IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(dataCol) : Catch : End Try
                If headerCell IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(headerCell) : Catch : End Try
            End Try
        Next
    End Sub

    ''' <summary>
    ''' Creates native Excel ListObjects. This gives users structured references, sorting/filtering,
    ''' table expansion, banded rows, and a much more product-like workbook experience.
    ''' </summary>
    Private Shared Sub ApplyExcelTables(wb As Microsoft.Office.Interop.Excel.Workbook, tables As List(Of JObject))
        If wb Is Nothing OrElse tables Is Nothing Then Exit Sub

        ' IMPORTANT FOR OFFICE VERSION COMPATIBILITY:
        ' Keep the native-table feature late-bound. Do not introduce compile-time
        ' references to Excel.ListObject / XlListObjectSourceType / XlYesNoGuess,
        ' because the host project intentionally targets the Office 15 interop baseline.
        ' Numeric COM constants used here are stable across Office versions:
        '   xlSrcRange = 1, xlYes = 1.
        For Each tableObj As JObject In tables
            Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            Dim rng As Microsoft.Office.Interop.Excel.Range = Nothing
            Dim listObjects As Object = Nothing
            Dim lo As Object = Nothing
            Try
                Dim sheetName As String = tableObj.Value(Of String)("sheet_name")
                If String.IsNullOrWhiteSpace(sheetName) Then
                    ws = CType(wb.Sheets(1), Microsoft.Office.Interop.Excel.Worksheet)
                Else
                    ws = CType(wb.Sheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                End If

                Dim rangeName As String = tableObj.Value(Of String)("range")
                If String.IsNullOrWhiteSpace(rangeName) Then Continue For
                rng = ws.Range(rangeName)

                Try
                    If ws.AutoFilterMode Then ws.AutoFilterMode = False
                Catch
                End Try

                Dim wsLate As Object = DirectCast(ws, Object)
                listObjects = wsLate.ListObjects
                lo = listObjects.Add(1, rng, Nothing, 1)

                Dim tableName As String = SanitizeExcelTableName(tableObj.Value(Of String)("name"))
                If Not String.IsNullOrWhiteSpace(tableName) Then Try : lo.Name = tableName : Catch : End Try

                Dim tableStyle As String = tableObj.Value(Of String)("style")
                If String.IsNullOrWhiteSpace(tableStyle) Then tableStyle = "TableStyleMedium2"
                Try : lo.TableStyle = tableStyle : Catch : End Try

                If tableObj("show_totals") IsNot Nothing Then Try : lo.ShowTotals = GetJBool(tableObj, "show_totals") : Catch : End Try
                If tableObj("show_row_stripes") IsNot Nothing Then Try : lo.ShowTableStyleRowStripes = GetJBool(tableObj, "show_row_stripes") : Catch : End Try
                If tableObj("show_first_column") IsNot Nothing Then Try : lo.ShowTableStyleFirstColumn = GetJBool(tableObj, "show_first_column") : Catch : End Try
                If tableObj("show_last_column") IsNot Nothing Then Try : lo.ShowTableStyleLastColumn = GetJBool(tableObj, "show_last_column") : Catch : End Try

            Catch ex As System.Exception
                Debug.WriteLine($"Excel table creation error: {ex.Message}")
            Finally
                If lo IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(lo) : Catch : End Try
                If listObjects IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(listObjects) : Catch : End Try
                If rng IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(rng) : Catch : End Try
                If ws IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(ws) : Catch : End Try
            End Try
        Next
    End Sub

    Private Shared Function SanitizeExcelTableName(value As String) As String
        If String.IsNullOrWhiteSpace(value) Then Return Nothing
        Dim sb As New System.Text.StringBuilder()
        For Each ch As Char In value.Trim()
            If Char.IsLetterOrDigit(ch) OrElse ch = "_"c Then sb.Append(ch) Else sb.Append("_"c)
        Next
        Dim result As String = sb.ToString()
        If result.Length = 0 Then Return Nothing
        If Char.IsDigit(result(0)) Then result = "T_" & result
        Return result
    End Function

    ''' <summary>
    ''' Applies data validation rules to a worksheet.
    ''' </summary>
    Private Shared Sub ApplyDataValidations(ws As Microsoft.Office.Interop.Excel.Worksheet,
                                              validations As List(Of JObject))
        For Each dvObj In validations
            Dim dvRange As Microsoft.Office.Interop.Excel.Range = Nothing
            Try
                Dim rangeName = dvObj.Value(Of String)("range")
                If String.IsNullOrWhiteSpace(rangeName) Then Continue For

                dvRange = ws.Range(rangeName)
                dvRange.Validation.Delete() ' Clear existing validation

                Dim dvType = If(dvObj.Value(Of String)("type"), "list").ToLowerInvariant()
                Dim formula1 = dvObj.Value(Of String)("formula1")
                Dim formula2 = dvObj.Value(Of String)("formula2")
                Dim operatorStr = If(dvObj.Value(Of String)("operator"), "between").ToLowerInvariant()

                ' Map type to Excel constant
                Dim xlType As Integer
                Select Case dvType
                    Case "list" : xlType = 3 ' xlValidateList
                    Case "whole_number" : xlType = 1 ' xlValidateWholeNumber
                    Case "decimal" : xlType = 2 ' xlValidateDecimal
                    Case "date" : xlType = 4 ' xlValidateDate
                    Case "text_length" : xlType = 6 ' xlValidateTextLength
                    Case "custom" : xlType = 7 ' xlValidateCustom
                    Case Else : xlType = 3
                End Select

                ' Map operator to Excel constant
                Dim xlOp As Integer = 1 ' xlBetween
                Select Case operatorStr
                    Case "between" : xlOp = 1
                    Case "not_between" : xlOp = 2
                    Case "equal" : xlOp = 3
                    Case "not_equal" : xlOp = 4
                    Case "greater_than" : xlOp = 5
                    Case "less_than" : xlOp = 6
                    Case "greater_than_or_equal" : xlOp = 7
                    Case "less_than_or_equal" : xlOp = 8
                End Select

                If dvType = "list" Then
                    Dim cleanedFormula1 = formula1
                    If Not String.IsNullOrWhiteSpace(cleanedFormula1) Then
                        Dim parts = cleanedFormula1.Split(","c)
                        For i = 0 To parts.Length - 1
                            parts(i) = parts(i).Trim().Trim(""""c).Trim("'"c)
                        Next
                        cleanedFormula1 = String.Join(",", parts)
                    End If
                    dvRange.Validation.Add(Type:=xlType, AlertStyle:=1,
                                           Formula1:=cleanedFormula1)
                ElseIf Not String.IsNullOrWhiteSpace(formula2) Then
                    dvRange.Validation.Add(Type:=xlType, AlertStyle:=1,
                                           Operator:=xlOp,
                                           Formula1:=formula1, Formula2:=formula2)
                Else
                    dvRange.Validation.Add(Type:=xlType, AlertStyle:=1,
                                           Operator:=xlOp,
                                           Formula1:=formula1)
                End If

                ' Show dropdown for list type
                Dim showDropdown = dvObj("show_dropdown")
                If showDropdown IsNot Nothing AndAlso showDropdown.Type = JTokenType.Boolean Then
                    dvRange.Validation.InCellDropdown = CBool(showDropdown)
                End If

                ' Input message
                Dim inputTitle = dvObj.Value(Of String)("input_title")
                Dim inputMsg = dvObj.Value(Of String)("input_message")
                If Not String.IsNullOrWhiteSpace(inputTitle) OrElse Not String.IsNullOrWhiteSpace(inputMsg) Then
                    dvRange.Validation.ShowInput = True
                    If Not String.IsNullOrWhiteSpace(inputTitle) Then dvRange.Validation.InputTitle = inputTitle
                    If Not String.IsNullOrWhiteSpace(inputMsg) Then dvRange.Validation.InputMessage = inputMsg
                End If

                ' Error message
                Dim errorTitle = dvObj.Value(Of String)("error_title")
                Dim errorMsg = dvObj.Value(Of String)("error_message")
                If Not String.IsNullOrWhiteSpace(errorTitle) OrElse Not String.IsNullOrWhiteSpace(errorMsg) Then
                    dvRange.Validation.ShowError = True
                    If Not String.IsNullOrWhiteSpace(errorTitle) Then dvRange.Validation.ErrorTitle = errorTitle
                    If Not String.IsNullOrWhiteSpace(errorMsg) Then dvRange.Validation.ErrorMessage = errorMsg
                End If

            Catch ex As System.Exception
                Debug.WriteLine($"Data validation error: {ex.Message}")
            Finally
                If dvRange IsNot Nothing Then
                    Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(dvRange) : Catch : End Try
                End If
            End Try
        Next
    End Sub

    ''' <summary>
    ''' Applies conditional formatting rules to a worksheet.
    ''' </summary>
    Private Shared Sub ApplyConditionalFormats(ws As Microsoft.Office.Interop.Excel.Worksheet,
                                                formats As List(Of JObject))
        For Each cfObj In formats
            Dim cfRange As Microsoft.Office.Interop.Excel.Range = Nothing
            Try
                Dim rangeName = cfObj.Value(Of String)("range")
                If String.IsNullOrWhiteSpace(rangeName) Then Continue For

                cfRange = ws.Range(rangeName)
                Dim cfType = If(cfObj.Value(Of String)("type"), "cell_value").ToLowerInvariant()
                Dim operatorStr = If(cfObj.Value(Of String)("operator"), "greater_than").ToLowerInvariant()
                Dim formula1 = cfObj.Value(Of String)("formula1")
                Dim formula2 = cfObj.Value(Of String)("formula2")

                ' Map operator
                Dim xlOp As Integer = 5 ' xlGreater
                Select Case operatorStr
                    Case "between" : xlOp = 1
                    Case "not_between" : xlOp = 2
                    Case "equal" : xlOp = 3
                    Case "not_equal" : xlOp = 4
                    Case "greater_than" : xlOp = 5
                    Case "less_than" : xlOp = 6
                    Case "greater_than_or_equal" : xlOp = 7
                    Case "less_than_or_equal" : xlOp = 8
                End Select

                Dim fc As Microsoft.Office.Interop.Excel.FormatCondition = Nothing

                Select Case cfType
                    Case "cell_value"
                        If Not String.IsNullOrWhiteSpace(formula2) Then
                            fc = CType(cfRange.FormatConditions.Add(
                                Type:=Microsoft.Office.Interop.Excel.XlFormatConditionType.xlCellValue,
                                Operator:=xlOp, Formula1:=formula1, Formula2:=formula2),
                                Microsoft.Office.Interop.Excel.FormatCondition)
                        Else
                            fc = CType(cfRange.FormatConditions.Add(
                                Type:=Microsoft.Office.Interop.Excel.XlFormatConditionType.xlCellValue,
                                Operator:=xlOp, Formula1:=formula1),
                                Microsoft.Office.Interop.Excel.FormatCondition)
                        End If

                    Case "text_contains"
                        fc = CType(cfRange.FormatConditions.Add(
                            Type:=Microsoft.Office.Interop.Excel.XlFormatConditionType.xlTextString,
                            TextOperator:=Microsoft.Office.Interop.Excel.XlContainsOperator.xlContains,
                            String:=formula1),
                            Microsoft.Office.Interop.Excel.FormatCondition)

                    Case "duplicate"
                        fc = CType(cfRange.FormatConditions.AddUniqueValues(),
                            Microsoft.Office.Interop.Excel.UniqueValues)
                        CType(fc, Microsoft.Office.Interop.Excel.UniqueValues).DupeUnique = Microsoft.Office.Interop.Excel.XlDupeUnique.xlDuplicate
                        Dim fmtBgColor = ParseHexColor(cfObj.Value(Of String)("format_bg_color"))
                        If fmtBgColor.HasValue Then
                            Try : CType(fc, Microsoft.Office.Interop.Excel.UniqueValues).Interior.Color = fmtBgColor.Value : Catch : End Try
                        End If
                        Dim fmtFontColor = ParseHexColor(cfObj.Value(Of String)("format_font_color"))
                        If fmtFontColor.HasValue Then
                            Try : CType(fc, Microsoft.Office.Interop.Excel.UniqueValues).Font.Color = fmtFontColor.Value : Catch : End Try
                        End If
                        Continue For

                    Case "unique"
                        fc = CType(cfRange.FormatConditions.AddUniqueValues(),
                            Microsoft.Office.Interop.Excel.UniqueValues)
                        CType(fc, Microsoft.Office.Interop.Excel.UniqueValues).DupeUnique = Microsoft.Office.Interop.Excel.XlDupeUnique.xlUnique
                        Dim fmtBgColorU = ParseHexColor(cfObj.Value(Of String)("format_bg_color"))
                        If fmtBgColorU.HasValue Then
                            Try : CType(fc, Microsoft.Office.Interop.Excel.UniqueValues).Interior.Color = fmtBgColorU.Value : Catch : End Try
                        End If
                        Continue For

                    Case "color_scale"
                        cfRange.FormatConditions.AddColorScale(ColorScaleType:=3)
                        Continue For

                    Case "data_bar"
                        cfRange.FormatConditions.AddDatabar()
                        Continue For

                    Case "icon_set"
                        cfRange.FormatConditions.AddIconSetCondition()
                        Continue For

                    Case "top_10"
                        fc = CType(cfRange.FormatConditions.AddTop10(),
                            Microsoft.Office.Interop.Excel.Top10)
                        Dim rank As Integer = 10
                        If Not String.IsNullOrWhiteSpace(formula1) Then
                            Integer.TryParse(formula1, rank)
                        End If
                        CType(fc, Microsoft.Office.Interop.Excel.Top10).Rank = rank
                        Dim fmtBgColorT = ParseHexColor(cfObj.Value(Of String)("format_bg_color"))
                        If fmtBgColorT.HasValue Then
                            Try : CType(fc, Microsoft.Office.Interop.Excel.Top10).Interior.Color = fmtBgColorT.Value : Catch : End Try
                        End If
                        Continue For

                    Case Else
                        Continue For
                End Select

                ' Apply formatting to the FormatCondition
                If fc IsNot Nothing Then
                    Dim fmtFontColor = ParseHexColor(cfObj.Value(Of String)("format_font_color"))
                    If fmtFontColor.HasValue Then Try : fc.Font.Color = fmtFontColor.Value : Catch : End Try

                    Dim fmtBgColor = ParseHexColor(cfObj.Value(Of String)("format_bg_color"))
                    If fmtBgColor.HasValue Then
                        Try
                            fc.Interior.Color = fmtBgColor.Value
                            fc.Interior.Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternSolid
                        Catch
                        End Try
                    End If

                    If GetJBool(cfObj, "format_bold") Then Try : fc.Font.Bold = True : Catch : End Try
                End If

            Catch ex As System.Exception
                Debug.WriteLine($"Conditional format error: {ex.Message}")
            Finally
                If cfRange IsNot Nothing Then
                    Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(cfRange) : Catch : End Try
                End If
            End Try
        Next
    End Sub

    ''' <summary>
    ''' Creates charts and places them on worksheets.
    ''' </summary>
    Private Shared Sub ApplyCharts(wb As Microsoft.Office.Interop.Excel.Workbook,
                                    charts As List(Of JObject),
                                    sheetDefs As List(Of (SheetName As String, Cells As JArray)),
                                    workbookArgs As Dictionary(Of String, Object))
        For Each chartObj In charts
            Dim targetWs As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            Dim posCell As Microsoft.Office.Interop.Excel.Range = Nothing
            Dim chartObjects As Microsoft.Office.Interop.Excel.ChartObjects = Nothing
            Dim chartObject As Microsoft.Office.Interop.Excel.ChartObject = Nothing
            Dim chart As Microsoft.Office.Interop.Excel.Chart = Nothing
            Dim dataRangeObj As Microsoft.Office.Interop.Excel.Range = Nothing

            Try
                Dim chartType = If(chartObj.Value(Of String)("type"), "column").ToLowerInvariant()
                Dim dataRange = chartObj.Value(Of String)("data_range")
                Dim chartTitle = chartObj.Value(Of String)("title")
                Dim position = If(chartObj.Value(Of String)("position"), "E2")
                Dim chartSheetName = chartObj.Value(Of String)("sheet_name")

                If String.IsNullOrWhiteSpace(dataRange) Then Continue For

                ' Determine target worksheet
                If Not String.IsNullOrWhiteSpace(chartSheetName) Then
                    Try
                        targetWs = CType(wb.Sheets(chartSheetName), Microsoft.Office.Interop.Excel.Worksheet)
                    Catch
                        targetWs = CType(wb.Sheets(1), Microsoft.Office.Interop.Excel.Worksheet)
                    End Try
                Else
                    targetWs = CType(wb.Sheets(1), Microsoft.Office.Interop.Excel.Worksheet)
                End If

                ' Parse width/height with normalization
                Dim chartWidth As Double = NormalizeChartDimension(chartObj("width"), 480, 320)
                Dim chartHeight As Double = NormalizeChartDimension(chartObj("height"), 300, 220)

                ' Get position from cell
                posCell = targetWs.Range(position)
                Dim posLeft As Double = CDbl(posCell.Left)
                Dim posTop As Double = CDbl(posCell.Top)

                ' Map chart type to Excel constant
                Dim xlChartType As Microsoft.Office.Interop.Excel.XlChartType
                Select Case chartType
                    Case "column" : xlChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered
                    Case "bar" : xlChartType = Microsoft.Office.Interop.Excel.XlChartType.xlBarClustered
                    Case "line" : xlChartType = Microsoft.Office.Interop.Excel.XlChartType.xlLine
                    Case "pie" : xlChartType = Microsoft.Office.Interop.Excel.XlChartType.xlPie
                    Case "area" : xlChartType = Microsoft.Office.Interop.Excel.XlChartType.xlArea
                    Case "scatter" : xlChartType = Microsoft.Office.Interop.Excel.XlChartType.xlXYScatter
                    Case "doughnut" : xlChartType = Microsoft.Office.Interop.Excel.XlChartType.xlDoughnut
                    Case Else : xlChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered
                End Select

                ' Add chart as embedded ChartObject
                chartObjects = CType(targetWs.ChartObjects(), Microsoft.Office.Interop.Excel.ChartObjects)
                chartObject = chartObjects.Add(posLeft, posTop, chartWidth, chartHeight)

                Try
                    chartObject.Placement = Microsoft.Office.Interop.Excel.XlPlacement.xlFreeFloating
                Catch
                End Try

                chart = chartObject.Chart

                dataRangeObj = targetWs.Range(dataRange)
                chart.SetSourceData(dataRangeObj)
                chart.ChartType = xlChartType

                Try
                    chartObject.Width = chartWidth
                    chartObject.Height = chartHeight
                Catch
                End Try

                If Not String.IsNullOrWhiteSpace(chartTitle) Then
                    chart.HasTitle = True
                    chart.ChartTitle.Text = chartTitle
                End If

                ApplyProfessionalExcelChartBase(chart, chartObj, workbookArgs)

                ' ── Series / point colors ──
                Dim seriesColorsArr = TryCast(chartObj("series_colors"), JArray)
                Dim singleSeriesColor = ParseHexColor(chartObj.Value(Of String)("color"))
                If (seriesColorsArr Is Nothing OrElse seriesColorsArr.Count = 0) AndAlso Not singleSeriesColor.HasValue Then
                    Dim accentHex As String = GetArgString(workbookArgs, "accent_color")
                    Dim secondaryHex As String = GetArgString(workbookArgs, "secondary_color")
                    If String.IsNullOrWhiteSpace(accentHex) Then accentHex = "#17365D"
                    If String.IsNullOrWhiteSpace(secondaryHex) Then secondaryHex = "#2F75B5"
                    seriesColorsArr = New JArray(accentHex, secondaryHex, "#7F8C8D", "#D98E04", "#70AD47")
                End If
                ' Pie and doughnut charts have a single series whose slices are POINTS,
                ' so per-slice colors must be applied point-by-point rather than per series.
                Dim isPointColored As Boolean = (chartType = "pie" OrElse chartType = "doughnut")
                If (seriesColorsArr IsNot Nothing AndAlso seriesColorsArr.Count > 0) OrElse singleSeriesColor.HasValue Then
                    Dim seriesCol As Object = Nothing
                    Try
                        seriesCol = chart.SeriesCollection()
                        Dim seriesCount As Integer = CInt(seriesCol.Count)
                        For si As Integer = 1 To seriesCount
                            Dim ser As Object = Nothing
                            Try
                                ser = seriesCol.Item(si)

                                If isPointColored AndAlso seriesColorsArr IsNot Nothing AndAlso seriesColorsArr.Count > 0 Then
                                    ' Color each slice/point of the pie or doughnut individually
                                    Dim pts As Object = Nothing
                                    Try
                                        pts = ser.Points()
                                        Dim ptCount As Integer = CInt(pts.Count)
                                        For pi As Integer = 1 To ptCount
                                            Dim pt As Object = Nothing
                                            Try
                                                pt = pts.Item(pi)
                                                Dim pClr = ParseHexColor(seriesColorsArr((pi - 1) Mod seriesColorsArr.Count).ToString())
                                                If pClr.HasValue Then
                                                    Try : pt.Format.Fill.ForeColor.RGB = pClr.Value : Catch : End Try
                                                End If
                                            Catch
                                            Finally
                                                If pt IsNot Nothing Then
                                                    Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pt) : Catch : End Try
                                                End If
                                            End Try
                                        Next
                                    Catch
                                    Finally
                                        If pts IsNot Nothing Then
                                            Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pts) : Catch : End Try
                                        End If
                                    End Try
                                Else
                                    ' Standard per-series coloring
                                    Dim clr As Integer? = Nothing
                                    If seriesColorsArr IsNot Nothing AndAlso seriesColorsArr.Count > 0 Then
                                        clr = ParseHexColor(seriesColorsArr((si - 1) Mod seriesColorsArr.Count).ToString())
                                    ElseIf singleSeriesColor.HasValue Then
                                        clr = singleSeriesColor
                                    End If
                                    If clr.HasValue Then
                                        Try : ser.Format.Fill.ForeColor.RGB = clr.Value : Catch : End Try
                                        Try : ser.Format.Line.ForeColor.RGB = clr.Value : Catch : End Try
                                    End If
                                End If
                            Catch
                            Finally
                                If ser IsNot Nothing Then
                                    Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(ser) : Catch : End Try
                                End If
                            End Try
                        Next
                    Catch
                    Finally
                        If seriesCol IsNot Nothing Then
                            Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(seriesCol) : Catch : End Try
                        End If
                    End Try
                End If

                ' ── Legend ──
                Dim legendToken = chartObj("show_legend")
                If legendToken IsNot Nothing AndAlso legendToken.Type = JTokenType.Boolean Then
                    Try : chart.HasLegend = CBool(legendToken) : Catch : End Try
                End If
                Dim legendPos = chartObj.Value(Of String)("legend_position")
                If Not String.IsNullOrWhiteSpace(legendPos) Then
                    Try
                        chart.HasLegend = True
                        Select Case legendPos.ToLowerInvariant()
                            Case "bottom" : chart.Legend.Position = Microsoft.Office.Interop.Excel.XlLegendPosition.xlLegendPositionBottom
                            Case "top" : chart.Legend.Position = Microsoft.Office.Interop.Excel.XlLegendPosition.xlLegendPositionTop
                            Case "left" : chart.Legend.Position = Microsoft.Office.Interop.Excel.XlLegendPosition.xlLegendPositionLeft
                            Case "right" : chart.Legend.Position = Microsoft.Office.Interop.Excel.XlLegendPosition.xlLegendPositionRight
                            Case "corner" : chart.Legend.Position = Microsoft.Office.Interop.Excel.XlLegendPosition.xlLegendPositionCorner
                        End Select
                    Catch
                    End Try
                End If

                ' ── Data labels ──
                If GetJBool(chartObj, "show_data_labels") Then
                    Try : chart.ApplyDataLabels() : Catch : End Try
                End If

                ' ── Axis titles ──
                Dim xAxisTitle = chartObj.Value(Of String)("x_axis_title")
                If Not String.IsNullOrWhiteSpace(xAxisTitle) Then
                    Dim xAxis As Object = Nothing
                    Try
                        xAxis = chart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory)
                        xAxis.HasTitle = True
                        xAxis.AxisTitle.Text = xAxisTitle
                    Catch
                    Finally
                        If xAxis IsNot Nothing Then
                            Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xAxis) : Catch : End Try
                        End If
                    End Try
                End If
                Dim yAxisTitle = chartObj.Value(Of String)("y_axis_title")
                If Not String.IsNullOrWhiteSpace(yAxisTitle) Then
                    Dim yAxis As Object = Nothing
                    Try
                        yAxis = chart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue)
                        yAxis.HasTitle = True
                        yAxis.AxisTitle.Text = yAxisTitle
                    Catch
                    Finally
                        If yAxis IsNot Nothing Then
                            Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(yAxis) : Catch : End Try
                        End If
                    End Try
                End If

            Catch ex As System.Exception
                Debug.WriteLine($"Chart creation error: {ex.Message}")
            Finally
                If dataRangeObj IsNot Nothing Then
                    Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(dataRangeObj) : Catch : End Try
                End If
                If chart IsNot Nothing Then
                    Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(chart) : Catch : End Try
                End If
                If chartObject IsNot Nothing Then
                    Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(chartObject) : Catch : End Try
                End If
                If chartObjects IsNot Nothing Then
                    Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(chartObjects) : Catch : End Try
                End If
                If posCell IsNot Nothing Then
                    Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(posCell) : Catch : End Try
                End If
                If targetWs IsNot Nothing Then
                    Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(targetWs) : Catch : End Try
                End If
            End Try
        Next
    End Sub

    Private Shared Sub ApplyProfessionalExcelChartBase(chart As Object,
                                                       chartObj As JObject,
                                                       workbookArgs As Dictionary(Of String, Object))
        If chart Is Nothing Then Exit Sub

        ' Keep all professional chart-decoration code late-bound. The core chart
        ' creation path already works with the Office 15 Excel PIA; decorative
        ' Format/TextFrame2 members must not expand the compile-time Office surface.
        Dim fontName As String = GetArgString(workbookArgs, "font_name")
        If String.IsNullOrWhiteSpace(fontName) Then fontName = "Aptos"
        Dim textColor As Integer = PptHexColor("#202124", "#202124")
        Dim muted As Integer = PptHexColor("#667085", "#667085")
        Dim grid As Integer = PptHexColor("#E6EAF0", "#E6EAF0")
        Dim white As Integer = PptHexColor("#FFFFFF", "#FFFFFF")

        Try
            chart.ChartArea.Format.Fill.Solid()
            chart.ChartArea.Format.Fill.ForeColor.RGB = white
        Catch
        End Try
        Try : chart.ChartArea.Format.Line.Visible = 0 : Catch : End Try
        Try
            chart.PlotArea.Format.Fill.Solid()
            chart.PlotArea.Format.Fill.ForeColor.RGB = white
        Catch
        End Try
        Try : chart.PlotArea.Format.Line.Visible = 0 : Catch : End Try

        ' Prefer the legacy ChartTitle.Font path first; it is sufficient for the
        ' professional appearance and avoids a hard TextFrame2 dependency.
        Try
            If CBool(chart.HasTitle) Then
                chart.ChartTitle.Font.Name = fontName
                chart.ChartTitle.Font.Size = 13.0F
                chart.ChartTitle.Font.Bold = True
                chart.ChartTitle.Font.Color = textColor
            End If
        Catch
            ' Legacy title-font formatting is optional; leave the Office default if unavailable.
        End Try

        Try
            If CBool(chart.HasLegend) Then
                chart.Legend.Font.Name = fontName
                chart.Legend.Font.Size = 9.0F
                chart.Legend.Font.Color = muted
                Try : chart.Legend.Format.Line.Visible = 0 : Catch : End Try
            End If
        Catch
        End Try

        ' Excel COM constants are used numerically here to avoid adding enum/type
        ' references that are unnecessary for rendering:
        '   xlCategory = 1, xlValue = 2.
        For Each axisType As Integer In {1, 2}
            Dim ax As Object = Nothing
            Try
                ax = chart.Axes(axisType)
                ax.TickLabels.Font.Name = fontName
                ax.TickLabels.Font.Size = 9.0F
                ax.TickLabels.Font.Color = muted
                Try : ax.Format.Line.ForeColor.RGB = grid : Catch : End Try
                If axisType = 2 Then
                    Try : ax.MajorGridlines.Format.Line.ForeColor.RGB = grid : Catch : End Try
                    Try : ax.MajorGridlines.Format.Line.Weight = 0.75F : Catch : End Try
                End If
                If CBool(ax.HasTitle) Then
                    Try
                        ax.AxisTitle.Font.Name = fontName
                        ax.AxisTitle.Font.Size = 9.0F
                        ax.AxisTitle.Font.Color = muted
                    Catch
                    End Try
                End If
            Catch
            Finally
                If ax IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(ax) : Catch : End Try
            End Try
        Next
    End Sub

    ''' <summary>
    ''' Normalizes chart dimensions for Excel.
    ''' Excel expects points. Very small values are usually intended as inches.
    ''' </summary>
    Private Shared Function NormalizeChartDimension(
            token As JToken,
            defaultPoints As Double,
            minPoints As Double) As Double

        Dim value As Double = defaultPoints

        If token IsNot Nothing Then
            Dim parsed As Double
            If Double.TryParse(token.ToString(),
                               Globalization.NumberStyles.Any,
                               Globalization.CultureInfo.InvariantCulture,
                               parsed) Then
                value = parsed
            End If
        End If

        If value <= 0 Then value = defaultPoints

        ' Heuristic:
        ' Values like 4, 5, 6 are usually meant as inches, not points.
        If value <= 24 Then
            value *= 72.0
        End If

        If value < minPoints Then
            value = minPoints
        End If

        Return value
    End Function

    ''' <summary>
    ''' Applies print/page setup to a worksheet.
    ''' </summary>
    Private Shared Sub ApplyPrintSetup(ws As Microsoft.Office.Interop.Excel.Worksheet, setup As JObject)
        Try
            Dim orientation = setup.Value(Of String)("orientation")
            If Not String.IsNullOrWhiteSpace(orientation) Then
                Select Case orientation.ToLowerInvariant()
                    Case "landscape" : ws.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape
                    Case "portrait" : ws.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlPortrait
                End Select
            End If

            Dim fitWideToken = setup("fit_to_pages_wide")
            If fitWideToken IsNot Nothing Then
                ws.PageSetup.Zoom = False
                ws.PageSetup.FitToPagesWide = CInt(fitWideToken)
            End If

            Dim fitTallToken = setup("fit_to_pages_tall")
            If fitTallToken IsNot Nothing Then
                ws.PageSetup.Zoom = False
                ws.PageSetup.FitToPagesTall = CInt(fitTallToken)
            End If

            Dim headerText = setup.Value(Of String)("header_text")
            If Not String.IsNullOrWhiteSpace(headerText) Then ws.PageSetup.CenterHeader = headerText

            Dim footerText = setup.Value(Of String)("footer_text")
            If Not String.IsNullOrWhiteSpace(footerText) Then ws.PageSetup.CenterFooter = footerText
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Injects VBA code modules into the workbook using late binding to avoid
    ''' a hard reference to Microsoft.Vbe.Interop.
    ''' Requires "Trust access to the VBA project object model" to be enabled in Excel Trust Center settings.
    ''' </summary>
    Private Shared Sub ApplyVbaModules(wb As Microsoft.Office.Interop.Excel.Workbook, modules As JArray)
        For Each modObj As JObject In modules
            Try
                Dim modName = If(modObj.Value(Of String)("name"), "Module1")
                Dim modCode = modObj.Value(Of String)("code")
                Dim modType = If(modObj.Value(Of String)("type"), "module").ToLowerInvariant()

                If String.IsNullOrWhiteSpace(modCode) Then Continue For

                ' Use CallByName to fully late-bind and avoid requiring Microsoft.Vbe.Interop reference.
                ' Even with Option Strict Off, wb.VBProject resolves via the typed Workbook interface
                ' which pulls in the Vbe.Interop assembly at compile time.
                Dim vbProj As Object = Microsoft.VisualBasic.Interaction.CallByName(wb, "VBProject", CallType.Get)
                Dim vbComponents As Object = Microsoft.VisualBasic.Interaction.CallByName(vbProj, "VBComponents", CallType.Get)

                If modType = "thisworkbook" Then
                    ' Insert code into the ThisWorkbook module
                    Dim tbComponent As Object = vbComponents("ThisWorkbook")
                    Dim codeMod As Object = Microsoft.VisualBasic.Interaction.CallByName(tbComponent, "CodeModule", CallType.Get)
                    Microsoft.VisualBasic.Interaction.CallByName(codeMod, "AddFromString", CallType.Method, modCode)
                Else
                    ' vbext_ct_StdModule = 1, vbext_ct_ClassModule = 2
                    Dim componentType As Integer = If(modType = "class", 2, 1)
                    Dim newMod As Object = Microsoft.VisualBasic.Interaction.CallByName(vbComponents, "Add", CallType.Method, componentType)
                    Microsoft.VisualBasic.Interaction.CallByName(newMod, "Name", CallType.Let, modName)
                    Dim codeMod As Object = Microsoft.VisualBasic.Interaction.CallByName(newMod, "CodeModule", CallType.Get)
                    Microsoft.VisualBasic.Interaction.CallByName(codeMod, "AddFromString", CallType.Method, modCode)
                End If
            Catch ex As System.Exception
                Debug.WriteLine($"VBA module insertion error: {ex.Message}")
            End Try
        Next
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: create_word_document
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Shared Function GetArgSingleInvariant(args As Dictionary(Of String, Object),
                                                  key As String,
                                                  defaultVal As Single) As Single
        Dim raw As String = GetArgString(args, key)
        If String.IsNullOrWhiteSpace(raw) Then Return defaultVal

        Dim parsed As Single
        If Single.TryParse(raw, Globalization.NumberStyles.Any,
                           Globalization.CultureInfo.InvariantCulture, parsed) Then
            Return parsed
        End If

        If Single.TryParse(raw, parsed) Then
            Return parsed
        End If

        Return defaultVal
    End Function

    Private Shared Function TryApplyPreferredWordTableStyle(tbl As Microsoft.Office.Interop.Word.Table,
                                                       preferredStyleName As String) As Boolean
        If tbl Is Nothing Then Return False

        If Not String.IsNullOrWhiteSpace(preferredStyleName) Then
            Try
                tbl.Style = preferredStyleName.Trim()
                Return True
            Catch
            End Try
        End If

        Try
            tbl.Style = "Table Grid"
            Return True
        Catch
        End Try

        Return False
    End Function

    Private Shared Sub InsertAutoPilotWordCoverPage(sel As Microsoft.Office.Interop.Word.Selection,
                                                     title As String,
                                                     subtitle As String,
                                                     kicker As String,
                                                     accentColor As Integer,
                                                     textColor As Integer,
                                                     mutedColor As Integer,
                                                     fontName As String)
        If sel Is Nothing OrElse String.IsNullOrWhiteSpace(title) Then Exit Sub
        Try
            sel.ParagraphFormat.Alignment = 0 ' wdAlignParagraphLeft
            sel.ParagraphFormat.SpaceAfter = 0
            sel.Font.Name = fontName

            ' Deliberate white space creates a clean executive-cover rhythm.
            sel.TypeParagraph()
            sel.TypeParagraph()
            sel.TypeParagraph()
            sel.TypeParagraph()

            If Not String.IsNullOrWhiteSpace(kicker) Then
                sel.Font.Size = 9.0F
                sel.Font.Bold = True
                sel.Font.Color = accentColor
                sel.TypeText(kicker.Trim().ToUpperInvariant())
                sel.TypeParagraph()
                sel.TypeParagraph()
            End If

            sel.Font.Size = 30.0F
            sel.Font.Bold = True
            sel.Font.Color = textColor
            sel.TypeText(title.Trim())
            sel.TypeParagraph()

            If Not String.IsNullOrWhiteSpace(subtitle) Then
                sel.Font.Size = 13.0F
                sel.Font.Bold = False
                sel.Font.Color = mutedColor
                sel.TypeText(subtitle.Trim())
                sel.TypeParagraph()
            End If

            sel.TypeParagraph()
            sel.Font.Size = 11.0F
            sel.Font.Bold = True
            sel.Font.Color = accentColor
            sel.TypeText("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
            sel.TypeParagraph()
            Dim selLate As Object = DirectCast(sel, Object)
            selLate.InsertBreak(7) ' wdPageBreak
        Catch
        End Try
    End Sub

    Private Shared Sub ApplyAutoPilotWordDocumentStyling(
            doc As Microsoft.Office.Interop.Word.Document,
            args As Dictionary(Of String, Object))

        If doc Is Nothing Then Exit Sub

        Dim documentTitle As String = GetArgString(args, "document_title")
        Dim documentAuthor As String = GetArgString(args, "document_author")
        Dim baseFontName As String = GetArgString(args, "base_font_name")
        Dim tableStyleName As String = GetArgString(args, "table_style_name")
        Dim pageOrientation As String = If(GetArgString(args, "page_orientation"), "").Trim().ToLowerInvariant()
        Dim professionalLayout As Boolean = GetArgBool(args, "professional_layout", True)
        Dim stylePreset As String = If(GetArgString(args, "style_preset"), "consulting").Trim().ToLowerInvariant()
        Dim useTemplateStyles As Boolean = GetArgBool(args, "use_template_styles", False)
        Dim accentHex As String = GetArgString(args, "accent_color")
        Dim secondaryHex As String = GetArgString(args, "secondary_color")
        Dim textHex As String = GetArgString(args, "text_color")
        Dim mutedHex As String = GetArgString(args, "muted_color")
        Dim lightHex As String = GetArgString(args, "light_color")
        Dim lineHex As String = GetArgString(args, "line_color")
        If String.IsNullOrWhiteSpace(accentHex) Then accentHex = "#17365D"
        If String.IsNullOrWhiteSpace(secondaryHex) Then secondaryHex = "#2F75B5"
        If String.IsNullOrWhiteSpace(textHex) Then textHex = "#202124"
        If String.IsNullOrWhiteSpace(mutedHex) Then mutedHex = "#667085"
        If String.IsNullOrWhiteSpace(lightHex) Then lightHex = "#F3F6F9"
        If String.IsNullOrWhiteSpace(lineHex) Then lineHex = "#D9E2EC"
        Dim accentColor As Integer = PptHexColor(accentHex, "#17365D")
        Dim secondaryColor As Integer = PptHexColor(secondaryHex, "#2F75B5")
        Dim textColor As Integer = PptHexColor(textHex, "#202124")
        Dim mutedColor As Integer = PptHexColor(mutedHex, "#667085")
        Dim lightColor As Integer = PptHexColor(lightHex, "#F3F6F9")
        Dim lineColor As Integer = PptHexColor(lineHex, "#D9E2EC")

        If String.IsNullOrWhiteSpace(baseFontName) Then baseFontName = "Aptos"
        Dim effectiveFontSize As Single = GetArgSingleInvariant(args, "base_font_size", 10.5F)
        If effectiveFontSize <= 0 Then effectiveFontSize = 10.5F

        Dim normalStyle As Microsoft.Office.Interop.Word.Style = Nothing
        Dim titleStyle As Microsoft.Office.Interop.Word.Style = Nothing
        Dim h1Style As Microsoft.Office.Interop.Word.Style = Nothing
        Dim h2Style As Microsoft.Office.Interop.Word.Style = Nothing
        Dim h3Style As Microsoft.Office.Interop.Word.Style = Nothing

        Try
            If Not String.IsNullOrWhiteSpace(documentTitle) Then
                Try : doc.BuiltInDocumentProperties("Title").Value = documentTitle.Trim() : Catch : End Try
            End If
            If Not String.IsNullOrWhiteSpace(documentAuthor) Then
                Try : doc.BuiltInDocumentProperties("Author").Value = documentAuthor.Trim() : Catch : End Try
            End If

            Select Case pageOrientation
                Case "landscape"
                    Try : doc.PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientLandscape : Catch : End Try
                Case "portrait"
                    Try : doc.PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientPortrait : Catch : End Try
            End Select

            If Not useTemplateStyles AndAlso professionalLayout AndAlso stylePreset <> "plain" Then
                Try : doc.PageSetup.TopMargin = 50.0F : Catch : End Try
                Try : doc.PageSetup.BottomMargin = 46.0F : Catch : End Try
                Try : doc.PageSetup.LeftMargin = 54.0F : Catch : End Try
                Try : doc.PageSetup.RightMargin = 54.0F : Catch : End Try
            End If

            If Not useTemplateStyles Then
                normalStyle = doc.Styles(Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleNormal)
            Try : normalStyle.Font.Name = baseFontName : Catch : End Try
            Try : normalStyle.Font.Size = effectiveFontSize : Catch : End Try
            Try : normalStyle.Font.Color = textColor : Catch : End Try
            Try : normalStyle.ParagraphFormat.SpaceBefore = 0.0F : Catch : End Try
            Try : normalStyle.ParagraphFormat.SpaceAfter = 6.0F : Catch : End Try
            Try : normalStyle.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceSingle : Catch : End Try
            Try : normalStyle.ParagraphFormat.WidowControl = -1 : Catch : End Try

            Try
                titleStyle = doc.Styles(Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleTitle)
                titleStyle.Font.Name = baseFontName
                titleStyle.Font.Size = 28.0F
                titleStyle.Font.Bold = True
                titleStyle.Font.Color = textColor
                titleStyle.ParagraphFormat.SpaceAfter = 14.0F
                titleStyle.ParagraphFormat.KeepWithNext = -1
            Catch
            End Try

            Try
                h1Style = doc.Styles(Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleHeading1)
                h1Style.Font.Name = baseFontName
                h1Style.Font.Size = 16.0F
                h1Style.Font.Bold = True
                h1Style.Font.Color = accentColor
                h1Style.ParagraphFormat.SpaceBefore = 16.0F
                h1Style.ParagraphFormat.SpaceAfter = 6.0F
                h1Style.ParagraphFormat.KeepWithNext = -1
            Catch
            End Try

            Try
                h2Style = doc.Styles(Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleHeading2)
                h2Style.Font.Name = baseFontName
                h2Style.Font.Size = 12.5F
                h2Style.Font.Bold = True
                h2Style.Font.Color = textColor
                h2Style.ParagraphFormat.SpaceBefore = 12.0F
                h2Style.ParagraphFormat.SpaceAfter = 4.0F
                h2Style.ParagraphFormat.KeepWithNext = -1
            Catch
            End Try

            Try
                h3Style = doc.Styles(Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleHeading3)
                h3Style.Font.Name = baseFontName
                h3Style.Font.Size = 10.5F
                h3Style.Font.Bold = True
                h3Style.Font.Color = secondaryColor
                h3Style.ParagraphFormat.SpaceBefore = 9.0F
                h3Style.ParagraphFormat.SpaceAfter = 3.0F
                h3Style.ParagraphFormat.KeepWithNext = -1
            Catch
            End Try

            ' Clean paragraph rhythm throughout the document.
            If professionalLayout Then
                For Each para As Object In doc.Paragraphs
                    Try
                        para.Range.Font.Name = baseFontName
                        If para.Range.Style Is normalStyle Then para.Range.Font.Size = effectiveFontSize
                        para.Format.WidowControl = -1
                    Catch
                    End Try
                Next
            End If

            ' Tables: dark strategic header, subtle rules, light banding, compact padding.
            For Each tbl As Microsoft.Office.Interop.Word.Table In doc.Tables
                Dim headerRange As Microsoft.Office.Interop.Word.Range = Nothing
                Try
                    TryApplyPreferredWordTableStyle(tbl, tableStyleName)
                    Try : tbl.Range.Font.Name = baseFontName : Catch : End Try
                    Try : tbl.Range.Font.Size = effectiveFontSize : Catch : End Try
                    Try : tbl.Range.Font.Color = textColor : Catch : End Try
                    Try : tbl.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter : Catch : End Try
                    Try : tbl.Rows.Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowLeft : Catch : End Try
                    Try : tbl.Range.ParagraphFormat.SpaceBefore = 0 : tbl.Range.ParagraphFormat.SpaceAfter = 0 : Catch : End Try

                    If professionalLayout Then
                        Try : tbl.AllowAutoFit = True : Catch : End Try
                        Try : tbl.AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow) : Catch : End Try
                        Try : tbl.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent : tbl.PreferredWidth = 100.0F : Catch : End Try
                        Try : tbl.TopPadding = 4.0F : tbl.BottomPadding = 4.0F : tbl.LeftPadding = 5.0F : tbl.RightPadding = 5.0F : Catch : End Try
                        Try
                            tbl.Borders.Enable = 1
                            tbl.Borders.OutsideColor = lineColor
                            tbl.Borders.InsideColor = lineColor
                        Catch
                        End Try
                    End If

                    If tbl.Rows.Count > 0 Then
                        Try : tbl.Rows(1).HeadingFormat = -1 : Catch : End Try
                        Try
                            headerRange = tbl.Rows(1).Range
                            headerRange.Font.Bold = True
                            headerRange.Font.Name = baseFontName
                            headerRange.Font.Size = effectiveFontSize
                            headerRange.Font.Color = PptHexColor("#FFFFFF", "#FFFFFF")
                            headerRange.Shading.BackgroundPatternColor = accentColor
                            headerRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        Catch
                        End Try
                    End If

                    If professionalLayout AndAlso tbl.Rows.Count > 2 AndAlso tbl.Rows.Count <= 500 Then
                        For r As Integer = 2 To tbl.Rows.Count
                            If r Mod 2 = 0 Then
                                Try : tbl.Rows(r).Range.Shading.BackgroundPatternColor = lightColor : Catch : End Try
                            End If
                        Next
                    End If
                Finally
                    If headerRange IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(headerRange) : Catch : End Try
                End Try
            Next

            Else
                ' A configured Office template is the design authority. Do not overwrite its
                ' Normal/Title/Heading/Table definitions with the generic renderer.
            End If

            If Not useTemplateStyles OrElse
               HasMeaningfulToolArgument(args, "header_text") OrElse
               HasMeaningfulToolArgument(args, "footer_text") OrElse
               HasMeaningfulToolArgument(args, "show_page_numbers") Then
                ApplyAutoPilotWordHeaderFooter(doc, args, baseFontName, mutedColor, lineColor)
            End If
            Try : doc.Repaginate() : Catch : End Try

        Finally
            If h3Style IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(h3Style) : Catch : End Try
            If h2Style IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(h2Style) : Catch : End Try
            If h1Style IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(h1Style) : Catch : End Try
            If titleStyle IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(titleStyle) : Catch : End Try
            If normalStyle IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(normalStyle) : Catch : End Try
        End Try
    End Sub

    Private Shared Sub ApplyAutoPilotWordHeaderFooter(doc As Microsoft.Office.Interop.Word.Document,
                                                       args As Dictionary(Of String, Object),
                                                       fontName As String,
                                                       mutedColor As Integer,
                                                       lineColor As Integer)
        Dim headerText As String = GetArgString(args, "header_text")
        Dim footerText As String = GetArgString(args, "footer_text")
        Dim showPageNumbers As Boolean = GetArgBool(args, "show_page_numbers", True)
        If String.IsNullOrWhiteSpace(headerText) AndAlso String.IsNullOrWhiteSpace(footerText) AndAlso Not showPageNumbers Then Exit Sub

        For Each section As Object In doc.Sections
            Dim header As Object = Nothing
            Dim footer As Object = Nothing
            Dim hr As Microsoft.Office.Interop.Word.Range = Nothing
            Dim fr As Microsoft.Office.Interop.Word.Range = Nothing
            Try
                header = section.Headers(1) ' wdHeaderFooterPrimary
                footer = section.Footers(1) ' wdHeaderFooterPrimary
                If Not String.IsNullOrWhiteSpace(headerText) Then
                    hr = header.Range
                    hr.Text = headerText.Trim()
                    hr.Font.Name = fontName
                    hr.Font.Size = 8.0F
                    hr.Font.Color = mutedColor
                    hr.ParagraphFormat.Alignment = 0 ' wdAlignParagraphLeft
                End If

                fr = footer.Range
                fr.Text = If(String.IsNullOrWhiteSpace(footerText), "", footerText.Trim())
                fr.Font.Name = fontName
                fr.Font.Size = 8.0F
                fr.Font.Color = mutedColor
                fr.ParagraphFormat.Alignment = 2 ' wdAlignParagraphRight
                If showPageNumbers Then
                    DirectCast(fr, Object).Collapse(0) ' wdCollapseEnd
                    If Not String.IsNullOrWhiteSpace(footerText) Then fr.InsertAfter("   •   ")
                    DirectCast(fr, Object).Collapse(0) ' wdCollapseEnd
                    Dim fieldsLate As Object = DirectCast(doc.Fields, Object)
                    fieldsLate.Add(fr, 33) ' wdFieldPage
                End If
            Catch
            Finally
                If fr IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(fr) : Catch : End Try
                If hr IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(hr) : Catch : End Try
                If footer IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(footer) : Catch : End Try
                If header IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(header) : Catch : End Try
            End Try
        Next
    End Sub

    Private Shared Function GetAutoPilotWordVisuals(args As Dictionary(Of String, Object)) As Newtonsoft.Json.Linq.JArray
        If args Is Nothing OrElse Not args.ContainsKey("visuals") OrElse args("visuals") Is Nothing Then
            Return New Newtonsoft.Json.Linq.JArray()
        End If

        Try
            Dim token As Newtonsoft.Json.Linq.JToken = TryCast(args("visuals"), Newtonsoft.Json.Linq.JToken)
            If token Is Nothing Then
                token = Newtonsoft.Json.Linq.JToken.FromObject(args("visuals"))
            End If
            If token IsNot Nothing AndAlso token.Type = Newtonsoft.Json.Linq.JTokenType.Array Then
                Return DirectCast(token, Newtonsoft.Json.Linq.JArray)
            End If
        Catch
        End Try

        Return New Newtonsoft.Json.Linq.JArray()
    End Function

    Private Shared Function ContainsLikelyWordPseudoGraphic(markdownContent As String) As Boolean
        If String.IsNullOrWhiteSpace(markdownContent) Then Return False

        Dim normalized As String = markdownContent.Replace(vbCrLf, vbLf)
        If System.Text.RegularExpressions.Regex.IsMatch(normalized,
                                                        "```\s*(mermaid|graphviz|dot)\b",
                                                        System.Text.RegularExpressions.RegexOptions.IgnoreCase) Then
            Return True
        End If

        ' Reject fenced pseudo-diagrams even when they use only ordinary brackets plus
        ' Unicode/ASCII arrows (for example: [A] -> [B] -> [C]). The previous detector
        ' missed these one-line flow diagrams, allowing character-based graphics into Word
        ' instead of forcing the model to use visuals type='process'.
        If System.Text.RegularExpressions.Regex.IsMatch(normalized,
                                                        "```[\s\S]*?(?:\[[^\]\r\n]{1,100}\]\s*(?:-{1,2}>|={1,2}>|→|➔|⇒)\s*\[[^\]\r\n]{1,100}\])(?:[\s\S]*?)```",
                                                        System.Text.RegularExpressions.RegexOptions.IgnoreCase) Then
            Return True
        End If

        Dim bracketArrowChains As Integer = System.Text.RegularExpressions.Regex.Matches(
            normalized,
            "\[[^\]\r\n]{1,100}\]\s*(?:-{1,2}>|={1,2}>|→|➔|⇒)\s*\[[^\]\r\n]{1,100}\]",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase).Count
        If bracketArrowChains >= 1 Then Return True

        ' ASCII box diagrams frequently contain border fragments on the same line as
        ' labels after Markdown normalization. Detect repeated +-----+ fragments anywhere
        ' in the content, not only lines made exclusively from borders. Markdown tables do
        ' not use plus-corner borders, so this does not reject normal pipe tables.
        Dim asciiBoxFragments As Integer = System.Text.RegularExpressions.Regex.Matches(
            normalized,
            "\+(?:[-=]{4,})\+",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase).Count
        If asciiBoxFragments >= 2 Then Return True

        ' Reject block/square character bar charts such as ▰▰▰▰▰ or █████. These can be
        ' hidden inside an otherwise valid Markdown table, so line-oriented box detection
        ' alone is not sufficient. A real requested chart must use the native visuals array.
        Dim unicodeBarRuns As Integer = System.Text.RegularExpressions.Regex.Matches(
            normalized,
            "(?:[\u2580-\u259F\u25A0-\u25FF]\s*){4,}",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase).Count
        If unicodeBarRuns >= 1 Then Return True

        Dim lines() As String = normalized.Split({vbLf}, StringSplitOptions.None)
        Dim suspiciousLines As Integer = 0
        For Each rawLine As String In lines
            Dim line As String = If(rawLine, String.Empty)
            If line.IndexOfAny(New Char() {"┌"c, "┐"c, "└"c, "┘"c, "├"c, "┤"c, "┬"c, "┴"c, "┼"c, "─"c, "│"c, "╔"c, "╗"c, "╚"c, "╝"c, "═"c, "║"c}) >= 0 Then
                suspiciousLines += 1
            ElseIf System.Text.RegularExpressions.Regex.IsMatch(line, "^\s*\+(?:[-=]{3,}\+)+\s*$") Then
                suspiciousLines += 1
            ElseIf System.Text.RegularExpressions.Regex.IsMatch(line, "^\s*(?:[-=]{3,}>|<[-=]{3,}|\|\s*\^\s*\||\|\s*[vV]\s*\|)\s*$") Then
                suspiciousLines += 1
            ElseIf System.Text.RegularExpressions.Regex.IsMatch(line, "^\s*\|.*(?:-->|==>|<--|<==).*\|\s*$") Then
                suspiciousLines += 1
            End If

            If suspiciousLines >= 2 Then Return True
        Next

        Return False
    End Function

    Private Shared Function CountAutoPilotWordVisualsOfType(visuals As Newtonsoft.Json.Linq.JArray,
                                                                ParamArray acceptedTypes() As System.String) As Integer
        If visuals Is Nothing OrElse acceptedTypes Is Nothing OrElse acceptedTypes.Length = 0 Then Return 0

        Dim accepted As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
        For Each acceptedType As System.String In acceptedTypes
            If Not System.String.IsNullOrWhiteSpace(acceptedType) Then accepted.Add(acceptedType.Trim())
        Next

        Dim count As Integer = 0
        For Each token As Newtonsoft.Json.Linq.JToken In visuals
            If token Is Nothing OrElse token.Type <> Newtonsoft.Json.Linq.JTokenType.Object Then Continue For
            Dim visual As Newtonsoft.Json.Linq.JObject = DirectCast(token, Newtonsoft.Json.Linq.JObject)
            Dim visualType As System.String = GetVisualText(visual, "type", "process")
            If accepted.Contains(visualType) Then count += 1
        Next
        Return count
    End Function

    Private Shared Function CountOrdinalOccurrences(text As System.String, value As System.String) As Integer
        If System.String.IsNullOrEmpty(text) OrElse System.String.IsNullOrEmpty(value) Then Return 0
        Dim count As Integer = 0
        Dim startIndex As Integer = 0
        Do
            Dim foundIndex As Integer = text.IndexOf(value, startIndex, System.StringComparison.Ordinal)
            If foundIndex < 0 Then Exit Do
            count += 1
            startIndex = foundIndex + value.Length
        Loop
        Return count
    End Function

    Private Shared Function ValidateAutoPilotWordVisualContract(markdownContent As System.String,
                                                                 visuals As Newtonsoft.Json.Linq.JArray,
                                                                 context As ToolExecutionContext,
                                                                 ByRef validationError As System.String) As Boolean
        validationError = System.String.Empty
        If visuals Is Nothing Then visuals = New Newtonsoft.Json.Linq.JArray()

        Dim ids As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
        For Each token As Newtonsoft.Json.Linq.JToken In visuals
            If token Is Nothing OrElse token.Type <> Newtonsoft.Json.Linq.JTokenType.Object Then
                validationError = "Every create_word_document visuals entry must be an object."
                Return False
            End If

            Dim visual As Newtonsoft.Json.Linq.JObject = DirectCast(token, Newtonsoft.Json.Linq.JObject)
            Dim id As System.String = GetVisualText(visual, "id")
            If System.String.IsNullOrWhiteSpace(id) OrElse
               Not System.Text.RegularExpressions.Regex.IsMatch(id, "^[A-Za-z0-9_.-]{1,64}$") Then
                validationError = "Every create_word_document visual requires a valid id."
                Return False
            End If
            If Not ids.Add(id) Then
                validationError = "Duplicate create_word_document visual id '" & id & "' is not allowed."
                Return False
            End If

            Dim placeholder As System.String = "[[visual:" & id & "]]"
            Dim placeholderCount As Integer = CountOrdinalOccurrences(If(markdownContent, System.String.Empty), placeholder)
            If placeholderCount <> 1 Then
                validationError = "Editable visual '" & id & "' requires exactly one " & placeholder & " placeholder in markdown_content; found " & placeholderCount.ToString() & "."
                Return False
            End If
        Next

        ' Reject orphan placeholders as well. They otherwise survive into the document or
        ' make a later retry appear successful after the corresponding visual was dropped.
        Dim placeholderMatches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(
            If(markdownContent, System.String.Empty),
            "\[\[visual:([A-Za-z0-9_.-]{1,64})\]\]",
            System.Text.RegularExpressions.RegexOptions.CultureInvariant)
        For Each placeholderMatch As System.Text.RegularExpressions.Match In placeholderMatches
            Dim placeholderId As System.String = placeholderMatch.Groups(1).Value
            If Not ids.Contains(placeholderId) Then
                validationError = "markdown_content contains [[visual:" & placeholderId & "]] but no matching visuals entry."
                Return False
            End If
        Next

        Dim requestText As System.String = System.String.Empty
        If context IsNot Nothing Then
            requestText = If(context.LatestUserRequestRaw, System.String.Empty)
            If Not System.String.IsNullOrWhiteSpace(context.HostTaskSummary) Then
                requestText &= vbLf & context.HostTaskSummary
            End If
        End If
        Dim markdown As System.String = If(markdownContent, System.String.Empty)

        Dim requiresOrgChart As Boolean =
            System.Text.RegularExpressions.Regex.IsMatch(requestText,
                "\b(?:organigramm|org\s*chart|organi[sz]ation(?:al)?\s+chart|organi[sz]ational\s+chart)\b",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.CultureInvariant) OrElse
            System.Text.RegularExpressions.Regex.IsMatch(markdown,
                "(?im)^\s{0,3}(?:#{1,6}\s*)?.*\b(?:organigramm|org\s*chart|organi[sz]ation(?:al)?\s+chart|organi[sz]ational\s+chart)\b.*$",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.CultureInvariant)

        If requiresOrgChart AndAlso CountAutoPilotWordVisualsOfType(visuals, "org_chart", "hierarchy") = 0 Then
            validationError = "The current request/document explicitly requires an organization chart, but create_word_document contains no editable org_chart/hierarchy visual. A table is not a substitute."
            Return False
        End If

        Dim chartIntentPattern As System.String =
            "(?:\b(?:umsatz|umsatzentwicklung|revenue|sales|financial|finanz(?:en|entwicklung)?|trend)\b.{0,80}\b(?:chart|diagramm|grafik|graph|visual(?:isierung|ization)?)\b|" &
            "\b(?:chart|diagramm|grafik|graph|visual(?:isierung|ization)?)\b.{0,80}\b(?:umsatz|umsatzentwicklung|revenue|sales|financial|finanz(?:en|entwicklung)?|trend)\b)"
        Dim requiresQuantitativeChart As Boolean =
            System.Text.RegularExpressions.Regex.IsMatch(requestText,
                chartIntentPattern,
                System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.CultureInvariant) OrElse
            System.Text.RegularExpressions.Regex.IsMatch(markdown,
                "(?im)^\s{0,3}(?:#{1,6}\s*)?.*(?:grafische\s+darstellung|visualisierung|visualization|umsatzdiagramm|revenue\s+chart|sales\s+chart|financial\s+chart).*$",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.CultureInvariant)

        If requiresQuantitativeChart AndAlso CountAutoPilotWordVisualsOfType(visuals, "bar_chart", "column_chart", "line_chart", "area_chart", "pie_chart", "doughnut_chart") = 0 Then
            validationError = "The current request/document explicitly requires a quantitative chart, but create_word_document contains no editable native chart visual. Character bars and chart-like tables are not substitutes."
            Return False
        End If

        Dim requiresProcess As Boolean = System.Text.RegularExpressions.Regex.IsMatch(
            requestText,
            "\b(?:flowchart|process\s+diagram|prozessdiagramm|workflow\s+(?:diagram|graphic|chart)|prozessgrafik)\b",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.CultureInvariant)
        If requiresProcess AndAlso CountAutoPilotWordVisualsOfType(visuals, "process") = 0 Then
            validationError = "The current request explicitly requires a process/flow diagram, but create_word_document contains no editable process visual."
            Return False
        End If

        Dim requiresTimeline As Boolean = System.Text.RegularExpressions.Regex.IsMatch(
            requestText,
            "\b(?:timeline|zeitachse|zeitstrahl|chronology\s+graphic)\b",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.CultureInvariant)
        If requiresTimeline AndAlso CountAutoPilotWordVisualsOfType(visuals, "timeline") = 0 Then
            validationError = "The current request explicitly requires a timeline, but create_word_document contains no editable timeline visual."
            Return False
        End If

        Return True
    End Function

    Private Shared Function WordVisualColor(hexColor As String, fallbackHex As String) As System.Drawing.Color
        Dim value As String = If(String.IsNullOrWhiteSpace(hexColor), fallbackHex, hexColor.Trim())
        If Not value.StartsWith("#", StringComparison.Ordinal) Then value = "#" & value
        Try
            Return System.Drawing.ColorTranslator.FromHtml(value)
        Catch
            Return System.Drawing.ColorTranslator.FromHtml(fallbackHex)
        End Try
    End Function

    Private Shared Function CreateWordVisualFont(fontName As String,
                                                 size As Single,
                                                 style As System.Drawing.FontStyle) As System.Drawing.Font
        Dim requested As String = If(String.IsNullOrWhiteSpace(fontName), "Aptos", fontName.Trim())
        Dim normalized As String = requested.ToLowerInvariant()

        ' Never render business/document visuals in a code-oriented monospace face.
        ' If the surrounding request accidentally selected one (often a side effect of
        ' pseudo-diagram markdown), use a professional Office sans-serif instead.
        If normalized.Contains("courier") OrElse
           normalized.Contains("consolas") OrElse
           normalized.Contains("cascadia mono") OrElse
           normalized.Contains("lucida console") OrElse
           normalized.Contains("source code") OrElse
           normalized.Contains("mono") Then

            requested = "Aptos"
        End If

        Try
            Return New System.Drawing.Font(requested, size, style, System.Drawing.GraphicsUnit.Point)
        Catch
            Return New System.Drawing.Font("Arial", size, style, System.Drawing.GraphicsUnit.Point)
        End Try
    End Function

    Private Shared Function GetVisualText(obj As Newtonsoft.Json.Linq.JObject,
                                         key As String,
                                         Optional fallback As String = "") As String
        If obj Is Nothing Then Return fallback
        Dim token As Newtonsoft.Json.Linq.JToken = obj(key)
        If token Is Nothing OrElse token.Type = Newtonsoft.Json.Linq.JTokenType.Null Then Return fallback
        Return token.ToString().Trim()
    End Function

    Private Shared Function GetVisualNumber(obj As Newtonsoft.Json.Linq.JObject,
                                           key As String,
                                           fallback As Double,
                                           minimum As Double,
                                           maximum As Double) As Double
        If obj Is Nothing Then Return fallback
        Try
            Dim token As Newtonsoft.Json.Linq.JToken = obj(key)
            If token Is Nothing OrElse token.Type = Newtonsoft.Json.Linq.JTokenType.Null Then Return fallback
            Dim value As Double = token.Value(Of Double)()
            Return Math.Max(minimum, Math.Min(maximum, value))
        Catch
            Return fallback
        End Try
    End Function

    Private Shared Function GetWordVisualItems(visual As Newtonsoft.Json.Linq.JObject) As List(Of System.Tuple(Of String, String))
        Dim result As New List(Of System.Tuple(Of String, String))()
        If visual Is Nothing Then Return result

        Dim token As Newtonsoft.Json.Linq.JToken = visual("items")
        If token Is Nothing OrElse token.Type <> Newtonsoft.Json.Linq.JTokenType.Array Then Return result

        For Each item As Newtonsoft.Json.Linq.JToken In DirectCast(token, Newtonsoft.Json.Linq.JArray)
            If item Is Nothing OrElse item.Type = Newtonsoft.Json.Linq.JTokenType.Null Then Continue For
            If item.Type = Newtonsoft.Json.Linq.JTokenType.Object Then
                Dim obj As Newtonsoft.Json.Linq.JObject = DirectCast(item, Newtonsoft.Json.Linq.JObject)
                Dim label As String = GetVisualText(obj, "label")
                If String.IsNullOrWhiteSpace(label) Then label = GetVisualText(obj, "title")
                Dim detail As String = GetVisualText(obj, "detail")
                If String.IsNullOrWhiteSpace(detail) Then detail = GetVisualText(obj, "description")
                If Not String.IsNullOrWhiteSpace(label) Then result.Add(System.Tuple.Create(label, detail))
            Else
                Dim label As String = item.ToString().Trim()
                If Not String.IsNullOrWhiteSpace(label) Then result.Add(System.Tuple.Create(label, String.Empty))
            End If
        Next

        Return result
    End Function

    Private Shared Function GetWordVisualCategories(visual As Newtonsoft.Json.Linq.JObject) As List(Of String)
        Dim result As New List(Of String)()
        If visual Is Nothing Then Return result
        Dim token As Newtonsoft.Json.Linq.JToken = visual("categories")
        If token Is Nothing OrElse token.Type <> Newtonsoft.Json.Linq.JTokenType.Array Then Return result
        For Each item As Newtonsoft.Json.Linq.JToken In DirectCast(token, Newtonsoft.Json.Linq.JArray)
            Dim value As String = If(item Is Nothing, String.Empty, item.ToString().Trim())
            If Not String.IsNullOrWhiteSpace(value) Then result.Add(value)
        Next
        Return result
    End Function

    Private Shared Function GetWordVisualSeries(visual As Newtonsoft.Json.Linq.JObject) As List(Of System.Tuple(Of String, List(Of Double)))
        Dim result As New List(Of System.Tuple(Of String, List(Of Double)))()
        If visual Is Nothing Then Return result
        Dim token As Newtonsoft.Json.Linq.JToken = visual("series")
        If token Is Nothing OrElse token.Type <> Newtonsoft.Json.Linq.JTokenType.Array Then Return result

        For Each item As Newtonsoft.Json.Linq.JToken In DirectCast(token, Newtonsoft.Json.Linq.JArray)
            If item Is Nothing OrElse item.Type <> Newtonsoft.Json.Linq.JTokenType.Object Then Continue For
            Dim obj As Newtonsoft.Json.Linq.JObject = DirectCast(item, Newtonsoft.Json.Linq.JObject)
            Dim name As String = GetVisualText(obj, "name", "Series")
            Dim values As New List(Of Double)()
            Dim valuesToken As Newtonsoft.Json.Linq.JToken = obj("values")
            If valuesToken IsNot Nothing AndAlso valuesToken.Type = Newtonsoft.Json.Linq.JTokenType.Array Then
                For Each valueToken As Newtonsoft.Json.Linq.JToken In DirectCast(valuesToken, Newtonsoft.Json.Linq.JArray)
                    Try
                        values.Add(valueToken.Value(Of Double)())
                    Catch
                        values.Add(0.0R)
                    End Try
                Next
            End If
            If values.Count > 0 Then result.Add(System.Tuple.Create(name, values))
        Next
        Return result
    End Function

    Private Shared Sub DrawWordVisualTitle(g As System.Drawing.Graphics,
                                           title As String,
                                           caption As String,
                                           fontName As String,
                                           canvasWidth As Integer)
        Using titleFont As System.Drawing.Font = CreateWordVisualFont(fontName, 28.0F, System.Drawing.FontStyle.Bold),
              captionFont As System.Drawing.Font = CreateWordVisualFont(fontName, 14.0F, System.Drawing.FontStyle.Regular),
              textBrush As New System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(30, 35, 42)),
              captionBrush As New System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(90, 98, 108))
            Dim titleRect As New System.Drawing.RectangleF(70.0F, 45.0F, CSng(canvasWidth - 140), 55.0F)
            g.DrawString(If(String.IsNullOrWhiteSpace(title), "Visual", title), titleFont, textBrush, titleRect)
            If Not String.IsNullOrWhiteSpace(caption) Then
                Dim captionRect As New System.Drawing.RectangleF(70.0F, 100.0F, CSng(canvasWidth - 140), 48.0F)
                g.DrawString(caption, captionFont, captionBrush, captionRect)
            End If
        End Using
    End Sub

    Private Shared Function WordVisualSeriesColor(index As Integer, accent As System.Drawing.Color) As System.Drawing.Color
        If index <= 0 Then Return accent
        Dim palette() As System.Drawing.Color = {
            System.Drawing.Color.FromArgb(68, 114, 196),
            System.Drawing.Color.FromArgb(112, 173, 71),
            System.Drawing.Color.FromArgb(237, 125, 49),
            System.Drawing.Color.FromArgb(165, 165, 165),
            System.Drawing.Color.FromArgb(91, 155, 213)
        }
        Return palette((index - 1) Mod palette.Length)
    End Function

    Private Shared Sub DrawProcessWordVisual(g As System.Drawing.Graphics,
                                             visual As Newtonsoft.Json.Linq.JObject,
                                             fontName As String,
                                             accent As System.Drawing.Color,
                                             canvasWidth As Integer,
                                             canvasHeight As Integer)
        Dim items As List(Of System.Tuple(Of String, String)) = GetWordVisualItems(visual)
        If items.Count = 0 Then items.Add(System.Tuple.Create("Process", String.Empty))
        If items.Count > 7 Then items = items.GetRange(0, 7)

        Dim left As Single = 80.0F
        Dim right As Single = 80.0F
        Dim gap As Single = 34.0F
        Dim usable As Single = CSng(canvasWidth) - left - right
        Dim boxWidth As Single = Math.Max(135.0F, (usable - gap * (items.Count - 1)) / Math.Max(1, items.Count))
        Dim totalWidth As Single = boxWidth * items.Count + gap * (items.Count - 1)
        If totalWidth > usable Then
            gap = 20.0F
            boxWidth = (usable - gap * (items.Count - 1)) / Math.Max(1, items.Count)
        End If
        Dim boxHeight As Single = 170.0F
        Dim y As Single = CSng(canvasHeight) / 2.0F - boxHeight / 2.0F + 35.0F

        Using borderPen As New System.Drawing.Pen(accent, 3.0F),
              arrowPen As New System.Drawing.Pen(accent, 4.0F),
              fillBrush As New System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(18, accent)),
              titleFont As System.Drawing.Font = CreateWordVisualFont(fontName, 17.0F, System.Drawing.FontStyle.Bold),
              detailFont As System.Drawing.Font = CreateWordVisualFont(fontName, 12.0F, System.Drawing.FontStyle.Regular),
              textBrush As New System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(30, 35, 42)),
              detailBrush As New System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(85, 92, 102))

            For i As Integer = 0 To items.Count - 1
                Dim x As Single = left + i * (boxWidth + gap)
                Dim rect As New System.Drawing.RectangleF(x, y, boxWidth, boxHeight)
                g.FillRectangle(fillBrush, rect)
                g.DrawRectangle(borderPen, x, y, boxWidth, boxHeight)

                Dim sf As New System.Drawing.StringFormat() With {
                    .Alignment = System.Drawing.StringAlignment.Center,
                    .LineAlignment = System.Drawing.StringAlignment.Near,
                    .Trimming = System.Drawing.StringTrimming.EllipsisWord
                }
                Dim titleRect As New System.Drawing.RectangleF(x + 14.0F, y + 26.0F, boxWidth - 28.0F, 58.0F)
                g.DrawString(items(i).Item1, titleFont, textBrush, titleRect, sf)
                If Not String.IsNullOrWhiteSpace(items(i).Item2) Then
                    Dim detailRect As New System.Drawing.RectangleF(x + 14.0F, y + 88.0F, boxWidth - 28.0F, 60.0F)
                    g.DrawString(items(i).Item2, detailFont, detailBrush, detailRect, sf)
                End If
                sf.Dispose()

                If i < items.Count - 1 Then
                    Dim x1 As Single = x + boxWidth + 5.0F
                    Dim x2 As Single = x + boxWidth + gap - 5.0F
                    Dim midY As Single = y + boxHeight / 2.0F
                    g.DrawLine(arrowPen, x1, midY, x2, midY)
                    Dim arrow() As System.Drawing.PointF = {
                        New System.Drawing.PointF(x2, midY),
                        New System.Drawing.PointF(x2 - 13.0F, midY - 8.0F),
                        New System.Drawing.PointF(x2 - 13.0F, midY + 8.0F)
                    }
                    Using arrowBrush As New System.Drawing.SolidBrush(accent)
                        g.FillPolygon(arrowBrush, arrow)
                    End Using
                End If
            Next
        End Using
    End Sub

    Private Shared Function GetWordOrgChartNodes(visual As Newtonsoft.Json.Linq.JObject) As List(Of Newtonsoft.Json.Linq.JObject)
        Dim result As New List(Of Newtonsoft.Json.Linq.JObject)()
        If visual Is Nothing Then Return result
        Dim token As Newtonsoft.Json.Linq.JToken = visual("nodes")
        If token Is Nothing OrElse token.Type <> Newtonsoft.Json.Linq.JTokenType.Array Then Return result

        Dim seen As New HashSet(Of System.String)(StringComparer.Ordinal)
        For Each item As Newtonsoft.Json.Linq.JToken In DirectCast(token, Newtonsoft.Json.Linq.JArray)
            If item Is Nothing OrElse item.Type <> Newtonsoft.Json.Linq.JTokenType.Object Then Continue For
            Dim node As Newtonsoft.Json.Linq.JObject = DirectCast(item, Newtonsoft.Json.Linq.JObject)
            Dim id As String = GetVisualText(node, "id")
            Dim label As String = GetVisualText(node, "label")
            If String.IsNullOrWhiteSpace(id) OrElse String.IsNullOrWhiteSpace(label) Then Continue For
            If seen.Add(id) Then result.Add(node)
            If result.Count >= 24 Then Exit For
        Next
        Return result
    End Function

    Private Shared Function GetWordOrgChartDepth(node As Newtonsoft.Json.Linq.JObject,
                                                  byId As Dictionary(Of String, Newtonsoft.Json.Linq.JObject)) As Integer
        If node Is Nothing Then Return 0
        Dim depth As Integer = 0
        Dim current As Newtonsoft.Json.Linq.JObject = node
        Dim visited As New HashSet(Of System.String)(StringComparer.Ordinal)
        While current IsNot Nothing AndAlso depth < 6
            Dim parentId As String = GetVisualText(current, "parent_id")
            If String.IsNullOrWhiteSpace(parentId) Then Exit While
            If Not visited.Add(parentId) Then Exit While
            Dim parent As Newtonsoft.Json.Linq.JObject = Nothing
            If Not byId.TryGetValue(parentId, parent) Then Exit While
            depth += 1
            current = parent
        End While
        Return depth
    End Function

    Private Shared Sub DrawOrgChartWordVisual(g As System.Drawing.Graphics,
                                               visual As Newtonsoft.Json.Linq.JObject,
                                               fontName As String,
                                               accent As System.Drawing.Color,
                                               canvasWidth As Integer,
                                               canvasHeight As Integer)
        Dim nodes As List(Of Newtonsoft.Json.Linq.JObject) = GetWordOrgChartNodes(visual)
        If nodes.Count = 0 Then
            DrawProcessWordVisual(g, visual, fontName, accent, canvasWidth, canvasHeight)
            Return
        End If

        Dim byId As New Dictionary(Of String, Newtonsoft.Json.Linq.JObject)(StringComparer.Ordinal)
        For Each node As Newtonsoft.Json.Linq.JObject In nodes
            byId(GetVisualText(node, "id")) = node
        Next

        Dim levels As New SortedDictionary(Of Integer, List(Of Newtonsoft.Json.Linq.JObject))()
        Dim maxDepth As Integer = 0
        For Each node As Newtonsoft.Json.Linq.JObject In nodes
            Dim depth As Integer = GetWordOrgChartDepth(node, byId)
            maxDepth = Math.Max(maxDepth, depth)
            If Not levels.ContainsKey(depth) Then levels(depth) = New List(Of Newtonsoft.Json.Linq.JObject)()
            levels(depth).Add(node)
        Next

        ' Keep raster fallback readable at the final Word display size. The former
        ' implementation forced every hierarchy level into one row and could shrink
        ' node boxes to only 78 px, which made labels crowded and effectively tiny
        ' after Word scaled the image to the printable page width.
        Dim left As Single = 70.0F
        Dim right As Single = 70.0F
        Dim top As Single = 170.0F
        Dim bottom As Single = 55.0F
        Dim usableWidth As Single = Math.Max(420.0F, CSng(canvasWidth) - left - right)
        Dim usableHeight As Single = Math.Max(320.0F, CSng(canvasHeight) - top - bottom)
        Dim minBoxWidth As Single = 205.0F
        Dim columnGap As Single = 30.0F
        Dim rowGap As Single = 26.0F
        Dim levelGap As Single = 48.0F
        Dim boxHeight As Single = 112.0F

        Dim columnsPerRow As Integer = Math.Max(1, CInt(Math.Floor((usableWidth + columnGap) / (minBoxWidth + columnGap))))
        columnsPerRow = Math.Min(4, columnsPerRow)

        Dim levelRows As New Dictionary(Of Integer, Integer)()
        Dim requiredHeight As Single = 0.0F
        For Each kvp As KeyValuePair(Of Integer, List(Of Newtonsoft.Json.Linq.JObject)) In levels
            Dim rows As Integer = Math.Max(1, CInt(Math.Ceiling(kvp.Value.Count / CDbl(columnsPerRow))))
            levelRows(kvp.Key) = rows
            requiredHeight += rows * boxHeight
            requiredHeight += Math.Max(0, rows - 1) * rowGap
            If kvp.Key < maxDepth Then requiredHeight += levelGap
        Next

        ' If a very large hierarchy must fit into the requested image height, compact
        ' geometry moderately but preserve readable text and meaningful white space.
        If requiredHeight > usableHeight Then
            boxHeight = 96.0F
            rowGap = 20.0F
            levelGap = 34.0F
            requiredHeight = 0.0F
            For Each kvp As KeyValuePair(Of Integer, List(Of Newtonsoft.Json.Linq.JObject)) In levels
                requiredHeight += levelRows(kvp.Key) * boxHeight
                requiredHeight += Math.Max(0, levelRows(kvp.Key) - 1) * rowGap
                If kvp.Key < maxDepth Then requiredHeight += levelGap
            Next
        End If

        Dim rects As New Dictionary(Of String, System.Drawing.RectangleF)(StringComparer.Ordinal)
        Dim currentY As Single = top + Math.Max(0.0F, (usableHeight - Math.Min(requiredHeight, usableHeight)) / 2.0F)

        For Each kvp As KeyValuePair(Of Integer, List(Of Newtonsoft.Json.Linq.JObject)) In levels
            Dim levelNodes As List(Of Newtonsoft.Json.Linq.JObject) = kvp.Value
            Dim rows As Integer = levelRows(kvp.Key)

            For row As Integer = 0 To rows - 1
                Dim startIndex As Integer = row * columnsPerRow
                Dim rowCount As Integer = Math.Min(columnsPerRow, levelNodes.Count - startIndex)
                If rowCount <= 0 Then Continue For

                Dim boxWidth As Single = Math.Min(255.0F, (usableWidth - columnGap * Math.Max(0, rowCount - 1)) / rowCount)
                boxWidth = Math.Max(minBoxWidth, boxWidth)
                Dim rowWidth As Single = boxWidth * rowCount + columnGap * Math.Max(0, rowCount - 1)
                Dim rowLeft As Single = left + Math.Max(0.0F, (usableWidth - rowWidth) / 2.0F)

                For column As Integer = 0 To rowCount - 1
                    Dim node As Newtonsoft.Json.Linq.JObject = levelNodes(startIndex + column)
                    Dim id As String = GetVisualText(node, "id")
                    rects(id) = New System.Drawing.RectangleF(
                        rowLeft + column * (boxWidth + columnGap),
                        currentY,
                        boxWidth,
                        boxHeight)
                Next

                currentY += boxHeight
                If row < rows - 1 Then currentY += rowGap
            Next

            If kvp.Key < maxDepth Then currentY += levelGap
        Next

        Using connectorPen As New System.Drawing.Pen(System.Drawing.Color.FromArgb(145, 154, 165), 2.5F),
              borderPen As New System.Drawing.Pen(accent, 2.8F),
              rootBrush As New System.Drawing.SolidBrush(accent),
              childBrush As New System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(18, accent)),
              titleFont As System.Drawing.Font = CreateWordVisualFont(fontName, 13.0F, System.Drawing.FontStyle.Bold),
              detailFont As System.Drawing.Font = CreateWordVisualFont(fontName, 9.75F, System.Drawing.FontStyle.Regular),
              rootTextBrush As New System.Drawing.SolidBrush(System.Drawing.Color.White),
              textBrush As New System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(30, 35, 42)),
              detailBrush As New System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(80, 88, 98))

            ' Draw reporting lines first so node boxes remain visually dominant.
            For Each node As Newtonsoft.Json.Linq.JObject In nodes
                Dim id As String = GetVisualText(node, "id")
                Dim parentId As String = GetVisualText(node, "parent_id")
                If String.IsNullOrWhiteSpace(parentId) OrElse Not rects.ContainsKey(id) OrElse Not rects.ContainsKey(parentId) Then Continue For
                Dim childRect As System.Drawing.RectangleF = rects(id)
                Dim parentRect As System.Drawing.RectangleF = rects(parentId)
                Dim px As Single = parentRect.Left + parentRect.Width / 2.0F
                Dim py As Single = parentRect.Bottom
                Dim cx As Single = childRect.Left + childRect.Width / 2.0F
                Dim cy As Single = childRect.Top
                Dim midY As Single = py + Math.Max(15.0F, (cy - py) / 2.0F)
                g.DrawLine(connectorPen, px, py, px, midY)
                g.DrawLine(connectorPen, px, midY, cx, midY)
                g.DrawLine(connectorPen, cx, midY, cx, cy)
            Next

            For Each node As Newtonsoft.Json.Linq.JObject In nodes
                Dim id As String = GetVisualText(node, "id")
                If Not rects.ContainsKey(id) Then Continue For
                Dim rect As System.Drawing.RectangleF = rects(id)
                Dim isRoot As Boolean = String.IsNullOrWhiteSpace(GetVisualText(node, "parent_id"))
                If isRoot Then
                    g.FillRectangle(rootBrush, rect)
                Else
                    g.FillRectangle(childBrush, rect)
                    g.DrawRectangle(borderPen, rect.X, rect.Y, rect.Width, rect.Height)
                End If

                Dim sf As New System.Drawing.StringFormat() With {
                    .Alignment = System.Drawing.StringAlignment.Center,
                    .LineAlignment = System.Drawing.StringAlignment.Near,
                    .Trimming = System.Drawing.StringTrimming.EllipsisWord
                }
                Dim label As String = GetVisualText(node, "label")
                Dim detail As String = GetVisualText(node, "detail")
                Dim labelHeight As Single = If(String.IsNullOrWhiteSpace(detail), rect.Height - 24.0F, Math.Min(48.0F, rect.Height * 0.46F))
                Dim labelTop As Single = If(String.IsNullOrWhiteSpace(detail),
                                            rect.Y + Math.Max(12.0F, (rect.Height - labelHeight) / 2.0F),
                                            rect.Y + 14.0F)
                Dim labelRect As New System.Drawing.RectangleF(rect.X + 14.0F, labelTop, rect.Width - 28.0F, labelHeight)
                g.DrawString(label, titleFont, If(isRoot, rootTextBrush, textBrush), labelRect, sf)

                If Not String.IsNullOrWhiteSpace(detail) Then
                    Dim detailTop As Single = rect.Y + Math.Max(56.0F, rect.Height * 0.50F)
                    Dim detailRect As New System.Drawing.RectangleF(
                        rect.X + 14.0F,
                        detailTop,
                        rect.Width - 28.0F,
                        Math.Max(22.0F, rect.Bottom - detailTop - 12.0F))
                    g.DrawString(detail, detailFont, If(isRoot, rootTextBrush, detailBrush), detailRect, sf)
                End If
                sf.Dispose()
            Next
        End Using
    End Sub

    Private Shared Function MeasureWordVisualLegendHeight(g As System.Drawing.Graphics,
                                                           series As List(Of System.Tuple(Of String, List(Of Double))),
                                                           maxSeries As Integer,
                                                           legendFont As System.Drawing.Font,
                                                           availableWidth As Single) As Single
        If series Is Nothing OrElse maxSeries <= 0 Then Return 0.0F
        Dim x As Single = 0.0F
        Dim rows As Integer = 1
        For s As Integer = 0 To maxSeries - 1
            Dim name As String = If(series(s).Item1, String.Empty)
            Dim textWidth As Single = g.MeasureString(name, legendFont).Width
            Dim entryWidth As Single = Math.Min(availableWidth, 45.0F + textWidth + 18.0F)
            If x > 0.0F AndAlso x + entryWidth > availableWidth Then
                rows += 1
                x = 0.0F
            End If
            x += entryWidth + 18.0F
        Next
        Return rows * 30.0F
    End Function

    Private Shared Sub DrawWordVisualLegend(g As System.Drawing.Graphics,
                                             series As List(Of System.Tuple(Of String, List(Of Double))),
                                             maxSeries As Integer,
                                             legendFont As System.Drawing.Font,
                                             textBrush As System.Drawing.Brush,
                                             accent As System.Drawing.Color,
                                             left As Single,
                                             top As Single,
                                             availableWidth As Single)
        If series Is Nothing OrElse maxSeries <= 0 Then Return
        Dim x As Single = left
        Dim y As Single = top
        For s As Integer = 0 To maxSeries - 1
            Dim name As String = If(series(s).Item1, String.Empty)
            Dim textWidth As Single = g.MeasureString(name, legendFont).Width
            Dim entryWidth As Single = Math.Min(availableWidth, 45.0F + textWidth + 18.0F)
            If x > left AndAlso x + entryWidth > left + availableWidth Then
                x = left
                y += 30.0F
            End If
            Using b As New System.Drawing.SolidBrush(WordVisualSeriesColor(s, accent))
                g.FillRectangle(b, x, y + 4.0F, 16.0F, 16.0F)
            End Using
            Dim textRect As New System.Drawing.RectangleF(x + 24.0F, y, Math.Max(35.0F, entryWidth - 24.0F), 25.0F)
            g.DrawString(name, legendFont, textBrush, textRect)
            x += entryWidth + 18.0F
        Next
    End Sub

    Private Shared Sub DrawTimelineWordVisual(g As System.Drawing.Graphics,
                                              visual As Newtonsoft.Json.Linq.JObject,
                                              fontName As String,
                                              accent As System.Drawing.Color,
                                              canvasWidth As Integer,
                                              canvasHeight As Integer)
        Dim items As List(Of System.Tuple(Of String, String)) = GetWordVisualItems(visual)
        If items.Count = 0 Then items.Add(System.Tuple.Create("Milestone", String.Empty))
        If items.Count > 8 Then items = items.GetRange(0, 8)

        Dim left As Single = 100.0F
        Dim right As Single = 100.0F
        Dim y As Single = CSng(canvasHeight) / 2.0F + 45.0F
        Dim stepWidth As Single = If(items.Count <= 1, 0.0F, (CSng(canvasWidth) - left - right) / (items.Count - 1))

        Using linePen As New System.Drawing.Pen(System.Drawing.Color.FromArgb(170, 177, 187), 4.0F),
              nodeBrush As New System.Drawing.SolidBrush(accent),
              titleFont As System.Drawing.Font = CreateWordVisualFont(fontName, 15.0F, System.Drawing.FontStyle.Bold),
              detailFont As System.Drawing.Font = CreateWordVisualFont(fontName, 11.0F, System.Drawing.FontStyle.Regular),
              textBrush As New System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(30, 35, 42)),
              detailBrush As New System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(85, 92, 102))
            g.DrawLine(linePen, left, y, CSng(canvasWidth) - right, y)

            For i As Integer = 0 To items.Count - 1
                Dim x As Single = If(items.Count <= 1, CSng(canvasWidth) / 2.0F, left + i * stepWidth)
                g.FillEllipse(nodeBrush, x - 12.0F, y - 12.0F, 24.0F, 24.0F)
                Dim above As Boolean = (i Mod 2 = 0)
                Dim textY As Single = If(above, y - 150.0F, y + 34.0F)
                Dim rect As New System.Drawing.RectangleF(x - 100.0F, textY, 200.0F, 110.0F)
                Dim sf As New System.Drawing.StringFormat() With {
                    .Alignment = System.Drawing.StringAlignment.Center,
                    .Trimming = System.Drawing.StringTrimming.EllipsisWord
                }
                g.DrawString(items(i).Item1, titleFont, textBrush, rect, sf)
                If Not String.IsNullOrWhiteSpace(items(i).Item2) Then
                    Dim detailRect As New System.Drawing.RectangleF(rect.X, rect.Y + 42.0F, rect.Width, 62.0F)
                    g.DrawString(items(i).Item2, detailFont, detailBrush, detailRect, sf)
                End If
                sf.Dispose()
            Next
        End Using
    End Sub

    Private Shared Sub DrawBarWordVisual(g As System.Drawing.Graphics,
                                         visual As Newtonsoft.Json.Linq.JObject,
                                         fontName As String,
                                         accent As System.Drawing.Color,
                                         canvasWidth As Integer,
                                         canvasHeight As Integer)
        Dim categories As List(Of String) = GetWordVisualCategories(visual)
        Dim series As List(Of System.Tuple(Of String, List(Of Double))) = GetWordVisualSeries(visual)
        If categories.Count = 0 OrElse series.Count = 0 Then
            DrawProcessWordVisual(g, visual, fontName, accent, canvasWidth, canvasHeight)
            Return
        End If

        Dim maxCategories As Integer = Math.Min(10, categories.Count)
        Dim maxSeries As Integer = Math.Min(5, series.Count)
        Dim maxValue As Double = 0.0R
        For s As Integer = 0 To maxSeries - 1
            For c As Integer = 0 To Math.Min(maxCategories, series(s).Item2.Count) - 1
                maxValue = Math.Max(maxValue, Math.Abs(series(s).Item2(c)))
            Next
        Next
        If maxValue <= 0.0R Then maxValue = 1.0R

        Dim plotLeft As Single = 105.0F
        Dim plotRight As Single = 80.0F
        Dim plotBottom As Single = 125.0F
        Dim plotWidth As Single = CSng(canvasWidth) - plotLeft - plotRight

        Using axisPen As New System.Drawing.Pen(System.Drawing.Color.FromArgb(160, 168, 178), 2.0F),
              labelFont As System.Drawing.Font = CreateWordVisualFont(fontName, 11.0F, System.Drawing.FontStyle.Regular),
              legendFont As System.Drawing.Font = CreateWordVisualFont(fontName, 11.0F, System.Drawing.FontStyle.Bold),
              textBrush As New System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(60, 66, 74))
            Dim legendTop As Single = 148.0F
            Dim legendHeight As Single = MeasureWordVisualLegendHeight(g, series, maxSeries, legendFont, plotWidth)
            Dim plotTop As Single = legendTop + legendHeight + 18.0F
            Dim plotHeight As Single = Math.Max(150.0F, CSng(canvasHeight) - plotTop - plotBottom)
            DrawWordVisualLegend(g, series, maxSeries, legendFont, textBrush, accent, plotLeft, legendTop, plotWidth)
            g.DrawLine(axisPen, plotLeft, plotTop, plotLeft, plotTop + plotHeight)
            g.DrawLine(axisPen, plotLeft, plotTop + plotHeight, plotLeft + plotWidth, plotTop + plotHeight)

            Dim groupWidth As Single = plotWidth / maxCategories
            Dim innerWidth As Single = groupWidth * 0.72F
            Dim barWidth As Single = Math.Max(7.0F, innerWidth / maxSeries)

            For c As Integer = 0 To maxCategories - 1
                Dim groupStart As Single = plotLeft + c * groupWidth + (groupWidth - innerWidth) / 2.0F
                For s As Integer = 0 To maxSeries - 1
                    Dim value As Double = If(c < series(s).Item2.Count, series(s).Item2(c), 0.0R)
                    Dim h As Single = CSng(Math.Abs(value) / maxValue) * (plotHeight - 20.0F)
                    Dim x As Single = groupStart + s * barWidth
                    Dim y As Single = plotTop + plotHeight - h
                    Using b As New System.Drawing.SolidBrush(WordVisualSeriesColor(s, accent))
                        g.FillRectangle(b, x, y, Math.Max(4.0F, barWidth - 3.0F), h)
                    End Using
                Next

                Dim sf As New System.Drawing.StringFormat() With {
                    .Alignment = System.Drawing.StringAlignment.Center,
                    .Trimming = System.Drawing.StringTrimming.EllipsisCharacter
                }
                Dim labelRect As New System.Drawing.RectangleF(plotLeft + c * groupWidth, plotTop + plotHeight + 10.0F, groupWidth, 48.0F)
                g.DrawString(categories(c), labelFont, textBrush, labelRect, sf)
                sf.Dispose()
            Next

        End Using
    End Sub

    Private Shared Sub DrawLineWordVisual(g As System.Drawing.Graphics,
                                          visual As Newtonsoft.Json.Linq.JObject,
                                          fontName As String,
                                          accent As System.Drawing.Color,
                                          canvasWidth As Integer,
                                          canvasHeight As Integer)
        Dim categories As List(Of String) = GetWordVisualCategories(visual)
        Dim series As List(Of System.Tuple(Of String, List(Of Double))) = GetWordVisualSeries(visual)
        If categories.Count = 0 OrElse series.Count = 0 Then
            DrawProcessWordVisual(g, visual, fontName, accent, canvasWidth, canvasHeight)
            Return
        End If

        Dim maxCategories As Integer = Math.Min(12, categories.Count)
        Dim maxSeries As Integer = Math.Min(5, series.Count)
        Dim minValue As Double = Double.MaxValue
        Dim maxValue As Double = Double.MinValue
        For s As Integer = 0 To maxSeries - 1
            For c As Integer = 0 To Math.Min(maxCategories, series(s).Item2.Count) - 1
                minValue = Math.Min(minValue, series(s).Item2(c))
                maxValue = Math.Max(maxValue, series(s).Item2(c))
            Next
        Next
        If minValue = Double.MaxValue Then minValue = 0.0R
        If maxValue = Double.MinValue Then maxValue = 1.0R
        If Math.Abs(maxValue - minValue) < 0.000001R Then
            minValue -= 1.0R
            maxValue += 1.0R
        End If

        Dim plotLeft As Single = 105.0F
        Dim plotRight As Single = 80.0F
        Dim plotBottom As Single = 125.0F
        Dim plotWidth As Single = CSng(canvasWidth) - plotLeft - plotRight

        Using axisPen As New System.Drawing.Pen(System.Drawing.Color.FromArgb(160, 168, 178), 2.0F),
              labelFont As System.Drawing.Font = CreateWordVisualFont(fontName, 11.0F, System.Drawing.FontStyle.Regular),
              legendFont As System.Drawing.Font = CreateWordVisualFont(fontName, 11.0F, System.Drawing.FontStyle.Bold),
              textBrush As New System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(60, 66, 74))
            Dim legendTop As Single = 148.0F
            Dim legendHeight As Single = MeasureWordVisualLegendHeight(g, series, maxSeries, legendFont, plotWidth)
            Dim plotTop As Single = legendTop + legendHeight + 18.0F
            Dim plotHeight As Single = Math.Max(150.0F, CSng(canvasHeight) - plotTop - plotBottom)
            DrawWordVisualLegend(g, series, maxSeries, legendFont, textBrush, accent, plotLeft, legendTop, plotWidth)
            g.DrawLine(axisPen, plotLeft, plotTop, plotLeft, plotTop + plotHeight)
            g.DrawLine(axisPen, plotLeft, plotTop + plotHeight, plotLeft + plotWidth, plotTop + plotHeight)

            Dim xStep As Single = If(maxCategories <= 1, 0.0F, plotWidth / (maxCategories - 1))
            For s As Integer = 0 To maxSeries - 1
                Dim points As New List(Of System.Drawing.PointF)()
                For c As Integer = 0 To Math.Min(maxCategories, series(s).Item2.Count) - 1
                    Dim x As Single = If(maxCategories <= 1, plotLeft + plotWidth / 2.0F, plotLeft + c * xStep)
                    Dim ratio As Double = (series(s).Item2(c) - minValue) / (maxValue - minValue)
                    Dim y As Single = plotTop + plotHeight - CSng(ratio) * plotHeight
                    points.Add(New System.Drawing.PointF(x, y))
                Next
                Dim seriesColor As System.Drawing.Color = WordVisualSeriesColor(s, accent)
                If points.Count > 1 Then
                    Using p As New System.Drawing.Pen(seriesColor, 4.0F)
                        g.DrawLines(p, points.ToArray())
                    End Using
                End If
                Using b As New System.Drawing.SolidBrush(seriesColor)
                    For Each point As System.Drawing.PointF In points
                        g.FillEllipse(b, point.X - 5.0F, point.Y - 5.0F, 10.0F, 10.0F)
                    Next
                End Using
            Next

            For c As Integer = 0 To maxCategories - 1
                Dim x As Single = If(maxCategories <= 1, plotLeft + plotWidth / 2.0F, plotLeft + c * xStep)
                Dim sf As New System.Drawing.StringFormat() With {
                    .Alignment = System.Drawing.StringAlignment.Center,
                    .Trimming = System.Drawing.StringTrimming.EllipsisCharacter
                }
                Dim labelRect As New System.Drawing.RectangleF(x - 60.0F, plotTop + plotHeight + 10.0F, 120.0F, 48.0F)
                g.DrawString(categories(c), labelFont, textBrush, labelRect, sf)
                sf.Dispose()
            Next

        End Using
    End Sub

    Private Shared Function RenderAutoPilotWordVisual(visual As Newtonsoft.Json.Linq.JObject,
                                                      outputPath As String,
                                                      fontName As String,
                                                      accentHex As String) As Boolean
        If visual Is Nothing Then Return False

        Dim widthInches As Double = GetVisualNumber(visual, "width_inches", 8.4R, 4.0R, 10.5R)
        Dim heightInches As Double = GetVisualNumber(visual, "height_inches", 4.7R, 2.5R, 7.0R)
        Dim canvasWidth As Integer = CInt(Math.Round(widthInches * 150.0R))
        Dim canvasHeight As Integer = CInt(Math.Round(heightInches * 150.0R))
        Dim accent As System.Drawing.Color = WordVisualColor(accentHex, "#17365D")
        Dim visualType As System.String = GetVisualText(visual, "type", "process").ToLowerInvariant()
        Dim title As String = GetVisualText(visual, "title")
        Dim caption As String = GetVisualText(visual, "caption")

        Try
            Using bitmap As New System.Drawing.Bitmap(canvasWidth, canvasHeight, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
                bitmap.SetResolution(150.0F, 150.0F)
                Using g As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(bitmap)
                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias
                    g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit
                    g.Clear(System.Drawing.Color.White)
                    DrawWordVisualTitle(g, title, caption, fontName, canvasWidth)

                    Select Case visualType
                        Case "timeline"
                            DrawTimelineWordVisual(g, visual, fontName, accent, canvasWidth, canvasHeight)
                        Case "org_chart"
                            DrawOrgChartWordVisual(g, visual, fontName, accent, canvasWidth, canvasHeight)
                        Case "bar_chart"
                            DrawBarWordVisual(g, visual, fontName, accent, canvasWidth, canvasHeight)
                        Case "line_chart"
                            DrawLineWordVisual(g, visual, fontName, accent, canvasWidth, canvasHeight)
                        Case Else
                            DrawProcessWordVisual(g, visual, fontName, accent, canvasWidth, canvasHeight)
                    End Select
                End Using
                bitmap.Save(outputPath, System.Drawing.Imaging.ImageFormat.Png)
            End Using
            Return File.Exists(outputPath) AndAlso New FileInfo(outputPath).Length > 0
        Catch ex As System.Exception
            Debug.WriteLine($"Word visual render error: {ex.Message}")
            Return False
        End Try
    End Function

    Private Shared Function GetAutoPilotWordVisualEditable(visual As Newtonsoft.Json.Linq.JObject) As Boolean
        If visual Is Nothing Then Return True
        Dim editableToken As Newtonsoft.Json.Linq.JToken = visual("editable")
        If editableToken IsNot Nothing AndAlso editableToken.Type = Newtonsoft.Json.Linq.JTokenType.Boolean Then
            Return CBool(editableToken)
        End If
        Return True
    End Function

    Private Shared Function GetAutoPilotWordVisualInsertionMode(visual As Newtonsoft.Json.Linq.JObject) As String
        Dim mode As String = GetVisualText(visual, "insertion_mode", "auto").ToLowerInvariant()
        Select Case mode
            Case "inline", "floating", "auto"
                Return mode
            Case Else
                Return "auto"
        End Select
    End Function

    Private Shared Function FindAutoPilotWordSmartArtLayout(doc As Microsoft.Office.Interop.Word.Document,
                                                             visual As Newtonsoft.Json.Linq.JObject,
                                                             visualType As String,
                                                             ByRef warning As String) As Microsoft.Office.Core.SmartArtLayout
        warning = String.Empty
        If doc Is Nothing Then Return Nothing

        Dim layouts As Microsoft.Office.Core.SmartArtLayouts = Nothing
        Dim bestLayout As Microsoft.Office.Core.SmartArtLayout = Nothing
        Dim bestScore As Integer = Integer.MinValue
        Dim requestedLayoutRaw As String = GetVisualText(visual, "smartart_layout")
        Dim requestedLayout As String = requestedLayoutRaw.ToLowerInvariant()

        Try
            layouts = doc.Application.SmartArtLayouts
            If layouts Is Nothing OrElse layouts.Count <= 0 Then
                warning = "Word did not expose any SmartArt layouts."
                Return Nothing
            End If

            ' Prefer stable built-in SmartArt layout IDs where Microsoft Office exposes them.
            ' This avoids depending on localized display names (for example German vs. English Word).
            Dim directLayoutIds As New List(Of String)()
            If Not String.IsNullOrWhiteSpace(requestedLayoutRaw) Then directLayoutIds.Add(requestedLayoutRaw.Trim())
            Select Case visualType
                Case "org_chart", "hierarchy"
                    directLayoutIds.Add("urn:microsoft.com/office/officeart/2005/8/layout/orgChart1")
                Case "process"
                    directLayoutIds.Add("urn:microsoft.com/office/officeart/2005/8/layout/process1")
                Case "cycle"
                    directLayoutIds.Add("urn:microsoft.com/office/officeart/2005/8/layout/cycle1")
            End Select

            For Each layoutId As String In directLayoutIds
                If String.IsNullOrWhiteSpace(layoutId) Then Continue For
                Dim directLayout As Microsoft.Office.Core.SmartArtLayout = Nothing
                Try
                    directLayout = layouts.Item(layoutId)
                    If directLayout IsNot Nothing Then Return directLayout
                Catch ex As System.Exception
                    If directLayout IsNot Nothing Then
                        Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(directLayout) : Catch releaseEx As System.Exception : End Try
                    End If
                End Try
            Next

            ' Fall back to localized metadata scoring for layouts without a stable built-in ID
            ' mapping here, and for installations where a built-in ID is unavailable.
            For index As Integer = 1 To layouts.Count
                Dim layout As Microsoft.Office.Core.SmartArtLayout = Nothing
                Try
                    layout = layouts.Item(index)
                    If layout Is Nothing Then Continue For

                    Dim name As String = String.Empty
                    Dim category As String = String.Empty
                    Dim description As String = String.Empty
                    Dim layoutId As String = String.Empty
                    Try : name = If(layout.Name, String.Empty).ToLowerInvariant() : Catch ex As System.Exception : End Try
                    Try : category = If(layout.Category, String.Empty).ToLowerInvariant() : Catch ex As System.Exception : End Try
                    Try : description = If(layout.Description, String.Empty).ToLowerInvariant() : Catch ex As System.Exception : End Try
                    Try : layoutId = If(layout.Id, String.Empty).ToLowerInvariant() : Catch ex As System.Exception : End Try

                    Dim haystack As String = name & " " & category & " " & description & " " & layoutId
                    Dim score As Integer = 0

                    If Not String.IsNullOrWhiteSpace(requestedLayout) Then
                        If String.Equals(name, requestedLayout, StringComparison.OrdinalIgnoreCase) OrElse
                           String.Equals(layoutId, requestedLayout, StringComparison.OrdinalIgnoreCase) Then
                            score += 2000
                        ElseIf haystack.Contains(requestedLayout) Then
                            score += 1200
                        End If
                    End If

                    Select Case visualType
                        Case "org_chart", "hierarchy"
                            If name.Contains("organization chart") OrElse name.Contains("organisation chart") OrElse name.Contains("organigram") Then score += 900
                            If name.Contains("hierarchy") OrElse name.Contains("hierarchie") Then score += 500
                            If category.Contains("hierarchy") OrElse category.Contains("hierarchie") Then score += 450
                            If description.Contains("organization") OrElse description.Contains("organisation") OrElse description.Contains("hierarchy") OrElse description.Contains("hierarchie") Then score += 250
                        Case "timeline"
                            If name.Contains("timeline") OrElse name.Contains("zeitachse") OrElse name.Contains("zeitlinie") Then score += 900
                            If category.Contains("process") OrElse category.Contains("prozess") Then score += 250
                            If description.Contains("timeline") OrElse description.Contains("zeitachse") OrElse description.Contains("chronolog") Then score += 300
                        Case "cycle"
                            If name.Contains("cycle") OrElse name.Contains("zyklus") OrElse name.Contains("kreis") Then score += 900
                            If category.Contains("cycle") OrElse category.Contains("zyklus") Then score += 450
                        Case "relationship"
                            If name.Contains("relationship") OrElse name.Contains("beziehung") Then score += 900
                            If category.Contains("relationship") OrElse category.Contains("beziehung") Then score += 450
                        Case "matrix"
                            If name.Contains("matrix") Then score += 900
                            If category.Contains("matrix") Then score += 450
                        Case "pyramid"
                            If name.Contains("pyramid") OrElse name.Contains("pyramide") Then score += 900
                            If category.Contains("pyramid") OrElse category.Contains("pyramide") Then score += 450
                        Case "list"
                            If name.Contains("list") OrElse name.Contains("liste") Then score += 750
                            If category.Contains("list") OrElse category.Contains("liste") Then score += 450
                        Case Else
                            If name.Contains("basic process") OrElse name.Contains("einfacher prozess") Then score += 900
                            If name.Contains("process") OrElse name.Contains("prozess") Then score += 600
                            If category.Contains("process") OrElse category.Contains("prozess") Then score += 450
                            If description.Contains("process") OrElse description.Contains("prozess") Then score += 200
                    End Select

                    If score > bestScore Then
                        If bestLayout IsNot Nothing Then
                            Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(bestLayout) : Catch ex As System.Exception : End Try
                        End If
                        bestLayout = layout
                        layout = Nothing
                        bestScore = score
                    End If
                Finally
                    If layout IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(layout) : Catch ex As System.Exception : End Try
                End Try
            Next

            If bestLayout Is Nothing OrElse bestScore <= 0 Then
                warning = "No suitable SmartArt layout was found for visual type '" & visualType & "'."
                If bestLayout IsNot Nothing Then
                    Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(bestLayout) : Catch ex As System.Exception : End Try
                    bestLayout = Nothing
                End If
            End If

            Return bestLayout
        Catch ex As System.Exception
            warning = "SmartArt layout discovery failed: " & ex.Message
            If bestLayout IsNot Nothing Then
                Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(bestLayout) : Catch releaseEx As System.Exception : End Try
            End If
            Return Nothing
        Finally
            If layouts IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(layouts) : Catch ex As System.Exception : End Try
        End Try
    End Function

    Private Shared Sub SetAutoPilotWordSmartArtNodeText(node As Microsoft.Office.Core.SmartArtNode,
                                                        label As String,
                                                        detail As String,
                                                        fontName As String,
                                                        visualType As String)
        If node Is Nothing Then Return
        Dim textRange As Microsoft.Office.Core.TextRange2 = Nothing
        Try
            textRange = node.TextFrame2.TextRange
            Dim nodeText As String = If(label, String.Empty).Trim()
            If Not String.IsNullOrWhiteSpace(detail) Then
                If Not String.IsNullOrWhiteSpace(nodeText) Then nodeText &= vbCrLf
                nodeText &= detail.Trim()
            End If
            textRange.Text = nodeText
            Try : textRange.Font.Name = fontName : Catch ex As System.Exception : End Try
            Try
                If visualType = "org_chart" OrElse visualType = "hierarchy" Then
                    textRange.Font.Size = 9.0F
                Else
                    textRange.Font.Size = 10.0F
                End If
            Catch ex As System.Exception
            End Try
        Finally
            If textRange IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(textRange) : Catch ex As System.Exception : End Try
        End Try
    End Sub

    Private Shared Function PrepareAutoPilotWordSmartArtSeed(smartArt As Microsoft.Office.Core.SmartArt,
                                                               ByRef seedNode As Microsoft.Office.Core.SmartArtNode,
                                                               ByRef warning As String) As Boolean
        seedNode = Nothing
        warning = String.Empty
        If smartArt Is Nothing Then Return False

        Dim allNodes As Microsoft.Office.Core.SmartArtNodes = Nothing
        Dim topNodes As Microsoft.Office.Core.SmartArtNodes = Nothing
        Try
            allNodes = smartArt.AllNodes
            If allNodes Is Nothing Then
                warning = "SmartArt did not expose its node collection."
                Return False
            End If

            ' SmartArt layouts normally start with one or more seed nodes. Keep exactly one
            ' native seed node and reuse it for the first requested item. Deleting every seed
            ' and then adding new roots is layout-dependent and can leave an unusable model.
            While allNodes.Count > 1
                Dim node As Microsoft.Office.Core.SmartArtNode = Nothing
                Try
                    node = allNodes.Item(allNodes.Count)
                    node.Delete()
                Catch ex As System.Exception
                    warning = "SmartArt seed normalization failed: " & ex.Message
                    Return False
                Finally
                    If node IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(node) : Catch releaseEx As System.Exception : End Try
                End Try
            End While

            If allNodes.Count = 1 Then
                seedNode = allNodes.Item(1)
                Return seedNode IsNot Nothing
            End If

            topNodes = smartArt.Nodes
            If topNodes Is Nothing Then
                warning = "SmartArt did not expose top-level nodes."
                Return False
            End If
            seedNode = topNodes.Add()
            If seedNode Is Nothing Then warning = "SmartArt could not create its seed node."
            Return seedNode IsNot Nothing
        Catch ex As System.Exception
            warning = "SmartArt seed preparation failed: " & ex.Message
            Return False
        Finally
            If topNodes IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(topNodes) : Catch ex As System.Exception : End Try
            If allNodes IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(allNodes) : Catch ex As System.Exception : End Try
        End Try
    End Function

    Private Shared Function PopulateAutoPilotWordSmartArt(smartArt As Microsoft.Office.Core.SmartArt,
                                                           visual As Newtonsoft.Json.Linq.JObject,
                                                           visualType As String,
                                                           fontName As String,
                                                           ByRef warning As String) As Boolean
        warning = String.Empty
        If smartArt Is Nothing OrElse visual Is Nothing Then Return False

        Dim topNodes As Microsoft.Office.Core.SmartArtNodes = Nothing
        Dim seedNode As Microsoft.Office.Core.SmartArtNode = Nothing
        Dim createdNodes As New Dictionary(Of String, Microsoft.Office.Core.SmartArtNode)(StringComparer.Ordinal)
        Try
            Dim seedWarning As String = String.Empty
            If Not PrepareAutoPilotWordSmartArtSeed(smartArt, seedNode, seedWarning) Then
                warning = seedWarning
                Return False
            End If

            topNodes = smartArt.Nodes
            If topNodes Is Nothing Then
                warning = "SmartArt did not expose a node collection."
                Return False
            End If

            If visualType = "org_chart" OrElse visualType = "hierarchy" Then
                Dim nodes As List(Of Newtonsoft.Json.Linq.JObject) = GetWordOrgChartNodes(visual)
                If nodes.Count = 0 Then
                    warning = "No valid org-chart nodes were supplied."
                    Return False
                End If

                Dim byId As New Dictionary(Of String, Newtonsoft.Json.Linq.JObject)(StringComparer.Ordinal)
                For Each sourceNode As Newtonsoft.Json.Linq.JObject In nodes
                    byId(GetVisualText(sourceNode, "id")) = sourceNode
                Next

                Dim pending As New List(Of Newtonsoft.Json.Linq.JObject)()
                Dim usedSeed As Boolean = False
                For Each sourceNode As Newtonsoft.Json.Linq.JObject In nodes
                    Dim id As String = GetVisualText(sourceNode, "id")
                    Dim parentId As String = GetVisualText(sourceNode, "parent_id")
                    If String.IsNullOrWhiteSpace(parentId) OrElse Not byId.ContainsKey(parentId) Then
                        Dim smartNode As Microsoft.Office.Core.SmartArtNode = Nothing
                        If Not usedSeed Then
                            smartNode = seedNode
                            seedNode = Nothing
                            usedSeed = True
                        Else
                            smartNode = topNodes.Add()
                        End If
                        If smartNode Is Nothing Then
                            warning = "SmartArt could not create a top-level organization node."
                            Return False
                        End If
                        SetAutoPilotWordSmartArtNodeText(smartNode, GetVisualText(sourceNode, "label"), GetVisualText(sourceNode, "detail"), fontName, visualType)
                        createdNodes(id) = smartNode
                    Else
                        pending.Add(sourceNode)
                    End If
                Next

                If createdNodes.Count = 0 AndAlso nodes.Count > 0 Then
                    Dim sourceNode As Newtonsoft.Json.Linq.JObject = nodes(0)
                    Dim id As String = GetVisualText(sourceNode, "id")
                    Dim smartNode As Microsoft.Office.Core.SmartArtNode = seedNode
                    seedNode = Nothing
                    If smartNode Is Nothing Then smartNode = topNodes.Add()
                    If smartNode Is Nothing Then
                        warning = "SmartArt could not create the organization root node."
                        Return False
                    End If
                    SetAutoPilotWordSmartArtNodeText(smartNode, GetVisualText(sourceNode, "label"), GetVisualText(sourceNode, "detail"), fontName, visualType)
                    createdNodes(id) = smartNode
                    pending.Remove(sourceNode)
                End If

                Dim madeProgress As Boolean = True
                While pending.Count > 0 AndAlso madeProgress
                    madeProgress = False
                    For index As Integer = pending.Count - 1 To 0 Step -1
                        Dim sourceNode As Newtonsoft.Json.Linq.JObject = pending(index)
                        Dim id As String = GetVisualText(sourceNode, "id")
                        Dim parentId As String = GetVisualText(sourceNode, "parent_id")
                        If createdNodes.ContainsKey(parentId) Then
                            Dim parentNode As Microsoft.Office.Core.SmartArtNode = createdNodes(parentId)
                            Dim smartNode As Microsoft.Office.Core.SmartArtNode = parentNode.AddNode(
                                Microsoft.Office.Core.MsoSmartArtNodePosition.msoSmartArtNodeBelow,
                                Microsoft.Office.Core.MsoSmartArtNodeType.msoSmartArtNodeTypeDefault)
                            If smartNode Is Nothing Then
                                warning = "SmartArt could not create a child organization node."
                                Return False
                            End If
                            SetAutoPilotWordSmartArtNodeText(smartNode, GetVisualText(sourceNode, "label"), GetVisualText(sourceNode, "detail"), fontName, visualType)
                            createdNodes(id) = smartNode
                            pending.RemoveAt(index)
                            madeProgress = True
                        End If
                    Next
                End While

                ' Broken or cyclic parent links must not make a requested node disappear.
                ' Preserve editability by placing unresolved nodes at the top level.
                For Each sourceNode As Newtonsoft.Json.Linq.JObject In pending
                    Dim id As String = GetVisualText(sourceNode, "id")
                    Dim smartNode As Microsoft.Office.Core.SmartArtNode = topNodes.Add()
                    If smartNode Is Nothing Then
                        warning = "SmartArt could not preserve an unresolved organization node."
                        Return False
                    End If
                    SetAutoPilotWordSmartArtNodeText(smartNode, GetVisualText(sourceNode, "label"), GetVisualText(sourceNode, "detail"), fontName, visualType)
                    createdNodes(id) = smartNode
                Next

                If createdNodes.Count <> nodes.Count Then
                    warning = "SmartArt node population was incomplete. Expected " & nodes.Count.ToString() & ", created " & createdNodes.Count.ToString() & "."
                    Return False
                End If

                ' Keep wide management levels compact. Organization Chart SmartArt supports
                ' native hanging child layouts, so use them instead of manually routing lines
                ' or shrinking/overlapping boxes. Word remains responsible for the geometry.
                Dim childCounts As New Dictionary(Of String, Integer)(StringComparer.Ordinal)
                For Each sourceNode As Newtonsoft.Json.Linq.JObject In nodes
                    Dim parentId As String = GetVisualText(sourceNode, "parent_id")
                    If Not String.IsNullOrWhiteSpace(parentId) Then
                        If Not childCounts.ContainsKey(parentId) Then childCounts(parentId) = 0
                        childCounts(parentId) += 1
                    End If
                Next
                For Each childCount As KeyValuePair(Of String, Integer) In childCounts
                    If childCount.Value >= 4 AndAlso createdNodes.ContainsKey(childCount.Key) Then
                        Try
                            createdNodes(childCount.Key).OrgChartLayout = Microsoft.Office.Core.MsoOrgChartLayoutType.msoOrgChartLayoutBothHanging
                        Catch ex As System.Exception
                            ' Some hierarchy layouts do not expose organization-chart branch layout.
                        End Try
                    End If
                Next
            Else
                Dim items As List(Of System.Tuple(Of String, String)) = GetWordVisualItems(visual)
                If items.Count = 0 Then
                    warning = "No valid SmartArt items were supplied."
                    Return False
                End If

                For index As Integer = 0 To items.Count - 1
                    Dim smartNode As Microsoft.Office.Core.SmartArtNode = Nothing
                    If index = 0 Then
                        smartNode = seedNode
                        seedNode = Nothing
                    Else
                        smartNode = topNodes.Add()
                    End If
                    If smartNode Is Nothing Then
                        warning = "SmartArt could not create item " & (index + 1).ToString() & "."
                        Return False
                    End If
                    SetAutoPilotWordSmartArtNodeText(smartNode, items(index).Item1, items(index).Item2, fontName, visualType)
                    createdNodes("item_" & index.ToString()) = smartNode
                Next
            End If

            Dim verifyNodes As Microsoft.Office.Core.SmartArtNodes = Nothing
            Try
                verifyNodes = smartArt.AllNodes
                Dim expectedCount As Integer = If(visualType = "org_chart" OrElse visualType = "hierarchy", GetWordOrgChartNodes(visual).Count, GetWordVisualItems(visual).Count)
                If verifyNodes Is Nothing OrElse verifyNodes.Count < expectedCount Then
                    warning = "SmartArt verification failed after node population."
                    Return False
                End If
            Finally
                If verifyNodes IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(verifyNodes) : Catch ex As System.Exception : End Try
            End Try

            Return True
        Catch ex As System.Exception
            warning = "SmartArt population failed: " & ex.Message
            Return False
        Finally
            If seedNode IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(seedNode) : Catch ex As System.Exception : End Try
            For Each kvp As KeyValuePair(Of String, Microsoft.Office.Core.SmartArtNode) In createdNodes
                If kvp.Value IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(kvp.Value) : Catch ex As System.Exception : End Try
            Next
            If topNodes IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(topNodes) : Catch ex As System.Exception : End Try
        End Try
    End Function

    Private Shared Function TryInsertEditableWordSmartArt(doc As Microsoft.Office.Interop.Word.Document,
                                                           visual As Newtonsoft.Json.Linq.JObject,
                                                           anchorRange As Microsoft.Office.Interop.Word.Range,
                                                           fontName As String,
                                                           visualType As String,
                                                           asInline As Boolean,
                                                           ByRef warning As String) As Boolean
        warning = String.Empty
        If doc Is Nothing OrElse visual Is Nothing OrElse anchorRange Is Nothing Then Return False

        Dim layout As Microsoft.Office.Core.SmartArtLayout = Nothing
        Dim inlineShape As Microsoft.Office.Interop.Word.InlineShape = Nothing
        Dim floatingShape As Microsoft.Office.Interop.Word.Shape = Nothing
        Dim smartArt As Microsoft.Office.Core.SmartArt = Nothing
        Try
            Dim layoutWarning As String = String.Empty
            layout = FindAutoPilotWordSmartArtLayout(doc, visual, visualType, layoutWarning)
            If layout Is Nothing Then
                warning = layoutWarning
                Return False
            End If

            Dim availableWidth As Single = Math.Max(240.0F, doc.PageSetup.PageWidth - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin)
            Dim availableHeight As Single = Math.Max(180.0F, doc.PageSetup.PageHeight - doc.PageSetup.TopMargin - doc.PageSetup.BottomMargin)
            Dim defaultWidthInches As Double = If(visualType = "org_chart" OrElse visualType = "hierarchy", 5.8R, 6.4R)
            Dim defaultHeightInches As Double = If(visualType = "org_chart" OrElse visualType = "hierarchy", 2.9R, If(visualType = "timeline", 2.2R, 2.4R))
            Dim requestedWidth As Single = CSng(GetVisualNumber(visual, "width_inches", defaultWidthInches, 3.0R, 10.5R) * 72.0R)
            Dim requestedHeight As Single = CSng(GetVisualNumber(visual, "height_inches", defaultHeightInches, 1.6R, 7.0R) * 72.0R)
            Dim maxTypeWidth As Single = If(visualType = "org_chart" OrElse visualType = "hierarchy", 430.0F, 475.0F)
            Dim displayWidth As Single = Math.Min(Math.Min(requestedWidth, availableWidth), maxTypeWidth)
            Dim displayHeight As Single = Math.Min(requestedHeight, Math.Min(availableHeight, If(visualType = "org_chart" OrElse visualType = "hierarchy", 250.0F, 250.0F)))

            If asInline Then
                Dim rangeArgument As Object = anchorRange
                inlineShape = doc.InlineShapes.AddSmartArt(layout, rangeArgument)
                If inlineShape Is Nothing Then
                    warning = "Word did not return an inline SmartArt object."
                    Return False
                End If
                Try : inlineShape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse : Catch ex As System.Exception : End Try
                Try : inlineShape.Width = displayWidth : Catch ex As System.Exception : End Try
                Try : inlineShape.Height = displayHeight : Catch ex As System.Exception : End Try
                smartArt = inlineShape.SmartArt
            Else
                Dim leftArgument As Object = 0.0F
                Dim topArgument As Object = 0.0F
                Dim widthArgument As Object = displayWidth
                Dim heightArgument As Object = displayHeight
                Dim anchorArgument As Object = anchorRange
                floatingShape = doc.Shapes.AddSmartArt(layout, leftArgument, topArgument, widthArgument, heightArgument, anchorArgument)
                If floatingShape Is Nothing Then
                    warning = "Word did not return a floating SmartArt object."
                    Return False
                End If
                floatingShape.RelativeHorizontalPosition = Microsoft.Office.Interop.Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin
                floatingShape.RelativeVerticalPosition = Microsoft.Office.Interop.Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph
                floatingShape.Left = Math.Max(0.0F, (availableWidth - displayWidth) / 2.0F)
                floatingShape.Top = 0.0F
                floatingShape.WrapFormat.Type = Microsoft.Office.Interop.Word.WdWrapType.wdWrapTopBottom
                floatingShape.LockAnchor = True
                smartArt = floatingShape.SmartArt
            End If

            If smartArt Is Nothing Then
                warning = "The inserted Word object does not expose editable SmartArt."
                Return False
            End If

            Dim populateWarning As String = String.Empty
            If Not PopulateAutoPilotWordSmartArt(smartArt, visual, visualType, fontName, populateWarning) Then
                warning = populateWarning
                Return False
            End If

            Dim altText As String = GetVisualText(visual, "title")
            If String.IsNullOrWhiteSpace(altText) Then altText = "Editable " & visualType.Replace("_", " ")
            If inlineShape IsNot Nothing Then Try : inlineShape.AlternativeText = altText : Catch ex As System.Exception : End Try
            If floatingShape IsNot Nothing Then Try : floatingShape.AlternativeText = altText : Catch ex As System.Exception : End Try
            Return True
        Catch ex As System.Exception
            warning = "Editable Word SmartArt insertion failed: " & ex.Message
            Return False
        Finally
            If Not String.IsNullOrWhiteSpace(warning) Then
                If inlineShape IsNot Nothing Then Try : inlineShape.Delete() : Catch ex As System.Exception : End Try
                If floatingShape IsNot Nothing Then Try : floatingShape.Delete() : Catch ex As System.Exception : End Try
            End If
            If smartArt IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(smartArt) : Catch ex As System.Exception : End Try
            If floatingShape IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(floatingShape) : Catch ex As System.Exception : End Try
            If inlineShape IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(inlineShape) : Catch ex As System.Exception : End Try
            If layout IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(layout) : Catch ex As System.Exception : End Try
        End Try
    End Function

    Private Shared Function GetAutoPilotWordChartType(visual As Newtonsoft.Json.Linq.JObject) As Integer
        Dim visualType As String = GetVisualText(visual, "type", "bar_chart").ToLowerInvariant()
        If visualType = "line_chart" Then Return 4 ' xlLine
        Return 51 ' xlColumnClustered
    End Function

    Private Shared Sub GetAutoPilotWordChartSize(doc As Microsoft.Office.Interop.Word.Document,
                                                  visual As Newtonsoft.Json.Linq.JObject,
                                                  ByRef displayWidth As Single,
                                                  ByRef displayHeight As Single)
        Dim requestedWidth As Single = CSng(GetVisualNumber(visual, "width_inches", 7.0R, 3.0R, 10.5R) * 72.0R)
        Dim requestedHeight As Single = CSng(GetVisualNumber(visual, "height_inches", 3.8R, 2.0R, 7.0R) * 72.0R)
        Dim availableWidth As Single = Math.Max(240.0F, doc.PageSetup.PageWidth - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin)
        Dim availableHeight As Single = Math.Max(180.0F, doc.PageSetup.PageHeight - doc.PageSetup.TopMargin - doc.PageSetup.BottomMargin)
        displayWidth = Math.Min(requestedWidth, availableWidth)
        displayHeight = Math.Min(requestedHeight, availableHeight)
        If requestedWidth > 0.0F AndAlso requestedHeight > 0.0F Then
            Dim aspectRatio As Single = requestedHeight / requestedWidth
            displayHeight = displayWidth * aspectRatio
            If displayHeight > availableHeight Then
                displayHeight = availableHeight
                displayWidth = displayHeight / Math.Max(0.01F, aspectRatio)
            End If
        End If
    End Sub

    Private Shared Function PopulateAutoPilotWordChart(chart As Object,
                                                        visual As Newtonsoft.Json.Linq.JObject,
                                                        fontName As String,
                                                        accentHex As String,
                                                        ByRef warning As String) As Boolean
        warning = String.Empty
        If chart Is Nothing OrElse visual Is Nothing Then Return False

        Dim categories As List(Of String) = GetWordVisualCategories(visual)
        Dim series As List(Of System.Tuple(Of String, List(Of Double))) = GetWordVisualSeries(visual)
        If categories.Count = 0 OrElse series.Count = 0 Then
            warning = "No valid chart categories/series were supplied."
            Return False
        End If

        Dim maxValues As Integer = 0
        For Each item As System.Tuple(Of String, List(Of Double)) In series
            If item IsNot Nothing AndAlso item.Item2 IsNot Nothing Then maxValues = Math.Max(maxValues, item.Item2.Count)
        Next
        Dim categoryCount As Integer = Math.Min(categories.Count, maxValues)
        If categoryCount <= 0 Then
            warning = "The chart does not contain matching category/value data."
            Return False
        End If

        Dim chartType As Integer = GetAutoPilotWordChartType(visual)
        Dim visualType As String = GetVisualText(visual, "type", "bar_chart").ToLowerInvariant()
        Dim chartData As Object = Nothing
        Dim workbook As Object = Nothing
        Dim worksheet As Object = Nothing
        Dim dataRange As Object = Nothing
        Dim accent As System.Drawing.Color = WordVisualColor(accentHex, "#17365D")
        Dim accentOle As Integer = System.Drawing.ColorTranslator.ToOle(accent)

        Try
            Dim chartDataWarning As String = String.Empty
            Dim populatedFromWorkbook As Boolean = False
            Try
                chartData = chart.ChartData
                chartData.Activate()
                workbook = chartData.Workbook
                If workbook Is Nothing Then Throw New System.InvalidOperationException("Word ChartData did not expose an embedded workbook.")
                Try : workbook.Application.Visible = False : Catch ex As System.Exception : End Try
                worksheet = workbook.Worksheets(1)
                If worksheet Is Nothing Then Throw New System.InvalidOperationException("Word ChartData did not expose a worksheet.")

                Try : worksheet.Cells.Clear() : Catch ex As System.Exception : End Try
                worksheet.Cells(1, 1).Value2 = String.Empty
                For categoryIndex As Integer = 0 To categoryCount - 1
                    worksheet.Cells(1, categoryIndex + 2).Value2 = categories(categoryIndex)
                Next

                For seriesIndex As Integer = 0 To series.Count - 1
                    Dim seriesItem As System.Tuple(Of String, List(Of Double)) = series(seriesIndex)
                    worksheet.Cells(seriesIndex + 2, 1).Value2 = If(String.IsNullOrWhiteSpace(seriesItem.Item1), "Series " & (seriesIndex + 1).ToString(), seriesItem.Item1)
                    For categoryIndex As Integer = 0 To categoryCount - 1
                        Dim value As Double = 0.0R
                        If seriesItem.Item2 IsNot Nothing AndAlso categoryIndex < seriesItem.Item2.Count Then value = seriesItem.Item2(categoryIndex)
                        worksheet.Cells(seriesIndex + 2, categoryIndex + 2).Value2 = value
                    Next
                Next

                dataRange = worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(series.Count + 1, categoryCount + 1))
                If dataRange Is Nothing Then Throw New System.InvalidOperationException("Word ChartData did not expose the populated data range.")
                Dim sourceAddress As String = CStr(dataRange.Address(True, True, 1, True))
                chart.SetSourceData(sourceAddress, 1)
                populatedFromWorkbook = True
            Catch chartDataEx As System.Exception
                chartDataWarning = "Embedded chart-data workbook unavailable (" & chartDataEx.Message & "); native series values were used instead."

                Dim directSeriesCollection As Object = Nothing
                Try
                    directSeriesCollection = chart.SeriesCollection()
                    If directSeriesCollection Is Nothing Then Throw New System.InvalidOperationException("Word chart did not expose a SeriesCollection.")

                    Do While CInt(directSeriesCollection.Count) > 0
                        Dim staleSeries As Object = Nothing
                        Try
                            staleSeries = directSeriesCollection.Item(1)
                            staleSeries.Delete()
                        Finally
                            If staleSeries IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(staleSeries) : Catch ex As System.Exception : End Try
                        End Try
                    Loop

                    Dim categoryValues(categoryCount - 1) As String
                    For categoryIndex As Integer = 0 To categoryCount - 1
                        categoryValues(categoryIndex) = categories(categoryIndex)
                    Next

                    For seriesIndex As Integer = 0 To series.Count - 1
                        Dim seriesItem As System.Tuple(Of String, List(Of Double)) = series(seriesIndex)
                        Dim valueValues(categoryCount - 1) As Double
                        For categoryIndex As Integer = 0 To categoryCount - 1
                            If seriesItem.Item2 IsNot Nothing AndAlso categoryIndex < seriesItem.Item2.Count Then
                                valueValues(categoryIndex) = seriesItem.Item2(categoryIndex)
                            Else
                                valueValues(categoryIndex) = 0.0R
                            End If
                        Next

                        Dim nativeSeries As Object = Nothing
                        Try
                            nativeSeries = directSeriesCollection.NewSeries()
                            nativeSeries.Name = If(String.IsNullOrWhiteSpace(seriesItem.Item1), "Series " & (seriesIndex + 1).ToString(), seriesItem.Item1)
                            nativeSeries.XValues = categoryValues
                            nativeSeries.Values = valueValues
                        Finally
                            If nativeSeries IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(nativeSeries) : Catch ex As System.Exception : End Try
                        End Try
                    Next
                Finally
                    If directSeriesCollection IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(directSeriesCollection) : Catch ex As System.Exception : End Try
                End Try
            End Try

            chart.ChartType = chartType
            If Not populatedFromWorkbook AndAlso Not String.IsNullOrWhiteSpace(chartDataWarning) Then warning = chartDataWarning

            Dim title As String = GetVisualText(visual, "title")
            If Not String.IsNullOrWhiteSpace(title) Then
                chart.HasTitle = True
                chart.ChartTitle.Text = title
                Try : chart.ChartTitle.Format.TextFrame2.TextRange.Font.Name = fontName : Catch ex As System.Exception : End Try
            Else
                Try : chart.HasTitle = False : Catch ex As System.Exception : End Try
            End If

            Dim showLegend As Boolean = series.Count > 1
            Dim legendToken As Newtonsoft.Json.Linq.JToken = visual("show_legend")
            If legendToken IsNot Nothing AndAlso legendToken.Type = Newtonsoft.Json.Linq.JTokenType.Boolean Then showLegend = CBool(legendToken)
            Try : chart.HasLegend = showLegend : Catch ex As System.Exception : End Try
            If showLegend Then Try : chart.Legend.Position = -4107 : Catch ex As System.Exception : End Try

            Try
                Dim seriesCollection As Object = chart.SeriesCollection()
                For seriesIndex As Integer = 1 To series.Count
                    Dim nativeSeries As Object = Nothing
                    Try
                        nativeSeries = seriesCollection.Item(seriesIndex)
                        nativeSeries.Format.Fill.ForeColor.RGB = accentOle
                        nativeSeries.Format.Line.ForeColor.RGB = accentOle
                        If visualType = "line_chart" Then
                            nativeSeries.Format.Line.Weight = 2.25F
                            Try : nativeSeries.MarkerStyle = 8 : Catch ex As System.Exception : End Try
                            Try : nativeSeries.MarkerSize = 6 : Catch ex As System.Exception : End Try
                        End If
                    Finally
                        If nativeSeries IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(nativeSeries) : Catch ex As System.Exception : End Try
                    End Try
                Next
                If seriesCollection IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(seriesCollection) : Catch ex As System.Exception : End Try
            Catch formatEx As System.Exception
                Debug.WriteLine($"Word native chart formatting warning: {formatEx.Message}")
            End Try

            Try : chart.ChartArea.Format.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse : Catch ex As System.Exception : End Try
            Try : chart.PlotArea.Format.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse : Catch ex As System.Exception : End Try
            Try : chart.Refresh() : Catch ex As System.Exception : End Try
            Try : workbook.Close(True) : Catch ex As System.Exception : End Try
            Return True
        Catch ex As System.Exception
            warning = "Editable Word chart population failed: " & ex.Message
            Return False
        Finally
            If dataRange IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(dataRange) : Catch ex As System.Exception : End Try
            If worksheet IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(worksheet) : Catch ex As System.Exception : End Try
            If workbook IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(workbook) : Catch ex As System.Exception : End Try
            If chartData IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(chartData) : Catch ex As System.Exception : End Try
        End Try
    End Function

    Private Shared Function TryInsertEditableWordChart(doc As Microsoft.Office.Interop.Word.Document,
                                                        visual As Newtonsoft.Json.Linq.JObject,
                                                        anchorRange As Microsoft.Office.Interop.Word.Range,
                                                        fontName As String,
                                                        accentHex As String,
                                                        ByRef warning As String) As Boolean
        warning = String.Empty
        If doc Is Nothing OrElse visual Is Nothing OrElse anchorRange Is Nothing Then Return False

        Dim inlineShapesObject As Object = Nothing
        Dim chartShape As Object = Nothing
        Dim chart As Object = Nothing
        Dim succeeded As Boolean = False
        Try
            Dim chartType As Integer = GetAutoPilotWordChartType(visual)
            Dim displayWidth As Single = 0.0F
            Dim displayHeight As Single = 0.0F
            GetAutoPilotWordChartSize(doc, visual, displayWidth, displayHeight)

            inlineShapesObject = doc.InlineShapes
            Try
                chartShape = inlineShapesObject.AddChart2(-1, chartType, anchorRange, True)
            Catch addChart2Ex As System.Exception
                chartShape = inlineShapesObject.AddChart(chartType, anchorRange)
            End Try
            If chartShape Is Nothing Then
                warning = "Word did not return an editable inline chart object."
                Return False
            End If

            Try : chartShape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse : Catch ex As System.Exception : End Try
            Try : chartShape.Width = displayWidth : Catch ex As System.Exception : End Try
            Try : chartShape.Height = displayHeight : Catch ex As System.Exception : End Try
            chart = chartShape.Chart
            If chart Is Nothing Then
                warning = "The inserted inline Word object does not expose an editable chart."
                Return False
            End If

            Dim populateWarning As String = String.Empty
            If Not PopulateAutoPilotWordChart(chart, visual, fontName, accentHex, populateWarning) Then
                warning = populateWarning
                Return False
            End If
            If Not String.IsNullOrWhiteSpace(populateWarning) Then warning = populateWarning
            succeeded = True
            Return True
        Catch ex As System.Exception
            warning = "Editable inline Word chart insertion failed: " & ex.Message
            Return False
        Finally
            If Not succeeded AndAlso chartShape IsNot Nothing Then Try : chartShape.Delete() : Catch ex As System.Exception : End Try
            If chart IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(chart) : Catch ex As System.Exception : End Try
            If chartShape IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(chartShape) : Catch ex As System.Exception : End Try
            If inlineShapesObject IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(inlineShapesObject) : Catch ex As System.Exception : End Try
        End Try
    End Function

    Private Shared Function TryInsertEditableWordFloatingChart(doc As Microsoft.Office.Interop.Word.Document,
                                                                visual As Newtonsoft.Json.Linq.JObject,
                                                                anchorRange As Microsoft.Office.Interop.Word.Range,
                                                                fontName As String,
                                                                accentHex As String,
                                                                ByRef warning As String) As Boolean
        warning = String.Empty
        If doc Is Nothing OrElse visual Is Nothing OrElse anchorRange Is Nothing Then Return False

        Dim shapesObject As Object = Nothing
        Dim chartShape As Object = Nothing
        Dim chart As Object = Nothing
        Dim succeeded As Boolean = False
        Try
            Dim chartType As Integer = GetAutoPilotWordChartType(visual)
            Dim displayWidth As Single = 0.0F
            Dim displayHeight As Single = 0.0F
            GetAutoPilotWordChartSize(doc, visual, displayWidth, displayHeight)
            Dim availableWidth As Single = Math.Max(240.0F, doc.PageSetup.PageWidth - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin)

            shapesObject = doc.Shapes
            chartShape = shapesObject.AddChart2(-1, chartType, 0.0F, 0.0F, displayWidth, displayHeight, anchorRange, True)
            If chartShape Is Nothing Then
                warning = "Word did not return an editable floating chart object."
                Return False
            End If

            Try : chartShape.RelativeHorizontalPosition = Microsoft.Office.Interop.Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin : Catch ex As System.Exception : End Try
            Try : chartShape.RelativeVerticalPosition = Microsoft.Office.Interop.Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph : Catch ex As System.Exception : End Try
            Try : chartShape.Left = Math.Max(0.0F, (availableWidth - displayWidth) / 2.0F) : Catch ex As System.Exception : End Try
            Try : chartShape.Top = 0.0F : Catch ex As System.Exception : End Try
            Try : chartShape.WrapFormat.Type = Microsoft.Office.Interop.Word.WdWrapType.wdWrapTopBottom : Catch ex As System.Exception : End Try
            Try : chartShape.LockAnchor = True : Catch ex As System.Exception : End Try

            chart = chartShape.Chart
            If chart Is Nothing Then
                warning = "The inserted floating Word object does not expose an editable chart."
                Return False
            End If

            Dim populateWarning As String = String.Empty
            If Not PopulateAutoPilotWordChart(chart, visual, fontName, accentHex, populateWarning) Then
                warning = populateWarning
                Return False
            End If
            If Not String.IsNullOrWhiteSpace(populateWarning) Then warning = populateWarning
            succeeded = True
            Return True
        Catch ex As System.Exception
            warning = "Editable floating Word chart insertion failed: " & ex.Message
            Return False
        Finally
            If Not succeeded AndAlso chartShape IsNot Nothing Then Try : chartShape.Delete() : Catch ex As System.Exception : End Try
            If chart IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(chart) : Catch ex As System.Exception : End Try
            If chartShape IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(chartShape) : Catch ex As System.Exception : End Try
            If shapesObject IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(shapesObject) : Catch ex As System.Exception : End Try
        End Try
    End Function

    Private Shared Function TryInsertEditableAutoPilotWordVisual(doc As Microsoft.Office.Interop.Word.Document,
                                                                  visual As Newtonsoft.Json.Linq.JObject,
                                                                  insertionPosition As Integer,
                                                                  fontName As String,
                                                                  accentHex As String,
                                                                  ByRef modeUsed As String,
                                                                  ByRef warning As String) As Boolean
        modeUsed = String.Empty
        warning = String.Empty
        If doc Is Nothing OrElse visual Is Nothing Then Return False

        Dim visualType As System.String = GetVisualText(visual, "type", "process").ToLowerInvariant()
        Dim requestedMode As String = GetAutoPilotWordVisualInsertionMode(visual)
        Dim firstWarning As String = String.Empty
        Dim secondWarning As String = String.Empty

        Select Case visualType
            Case "org_chart", "hierarchy", "process", "timeline", "cycle", "relationship", "matrix", "pyramid", "list", "smartart"
                If requestedMode = "auto" OrElse requestedMode = "inline" Then
                    Dim inlineRange As Microsoft.Office.Interop.Word.Range = Nothing
                    Try
                        inlineRange = doc.Range(insertionPosition, insertionPosition)
                        If TryInsertEditableWordSmartArt(doc, visual, inlineRange, fontName, visualType, True, firstWarning) Then
                            modeUsed = "inline_smartart"
                            warning = firstWarning
                            Return True
                        End If
                    Finally
                        If inlineRange IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(inlineRange) : Catch ex As System.Exception : End Try
                    End Try
                    If requestedMode = "inline" Then
                        warning = firstWarning
                        Return False
                    End If
                End If

                If requestedMode = "auto" OrElse requestedMode = "floating" Then
                    Dim floatingRange As Microsoft.Office.Interop.Word.Range = Nothing
                    Try
                        floatingRange = doc.Range(insertionPosition, insertionPosition)
                        If TryInsertEditableWordSmartArt(doc, visual, floatingRange, fontName, visualType, False, secondWarning) Then
                            modeUsed = "floating_smartart"
                            warning = secondWarning
                            Return True
                        End If
                    Finally
                        If floatingRange IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(floatingRange) : Catch ex As System.Exception : End Try
                    End Try
                End If

            Case "bar_chart", "line_chart"
                If requestedMode = "auto" OrElse requestedMode = "inline" Then
                    Dim inlineRange As Microsoft.Office.Interop.Word.Range = Nothing
                    Try
                        inlineRange = doc.Range(insertionPosition, insertionPosition)
                        If TryInsertEditableWordChart(doc, visual, inlineRange, fontName, accentHex, firstWarning) Then
                            modeUsed = "inline_chart"
                            warning = firstWarning
                            Return True
                        End If
                    Finally
                        If inlineRange IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(inlineRange) : Catch ex As System.Exception : End Try
                    End Try
                    If requestedMode = "inline" Then
                        warning = firstWarning
                        Return False
                    End If
                End If

                If requestedMode = "auto" OrElse requestedMode = "floating" Then
                    Dim floatingRange As Microsoft.Office.Interop.Word.Range = Nothing
                    Try
                        floatingRange = doc.Range(insertionPosition, insertionPosition)
                        If TryInsertEditableWordFloatingChart(doc, visual, floatingRange, fontName, accentHex, secondWarning) Then
                            modeUsed = "floating_chart"
                            warning = secondWarning
                            Return True
                        End If
                    Finally
                        If floatingRange IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(floatingRange) : Catch ex As System.Exception : End Try
                    End Try
                End If

            Case Else
                warning = "Unsupported editable Word visual type '" & visualType & "'."
                Return False
        End Select

        Dim warningParts As New List(Of String)()
        If Not String.IsNullOrWhiteSpace(firstWarning) Then warningParts.Add(firstWarning)
        If Not String.IsNullOrWhiteSpace(secondWarning) Then warningParts.Add(secondWarning)
        warning = String.Join(" | ", warningParts)
        If String.IsNullOrWhiteSpace(warning) Then warning = "No editable Word insertion mode succeeded for visual type '" & visualType & "'."
        Return False
    End Function

    Private Shared Function InsertAutoPilotWordVisuals(doc As Microsoft.Office.Interop.Word.Document,
                                                       args As Dictionary(Of String, Object),
                                                       fontName As String,
                                                       accentHex As String,
                                                       tempDirectory As String,
                                                       ByRef embeddedCount As Integer,
                                                       ByRef warnings As List(Of String)) As Boolean
        embeddedCount = 0
        warnings = New List(Of String)()
        Dim visuals As Newtonsoft.Json.Linq.JArray = GetAutoPilotWordVisuals(args)
        If visuals.Count = 0 Then Return True

        For Each token As Newtonsoft.Json.Linq.JToken In visuals
            If token Is Nothing OrElse token.Type <> Newtonsoft.Json.Linq.JTokenType.Object Then
                warnings.Add("Ignored a visual entry because it was not an object.")
                Continue For
            End If

            Dim visual As Newtonsoft.Json.Linq.JObject = DirectCast(token, Newtonsoft.Json.Linq.JObject)
            Dim id As System.String = GetVisualText(visual, "id")
            If String.IsNullOrWhiteSpace(id) OrElse Not System.Text.RegularExpressions.Regex.IsMatch(id, "^[A-Za-z0-9_.-]{1,64}$") Then
                warnings.Add("Ignored a visual with a missing or invalid id.")
                Continue For
            End If

            Dim placeholder As System.String = "[[visual:" & id & "]]"
            Dim editable As Boolean = GetAutoPilotWordVisualEditable(visual)

            If editable Then
                ' Editable visuals use native Word objects only. Remove the placeholder BEFORE
                ' insertion and keep its original position. This avoids mutable Word Range
                ' expansion deleting a newly inserted object. If every editable insertion mode
                ' fails, restore the exact placeholder and fail the visual pass; do not silently
                ' replace the requested editable object with a PNG.
                Dim target As Microsoft.Office.Interop.Word.Range = Nothing
                Dim finder As Microsoft.Office.Interop.Word.Find = Nothing
                Try
                    target = doc.Content.Duplicate
                    finder = target.Find
                    finder.ClearFormatting()
                    finder.Text = placeholder
                    finder.Forward = True
                    finder.Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                    If Not finder.Execute() Then
                        warnings.Add("Visual placeholder '" & placeholder & "' was not found; the editable visual was not inserted.")
                        Continue For
                    End If

                    Dim insertionPosition As Integer = target.Start
                    target.Text = String.Empty
                    Dim modeUsed As String = String.Empty
                    Dim nativeWarning As String = String.Empty
                    If TryInsertEditableAutoPilotWordVisual(doc, visual, insertionPosition, fontName, accentHex, modeUsed, nativeWarning) Then
                        embeddedCount += 1
                        If Not String.IsNullOrWhiteSpace(nativeWarning) Then
                            warnings.Add("Visual '" & id & "' was inserted as " & modeUsed & " with warning: " & nativeWarning)
                        End If
                        Continue For
                    End If

                    Dim restoreRange As Microsoft.Office.Interop.Word.Range = Nothing
                    Try
                        restoreRange = doc.Range(insertionPosition, insertionPosition)
                        restoreRange.Text = placeholder
                    Finally
                        If restoreRange IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(restoreRange) : Catch ex As System.Exception : End Try
                    End Try

                    warnings.Add("Visual '" & id & "' could not be inserted as an editable native Word object. No PNG fallback was used. " & nativeWarning)
                Finally
                    If finder IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(finder) : Catch ex As System.Exception : End Try
                    If target IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(target) : Catch ex As System.Exception : End Try
                End Try

                Continue For
            End If

            ' Raster output is now opt-in only through editable=false. It is retained for
            ' callers that explicitly want a flattened visual, never as an implicit fallback.
            Dim imagePath As String = Path.Combine(tempDirectory, ".word_visual_" & Guid.NewGuid().ToString("N") & ".png")
            Try
                If Not RenderAutoPilotWordVisual(visual, imagePath, fontName, accentHex) Then
                    warnings.Add("Could not render non-editable visual '" & id & "'.")
                    Continue For
                End If

                Dim target As Microsoft.Office.Interop.Word.Range = Nothing
                Dim finder As Microsoft.Office.Interop.Word.Find = Nothing
                Dim inlineShape As Microsoft.Office.Interop.Word.InlineShape = Nothing
                Try
                    target = doc.Content.Duplicate
                    finder = target.Find
                    finder.ClearFormatting()
                    finder.Text = placeholder
                    finder.Forward = True
                    finder.Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                    If Not finder.Execute() Then
                        warnings.Add("Visual placeholder '" & placeholder & "' was not found; the visual was not inserted.")
                        Continue For
                    End If

                    Dim insertionPosition As Integer = target.Start
                    target.Text = String.Empty
                    Dim pictureRange As Microsoft.Office.Interop.Word.Range = Nothing
                    Try
                        pictureRange = doc.Range(insertionPosition, insertionPosition)
                        inlineShape = doc.InlineShapes.AddPicture(imagePath, False, True, pictureRange)
                    Finally
                        If pictureRange IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pictureRange) : Catch ex As System.Exception : End Try
                    End Try

                    Dim requestedWidth As Single = CSng(GetVisualNumber(visual, "width_inches", 8.4R, 4.0R, 10.5R) * 72.0R)
                    Dim requestedHeight As Single = CSng(GetVisualNumber(visual, "height_inches", 4.7R, 2.5R, 7.0R) * 72.0R)
                    Try
                        Dim availableWidth As Single = doc.PageSetup.PageWidth - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin
                        Dim availableHeight As Single = doc.PageSetup.PageHeight - doc.PageSetup.TopMargin - doc.PageSetup.BottomMargin
                        Dim displayWidth As Single = Math.Min(requestedWidth, Math.Max(72.0F, availableWidth))
                        Dim aspectRatio As Single = If(requestedWidth > 0.0F, requestedHeight / requestedWidth, 1.0F)
                        Dim displayHeight As Single = displayWidth * aspectRatio
                        If displayHeight > availableHeight AndAlso availableHeight > 72.0F Then
                            displayHeight = availableHeight
                            displayWidth = displayHeight / Math.Max(0.01F, aspectRatio)
                        End If
                        inlineShape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue
                        inlineShape.Width = displayWidth
                    Catch ex As System.Exception
                    End Try

                    Dim altText As String = GetVisualText(visual, "title")
                    If String.IsNullOrWhiteSpace(altText) Then altText = "Document visual " & id
                    Try : inlineShape.AlternativeText = altText : Catch ex As System.Exception : End Try
                    embeddedCount += 1
                Catch ex As System.Exception
                    warnings.Add("Non-editable visual '" & id & "' could not be inserted: " & ex.Message)
                Finally
                    If inlineShape IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(inlineShape) : Catch ex As System.Exception : End Try
                    If finder IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(finder) : Catch ex As System.Exception : End Try
                    If target IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(target) : Catch ex As System.Exception : End Try
                End Try
            Finally
                Try
                    If File.Exists(imagePath) Then File.Delete(imagePath)
                Catch ex As System.Exception
                End Try
            End Try
        Next

        Return embeddedCount = visuals.Count
    End Function

    Private Shared Function ValidateSavedAutoPilotWordVisualPersistence(outputPath As String,
                                                                         visuals As Newtonsoft.Json.Linq.JArray,
                                                                         ByRef validationError As System.String) As Boolean
        validationError = String.Empty
        If visuals Is Nothing OrElse visuals.Count = 0 Then Return True
        If String.IsNullOrWhiteSpace(outputPath) OrElse Not System.IO.File.Exists(outputPath) Then
            validationError = "The saved Word file is unavailable for visual persistence validation."
            Return False
        End If

        Dim expectedCharts As Integer = 0
        Dim expectedSmartArt As Integer = 0
        For Each token As Newtonsoft.Json.Linq.JToken In visuals
            If token Is Nothing OrElse token.Type <> Newtonsoft.Json.Linq.JTokenType.Object Then Continue For
            Dim visual As Newtonsoft.Json.Linq.JObject = DirectCast(token, Newtonsoft.Json.Linq.JObject)
            If Not GetAutoPilotWordVisualEditable(visual) Then Continue For
            Select Case GetVisualText(visual, "type", "process").ToLowerInvariant()
                Case "bar_chart", "line_chart"
                    expectedCharts += 1
                Case "org_chart", "hierarchy", "process", "timeline", "cycle", "relationship", "matrix", "pyramid", "list", "smartart"
                    expectedSmartArt += 1
            End Select
        Next

        If expectedCharts = 0 AndAlso expectedSmartArt = 0 Then Return True

        Try
            Using archive As System.IO.Compression.ZipArchive = System.IO.Compression.ZipFile.OpenRead(outputPath)
                Dim persistedCharts As Integer = 0
                Dim persistedSmartArt As Integer = 0
                For Each entry As System.IO.Compression.ZipArchiveEntry In archive.Entries
                    Dim normalizedName As String = entry.FullName.Replace("\\", "/")
                    If System.Text.RegularExpressions.Regex.IsMatch(normalizedName, "^word/charts/chart[0-9]+\\.xml$", System.Text.RegularExpressions.RegexOptions.IgnoreCase) Then
                        persistedCharts += 1
                    ElseIf System.Text.RegularExpressions.Regex.IsMatch(normalizedName, "^word/diagrams/data[0-9]+\\.xml$", System.Text.RegularExpressions.RegexOptions.IgnoreCase) Then
                        persistedSmartArt += 1
                    End If
                Next

                If persistedCharts < expectedCharts OrElse persistedSmartArt < expectedSmartArt Then
                    validationError = "Saved DOCX visual persistence check failed. Expected editable charts=" & expectedCharts.ToString() &
                                      ", persisted=" & persistedCharts.ToString() &
                                      "; expected SmartArt visuals=" & expectedSmartArt.ToString() &
                                      ", persisted=" & persistedSmartArt.ToString() & "."
                    Return False
                End If
            End Using
            Return True
        Catch ex As System.Exception
            validationError = "Saved DOCX visual persistence validation failed: " & ex.Message
            Return False
        End Try
    End Function

    Private Async Function ExecuteCreateWordDocTool(
            toolCall As ToolCall, context As ToolExecutionContext, ct As CancellationToken) As System.Threading.Tasks.Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim markdownContent = GetArgString(toolCall.Arguments, "markdown_content")
            If String.IsNullOrWhiteSpace(markdownContent) Then
                response.Success = False
                response.Response = "Missing required parameter: markdown_content"
                Return response
            End If

            Dim design As AutoPilotDesignResolution = ResolveAutoPilotDocumentDesign(
                toolCall.Arguments,
                "Word",
                New String() {"base_font_name", "base_font_size", "page_orientation", "professional_layout", "style_preset", "accent_color", "secondary_color", "text_color", "muted_color", "light_color", "line_color", "table_style_name", "header_text", "footer_text", "show_page_numbers", "use_template_styles"},
                New String() {".dotx", ".dotm", ".docx"},
                context)

            Dim visuals As Newtonsoft.Json.Linq.JArray = GetAutoPilotWordVisuals(toolCall.Arguments)
            Dim visualContractError As System.String = System.String.Empty
            If Not ValidateAutoPilotWordVisualContract(markdownContent, visuals, context, visualContractError) Then
                response.Success = False
                response.ErrorMessage = visualContractError
                response.Response = response.ErrorMessage
                Return response
            End If
            If ContainsLikelyWordPseudoGraphic(markdownContent) Then
                response.Success = False
                response.ErrorMessage = "Diagram-like ASCII/Unicode/Mermaid/block-character content is not allowed in create_word_document. Requested graphics must use native editable visuals with one exact [[visual:ID]] placeholder each. For an organization chart use type='org_chart' with nodes [{id,label,detail,parent_id}]; for quantitative graphics use bar_chart/column_chart/line_chart/area_chart/pie_chart/doughnut_chart."
                response.Response = response.ErrorMessage
                Return response
            End If

            Dim fileName = GetArgString(toolCall.Arguments, "file_name")
            If String.IsNullOrWhiteSpace(fileName) Then fileName = "Document"

            For Each c In Path.GetInvalidFileNameChars()
                fileName = fileName.Replace(c, "_"c)
            Next
            If Not fileName.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) Then
                fileName &= ".docx"
            End If

            Dim outputPath = Path.Combine(_apCurrentTempDir, fileName)

            Dim counter = 1
            While File.Exists(outputPath)
                Dim baseName = Path.GetFileNameWithoutExtension(fileName)
                fileName = baseName & $"_{counter}.docx"
                outputPath = Path.Combine(_apCurrentTempDir, fileName)
                counter += 1
            End While

            context.Log($"Creating Word document: {fileName}")
            ApDashboardLog($"📝 Creating Word document: {fileName}", "step")

            Dim embeddedVisualCount As Integer = 0
            Dim visualWarnings As New List(Of String)()
            Dim creationError As String = String.Empty

            Dim success = Await SwitchToUi(Function()
                                               Dim wordApp As Microsoft.Office.Interop.Word.Application = Nothing
                                               Dim doc As Microsoft.Office.Interop.Word.Document = Nothing
                                               Dim weCreated As Boolean = False
                                               Dim sel As Microsoft.Office.Interop.Word.Selection = Nothing

                                               Try
                                                   Try
                                                       wordApp = DirectCast(GetObject(, "Word.Application"), Microsoft.Office.Interop.Word.Application)
                                                   Catch
                                                       wordApp = New Microsoft.Office.Interop.Word.Application()
                                                       wordApp.Visible = False
                                                       weCreated = True
                                                   End Try

                                                   wordApp.ScreenUpdating = False
                                                   If design IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(design.TemplatePath) Then
                                                       Dim designExt As String = System.IO.Path.GetExtension(design.TemplatePath).ToLowerInvariant()
                                                       If designExt = ".dotx" OrElse designExt = ".dotm" Then
                                                           doc = wordApp.Documents.Add(Template:=design.TemplatePath, NewTemplate:=False)
                                                       Else
                                                           ' A .docx design source is cloned before use and its body is cleared.
                                                           ' Styles, theme, sections, headers and footers remain available, while
                                                           ' sample document content cannot leak into the generated deliverable.
                                                           System.IO.File.Copy(design.TemplatePath, outputPath, overwrite:=False)
                                                           doc = wordApp.Documents.Open(outputPath, ReadOnly:=False, AddToRecentFiles:=False, Visible:=False)
                                                           Try
                                                               Dim bodyRange As Microsoft.Office.Interop.Word.Range = doc.Content
                                                               Try
                                                                   If bodyRange.End > bodyRange.Start Then bodyRange.Text = ""
                                                               Finally
                                                                   Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(bodyRange) : Catch ex As System.Exception : End Try
                                                               End Try
                                                           Catch ex As System.Exception
                                                           End Try
                                                       End If
                                                   Else
                                                       doc = wordApp.Documents.Add()
                                                   End If
                                                   doc.Activate()

                                                   sel = wordApp.Selection

                                                   Dim includeCover As Boolean = GetArgBool(toolCall.Arguments, "include_cover", False)
                                                   Dim coverTitle As String = GetArgString(toolCall.Arguments, "cover_title")
                                                   If includeCover AndAlso String.IsNullOrWhiteSpace(coverTitle) Then coverTitle = GetArgString(toolCall.Arguments, "document_title")
                                                   If Not String.IsNullOrWhiteSpace(coverTitle) Then
                                                       Dim coverSubtitle As String = GetArgString(toolCall.Arguments, "cover_subtitle")
                                                       Dim coverKicker As String = GetArgString(toolCall.Arguments, "cover_kicker")
                                                       Dim coverAccent As Integer = PptHexColor(GetArgString(toolCall.Arguments, "accent_color"), "#17365D")
                                                       Dim coverText As Integer = PptHexColor(GetArgString(toolCall.Arguments, "text_color"), "#202124")
                                                       Dim coverMuted As Integer = PptHexColor(GetArgString(toolCall.Arguments, "muted_color"), "#667085")
                                                       Dim coverFont As String = GetArgString(toolCall.Arguments, "base_font_name")
                                                       If String.IsNullOrWhiteSpace(coverFont) Then coverFont = "Aptos"
                                                       InsertAutoPilotWordCoverPage(sel, coverTitle, coverSubtitle, coverKicker, coverAccent, coverText, coverMuted, coverFont)
                                                   End If

                                                   SharedMethods.InsertTextWithMarkdown(sel, markdownContent, TrailingCR:=False)

                                                   Dim baseFontName As String = GetArgString(toolCall.Arguments, "base_font_name")
                                                   If String.IsNullOrWhiteSpace(baseFontName) Then baseFontName = "Aptos"
                                                   Dim accentHex As String = GetArgString(toolCall.Arguments, "accent_color")
                                                   If String.IsNullOrWhiteSpace(accentHex) Then accentHex = "#17365D"

                                                   ' Establish final page geometry and styles before inserting graphics so
                                                   ' every visual can be constrained to the actual printable page area.
                                                   ApplyAutoPilotWordDocumentStyling(doc, toolCall.Arguments)

                                                   ' V8: Visual objects are deliberately NOT created through Word/Excel COM.
                                                   ' The formatted DOCX is saved first with exact [[visual:ID]] marker paragraphs.
                                                   ' After Word is closed, InsertAutoPilotWordVisualsOpenXml writes native
                                                   ' chart parts / embedded workbooks / editable DrawingML directly into
                                                   ' the package. This removes Range/anchor/ChartData COM failure modes.

                                                   Dim currentDocPath As String = ""
                                                   Try : currentDocPath = doc.FullName : Catch ex As System.Exception : End Try
                                                   If Not String.IsNullOrWhiteSpace(currentDocPath) AndAlso
                                                      String.Equals(System.IO.Path.GetFullPath(currentDocPath), System.IO.Path.GetFullPath(outputPath), StringComparison.OrdinalIgnoreCase) Then
                                                       doc.Save()
                                                   Else
                                                       doc.SaveAs2(outputPath, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument)
                                                   End If
                                                   Return True

                                               Catch ex As System.Exception
                                                   creationError = ex.Message
                                                   Debug.WriteLine($"CreateWordDoc error: {ex.Message}")
                                                   Return False

                                               Finally
                                                   If sel IsNot Nothing Then
                                                       Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(sel) : Catch : End Try
                                                   End If

                                                   If doc IsNot Nothing Then
                                                       Try : doc.Close(False) : Catch : End Try
                                                       Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(doc) : Catch : End Try
                                                   End If

                                                   Try
                                                       If wordApp IsNot Nothing Then wordApp.ScreenUpdating = True
                                                   Catch
                                                   End Try

                                                   If weCreated AndAlso wordApp IsNot Nothing Then
                                                       Try : wordApp.Quit(False) : Catch : End Try
                                                   End If

                                                   If wordApp IsNot Nothing Then
                                                       Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(wordApp) : Catch : End Try
                                                   End If
                                               End Try
                                           End Function)

            If success AndAlso File.Exists(outputPath) AndAlso visuals.Count > 0 Then
                If Not InsertAutoPilotWordVisualsOpenXml(outputPath, visuals, fontName:=GetArgString(toolCall.Arguments, "base_font_name"), accentHexRaw:=GetArgString(toolCall.Arguments, "accent_color"), embeddedCount:=embeddedVisualCount, warnings:=visualWarnings) Then
                    success = False
                    creationError = "Native OOXML Word visual insertion failed: " & String.Join(" | ", visualWarnings)
                Else
                    Dim persistenceError As String = String.Empty
                    If Not ValidateSavedAutoPilotWordVisualPersistenceOpenXml(outputPath, visuals, persistenceError) Then
                        success = False
                        creationError = persistenceError
                    End If
                End If
            End If

            If success AndAlso File.Exists(outputPath) Then
                RegisterAutoPilotGeneratedOutputFile(outputPath)

                response.Success = True
                Dim designSummary As String = BuildDesignExecutionNote(design)
                Dim visualSummary As String = String.Empty
                If visuals.Count > 0 Then
                    visualSummary = $" Embedded {embeddedVisualCount}/{visuals.Count} requested visual(s)."
                    If visualWarnings.Count > 0 Then
                        visualSummary &= " Visual warnings: " & String.Join(" | ", visualWarnings)
                    End If
                End If
                response.Response = $"Word document created: {fileName} ({New FileInfo(outputPath).Length / 1024:F0} KB). The file will be attached to the reply.{designSummary}{visualSummary}"
                ApDashboardLog($"✓ Word document created: {fileName}", "info")
            Else
                response.Success = False
                response.ErrorMessage = If(String.IsNullOrWhiteSpace(creationError), "Failed to create Word document.", creationError)
                response.Response = response.ErrorMessage
            End If

        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation was cancelled."
            response.Response = response.ErrorMessage
        Catch ex As System.Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error creating Word document: {ex.Message}"
        End Try

        Return response
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: comment_word_document
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Async Function ExecuteCommentWordDocTool(
            toolCall As ToolCall,
            context As ToolExecutionContext,
            ct As CancellationToken) As System.Threading.Tasks.Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim instruction = GetArgString(toolCall.Arguments, "instruction")
            If String.IsNullOrWhiteSpace(instruction) Then
                response.Success = False
                response.ErrorMessage = "Missing required parameter: instruction"
                response.Response = response.ErrorMessage
                Return response
            End If

            Dim author = GetArgString(toolCall.Arguments, "author")
            Dim targetNames = GetArgStringArray(toolCall.Arguments, "attachment_names")

            Dim toProcess As List(Of AutoPilotAttachmentInfo)
            If targetNames.Count > 0 Then
                toProcess = _apCurrentAttachments?.Where(
                    Function(a) targetNames.Any(
                        Function(n) a.OriginalFileName.Equals(n, StringComparison.OrdinalIgnoreCase)
                    ) AndAlso Not a.IsOverSizeLimit AndAlso a.TempFilePath IsNot Nothing
                ).ToList()
            Else
                toProcess = _apCurrentAttachments?.Where(
                    Function(a) (a.Extension = ".docx" OrElse a.Extension = ".doc") AndAlso
                                Not a.IsOverSizeLimit AndAlso a.TempFilePath IsNot Nothing
                ).ToList()
            End If

            If toProcess Is Nothing OrElse toProcess.Count = 0 Then
                response.Success = False
                response.Response = "No processable Word document attachments found."
                Return response
            End If

            Dim effectiveAuthor = If(String.IsNullOrWhiteSpace(author), AN6, author.Trim())
            Dim authorNote = If(effectiveAuthor.Equals(AN6, StringComparison.OrdinalIgnoreCase), "", $" (author: {effectiveAuthor})")
            Dim resultMessages As New List(Of String)()

            For Each att In toProcess
                context.Log($"Adding comments to: {att.OriginalFileName} with instruction: {instruction}{authorNote}")
                ApDashboardLog($"💬 Adding comments to: {att.OriginalFileName}{authorNote}", "step")

                If Not att.TempFilePath.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) Then
                    resultMessages.Add($"✗ {att.OriginalFileName}: Only .docx files are supported for comment insertion.")
                    Continue For
                End If

                Dim outputName = Path.GetFileNameWithoutExtension(att.OriginalFileName) & "_commented.docx"
                Dim outputPath = Path.Combine(_apCurrentTempDir, outputName)

                Dim success = Await CommentDocxForAutoPilot(att.TempFilePath, outputPath, instruction, ct, author)

                If success Then
                    att.OutputFiles.Add(outputPath)
                    resultMessages.Add($"✓ {att.OriginalFileName}: Comments added successfully. Output: {outputName}")
                    ApDashboardLog($"✓ Comments added to: {att.OriginalFileName}", "info")
                Else
                    resultMessages.Add($"✗ {att.OriginalFileName}: Failed to add comments (document may be empty or unsupported).")
                    ApDashboardLog($"⚠ Failed to add comments to: {att.OriginalFileName}", "warn")
                End If
            Next

            response.Success = resultMessages.Any(Function(m) m.StartsWith("✓"))
            response.Response = String.Join(vbCrLf, resultMessages)

        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation was cancelled."
            response.Response = response.ErrorMessage
        Catch ex As System.Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error adding comments to Word document(s): {ex.Message}"
        End Try

        Return response
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: compare_word_documents
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Async Function ExecuteCompareWordDocsTool(
                toolCall As ToolCall,
                context As ToolExecutionContext,
                ct As CancellationToken) As System.Threading.Tasks.Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName,
            .Timestamp = DateTime.UtcNow, .OriginalCallJson = toolCall.RawJson
        }

        Try
            Dim originalFilename = GetArgString(toolCall.Arguments, "original_filename")
            Dim revisedFilename = GetArgString(toolCall.Arguments, "revised_filename")

            If String.IsNullOrWhiteSpace(originalFilename) OrElse String.IsNullOrWhiteSpace(revisedFilename) Then
                response.Success = False
                response.ErrorMessage = "Both 'original_filename' and 'revised_filename' are required."
                response.Response = response.ErrorMessage
                Return response
            End If

            ' Guard: need at least some attachments or output files to compare
            If _apCurrentAttachments Is Nothing OrElse _apCurrentAttachments.Count = 0 Then
                response.Success = False
                response.ErrorMessage = "No attachments available for comparison."
                response.Response = response.ErrorMessage
                Return response
            End If

            Dim originalAtt = FindAttachment(originalFilename)
            Dim revisedAtt = FindAttachment(revisedFilename)

            ' Use GetAllAvailableFileNames for better error messages
            If originalAtt Is Nothing Then
                response.Success = False
                response.ErrorMessage = $"Original attachment '{originalFilename}' not found. Available: {String.Join(", ", GetAllAvailableFileNames())}"
                response.Response = response.ErrorMessage
                Return response
            End If

            If revisedAtt Is Nothing Then
                response.Success = False
                response.ErrorMessage = $"Revised attachment '{revisedFilename}' not found. Available: {String.Join(", ", GetAllAvailableFileNames())}"
                response.Response = response.ErrorMessage
                Return response
            End If

            Dim origExt = Path.GetExtension(originalAtt.TempFilePath).ToLowerInvariant()
            Dim revExt = Path.GetExtension(revisedAtt.TempFilePath).ToLowerInvariant()
            Dim supportedExts = {".doc", ".docx"}

            If Not supportedExts.Contains(origExt) OrElse Not supportedExts.Contains(revExt) Then
                response.Success = False
                response.ErrorMessage = "Both documents must be Word files (.doc or .docx)."
                response.Response = response.ErrorMessage
                Return response
            End If

            Dim compareName = $"Comparison_{Path.GetFileNameWithoutExtension(originalFilename)}_vs_{Path.GetFileNameWithoutExtension(revisedFilename)}.docx"
            Dim comparePath = Path.Combine(_apCurrentTempDir, compareName)

            context.Log($"Comparing: {originalFilename} (original) vs {revisedFilename} (revised)")
            ApDashboardLog($"📊 Comparing: {originalFilename} vs {revisedFilename}", "step")

            Dim success As Boolean = Await SwitchToUi(Function() CreateWordCompareDocumentForAutoPilot(
                originalAtt.TempFilePath, revisedAtt.TempFilePath, comparePath))

            If success AndAlso File.Exists(comparePath) Then
                ' Register on a real attachment if possible; for transient objects the
                ' fallback directory scan in CollectResultAttachments will pick it up.
                Dim registrationTarget = _apCurrentAttachments.FirstOrDefault(
                    Function(a) a.OriginalFileName.Equals(originalFilename, StringComparison.OrdinalIgnoreCase))
                If registrationTarget IsNot Nothing Then
                    registrationTarget.OutputFiles.Add(comparePath)
                Else
                    ' Fallback: register on the first original attachment
                    _apCurrentAttachments(0).OutputFiles.Add(comparePath)
                End If

                Dim summaryText As String = ""
                Try
                    summaryText = Await SwitchToUi(Function()
                                                       Dim wordApp As Microsoft.Office.Interop.Word.Application = Nothing
                                                       Dim doc As Microsoft.Office.Interop.Word.Document = Nothing
                                                       Dim weCreated As Boolean = False
                                                       Try
                                                           Try
                                                               wordApp = DirectCast(GetObject(, "Word.Application"), Microsoft.Office.Interop.Word.Application)
                                                           Catch
                                                               wordApp = New Microsoft.Office.Interop.Word.Application() With {.Visible = False}
                                                               weCreated = True
                                                           End Try
                                                           doc = wordApp.Documents.Open(comparePath, ReadOnly:=True, Visible:=False, AddToRecentFiles:=False)
                                                           Dim revCount = doc.Revisions.Count
                                                           Dim sb As New StringBuilder()
                                                           sb.AppendLine($"Comparison complete: {revCount} revision(s) found between '{originalFilename}' (original) and '{revisedFilename}' (revised).")
                                                           sb.AppendLine()
                                                           Dim maxRevisions = Math.Min(revCount, 50)
                                                           For i As Integer = 1 To maxRevisions
                                                               Dim rev = doc.Revisions(i)
                                                               Dim revType = rev.Type.ToString()
                                                               Dim revText = rev.Range.Text
                                                               If revText IsNot Nothing AndAlso revText.Length > 200 Then
                                                                   revText = revText.Substring(0, 200) & "..."
                                                               End If
                                                               sb.AppendLine($"  [{revType}] {revText}")
                                                           Next
                                                           If revCount > maxRevisions Then
                                                               sb.AppendLine($"  ... and {revCount - maxRevisions} more revision(s).")
                                                           End If
                                                           Return sb.ToString()
                                                       Finally
                                                           If doc IsNot Nothing Then
                                                               Try : doc.Close(False) : Catch : End Try
                                                               Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(doc) : Catch : End Try
                                                           End If
                                                           If wordApp IsNot Nothing Then
                                                               If weCreated Then
                                                                   Try : wordApp.Quit(False) : Catch : End Try
                                                               End If
                                                               Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(wordApp) : Catch : End Try
                                                           End If
                                                       End Try
                                                   End Function)
                Catch ex As System.Exception
                    summaryText = $"Comparison document created successfully but could not extract revision summary: {ex.Message}"
                End Try

                response.Success = True
                response.Response = summaryText & vbCrLf & $"The comparison document '{compareName}' has been generated and will be attached to the reply."
                ApDashboardLog($"✓ Comparison complete: {compareName}", "info")
            Else
                response.Success = False
                response.ErrorMessage = "Word comparison failed. The documents may be incompatible or corrupted."
                response.Response = response.ErrorMessage
                ApDashboardLog($"⚠ Comparison failed for: {originalFilename} vs {revisedFilename}", "warn")
            End If

        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation cancelled."
            response.Response = response.ErrorMessage
        Catch ex As System.Exception
            response.Success = False
            response.ErrorMessage = $"Error comparing documents: {ex.Message}"
            response.Response = response.ErrorMessage
        End Try

        Return response
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: process_word_document
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Async Function ExecuteProcessWordDocTool(
            toolCall As ToolCall,
            context As ToolExecutionContext,
            ct As CancellationToken) As System.Threading.Tasks.Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim instruction = GetArgString(toolCall.Arguments, "instruction")
            If String.IsNullOrWhiteSpace(instruction) Then
                response.Success = False
                response.ErrorMessage = "Missing required parameter: instruction"
                response.Response = response.ErrorMessage
                Return response
            End If

            Dim targetNames = GetArgStringArray(toolCall.Arguments, "attachment_names")
            Dim sheetNames = GetArgStringArray(toolCall.Arguments, "sheet_names")

            ' Parse task_type: "translate", "correct", or "other" (default)
            Dim taskType = If(GetArgString(toolCall.Arguments, "task_type"), "other").Trim().ToLowerInvariant()
            Dim useOfflineDocs As Boolean = (taskType = "translate" OrElse taskType = "correct")

            Dim toProcess As List(Of AutoPilotAttachmentInfo)
            If targetNames.Count > 0 Then
                ' Resolve each requested name via FindAttachment (supports output files)
                toProcess = New List(Of AutoPilotAttachmentInfo)()
                For Each name In targetNames
                    Dim att = FindAttachment(name)
                    If att IsNot Nothing AndAlso Not att.IsOverSizeLimit AndAlso att.TempFilePath IsNot Nothing Then
                        toProcess.Add(att)
                    End If
                Next
            Else
                toProcess = _apCurrentAttachments?.Where(
                    Function(a) (a.Extension = ".docx" OrElse a.Extension = ".doc" OrElse
                                 a.Extension = ".pptx" OrElse a.Extension = ".xlsx") AndAlso
                                Not a.IsOverSizeLimit AndAlso a.TempFilePath IsNot Nothing
                ).ToList()
            End If

            If toProcess Is Nothing OrElse toProcess.Count = 0 Then
                response.Success = False
                response.Response = "No processable Word, PowerPoint, or Excel attachments found."
                Return response
            End If

            ' Guard against recursive re-processing: warn if all targets are tool outputs
            Dim allAreOutputs = toProcess.All(Function(a) a.IsToolOutput)
            If allAreOutputs Then
                ApDashboardLog($"⚠ process_word_document called on tool output file(s) — proceeding with caution", "warn")
            End If

            Dim resultMessages As New List(Of String)()

            For Each att In toProcess
                Dim truncatedInstruction = If(instruction.Length > 120, instruction.Substring(0, 117) & "...", instruction)
                context.Log($"Processing: {att.OriginalFileName} with instruction: {truncatedInstruction} (task_type={taskType})")

                Dim inputPath = att.TempFilePath
                Dim ext = att.Extension.ToLowerInvariant()
                Dim isPptx As Boolean = ext.Equals(".pptx", StringComparison.OrdinalIgnoreCase)
                Dim isXlsx As Boolean = ext.Equals(".xlsx", StringComparison.OrdinalIgnoreCase)
                Dim outputExt As String = If(isPptx, ".pptx", If(isXlsx, ".xlsx", ".docx"))
                Dim outputName = Path.GetFileNameWithoutExtension(att.OriginalFileName) & "_processed" & outputExt
                Dim outputPath = Path.Combine(_apCurrentTempDir, outputName)

                ' Prevent filename collision when re-processing
                Dim counter = 1
                While File.Exists(outputPath)
                    outputName = Path.GetFileNameWithoutExtension(att.OriginalFileName) & $"_processed_{counter}" & outputExt
                    outputPath = Path.Combine(_apCurrentTempDir, outputName)
                    counter += 1
                End While

                ' Pass sheet filter only for Excel files
                Dim sheetFilter As List(Of String) = If(isXlsx AndAlso sheetNames.Count > 0, sheetNames, Nothing)
                Dim success = Await ProcessDocumentForAutoPilot(inputPath, outputPath, instruction, ct, sheetFilter, useOfflineDocs)

                If success Then
                    ' Register output on the original attachment (not on a transient object)
                    Dim registrationTarget = If(att.IsToolOutput,
                        _apCurrentAttachments.FirstOrDefault(Function(a) a.OutputFiles IsNot Nothing AndAlso
                            a.OutputFiles.Any(Function(p) Path.GetFileName(p).Equals(att.OriginalFileName, StringComparison.OrdinalIgnoreCase))),
                        att)
                    If registrationTarget Is Nothing Then registrationTarget = _apCurrentAttachments(0)

                    registrationTarget.OutputFiles.Add(outputPath)

                    ' Compare document only for Word files (not PPTX or XLSX)
                    If Not isPptx AndAlso Not isXlsx Then
                        Dim comparePath = Path.Combine(_apCurrentTempDir,
                            Path.GetFileNameWithoutExtension(att.OriginalFileName) & "_compare.docx")
                        ' Prevent compare filename collision too
                        Dim cmpCounter = 1
                        While File.Exists(comparePath)
                            comparePath = Path.Combine(_apCurrentTempDir,
                                Path.GetFileNameWithoutExtension(att.OriginalFileName) & $"_compare_{cmpCounter}.docx")
                            cmpCounter += 1
                        End While

                        Dim compareSuccess = Await SwitchToUi(Function() CreateWordCompareDocumentForAutoPilot(inputPath, outputPath, comparePath))
                        If compareSuccess Then
                            registrationTarget.OutputFiles.Add(comparePath)
                            resultMessages.Add($"✓ {att.OriginalFileName}: Processed successfully. Output: {outputName} + compare document.")
                        Else
                            resultMessages.Add($"✓ {att.OriginalFileName}: Processed successfully. Output: {outputName} (compare document creation failed).")
                        End If
                    Else
                        resultMessages.Add($"✓ {att.OriginalFileName}: Processed successfully. Output: {outputName}")
                    End If
                Else
                    resultMessages.Add($"✗ {att.OriginalFileName}: Processing failed.")
                End If
            Next

            response.Success = True
            response.Response = String.Join(vbCrLf, resultMessages)

        Catch ex As System.Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error processing document(s): {ex.Message}"
        End Try

        Return response
    End Function

    ''' <summary>
    ''' Creates a Word tracked-changes comparison document from an original and revised file.
    ''' </summary>
    ''' <param name="originalPath">Path to the baseline/original Word file.</param>
    ''' <param name="processedPath">Path to the revised/processed Word file.</param>
    ''' <param name="comparePath">Destination path for the generated comparison document.</param>
    ''' <returns><c>True</c> if comparison output is created successfully; otherwise <c>False</c>.</returns>
    Private Function CreateWordCompareDocumentForAutoPilot(originalPath As String, processedPath As String, comparePath As String) As Boolean
        Dim wordApp As Microsoft.Office.Interop.Word.Application = Nothing
        Dim originalDoc As Microsoft.Office.Interop.Word.Document = Nothing
        Dim processedDoc As Microsoft.Office.Interop.Word.Document = Nothing
        Dim compareDoc As Microsoft.Office.Interop.Word.Document = Nothing
        Dim weCreatedWordApp As Boolean = False

        Try
            Try
                wordApp = DirectCast(GetObject(, "Word.Application"), Microsoft.Office.Interop.Word.Application)
            Catch
                wordApp = New Microsoft.Office.Interop.Word.Application()
                wordApp.Visible = False
                weCreatedWordApp = True
            End Try

            Dim wasScreenUpdating = wordApp.ScreenUpdating
            wordApp.ScreenUpdating = False

            originalDoc = wordApp.Documents.Open(originalPath, ReadOnly:=True, Visible:=False, AddToRecentFiles:=False)
            processedDoc = wordApp.Documents.Open(processedPath, ReadOnly:=True, Visible:=False, AddToRecentFiles:=False)

            compareDoc = wordApp.CompareDocuments(
                OriginalDocument:=originalDoc, RevisedDocument:=processedDoc,
                Destination:=Microsoft.Office.Interop.Word.WdCompareDestination.wdCompareDestinationNew,
                Granularity:=Microsoft.Office.Interop.Word.WdGranularity.wdGranularityWordLevel,
                CompareFormatting:=True, CompareCaseChanges:=True, CompareWhitespace:=True,
                CompareTables:=True, CompareHeaders:=True, CompareFootnotes:=True,
                CompareTextboxes:=True, CompareFields:=True, CompareComments:=True,
                RevisedAuthor:=AN6, IgnoreAllComparisonWarnings:=True)

            compareDoc.SaveAs2(comparePath, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument)
            compareDoc.Close(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges)
            Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(compareDoc) : Catch : End Try
            compareDoc = Nothing
            processedDoc.Close(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges)
            Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(processedDoc) : Catch : End Try
            processedDoc = Nothing
            originalDoc.Close(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges)
            Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(originalDoc) : Catch : End Try
            originalDoc = Nothing
            wordApp.ScreenUpdating = wasScreenUpdating
            Return True

        Catch ex As System.Exception
            Debug.WriteLine($"CreateWordCompareDocumentForAutoPilot error: {ex.Message}")
            Return False
        Finally
            If compareDoc IsNot Nothing Then
                Try : compareDoc.Close(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
                Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(compareDoc) : Catch : End Try
            End If
            If processedDoc IsNot Nothing Then
                Try : processedDoc.Close(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
                Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(processedDoc) : Catch : End Try
            End If
            If originalDoc IsNot Nothing Then
                Try : originalDoc.Close(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
                Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(originalDoc) : Catch : End Try
            End If
            If wordApp IsNot Nothing Then
                Try : wordApp.ScreenUpdating = True : Catch : End Try
                If weCreatedWordApp Then
                    Try : wordApp.Quit(False) : Catch : End Try
                End If
                Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(wordApp) : Catch : End Try
            End If
        End Try
    End Function



    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: read_word_document_details (OpenXML deep reader)
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Async Function ExecuteReadWordDocDetailsTool(
            toolCall As ToolCall, context As ToolExecutionContext, ct As CancellationToken) As System.Threading.Tasks.Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim fileName = GetArgString(toolCall.Arguments, "attachment_name")
            If String.IsNullOrWhiteSpace(fileName) Then
                response.Success = False
                response.Response = "Missing required parameter: attachment_name"
                Return response
            End If

            Dim att = FindAttachment(fileName)
            If att Is Nothing Then response.Success = False : response.Response = $"Attachment '{fileName}' not found." : Return response
            If att.IsOverSizeLimit Then response.Success = False : response.Response = $"Attachment '{fileName}' exceeds the size limit." : Return response
            If att.TempFilePath Is Nothing OrElse Not File.Exists(att.TempFilePath) Then
                response.Success = False : response.Response = $"Attachment '{fileName}' could not be read." : Return response
            End If
            If att.Extension <> ".docx" Then
                response.Success = False : response.Response = $"Only .docx files are supported. '{fileName}' is {att.Extension}." : Return response
            End If

            Dim includeComments = GetArgBool(toolCall.Arguments, "include_comments", True)
            Dim includeHeadersFooters = GetArgBool(toolCall.Arguments, "include_headers_footers", False)
            Dim includeFootnotesEndnotes = GetArgBool(toolCall.Arguments, "include_footnotes_endnotes", False)
            Dim includeTrackedChanges = GetArgBool(toolCall.Arguments, "include_tracked_changes", True)
            Dim filterAuthor = GetArgString(toolCall.Arguments, "tracked_changes_author")
            Dim filterSinceStr = GetArgString(toolCall.Arguments, "tracked_changes_since")

            Dim filterSince As DateTime? = Nothing
            If Not String.IsNullOrWhiteSpace(filterSinceStr) Then
                Dim parsed As DateTime
                If DateTime.TryParse(filterSinceStr, Globalization.CultureInfo.InvariantCulture,
                                     Globalization.DateTimeStyles.None, parsed) Then
                    filterSince = parsed
                End If
            End If

            context.Log($"Deep-reading Word document: {fileName}")
            ApDashboardLog($"📖 Deep-reading: {fileName}", "step")

            Dim result = Await System.Threading.Tasks.Task.Run(Function() ExtractWordDocumentDetails(
                att.TempFilePath, includeComments, includeHeadersFooters,
                includeFootnotesEndnotes, includeTrackedChanges, filterAuthor, filterSince))

            If result.Length > 300000 Then
                result = result.Substring(0, 300000) & vbCrLf & "[... content truncated at 300,000 characters (use read_attachment for more) ...]"
            End If

            response.Success = True
            response.Response = result
            ApDashboardLog($"✓ Deep-read complete: {fileName} ({result.Length:N0} chars)", "info")

        Catch ex As System.Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error reading Word document details: {ex.Message}"
        End Try

        Return response
    End Function

    ''' <summary>
    ''' Extracts detailed content from a .docx file using OpenXML, including body text
    ''' with inline tracked change markers, comments, headers/footers, and footnotes/endnotes.
    ''' </summary>
    Private Function ExtractWordDocumentDetails(
            filePath As String,
            includeComments As Boolean,
            includeHeadersFooters As Boolean,
            includeFootnotesEndnotes As Boolean,
            includeTrackedChanges As Boolean,
            filterAuthor As String,
            filterSince As DateTime?) As String

        Dim tempDir = Path.Combine(Path.GetTempPath(), "ap_detail_" & Guid.NewGuid().ToString("N"))
        Try
            ZipFile.ExtractToDirectory(filePath, tempDir)

            Dim nsMgr As XmlNamespaceManager = Nothing
            Dim docXml As XmlDocument = Nothing
            Dim docPath = Path.Combine(tempDir, "word", "document.xml")

            If File.Exists(docPath) Then
                docXml = New XmlDocument()
                docXml.Load(docPath)
                nsMgr = New XmlNamespaceManager(docXml.NameTable)
                nsMgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
                nsMgr.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
            End If

            Dim sb As New StringBuilder()

            ' ── BODY TEXT (with optional inline tracked changes) ──
            If docXml IsNot Nothing Then
                Dim bodyNode = docXml.SelectSingleNode("//w:body", nsMgr)
                If bodyNode IsNot Nothing Then
                    Dim headerLabel = If(includeTrackedChanges, "═══ DOCUMENT BODY (with tracked changes) ═══", "═══ DOCUMENT BODY ═══")
                    sb.AppendLine(headerLabel)
                    sb.AppendLine()

                    Dim revInsCount = 0
                    Dim revDelCount = 0
                    Dim revFmtCount = 0
                    Dim authorCounts As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

                    For Each paraNode As XmlNode In bodyNode.SelectNodes("w:p", nsMgr)
                        Dim paraText As New StringBuilder()

                        For Each child As XmlNode In paraNode.ChildNodes
                            ProcessDocBodyNode(child, nsMgr, paraText, includeTrackedChanges,
                                             filterAuthor, filterSince,
                                             revInsCount, revDelCount, revFmtCount, authorCounts)
                        Next

                        Dim line = paraText.ToString()
                        If Not String.IsNullOrWhiteSpace(line) Then sb.AppendLine(line)
                        sb.AppendLine()
                    Next

                    ' Summary
                    If includeTrackedChanges Then
                        Dim total = revInsCount + revDelCount + revFmtCount
                        sb.AppendLine($"═══ TRACKED CHANGES SUMMARY ═══")
                        sb.AppendLine($"Total: {total} revision(s) (Insertions: {revInsCount} | Deletions: {revDelCount} | Format changes: {revFmtCount})")
                        If authorCounts.Count > 0 Then
                            sb.AppendLine("By author: " & String.Join(", ", authorCounts.Select(Function(kv) $"{kv.Key}: {kv.Value}")))
                        End If
                        sb.AppendLine()
                    End If
                End If
            End If

            ' ── COMMENTS ──
            If includeComments Then
                Dim commentsPath = Path.Combine(tempDir, "word", "comments.xml")
                If File.Exists(commentsPath) Then
                    Dim commDoc As New XmlDocument()
                    commDoc.Load(commentsPath)
                    Dim cNsMgr As New XmlNamespaceManager(commDoc.NameTable)
                    cNsMgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

                    Dim commentNodes = commDoc.SelectNodes("//w:comment", cNsMgr)

                    ' Build comment-to-anchor mapping from document.xml
                    Dim commentAnchors As New Dictionary(Of String, String)()
                    If docXml IsNot Nothing Then
                        BuildCommentAnchorMap(docXml, nsMgr, commentAnchors)
                    End If

                    If commentNodes.Count > 0 Then
                        sb.AppendLine($"═══ COMMENTS ({commentNodes.Count}) ═══")
                        Dim idx = 1
                        For Each cNode As XmlElement In commentNodes
                            Dim author = cNode.GetAttribute("w:author")
                            Dim dateStr = cNode.GetAttribute("w:date")
                            Dim commentId = cNode.GetAttribute("w:id")
                            Dim commentText As New StringBuilder()
                            For Each tNode As XmlNode In cNode.SelectNodes(".//w:t", cNsMgr)
                                commentText.Append(tNode.InnerText)
                            Next

                            sb.AppendLine($"[Comment #{idx}] Author: {author} | Date: {dateStr}")
                            Dim anchorText As String = Nothing
                            If commentAnchors.TryGetValue(commentId, anchorText) AndAlso Not String.IsNullOrWhiteSpace(anchorText) Then
                                If anchorText.Length > 200 Then anchorText = anchorText.Substring(0, 200) & "..."
                                sb.AppendLine($"  Anchored to: ""{anchorText}""")
                            End If
                            sb.AppendLine($"  Comment: {commentText}")
                            sb.AppendLine()
                            idx += 1
                        Next
                    End If
                End If
            End If

            ' ── HEADERS & FOOTERS ──
            If includeHeadersFooters Then
                ExtractHeadersFooters(tempDir, sb, "header", "HEADERS")
                ExtractHeadersFooters(tempDir, sb, "footer", "FOOTERS")
            End If

            ' ── FOOTNOTES & ENDNOTES ──
            If includeFootnotesEndnotes Then
                ExtractNotesSection(tempDir, sb, "footnotes.xml", "FOOTNOTES")
                ExtractNotesSection(tempDir, sb, "endnotes.xml", "ENDNOTES")
            End If

            Return sb.ToString().TrimEnd()
        Finally
            Try : Directory.Delete(tempDir, True) : Catch : End Try
        End Try
    End Function

    ''' <summary>
    ''' Recursively processes a node in the document body, emitting text and inline change markers.
    ''' </summary>
    Private Sub ProcessDocBodyNode(
            node As XmlNode, nsMgr As XmlNamespaceManager, sb As StringBuilder,
            includeTrackedChanges As Boolean, filterAuthor As String, filterSince As DateTime?,
            ByRef insCount As Integer, ByRef delCount As Integer, ByRef fmtCount As Integer,
            authorCounts As Dictionary(Of String, Integer))

        If node Is Nothing Then Return

        Select Case node.LocalName
            Case "r" ' Normal run
                For Each tNode As XmlNode In node.SelectNodes("w:t", nsMgr)
                    sb.Append(tNode.InnerText)
                Next

            Case "ins" ' Insertion
                Dim author = If(DirectCast(node, XmlElement).GetAttribute("w:author"), "")
                Dim dateStr = If(DirectCast(node, XmlElement).GetAttribute("w:date"), "")
                Dim shortDate = If(dateStr.Length >= 10, dateStr.Substring(0, 10), dateStr)

                Dim passesFilter = PassesRevisionFilter(author, dateStr, filterAuthor, filterSince)

                If includeTrackedChanges AndAlso passesFilter Then
                    Dim innerText As New StringBuilder()
                    For Each child As XmlNode In node.ChildNodes
                        For Each tNode As XmlNode In child.SelectNodes(".//w:t", nsMgr)
                            innerText.Append(tNode.InnerText)
                        Next
                    Next
                    sb.Append($"«INS|{author}|{shortDate}»{innerText}«/INS»")
                    insCount += 1
                    IncrementAuthorCount(authorCounts, author)
                Else
                    ' When not showing changes or filtered out: show inserted text as accepted
                    For Each child As XmlNode In node.ChildNodes
                        For Each tNode As XmlNode In child.SelectNodes(".//w:t", nsMgr)
                            sb.Append(tNode.InnerText)
                        Next
                    Next
                End If

            Case "del" ' Deletion
                Dim author = If(DirectCast(node, XmlElement).GetAttribute("w:author"), "")
                Dim dateStr = If(DirectCast(node, XmlElement).GetAttribute("w:date"), "")
                Dim shortDate = If(dateStr.Length >= 10, dateStr.Substring(0, 10), dateStr)

                Dim passesFilter = PassesRevisionFilter(author, dateStr, filterAuthor, filterSince)

                If includeTrackedChanges AndAlso passesFilter Then
                    Dim innerText As New StringBuilder()
                    For Each child As XmlNode In node.ChildNodes
                        For Each tNode As XmlNode In child.SelectNodes(".//w:delText | .//w:t", nsMgr)
                            innerText.Append(tNode.InnerText)
                        Next
                    Next
                    sb.Append($"«DEL|{author}|{shortDate}»{innerText}«/DEL»")
                    delCount += 1
                    IncrementAuthorCount(authorCounts, author)
                End If
                ' When not showing changes or filtered out: omit deleted text (it was deleted)

            Case "rPrChange" ' Format change
                If includeTrackedChanges Then
                    Dim author = If(DirectCast(node, XmlElement).GetAttribute("w:author"), "")
                    Dim dateStr = If(DirectCast(node, XmlElement).GetAttribute("w:date"), "")
                    If PassesRevisionFilter(author, dateStr, filterAuthor, filterSince) Then
                        fmtCount += 1
                        IncrementAuthorCount(authorCounts, author)
                    End If
                End If

            Case Else
                ' Recurse into child nodes for structure elements like hyperlinks, smart tags, etc.
                For Each child As XmlNode In node.ChildNodes
                    ProcessDocBodyNode(child, nsMgr, sb, includeTrackedChanges,
                                     filterAuthor, filterSince, insCount, delCount, fmtCount, authorCounts)
                Next
        End Select
    End Sub

    Private Shared Function PassesRevisionFilter(author As String, dateStr As String,
                                                  filterAuthor As String, filterSince As DateTime?) As Boolean
        If Not String.IsNullOrWhiteSpace(filterAuthor) Then
            If Not author.IndexOf(filterAuthor, StringComparison.OrdinalIgnoreCase) >= 0 Then Return False
        End If
        If filterSince.HasValue AndAlso Not String.IsNullOrWhiteSpace(dateStr) Then
            Dim revDate As DateTime
            If DateTime.TryParse(dateStr, Globalization.CultureInfo.InvariantCulture,
                                 Globalization.DateTimeStyles.None, revDate) Then
                If revDate < filterSince.Value Then Return False
            End If
        End If
        Return True
    End Function

    Private Shared Sub IncrementAuthorCount(dict As Dictionary(Of String, Integer), author As String)
        If String.IsNullOrWhiteSpace(author) Then author = "(unknown)"
        If dict.ContainsKey(author) Then dict(author) += 1 Else dict(author) = 1
    End Sub

    ''' <summary>
    ''' Builds a mapping from comment ID to the text that the comment is anchored to.
    ''' </summary>
    Private Sub BuildCommentAnchorMap(docXml As XmlDocument, nsMgr As XmlNamespaceManager,
                                      anchors As Dictionary(Of String, String))
        ' Find all commentRangeStart / commentRangeEnd pairs
        Dim starts = docXml.SelectNodes("//w:commentRangeStart", nsMgr)
        For Each startNode As XmlElement In starts
            Dim commentId = startNode.GetAttribute("w:id")
            If String.IsNullOrEmpty(commentId) Then Continue For

            ' Collect text nodes between commentRangeStart and commentRangeEnd with same id
            Dim anchorText As New StringBuilder()
            Dim current = startNode.NextSibling
            Dim found = False
            Dim maxNodes = 500 ' Safety limit

            While current IsNot Nothing AndAlso maxNodes > 0
                maxNodes -= 1
                If current.LocalName = "commentRangeEnd" Then
                    Dim endId = DirectCast(current, XmlElement).GetAttribute("w:id")
                    If endId = commentId Then found = True : Exit While
                End If

                For Each tNode As XmlNode In current.SelectNodes(".//w:t", nsMgr)
                    anchorText.Append(tNode.InnerText)
                Next

                current = current.NextSibling
            End While

            ' If not found as sibling, might be across paragraphs — still use what we got
            If anchorText.Length > 0 Then anchors(commentId) = anchorText.ToString()
        Next
    End Sub

    ''' <summary>
    ''' Extracts header or footer content from the word directory.
    ''' </summary>
    Private Sub ExtractHeadersFooters(tempDir As String, sb As StringBuilder, prefix As String, label As String)
        Dim wordDir = Path.Combine(tempDir, "word")
        If Not Directory.Exists(wordDir) Then Return

        Dim files = Directory.GetFiles(wordDir, prefix & "*.xml")
        If files.Length = 0 Then Return

        Dim anyContent = False
        Dim tempSb As New StringBuilder()

        For Each f In files
            Try
                Dim doc As New XmlDocument()
                doc.Load(f)
                Dim ns As New XmlNamespaceManager(doc.NameTable)
                ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

                Dim text As New StringBuilder()
                For Each tNode As XmlNode In doc.SelectNodes("//w:t", ns)
                    text.Append(tNode.InnerText)
                Next

                If text.Length > 0 Then
                    Dim shortName = Path.GetFileNameWithoutExtension(f)
                    tempSb.AppendLine($"[{shortName}] {text}")
                    anyContent = True
                End If
            Catch
            End Try
        Next

        If anyContent Then
            sb.AppendLine($"═══ {label} ═══")
            sb.Append(tempSb)
            sb.AppendLine()
        End If
    End Sub

    ''' <summary>
    ''' Extracts footnotes or endnotes from the corresponding XML file.
    ''' </summary>
    Private Sub ExtractNotesSection(tempDir As String, sb As StringBuilder, xmlFileName As String, label As String)
        Dim notesPath = Path.Combine(tempDir, "word", xmlFileName)
        If Not File.Exists(notesPath) Then Return

        Try
            Dim doc As New XmlDocument()
            doc.Load(notesPath)
            Dim ns As New XmlNamespaceManager(doc.NameTable)
            ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

            ' Footnotes/endnotes have w:footnote or w:endnote elements; skip type="separator"/"continuationSeparator"
            Dim nodeName = If(xmlFileName.Contains("footnote"), "w:footnote", "w:endnote")
            Dim noteNodes = doc.SelectNodes($"//{nodeName}", ns)

            Dim entries As New List(Of String)()
            For Each noteNode As XmlElement In noteNodes
                Dim noteType = noteNode.GetAttribute("w:type")
                If noteType = "separator" OrElse noteType = "continuationSeparator" Then Continue For

                Dim noteId = noteNode.GetAttribute("w:id")
                Dim noteText As New StringBuilder()
                For Each tNode As XmlNode In noteNode.SelectNodes(".//w:t", ns)
                    noteText.Append(tNode.InnerText)
                Next

                If noteText.Length > 0 Then
                    entries.Add($"[{label.TrimEnd("S"c)} {noteId}] {noteText}")
                End If
            Next

            If entries.Count > 0 Then
                sb.AppendLine($"═══ {label} ({entries.Count}) ═══")
                For Each entry In entries
                    sb.AppendLine(entry)
                Next
                sb.AppendLine()
            End If
        Catch
        End Try
    End Sub



    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: extract_excel_data
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Function ExecuteExtractExcelDataTool(toolCall As ToolCall, context As ToolExecutionContext) As ToolResponse
        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim fileName = GetArgString(toolCall.Arguments, "attachment_name")
            If String.IsNullOrWhiteSpace(fileName) Then
                response.Success = False
                response.Response = "Missing required parameter: attachment_name"
                Return response
            End If

            Dim att = FindAttachment(fileName)
            If att Is Nothing Then response.Success = False : response.Response = $"Attachment '{fileName}' not found." : Return response
            If att.IsOverSizeLimit Then response.Success = False : response.Response = $"Attachment '{fileName}' exceeds the size limit." : Return response

            Dim sheetFilter = GetArgString(toolCall.Arguments, "sheet_name")

            context.Log($"Extracting Excel data: {fileName}")
            ApDashboardLog($"📊 Extracting Excel data: {fileName}", "step")

            ' Use the existing ExtractExcelText which handles interop
            Dim text = ExtractExcelText(att.TempFilePath)

            If String.IsNullOrWhiteSpace(text) OrElse text.StartsWith("Error") Then
                response.Success = False
                response.Response = $"Could not extract data from '{fileName}'."
                Return response
            End If

            ' Filter by sheet name if specified
            If Not String.IsNullOrWhiteSpace(sheetFilter) Then
                Dim sheetMarker = $"[Sheet: {sheetFilter}]"
                Dim idx = text.IndexOf(sheetMarker, StringComparison.OrdinalIgnoreCase)
                If idx >= 0 Then
                    ' Find the next sheet marker or end
                    Dim nextSheet = text.IndexOf("[Sheet: ", idx + sheetMarker.Length, StringComparison.OrdinalIgnoreCase)
                    text = If(nextSheet >= 0, text.Substring(idx, nextSheet - idx).TrimEnd(), text.Substring(idx).TrimEnd())
                End If
            End If

            If text.Length > 50000 Then
                text = text.Substring(0, 50000) & vbCrLf & "[... content truncated at 50,000 characters ...]"
            End If

            response.Success = True
            response.Response = text
            ApDashboardLog($"✓ Excel data extracted: {fileName} ({text.Length:N0} chars)", "info")

        Catch ex As System.Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error extracting Excel data: {ex.Message}"
        End Try

        Return response
    End Function




    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: word_to_pdf
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Async Function ExecuteWordToPdfTool(
            toolCall As ToolCall, context As ToolExecutionContext, ct As CancellationToken) As System.Threading.Tasks.Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim fileName = GetArgString(toolCall.Arguments, "attachment_name")
            If String.IsNullOrWhiteSpace(fileName) Then
                response.Success = False
                response.Response = "Missing required parameter: attachment_name"
                Return response
            End If

            Dim att = FindAttachment(fileName)
            If att Is Nothing Then response.Success = False : response.Response = $"Attachment '{fileName}' not found." : Return response

            Dim ext = Path.GetExtension(att.TempFilePath).ToLowerInvariant()
            If ext <> ".doc" AndAlso ext <> ".docx" Then
                response.Success = False
                response.Response = $"'{fileName}' is not a Word document ({ext})."
                Return response
            End If

            Dim outputName = Path.GetFileNameWithoutExtension(att.OriginalFileName) & ".pdf"
            Dim outputPath = Path.Combine(_apCurrentTempDir, outputName)

            context.Log($"Converting to PDF: {fileName}")
            ApDashboardLog($"📄 Converting to PDF: {fileName}", "step")

            Dim success = Await SwitchToUi(Function()
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
                                                   doc = wordApp.Documents.Open(att.TempFilePath, ReadOnly:=True, Visible:=False, AddToRecentFiles:=False)
                                                   doc.SaveAs2(outputPath, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF)
                                                   Return True
                                               Catch ex As System.Exception
                                                   Debug.WriteLine($"WordToPdf error: {ex.Message}")
                                                   Return False
                                               Finally
                                                   If doc IsNot Nothing Then
                                                       Try : doc.Close(False) : Catch : End Try
                                                       Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(doc) : Catch : End Try
                                                   End If
                                                   Try : If wordApp IsNot Nothing Then wordApp.ScreenUpdating = True
                                                   Catch : End Try
                                                   If weCreated AndAlso wordApp IsNot Nothing Then
                                                       Try : wordApp.Quit(False) : Catch : End Try
                                                   End If
                                                   If wordApp IsNot Nothing Then
                                                       Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(wordApp) : Catch : End Try
                                                   End If
                                               End Try
                                           End Function)

            If success AndAlso File.Exists(outputPath) Then
                att.OutputFiles.Add(outputPath)
                response.Success = True
                response.Response = $"Converted '{fileName}' to PDF: {outputName} ({New FileInfo(outputPath).Length / 1024:F0} KB)."
                ApDashboardLog($"✓ Converted to PDF: {outputName}", "info")
            Else
                response.Success = False
                response.Response = $"Failed to convert '{fileName}' to PDF."
            End If

        Catch ex As System.Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error converting to PDF: {ex.Message}"
        End Try

        Return response
    End Function




    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: pdf_to_word
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Async Function ExecutePdfToWordTool(
            toolCall As ToolCall, context As ToolExecutionContext, ct As CancellationToken) As System.Threading.Tasks.Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim fileName = GetArgString(toolCall.Arguments, "attachment_name")
            If String.IsNullOrWhiteSpace(fileName) Then
                response.Success = False
                response.Response = "Missing required parameter: attachment_name"
                Return response
            End If

            Dim att = FindAttachment(fileName)
            If att Is Nothing Then response.Success = False : response.Response = $"Attachment '{fileName}' not found." : Return response
            If att.IsOverSizeLimit Then response.Success = False : response.Response = $"Attachment '{fileName}' exceeds the size limit." : Return response

            Dim ext = Path.GetExtension(att.TempFilePath).ToLowerInvariant()
            If ext <> ".pdf" Then
                response.Success = False
                response.Response = $"'{fileName}' is not a PDF ({ext})."
                Return response
            End If

            Dim defaultOutput = Path.GetFileNameWithoutExtension(att.OriginalFileName) & ".docx"
            Dim outputName = If(GetArgString(toolCall.Arguments, "output_filename"), defaultOutput)
            Dim outputPath = Path.Combine(_apCurrentTempDir, outputName)

            context.Log($"Converting PDF to Word: {fileName}")
            ApDashboardLog($"📄 Converting PDF to Word: {fileName}", "step")

            ' Use a timeout to prevent indefinite UI thread blocking
            Dim uiTask = SwitchToUi(Function()
                                        Dim wordApp As Microsoft.Office.Interop.Word.Application = Nothing
                                        Dim doc As Microsoft.Office.Interop.Word.Document = Nothing
                                        Dim weCreated As Boolean = False
                                        Dim prevAlerts As Microsoft.Office.Interop.Word.WdAlertLevel =
                                            Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone
                                        Dim prevAutoSec As Microsoft.Office.Core.MsoAutomationSecurity =
                                            Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityByUI
                                        Dim prevFileConverters As Object = Nothing
                                        Dim prevScreenUpdating As Boolean = True
                                        Try
                                            Try
                                                wordApp = DirectCast(GetObject(, "Word.Application"), Microsoft.Office.Interop.Word.Application)
                                            Catch
                                                wordApp = New Microsoft.Office.Interop.Word.Application()
                                                wordApp.Visible = False
                                                weCreated = True
                                            End Try

                                            ' Capture current state BEFORE modifying
                                            prevAlerts = wordApp.DisplayAlerts
                                            prevAutoSec = wordApp.AutomationSecurity
                                            Try : prevScreenUpdating = wordApp.ScreenUpdating : Catch : End Try

                                            ' Suppress all alerts and macro execution
                                            wordApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone
                                            wordApp.ScreenUpdating = False
                                            wordApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable

                                            ' Disable third-party file format converters to prevent modal dialogs
                                            ' from Adobe Acrobat, Foxit, Nuance, etc.
                                            Try
                                                prevFileConverters = wordApp.Options.ConfirmConversions
                                                wordApp.Options.ConfirmConversions = False
                                            Catch
                                            End Try

                                            ' Word can open PDFs and convert them to editable .docx
                                            ' Using Format:=wdOpenFormatAuto (0) lets Word use its BUILT-IN
                                            ' PDF reflow engine rather than deferring to a third-party converter.
                                            doc = wordApp.Documents.Open(
                                                FileName:=att.TempFilePath,
                                                [ReadOnly]:=False,
                                                Visible:=False,
                                                AddToRecentFiles:=False,
                                                ConfirmConversions:=False,
                                                OpenAndRepair:=False,
                                                Format:=0) ' wdOpenFormatAuto = 0

                                            doc.SaveAs2(outputPath, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument)
                                            Return True
                                        Catch ex As System.Exception
                                            Debug.WriteLine($"PdfToWord error: {ex.Message}")
                                            Return False
                                        Finally
                                            ' Close the document and release its COM reference
                                            Try
                                                If doc IsNot Nothing Then
                                                    Try : doc.Close(False) : Catch : End Try
                                                    Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(doc) : Catch : End Try
                                                    doc = Nothing
                                                End If
                                            Catch : End Try
                                            ' Restore Word application state
                                            Try
                                                If wordApp IsNot Nothing Then
                                                    wordApp.DisplayAlerts = prevAlerts
                                                    wordApp.ScreenUpdating = prevScreenUpdating
                                                    wordApp.AutomationSecurity = prevAutoSec
                                                    Try
                                                        If prevFileConverters IsNot Nothing Then
                                                            wordApp.Options.ConfirmConversions = CBool(prevFileConverters)
                                                        End If
                                                    Catch
                                                    End Try
                                                End If
                                            Catch : End Try
                                            ' Quit only if we created this instance, then release COM reference
                                            If weCreated AndAlso wordApp IsNot Nothing Then
                                                Try : wordApp.Quit(False) : Catch : End Try
                                            End If
                                            If wordApp IsNot Nothing Then
                                                Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(wordApp) : Catch : End Try
                                                wordApp = Nothing
                                            End If
                                        End Try
                                    End Function)

            ' Apply a 120-second timeout to prevent indefinite UI thread blocking
            Dim timeoutTask = System.Threading.Tasks.Task.Delay(TimeSpan.FromSeconds(120), ct)
            Dim completedTask = Await System.Threading.Tasks.Task.WhenAny(uiTask, timeoutTask)

            Dim success As Boolean = False
            If completedTask Is uiTask Then
                success = Await uiTask
            Else
                ' Timeout or cancellation
                response.Success = False
                response.Response = $"PDF to Word conversion timed out for '{fileName}'. The PDF may be too large, corrupted, or a third-party converter dialog may be blocking. Check if any dialog is open in Word."
                ApDashboardLog($"⚠ PdfToWord timed out: {fileName}", "warn")
                Return response
            End If

            If success AndAlso File.Exists(outputPath) Then
                att.OutputFiles.Add(outputPath)
                response.Success = True
                response.Response = $"Converted '{fileName}' to Word: {outputName} ({New FileInfo(outputPath).Length / 1024:F0} KB). " &
                    "This file can now be used with compare_word_documents. " &
                    "Note: Word does NOT perform OCR — if the PDF is a scanned image, the resulting .docx will contain images without extracted text."
                ApDashboardLog($"✓ Converted to Word: {outputName}", "info")
            Else
                response.Success = False
                response.Response = $"Failed to convert '{fileName}' to Word. The PDF may be image-only, corrupted, or a third-party PDF converter add-in may have interfered. " &
                    "Ensure no PDF add-ins (Adobe Acrobat, Foxit, etc.) are registered as Word file converters."
            End If

        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation was cancelled."
            response.Response = response.ErrorMessage
        Catch ex As System.Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error converting PDF to Word: {ex.Message}"
        End Try

        Return response
    End Function


End Class
